import uno
import unohelper
import sqlite3
import urllib.request
import urllib.parse
import json
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
import os

class LLMWriterMacro(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx
        self.db_path = os.path.join(os.path.expanduser("~"), "llm_writer_params.db")
        self.init_db()
        
    def init_db(self):
        """Initialize SQLite database for parameters"""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('''CREATE TABLE IF NOT EXISTS parameters
                         (key TEXT PRIMARY KEY, value TEXT)''')
            conn.execute('''CREATE TABLE IF NOT EXISTS api_logs (
                         id INTEGER PRIMARY KEY AUTOINCREMENT,
                         timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                         endpoint TEXT,
                         request TEXT,
                         response TEXT,
                         status_code INTEGER)''')
            # Set default values if they don't exist
            defaults = {
                'OPENAI_ENDPOINT': 'https://api.openai.com/v1/chat/completions',
                'OPENAI_API_KEY': '',
                'MAX_GENERATION_TOKENS': '100',
                'AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS': 'Continue the text naturally',
                'CONTEXT_PREVIOUS_CHARS': '100',
                'CONTEXT_NEXT_CHARS': '100'
            }
            for key, value in defaults.items():
                conn.execute('INSERT OR IGNORE INTO parameters (key, value) VALUES (?, ?)', (key, value))
            conn.commit()

    def get_param(self, key):
        """Get parameter from SQLite database"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute('SELECT value FROM parameters WHERE key = ?', (key,))
            result = cursor.fetchone()
            return result[0] if result else None

    def set_param(self, key, value):
        """Set parameter in SQLite database"""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('INSERT OR REPLACE INTO parameters (key, value) VALUES (?, ?)', (key, value))
            conn.commit()

    def get_context(self, cursor):
        """Get previous and next tokens around cursor position"""
        text = cursor.getText()
        prev_chars = int(self.get_param('CONTEXT_PREVIOUS_CHARS'))
        next_chars = int(self.get_param('CONTEXT_NEXT_CHARS'))
        
        start = max(0, cursor.getStart() - prev_chars)
        end = min(len(text.getString()), cursor.getEnd() + next_chars)
        
        previous_context = text.getString()[start:cursor.getStart()]
        next_context = text.getString()[cursor.getEnd():end]
        
        return previous_context, next_context

    def autocomplete(self, cursor):
        """Generate autocomplete suggestions using LLM"""
        try:
            previous_context, next_context = self.get_context(cursor)
            
            prompt = f"{previous_context}[COMPLETE HERE]{next_context}\n\n" + \
                    self.get_param('AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS')
            
            data = {
                'prompt': prompt,
                'max_tokens': int(self.get_param('MAX_GENERATION_TOKENS')),
                'temperature': 0.7,
                'stop': ['\n'] 
            }
            
            response = self.call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['text'])
                
        except Exception as e:
            self.show_message(f"Error: {str(e)}")

    def transform_text(self, cursor, instruction=None):
        """Transform selected text based on instruction"""
        try:
            selected_text = cursor.getString()
            if not instruction:
                instruction = selected_text
                
            prompt = f"Original text: {selected_text}\n\nInstruction: {instruction}\n\nTransformed text:"
            
            data = {
                'prompt': prompt,
                'max_tokens': int(self.get_param('MAX_GENERATION_TOKENS')),
                'temperature': 0.7
            }
            
            response = self.call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['text'])
                
        except Exception as e:
            self.show_message(f"Error: {str(e)}")

    def call_llm(self, data):
        """Make API call to OpenAI-compatible endpoint"""
        url = self.get_param('OPENAI_ENDPOINT') 
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f"Bearer {self.get_param('OPENAI_API_KEY')}"
        }
        
        req = urllib.request.Request(url, 
                                   data=json.dumps(data).encode('utf-8'),
                                   headers=headers)
        
        try:
            with urllib.request.urlopen(req) as response:
                response_data = json.loads(response.read().decode('utf-8'))
                self._log_api_call(url, data, response_data, response.status)
                return response_data
        except urllib.error.HTTPError as e:
            self._log_api_call(url, data, str(e), e.code)
            raise

    def _log_api_call(self, endpoint, request, response, status_code):
        """Log API call details to database"""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('''INSERT INTO api_logs 
                         (endpoint, request, response, status_code)
                         VALUES (?, ?, ?, ?)''',
                         (endpoint, json.dumps(request), 
                          json.dumps(response) if isinstance(response, dict) else response,
                          status_code))
            conn.commit()

    def get_api_logs(self, limit=100):
        """Retrieve API logs from database"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.execute('''SELECT * FROM api_logs 
                                  ORDER BY timestamp DESC 
                                  LIMIT ?''', (limit,))
            return cursor.fetchall()

    def show_message(self, message):
        """Show message dialog"""
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx)
        msgbox = toolkit.createMessageBox(
            None, MSG_BUTTONS.BUTTONS_OK, "infobox", "LLM Writer", message)
        msgbox.execute()

    def trigger(self, args):
        """Handle macro triggers"""
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)
        model = desktop.getCurrentComponent()
        cursor = model.CurrentController.getSelection().getByIndex(0)
        
        if args == "Autocomplete":
            self.autocomplete(cursor)
        elif args == "Transform":
            instruction = self.show_input_dialog("Enter transformation instructions:")
            self.transform_text(cursor, instruction)
        elif args == "ShowLogs":
            self.show_logs()

    def show_input_dialog(self, message):
        """Show input dialog"""
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx)
        dialog = toolkit.createInputBox()
        dialog.setTitle("LLM Writer")
        dialog.setMessageText(message)
        if dialog.execute():
            return dialog.getValue()
        return None

    def show_logs(self):
        """Display API logs in message box"""
        logs = self.get_api_logs()
        if not logs:
            self.show_message("No API logs found")
            return
            
        log_text = "API Logs:\n\n"
        for log in logs:
            log_id, timestamp, endpoint, request, response, status_code = log
            log_text += f"[{timestamp}]\n"
            log_text += f"Endpoint: {endpoint}\n"
            log_text += f"Status: {status_code}\n"
            log_text += f"Request: {request[:200]}...\n"
            log_text += f"Response: {str(response)[:200]}...\n"
            log_text += "-" * 40 + "\n"
            
        self.show_message(log_text)

# Register the macro
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    LLMWriterMacro,
    "org.extension.llm_writer",
    ("com.sun.star.task.Job",))
