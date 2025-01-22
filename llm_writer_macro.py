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
        self.db_path = os.path.join(os.path.expanduser("~"), ".llm_writer_params.db")
        self.init_db()
        
    def init_db(self):
        """Initialize SQLite database for parameters"""
        with sqlite3.connect(self.db_path) as conn:
            conn.execute('''CREATE TABLE IF NOT EXISTS parameters
                         (key TEXT PRIMARY KEY, value TEXT)''')
            # Set default values if they don't exist
            defaults = {
                'OPENAI_ENDPOINT': 'http://127.0.0.1:5000',
                'OPENAI_API_KEY': '',
                'MAX_GENERATION_TOKENS': '100',
                'AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS': 'Continue the text naturally'
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

    def get_context(self, cursor, num_tokens=100):
        """Get previous and next tokens around cursor position"""
        text = cursor.getText()
        start = max(0, cursor.getStart() - num_tokens)
        end = min(len(text.getString()), cursor.getEnd() + num_tokens)
        
        previous = text.getString()[start:cursor.getStart()]
        next = text.getString()[cursor.getEnd():end]
        
        return previous, next

    def autocomplete(self, cursor):
        """Generate autocomplete suggestions using LLM"""
        try:
            previous, next = self.get_context(cursor)
            
            prompt = f"{previous}[COMPLETE HERE]{next}\n\n" + \
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
        url = self.get_param('OPENAI_ENDPOINT') + "/v1/completions"
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f"Bearer {self.get_param('OPENAI_API_KEY')}"
        }
        
        req = urllib.request.Request(url, 
                                   data=json.dumps(data).encode('utf-8'),
                                   headers=headers)
        
        with urllib.request.urlopen(req) as response:
            return json.loads(response.read().decode('utf-8'))

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

# Register the macro
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    LLMWriterMacro,
    "org.extension.llm_writer",
    ("com.sun.star.task.Job",))
