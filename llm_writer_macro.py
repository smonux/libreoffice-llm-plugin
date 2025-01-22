import uno
import unohelper
import shelve
import datetime
import urllib.request
import urllib.parse
import json
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
import os

# Change the database for a json file for the parameters, and the logs have to go to a file. ai!
DB_PATH = os.path.join(os.path.expanduser("~"), "llm_writer_params.db")  # .db extension kept for compatibility

def init_db():
        """Initialize shelve database for parameters"""
        with shelve.open(DB_PATH) as db:
            # Initialize parameters with defaults if missing
            defaults = {
                'OPENAI_ENDPOINT': 'https://api.openai.com/v1/chat/completions',
                'OPENAI_API_KEY': '',
                'MAX_GENERATION_TOKENS': '100',
                'AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS': 'Continue the text naturally',
                'CONTEXT_PREVIOUS_CHARS': '100',
                'CONTEXT_NEXT_CHARS': '100'
            }
            for key, value in defaults.items():
                if key not in db:
                    db[key] = value
            
            # Initialize api_logs if missing
            if 'api_logs' not in db:
                db['api_logs'] = []

def get_param(key):
        """Get parameter from shelve database"""
        with shelve.open(DB_PATH) as db:
            return db.get(key)

def set_param(key, value):
        """Set parameter in shelve database"""
        with shelve.open(DB_PATH) as db:
            db[key] = value

def get_context(cursor):
        """Get previous and next tokens around cursor position"""
        text = cursor.getText()
        prev_chars = int(get_param('CONTEXT_PREVIOUS_CHARS'))
        next_chars = int(get_param('CONTEXT_NEXT_CHARS'))
        
        start = max(0, cursor.getStart() - prev_chars)
        end = min(len(text.getString()), cursor.getEnd() + next_chars)
        
        previous_context = text.getString()[start:cursor.getStart()]
        next_context = text.getString()[cursor.getEnd():end]
        
        return previous_context, next_context

def autocomplete(cursor):
        """Generate autocomplete suggestions using LLM"""
        try:
            previous_context, next_context = get_context(cursor)
            
            prompt = f"{previous_context}[COMPLETE HERE]{next_context}\n\n" + \
                    get_param('AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS')
            
            data = {
                'prompt': prompt,
                'max_tokens': int(get_param('MAX_GENERATION_TOKENS')),
                'temperature': 0.7,
                'stop': ['\n'] 
            }
            
            response = call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['text'])
                
        except Exception as e:
            show_message(f"Error: {str(e)}")

def transform_text(cursor, instruction=None):
        """Transform selected text based on instruction"""
        try:
            selected_text = cursor.getString()
            if not instruction:
                instruction = selected_text
                
            prompt = f"Original text: {selected_text}\n\nInstruction: {instruction}\n\nTransformed text:"
            
            data = {
                'prompt': prompt,
                'max_tokens': int(get_param('MAX_GENERATION_TOKENS')),
                'temperature': 0.7
            }
            
            response = call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['text'])
                
        except Exception as e:
            show_message(f"Error: {str(e)}")

def call_llm(data): 
        """Make API call to OpenAI-compatible endpoint"""
        url = get_param('OPENAI_ENDPOINT') 
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f"Bearer {get_param('OPENAI_API_KEY')}"
        }
        
        req = urllib.request.Request(url, 
                                   data=json.dumps(data).encode('utf-8'),
                                   headers=headers)
        
        try:
            with urllib.request.urlopen(req) as response:
                response_data = json.loads(response.read().decode('utf-8'))
                _log_api_call(url, data, response_data, response.status)
                return response_data
        except urllib.error.HTTPError as e:
            _log_api_call(url, data, str(e), e.code)
            raise

def _log_api_call(endpoint, request, response, status_code):
        """Log API call details to database"""
        with shelve.open(DB_PATH) as db:
            logs = db.get('api_logs', [])
            logs.insert(0, {
                'timestamp': datetime.datetime.now().isoformat(),
                'endpoint': endpoint,
                'request': request,
                'response': response,
                'status_code': status_code,
                'id': len(logs) + 1
            })
            db['api_logs'] = logs

def get_api_logs(limit=100):
        """Retrieve API logs from database"""
        with shelve.open(DB_PATH) as db:
            return db.get('api_logs', [])[:limit]

def show_message(message):
        """Show message dialog"""
        ctx = uno.getComponentContext()
        sm = ctx.getServiceManager()
        toolkit = sm.createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx)
        parent = toolkit.getDesktopWindow()
        msgbox = toolkit.createMessageBox(
            parent,  "infobox", MSG_BUTTONS.BUTTONS_OK, "LLM Writer", str(message))
        msgbox.execute()

def _get_cursor():
        """Get text cursor from current selection"""
        xModel = uno.getComponentContext().getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", uno.getComponentContext()).getCurrentComponent()
        xSelectionSupplier = xModel.getCurrentController()
        xIndexAccess = xSelectionSupplier.getSelection()
        return xIndexAccess.getByIndex(0)

def autocomplete():
        """Generate autocomplete suggestions using LLM"""
        try:
            cursor = _get_cursor()
            previous_context, next_context = get_context(cursor)
            
            prompt = f"{previous_context}[COMPLETE HERE]{next_context}\n\n" + \
                    get_param('AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS')
            
            data = {
                'prompt': prompt,
                'max_tokens': int(get_param('MAX_GENERATION_TOKENS')),
                'temperature': 0.7,
                'stop': ['\n'] 
            }
            
            response = call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['text'])
                
        except Exception as e:
            show_message(f"Error: {str(e)}")

def transform_text():
        """Transform selected text based on instruction"""
        try:
            cursor = _get_cursor()
            instruction = show_input_dialog("Enter transformation instructions:")
            selected_text = cursor.getString()
            
            if not instruction:
                return
                
            prompt = f"Original text: {selected_text}\n\nInstruction: {instruction}\n\nTransformed text:"
            
            data = {
                'prompt': prompt,
                'max_tokens': int(get_param('MAX_GENERATION_TOKENS')),
                'temperature': 0.7
            }
            
            response = call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['text'])
                
        except Exception as e:
            show_message(f"Error: {str(e)}")

def show_logs():
        """Display API logs in message box"""
        logs = get_api_logs()
        if not logs:
            show_message("No API logs found")
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
            
        show_message(log_text)

def show_input_dialog(message):
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

# Export the macros properly
g_exportedScripts = (autocomplete, transform_text, show_logs)

