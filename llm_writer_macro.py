import uno
import unohelper
import datetime
import urllib.request
import urllib.parse
import json
import traceback
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
import os

PARAMS_PATH = os.path.join(os.path.expanduser("~"), "llm_writer_params.json")
LOG_PATH = os.path.join(os.path.expanduser("~"), "llm_writer_api_logs.json")

def init_db():
        """Initialize JSON files for parameters and logs"""
        # Initialize parameters file
        if not os.path.exists(PARAMS_PATH):
            with open(PARAMS_PATH, 'w') as f:
                json.dump({
                    'OPENAI_ENDPOINT': 'https://api.openai.com/v1/chat/completions',
                    'OPENAI_API_KEY': '',
                    'MAX_GENERATION_TOKENS': '100',
                    'AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS': 'Continue the text naturally',
                    'CONTEXT_PREVIOUS_CHARS': '100',
                    'CONTEXT_NEXT_CHARS': '100'
                }, f)
        
        # Initialize logs file
        if not os.path.exists(LOG_PATH):
            with open(LOG_PATH, 'w') as f:
                json.dump([], f)

def get_param(key):
        """Get parameter from JSON file"""
        with open(PARAMS_PATH, 'r') as f:
            params = json.load(f)
            return params.get(key)

def set_param(key, value):
        """Set parameter in JSON file atomically"""
        # Load existing params first
        with open(PARAMS_PATH, 'r') as f:
            params = json.load(f)
        
        # Update the requested parameter
        params[key] = value
        
        # Write to temporary file first to prevent corruption
        temp_path = PARAMS_PATH + ".tmp"
        with open(temp_path, 'w') as f:
            json.dump(params, f)
        
        # Atomically replace original file
        os.replace(temp_path, PARAMS_PATH)

def get_context(cursor):
        """Get previous and next tokens around cursor position"""
        text = cursor.getText()
        prev_chars = int(get_param('CONTEXT_PREVIOUS_CHARS'))
        next_chars = int(get_param('CONTEXT_NEXT_CHARS'))
        
        start = max(0, int(cursor.getStart().Value) - prev_chars)
        end = min(len(text.getString()), int(cursor.getEnd().Value) + next_chars)
        
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
            error_msg = f"ERROR: {str(e)}\n\nMost recent traceback lines:\n" + "\\n".join(traceback.format_exc().splitlines()[-3:])
            full_error = f"FULL ERROR: {str(e)}\n{traceback.format_exc()}"
            show_message(f"{error_msg}\n\nCheck ~/llm_writer_api_logs.json for complete details")
            _log_api_call("autocomplete", {"error": full_error}, {}, 500)

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
        """Log API call details to JSON file"""
        log_entry = {
            'timestamp': datetime.datetime.now().isoformat(),
            'endpoint': endpoint,
            'request': request,
            'response': response,
            'status_code': status_code
        }
        
        with open(LOG_PATH, 'r+') as f:
            logs = json.load(f)
            logs.insert(0, log_entry)
            f.seek(0)
            json.dump(logs, f)
            f.truncate()

def get_api_logs(limit=100):
        """Retrieve API logs from JSON file"""
        with open(LOG_PATH, 'r') as f:
            return json.load(f)[:limit]

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
            
            if 'error' in request:
                log_text += f"Error Details:\n{request['error']}\n"
            else:
                log_text += f"Request: {json.dumps(request, indent=2)[:500]}...\n"
                
            log_text += f"Response: {json.dumps(response, indent=2)[:500]}...\n"
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
init_db()
g_exportedScripts = (autocomplete, transform_text, show_logs)

