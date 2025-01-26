import uno
import unohelper
import datetime
import urllib.request
import urllib.parse
import json
import traceback
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from msgbox import MsgBox 
import os

PARAMS_PATH = os.path.join(os.path.expanduser("~"), "llm_writer_params.json")
LOG_PATH = os.path.join(os.path.expanduser("~"), "llm_writer_api_logs.log")

def init_db():
        """Initialize JSON files for parameters and logs"""
        # Initialize parameters file
        if not os.path.exists(PARAMS_PATH):
            with open(PARAMS_PATH, 'w') as f:
                json.dump({
                    'OPENAI_ENDPOINT': 'https://api.openai.com/v1/chat/completions',
                    'OPENAI_API_KEY': '',
                    'MODEL': 'gpt-4o-mini',
                    'MAX_GENERATION_TOKENS': '30',
                    'AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS': 'Continue the text naturally in the [COMPLETE] ',
                    'CONTEXT_PREVIOUS_CHARS': '100',
                    'CONTEXT_NEXT_CHARS': '100',
                    'TEMPERATURE': '0.7'
                }, f, indent=4)
        
        # Initialize logs file
        if not os.path.exists(LOG_PATH):
            with open(LOG_PATH, 'w') as f:
                f.write("initializing...\n")

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

        text_cursor = cursor.getText().createTextCursorByRange(cursor)
        text_cursor.goLeft(prev_chars, True)
        previous_context = text_cursor.getString()

        text_cursor = cursor.getText().createTextCursorByRange(cursor)
        text_cursor.goRight(next_chars, True)
        next_context = text_cursor.getString()

        #show_message(previous_context + next_context)
        
        return previous_context, next_context


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
            # Read the response body to get the detailed error message
            error_response = e.read().decode('utf-8')
            _log_api_call(url, data, error_response, e.code)
            raise

def _log_api_call(endpoint, request, response, status_code):
        """Log API call details to a regular text file"""
        with open(LOG_PATH, 'a') as f:
            f.write(f"Timestamp: {datetime.datetime.now().isoformat()}\n")
            f.write(f"Endpoint: {endpoint}\n")
            f.write(f"Status Code: {status_code}\n")
            f.write(f"Request: {request}\n")
            f.write(f"Response: {response}\n")
            f.write("-" * 40 + "\n")

def get_api_logs(limit=100):
        """Retrieve API logs from text file"""
        with open(LOG_PATH, 'r') as f:
            logs = f.readlines()[:limit]
            return logs

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
            
            data = {
                'model': get_param('MODEL'),
                'messages': [
                    {"role": "system", "content": get_param('AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS')},
                    {"role": "user", "content": f"{previous_context}[COMPLETE HERE]{next_context}"}
                ],
                'max_tokens': int(get_param('MAX_GENERATION_TOKENS')),
                'temperature': float(get_param('TEMPERATURE')),
                'stop': ['\n'] 
            }
            
            response = call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['message']['content'])
                
        except Exception as e:
            error_msg = f"ERROR: {str(e)}"
            full_error = f"FULL ERROR: {str(e)}\n{traceback.format_exc()}"
            show_message(full_error) 

def transform_text():
        """Transform selected text based on instruction"""
        try:
            cursor = _get_cursor()
            selected_text = cursor.getString()
            
            if not selected_text:
                return

            instruction = show_input_dialog("Enter transformation instructions:")
                
            data = {
                'model': get_param('MODEL'),
                'messages': [
                    {"role": "system", "content": instruction},
                    {"role": "user", "content": f"Original text: {selected_text}\n\nTransformed text:"}
                ],
                'max_tokens': int(get_param('MAX_GENERATION_TOKENS')),
                'temperature': float(get_param('TEMPERATURE'))
            }
            
            response = call_llm(data)
            if response:
                cursor.setString(response['choices'][0]['message']['content'])
                
        except Exception as e:
            error_msg = f"ERROR: {str(e)}"
            full_error = f"FULL ERROR: {str(e)}\n{traceback.format_exc()}"
            show_message(full_error) 

def show_logs():
        """Display API logs in message box"""
        logs = get_api_logs(25)
        if not logs:
            show_message("No API logs found")
            return
            
        log_text = "API Logs:\n\n"
        for log in logs:
            log_text += log + "\n"
        
        show_message(log_text)

def show_input_dialog(message, title="", default="", x=None, y=None):
    """ Shows dialog with input box.
        @param message message to show on the dialog
        @param title window title
        @param default default value
        @param x optional dialog position in twips
        @param y optional dialog position in twips
        @return string if OK button pushed, otherwise zero length string
    """
    WIDTH = 600
    HORI_MARGIN = VERT_MARGIN = 8
    BUTTON_WIDTH = 100
    BUTTON_HEIGHT = 26
    HORI_SEP = VERT_SEP = 8
    LABEL_HEIGHT = BUTTON_HEIGHT * 2 + 5
    EDIT_HEIGHT = 24
    HEIGHT = VERT_MARGIN * 2 + LABEL_HEIGHT + VERT_SEP + EDIT_HEIGHT
    import uno
    from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
    from com.sun.star.awt.PushButtonType import OK, CANCEL
    from com.sun.star.util.MeasureUnit import TWIP
    ctx = uno.getComponentContext()
    def create(name):
        return ctx.getServiceManager().createInstanceWithContext(name, ctx)
    dialog = create("com.sun.star.awt.UnoControlDialog")
    dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
    dialog.setModel(dialog_model)
    dialog.setVisible(False)
    dialog.setTitle(title)
    dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)
    def add(name, type, x_, y_, width_, height_, props):
        model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
        dialog_model.insertByName(name, model)
        control = dialog.getControl(name)
        control.setPosSize(x_, y_, width_, height_, POSSIZE)
        for key, value in props.items():
            setattr(model, key, value)
    label_width = WIDTH - BUTTON_WIDTH - HORI_SEP - HORI_MARGIN * 2
    add("label", "FixedText", HORI_MARGIN, VERT_MARGIN, label_width, LABEL_HEIGHT, 
        {"Label": str(message), "NoLabel": True})
    add("btn_ok", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN, 
            BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
    add("btn_cancel", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN + BUTTON_HEIGHT + 5, 
            BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": CANCEL})
    add("edit", "Edit", HORI_MARGIN, LABEL_HEIGHT + VERT_MARGIN + VERT_SEP, 
            WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(default)})
    frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
    window = frame.getContainerWindow() if frame else None
    dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
    if not x is None and not y is None:
        ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
        _x, _y = ps.Width, ps.Height
    elif window:
        ps = window.getPosSize()
        _x = ps.Width / 2 - WIDTH / 2
        _y = ps.Height / 2 - HEIGHT / 2
    dialog.setPosSize(_x, _y, 0, 0, POS)
    edit = dialog.getControl("edit")
    edit.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(default))))
    edit.setFocus()
    ret = edit.getModel().Text if dialog.execute() else ""
    dialog.dispose()
    return ret


def XXXXshow_input_dialog(message):
    """Show input dialog"""
    ctx = uno.getComponentContext()
    sm = ctx.getServiceManager()

    # Create a dialog model
    dialog_model = sm.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialogModel", ctx)
    dialog_model.Width = 200
    dialog_model.Height = 100
    dialog_model.Title = "LLM Writer"
    dialog_model.PositionX = 100
    dialog_model.PositionY = 100

    # Add a text field
    text_field_model = sm.createInstanceWithContext(
        "com.sun.star.awt.UnoControlEditModel", ctx)
    text_field_model.setPropertyValue("Name", "TextField")
    text_field_model.setPropertyValue("PositionX", 10)
    text_field_model.setPropertyValue("PositionY", 40)
    text_field_model.setPropertyValue("Width", 180)
    text_field_model.setPropertyValue("Height", 20)
    dialog_model.insertByName("TextField", text_field_model)

    # Add a label
    label_model = sm.createInstanceWithContext(
        "com.sun.star.awt.UnoControlFixedTextModel", ctx)
    label_model.setPropertyValue("Name", "Label")
    label_model.setPropertyValue("PositionX", 10)
    label_model.setPropertyValue("PositionY", 10)
    label_model.setPropertyValue("Width", 180)
    label_model.setPropertyValue("Height", 20)
    label_model.setPropertyValue("Label", message)
    dialog_model.insertByName("Label", label_model)

    # Add an OK button
    ok_button_model = sm.createInstanceWithContext(
        "com.sun.star.awt.UnoControlButtonModel", ctx)
    ok_button_model.setPropertyValue("Name", "OKButton")
    ok_button_model.setPropertyValue("PositionX", 60)
    ok_button_model.setPropertyValue("PositionY", 70)
    ok_button_model.setPropertyValue("Width", 80)
    ok_button_model.setPropertyValue("Height", 20)
    ok_button_model.setPropertyValue("Label", "OK")
    dialog_model.insertByName("OKButton", ok_button_model)

    # Create the dialog
    dialog = sm.createInstanceWithContext(
        "com.sun.star.awt.UnoControlDialog", ctx)

    dialog.setModel(dialog_model)
    # Event handler for the OK button
    def button_action_handler(event):
        dialog.endExecute(1)

    ok_button = dialog.getControl("OKButton")
    ok_button.addActionListener(lambda x: button_action_handler(x))

    # Show the dialog
    dialog.execute()

    # Get the text from the text field
    text_field = dialog.getControl("TextField")
    return text_field.getText() if dialog.getResult() == 1 else None

# Export the macros properly
init_db()
g_exportedScripts = (autocomplete, transform_text, show_logs, modify_config)

