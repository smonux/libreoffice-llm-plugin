import uno
import unohelper
import datetime
import urllib.request
import urllib.parse
import json
import traceback
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
from com.sun.star.awt.PushButtonType import OK, CANCEL
from msgbox import MsgBox
import os

PARAMS_PATH = os.path.join(os.path.expanduser("~"), ".llm_writer", "llm_writer_params.json")
LOG_PATH = os.path.join(os.path.expanduser("~"), ".llm_writer", "llm_writer_api_logs.log")

AUTOCOMPLETE_DEFAULT_PROMPT = """
Continue the text naturally at the [COMPLETE HERE] position using the surrounding context.
Maintain consistent style/tone and preserve narrative flow.
Generate only the next logical sequence of words without repeating existing content. 
Prioritize grammatical correctness and contextual coherence with both preceding and following text.
Respect punctuacion and write consistently with best practices (for example: space and first letter of first word capitalized). 
Use the same language the text is written on.
Don't announce what you are going to do, just do it (e.g: here you have, etc..).
"""


def init_db_maybe():
    """Initialize JSON files for parameters and logs"""
    # Ensure the directory exists
    directory_path = os.path.dirname(PARAMS_PATH)
    os.makedirs(directory_path, exist_ok=True)

    # Initialize parameters file
    if not os.path.exists(PARAMS_PATH):
        with open(PARAMS_PATH, "w") as f:
            json.dump(
                # The order is important for displaying the config
                {
                    "OPENAI_ENDPOINT": "https://api.openai.com/v1/chat/completions",
                    "OPENAI_API_KEY": "",
                    "MODEL": "gpt-4o",
                    "MAX_GENERATION_WORDS": "10",
                    "CONTEXT_PREVIOUS_CHARS": "100",
                    "CONTEXT_NEXT_CHARS": "100",
                    "TEMPERATURE": "0.7",
                    "AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS": AUTOCOMPLETE_DEFAULT_PROMPT,
                },
                f,
                indent=4,
            )

    # Initialize logs file
    if not os.path.exists(LOG_PATH):
        with open(LOG_PATH, "w") as f:
            f.write("initializing...\n")


def get_param(key):
    """Get parameter from JSON file"""
    init_db_maybe()
    with open(PARAMS_PATH, "r") as f:
        params = json.load(f)
        return params.get(key)


def set_param(key, value):
    """Set parameter in JSON file atomically"""

    init_db_maybe()
    # Load existing params first
    with open(PARAMS_PATH, "r") as f:
        params = json.load(f)

    # Update the requested parameter
    params[key] = value

    # Write to temporary file first to prevent corruption
    temp_path = PARAMS_PATH + ".tmp"
    with open(temp_path, "w") as f:
        json.dump(params, f, indent=4)

    # Atomically replace original file
    os.replace(temp_path, PARAMS_PATH)


def get_context(cursor):
    """Get previous and next tokens around cursor position"""
    text = cursor.getText()
    prev_chars = int(get_param("CONTEXT_PREVIOUS_CHARS"))
    next_chars = int(get_param("CONTEXT_NEXT_CHARS"))

    text_cursor = cursor.getText().createTextCursorByRange(cursor)
    text_cursor.goLeft(prev_chars, True)
    previous_context = text_cursor.getString()

    text_cursor = cursor.getText().createTextCursorByRange(cursor)
    text_cursor.goRight(next_chars, True)
    next_context = text_cursor.getString()

    # show_message(previous_context + next_context)

    return previous_context, next_context


def call_llm(data):
    """Make API call to OpenAI-compatible endpoint"""
    url = get_param("OPENAI_ENDPOINT")
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {get_param('OPENAI_API_KEY')}",
    }

    req = urllib.request.Request(
        url, data=json.dumps(data).encode("utf-8"), headers=headers
    )

    try:
        with urllib.request.urlopen(req) as response:
            response_data = json.loads(response.read().decode("utf-8"))
            _log_api_call(url, data, response_data, response.status)
            return response_data
    except urllib.error.HTTPError as e:
        # Read the response body to get the detailed error message
        error_response = e.read().decode("utf-8")
        _log_api_call(url, data, error_response, e.code)
        raise


def _log_api_call(endpoint, request, response, status_code):
    """Log API call details to a regular text file"""
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(
            f"Timestamp: {datetime.datetime.now().isoformat()}\n Endpoint: {endpoint} Status Code: {status_code}\n"
        )
        f.write(f"Request: {request}\n")
        f.write(f"Response: {response}\n")
        f.write("-" * 40 + "\n")


def get_api_logs(limit=100):
    """Retrieve API logs from text file"""
    if not os.path.exists(LOG_PATH):
        return []
    with open(LOG_PATH, "r", encoding="utf-8") as f:
        logs = f.readlines()[-limit:]
        return logs


def show_message(message):
    """Show message dialog"""
    ctx = uno.getComponentContext()
    sm = ctx.getServiceManager()
    toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent = toolkit.getDesktopWindow()
    msgbox = toolkit.createMessageBox(
        parent, "infobox", MSG_BUTTONS.BUTTONS_OK, "LLM Writer", str(message)
    )
    msgbox.execute()


def _get_cursor():
    """Get text cursor from current selection"""
    xModel = (
        uno.getComponentContext()
        .getServiceManager()
        .createInstanceWithContext(
            "com.sun.star.frame.Desktop", uno.getComponentContext()
        )
        .getCurrentComponent()
    )
    xSelectionSupplier = xModel.getCurrentController()
    xIndexAccess = xSelectionSupplier.getSelection()
    return xIndexAccess.getByIndex(0)


def autocomplete(*args):
    """Generate autocomplete suggestions using LLM"""
    try:
        # Check if API key is set
        api_key = get_param("OPENAI_API_KEY")
        if not api_key:
            modify_config()
            return

        cursor = _get_cursor()
        previous_context, next_context = get_context(cursor)

        max_words_prompt = (
            f"\nGenerate at most {get_param('MAX_GENERATION_WORDS')} words\n"
        )

        data = {
            "model": get_param("MODEL"),
            "messages": [
                {
                    "role": "system",
                    "content": get_param("AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS"),
                },
                {
                    "role": "user",
                    "content": f"{previous_context}[COMPLETE HERE]{next_context}",
                },
            ],
            "temperature": float(get_param("TEMPERATURE")),
            "max_tokens": int(get_param("MAX_GENERATION_WORDS")) * 4,
        }

        response = call_llm(data)
        if response:
            cursor.setString(response["choices"][0]["message"]["content"])

    except Exception as e:
        error_msg = f"ERROR: {str(e)}"
        full_error = f"FULL ERROR: {str(e)}\n{traceback.format_exc()}"
        show_message(full_error)


def transform_text(*args):
    """Transform selected text based on instruction"""
    try:
        # Check if API key is set
        api_key = get_param("OPENAI_API_KEY")
        if not api_key:
            modify_config()
            return

        cursor = _get_cursor()
        selected_text = cursor.getString()
        previous_context, next_context = get_context(cursor)

        if not selected_text:
            return

        instruction, keep_original = show_input_dialog_with_checkbox(
            "Enter transformation instructions:", "Keep Original Text", True
        )
        if instruction is None:
            return
        if not instruction:
            instruction = "Perform the task present after the 'Original Text:'"

        data = {
            "model": get_param("MODEL"),
            "messages": [
                {"role": "system", "content": instruction},
                {
                    "role": "user",
                    "content": f"Previous context: {previous_context}\n"
                    f"Original text: {selected_text}\n"
                    f"Next context: {next_context}\n\n"
                    f"Transformed text:",
                },
            ],
            "temperature": float(get_param("TEMPERATURE")),
        }

        response = call_llm(data)
        if response:
            if keep_original:
                cursor.setString(
                    selected_text
                    + "\n\n\u21a6"
                    + response["choices"][0]["message"]["content"]
                    + "\u21a4"
                )
            else:
                cursor.setString(
                    "\u21a6" + response["choices"][0]["message"]["content"] + "\u21a4"
                )

    except Exception as e:
        error_msg = f"ERROR: {str(e)}"
        full_error = f"FULL ERROR: {str(e)}\n{traceback.format_exc()}"
        show_message(full_error)


def show_logs(*args):
    """Display API logs in message box"""
    logs = get_api_logs(10)
    if not logs:
        show_message("No API logs found")
        return

    log_text = "API Logs:\n\n"
    for log in logs:
        log_text += log + "\n"

    show_message(log_text)


def modify_config(*args):
    """Show configuration dialog to modify parameters"""
    WIDTH = 600
    HEIGHT = 400
    HORI_MARGIN = VERT_MARGIN = 8
    BUTTON_WIDTH = 100
    BUTTON_HEIGHT = 26
    HORI_SEP = VERT_SEP = 8
    LABEL_WIDTH = 200
    EDIT_WIDTH = WIDTH - LABEL_WIDTH - HORI_MARGIN * 2 - HORI_SEP
    ROW_HEIGHT = 24
    ROW_SPACING = 4

    init_db_maybe()
    # Load current parameters
    with open(PARAMS_PATH, "r") as f:
        params = json.load(f)

    # Create dialog
    ctx = uno.getComponentContext()

    def create(name):
        return ctx.getServiceManager().createInstanceWithContext(name, ctx)

    dialog = create("com.sun.star.awt.UnoControlDialog")
    dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
    dialog.setModel(dialog_model)
    dialog.setVisible(False)
    dialog.setTitle("Modify Configuration")
    dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)

    # Function to add controls
    def add(name, type, x_, y_, width_, height_, props):
        model = dialog_model.createInstance(
            "com.sun.star.awt.UnoControl" + type + "Model"
        )
        dialog_model.insertByName(name, model)
        control = dialog.getControl(name)
        control.setPosSize(x_, y_, width_, height_, POSSIZE)
        for key, value in props.items():
            setattr(model, key, value)

    # Add parameter fields
    y_pos = VERT_MARGIN
    for i, (key, value) in enumerate(params.items()):
        # Add label
        add(
            f"label_{i}",
            "FixedText",
            HORI_MARGIN,
            y_pos,
            LABEL_WIDTH,
            ROW_HEIGHT,
            {"Label": key, "NoLabel": True},
        )

        # Add edit field 
        add(
            f"edit_{i}",
            "Edit",
            HORI_MARGIN + LABEL_WIDTH + HORI_SEP,
            y_pos,
            EDIT_WIDTH,
        # HACK 
            ROW_HEIGHT + 80 * (key == "AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS"),
            {
                "MultiLine": key == "AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS",
                "VScroll": key == "AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS",
                "Text": str(value),
            },
        )

        y_pos += ROW_HEIGHT + ROW_SPACING + 80 * (key == "AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS")

    # Add buttons
    add(
        "btn_ok",
        "Button",
        WIDTH - BUTTON_WIDTH - HORI_MARGIN,
        y_pos,
        BUTTON_WIDTH,
        BUTTON_HEIGHT,
        {"PushButtonType": OK, "DefaultButton": True},
    )
    add(
        "btn_cancel",
        "Button",
        WIDTH - BUTTON_WIDTH * 2 - HORI_MARGIN - HORI_SEP,
        y_pos,
        BUTTON_WIDTH,
        BUTTON_HEIGHT,
        {"PushButtonType": CANCEL},
    )

    # Position dialog
    frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
    window = frame.getContainerWindow() if frame else None
    dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
    if window:
        ps = window.getPosSize()
        _x = ps.Width / 2 - WIDTH / 2
        _y = ps.Height / 2 - HEIGHT / 2
        dialog.setPosSize(_x, _y, 0, 0, POS)

    # Show dialog and process results
    if dialog.execute():
        # Update parameters with new values
        for i, key in enumerate(params.keys()):
            edit = dialog.getControl(f"edit_{i}")
            params[key] = edit.getModel().Text

        # Save updated parameters
        set_param("OPENAI_ENDPOINT", params["OPENAI_ENDPOINT"])
        set_param("OPENAI_API_KEY", params["OPENAI_API_KEY"])
        set_param("MODEL", params["MODEL"])
        set_param("MAX_GENERATION_WORDS", params["MAX_GENERATION_WORDS"])
        set_param("CONTEXT_PREVIOUS_CHARS", params["CONTEXT_PREVIOUS_CHARS"])
        set_param("CONTEXT_NEXT_CHARS", params["CONTEXT_NEXT_CHARS"])
        set_param("TEMPERATURE", params["TEMPERATURE"])
        set_param(
            "AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS",
            params["AUTOCOMPLETE_ADDITIONAL_INSTRUCTIONS"],
        )

        show_message("Configuration updated successfully!")

    dialog.dispose()


def show_input_dialog(message, title="", default="", x=None, y=None):
    """Shows dialog with input box.
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
    EDIT_HEIGHT = 70
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
        model = dialog_model.createInstance(
            "com.sun.star.awt.UnoControl" + type + "Model"
        )
        dialog_model.insertByName(name, model)
        control = dialog.getControl(name)
        control.setPosSize(x_, y_, width_, height_, POSSIZE)
        for key, value in props.items():
            setattr(model, key, value)

    label_width = WIDTH - BUTTON_WIDTH - HORI_SEP - HORI_MARGIN * 2
    add(
        "label",
        "FixedText",
        HORI_MARGIN,
        VERT_MARGIN,
        label_width,
        LABEL_HEIGHT,
        {"Label": str(message), "NoLabel": True},
    )
    add(
        "btn_ok",
        "Button",
        HORI_MARGIN + label_width + HORI_SEP,
        VERT_MARGIN,
        BUTTON_WIDTH,
        BUTTON_HEIGHT,
        {"PushButtonType": OK, "DefaultButton": True},
    )
    add(
        "btn_cancel",
        "Button",
        HORI_MARGIN + label_width + HORI_SEP,
        VERT_MARGIN + BUTTON_HEIGHT + 5,
        BUTTON_WIDTH,
        BUTTON_HEIGHT,
        {"PushButtonType": CANCEL},
    )
    add(
        "edit",
        "Edit",
        HORI_MARGIN,
        LABEL_HEIGHT + VERT_MARGIN + VERT_SEP,
        WIDTH - HORI_MARGIN * 2,
        EDIT_HEIGHT,
        {"MultiLine": True, "Text": str(default), "VScroll": True},
    )
    frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
    window = frame.getContainerWindow() if frame else None
    dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
    if not x is None and not y is None:
        ps = dialog.convertSizeToPixel(
            uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP
        )
        _x, _y = ps.Width, ps.Height
    elif window:
        ps = window.getPosSize()
        _x = ps.Width / 2 - WIDTH / 2
        _y = ps.Height / 2 - HEIGHT / 2
    dialog.setPosSize(_x, _y, 0, 0, POS)
    edit = dialog.getControl("edit")
    edit.setSelection(
        uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(default)))
    )
    edit.setFocus()
    ret = edit.getModel().Text if dialog.execute() else ""
    dialog.dispose()
    return ret


def show_input_dialog_with_checkbox(
    message, checkbox_label, checkbox_default, title="", default="", x=None, y=None
):
    """Shows dialog with input box and a checkbox.
    @param message message to show on the dialog
    @param checkbox_label label for the checkbox
    @param checkbox_default default state of the checkbox
    @param title window title
    @param default default value for the input box
    @param x optional dialog position in twips
    @param y optional dialog position in twips
    @return tuple (string, boolean) if OK button pushed, otherwise ("", False)
    """
    WIDTH = 600
    HORI_MARGIN = VERT_MARGIN = 8
    BUTTON_WIDTH = 100
    BUTTON_HEIGHT = 26
    HORI_SEP = VERT_SEP = 8
    LABEL_HEIGHT = BUTTON_HEIGHT * 2 + 5
    EDIT_HEIGHT = 70
    CHECKBOX_HEIGHT = 24
    HEIGHT = (
        VERT_MARGIN * 2
        + LABEL_HEIGHT
        + VERT_SEP
        + EDIT_HEIGHT
        + VERT_SEP
        + CHECKBOX_HEIGHT
    )
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
        model = dialog_model.createInstance(
            "com.sun.star.awt.UnoControl" + type + "Model"
        )
        dialog_model.insertByName(name, model)
        control = dialog.getControl(name)
        control.setPosSize(x_, y_, width_, height_, POSSIZE)
        for key, value in props.items():
            setattr(model, key, value)

    label_width = WIDTH - BUTTON_WIDTH - HORI_SEP - HORI_MARGIN * 2
    add(
        "label",
        "FixedText",
        HORI_MARGIN,
        VERT_MARGIN,
        label_width,
        LABEL_HEIGHT,
        {"Label": str(message), "NoLabel": True},
    )
    add(
        "btn_ok",
        "Button",
        HORI_MARGIN + label_width + HORI_SEP,
        VERT_MARGIN,
        BUTTON_WIDTH,
        BUTTON_HEIGHT,
        {"PushButtonType": OK, "DefaultButton": True},
    )
    add(
        "btn_cancel",
        "Button",
        HORI_MARGIN + label_width + HORI_SEP,
        VERT_MARGIN + BUTTON_HEIGHT + 5,
        BUTTON_WIDTH,
        BUTTON_HEIGHT,
        {"PushButtonType": CANCEL},
    )
    add(
        "edit",
        "Edit",
        HORI_MARGIN,
        LABEL_HEIGHT + VERT_MARGIN + VERT_SEP,
        WIDTH - HORI_MARGIN * 2,
        EDIT_HEIGHT,
        {"MultiLine": True, "Text": str(default), "VScroll": True},
    )
    add(
        "checkbox",
        "CheckBox",
        HORI_MARGIN,
        LABEL_HEIGHT + VERT_MARGIN + VERT_SEP + EDIT_HEIGHT + VERT_SEP,
        WIDTH - HORI_MARGIN * 2,
        CHECKBOX_HEIGHT,
        {"Label": checkbox_label, "State": checkbox_default},
    )
    frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
    window = frame.getContainerWindow() if frame else None
    dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
    if not x is None and not y is None:
        ps = dialog.convertSizeToPixel(
            uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP
        )
        _x, _y = ps.Width, ps.Height
    elif window:
        ps = window.getPosSize()
        _x = ps.Width / 2 - WIDTH / 2
        _y = ps.Height / 2 - HEIGHT / 2
    dialog.setPosSize(_x, _y, 0, 0, POS)
    edit = dialog.getControl("edit")
    edit.setSelection(
        uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(default)))
    )
    edit.setFocus()

    if dialog.execute():
        return edit.getModel().Text, dialog.getControl("checkbox").getModel().State
    else:
        return None, False

    dialog.dispose()


# Export the macros properly
init_db_maybe()
g_exportedScripts = (autocomplete, transform_text, show_logs, modify_config)
