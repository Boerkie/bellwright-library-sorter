import pyautogui
import keyboard
import time
from PIL import Image, ImageGrab, ImageOps, ImageDraw # ImageEnhance removed
import numpy as np
import pygetwindow as gw
from datetime import datetime
import configparser
import os

# --- Tesseract Configuration (Module Level Import) ---
try:
    import pytesseract
except ImportError:
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}] WARNING: Pytesseract library not found. OCR will be disabled.")
    pytesseract = None
except Exception as e:
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}] WARNING: Error during initial pytesseract import: {e}. OCR will be disabled.")
    pytesseract = None

# --- CONFIGURATION LOADING ---
CONFIG_FILE = 'config.ini'
config = configparser.ConfigParser() # Global config object

DEFAULT_CONFIG = { # All keys here should be lowercase
    'general': {'gamewindowtitle': 'Bellwright', 'tesseractcmdpath': r'C:\Program Files\Tesseract-OCR\tesseract.exe'},
    'hotkeys': {
        'calculateandplan': 'num_1', 'executesort': 'num_2',
        'startfulluicalibration': 'num_4',
        'setgridorigin': 'num_3', 'calibrateslotdimensions': 'num_5', 'calibrateslotxgap': 'num_6',
        'calibrateslotygap': 'num_plus', 'calibratetiercolorpoint': 'num_7', 'calibrateocrregion': 'num_8',
        'savecalibratedvalues': 'num_9', 'exitscript': 'num_0',
    },
    'gridstructure': {
        'gridoffsetx': '310', 'gridoffsety': '170', 'numcols': '6', 'maxnumrows': '10',
        'slotwidth': '83', 'slotheight': '83', 'slotgapx': '10', 'slotgapy': '10',
    },
    'colorpatch': {'relativex': '15', 'relativey': '68', 'size': '10', 'tolerance': '30'},
    'tiercolors': {'tier1': '47,67,81', 'tier2': '81,89,42', 'tier3': '95,64,40',
                   'tier4': '102,41,35', 'tier5': '61,50,85'},
    'ocr': {'relativex': '8', 'relativey': '8', 'width': '30', 'height': '25',
            'upscalefactor': '2', 'thresholdvalue': '180'},
    'mousemovement': {'moveduration': '0.20', 'dragduration': '0.30', 'postactiondelay': '0.30'}
}

# --- Global config variables ---
GAME_WINDOW_TITLE, TESSERACT_CMD_PATH = '', ''
CALCULATE_HOTKEY, EXECUTE_SORT_HOTKEY, FULL_UI_CALIBRATION_HOTKEY, SAVE_CONFIG_HOTKEY, EXIT_SCRIPT_HOTKEY = '', '', '', '', ''
# Individual calibration hotkeys
SET_GRID_ORIGIN_IND_HOTKEY, CALIBRATE_SLOT_DIM_IND_HOTKEY, CALIBRATE_SLOT_X_GAP_IND_HOTKEY = '', '', ''
CALIBRATE_SLOT_Y_GAP_IND_HOTKEY, CALIBRATE_TIER_COLOR_IND_HOTKEY, CALIBRATE_OCR_REGION_IND_HOTKEY = '', '', ''

GRID_OFFSET_X, GRID_OFFSET_Y, NUM_COLS, MAX_NUM_ROWS = 0, 0, 0, 0
SLOT_WIDTH, SLOT_HEIGHT, SLOT_GAP_X, SLOT_GAP_Y = 0, 0, 0, 0
COLOR_PATCH_RELATIVE_X, COLOR_PATCH_RELATIVE_Y, COLOR_PATCH_SIZE, COLOR_TOLERANCE = 0, 0, 0, 0
TIER_COLORS = {}
OCR_RELATIVE_X, OCR_RELATIVE_Y, OCR_WIDTH, OCR_HEIGHT = 0, 0, 0, 0
OCR_UPSCALE_FACTOR, OCR_THRESHOLD_VALUE = 0, 0
MOUSE_MOVE_DURATION, DRAG_DURATION, POST_ACTION_DELAY = 0.0, 0.0, 0.0

def log_message(message):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]}] {message}")

def load_config():
    global GAME_WINDOW_TITLE, TESSERACT_CMD_PATH, CALCULATE_HOTKEY, EXECUTE_SORT_HOTKEY, \
           FULL_UI_CALIBRATION_HOTKEY, SAVE_CONFIG_HOTKEY, EXIT_SCRIPT_HOTKEY, \
           SET_GRID_ORIGIN_IND_HOTKEY, CALIBRATE_SLOT_DIM_IND_HOTKEY, CALIBRATE_SLOT_X_GAP_IND_HOTKEY, \
           CALIBRATE_SLOT_Y_GAP_IND_HOTKEY, CALIBRATE_TIER_COLOR_IND_HOTKEY, CALIBRATE_OCR_REGION_IND_HOTKEY, \
           GRID_OFFSET_X, GRID_OFFSET_Y, NUM_COLS, MAX_NUM_ROWS, SLOT_WIDTH, SLOT_HEIGHT, \
           SLOT_GAP_X, SLOT_GAP_Y, COLOR_PATCH_RELATIVE_X, COLOR_PATCH_RELATIVE_Y, \
           COLOR_PATCH_SIZE, COLOR_TOLERANCE, TIER_COLORS, OCR_RELATIVE_X, OCR_RELATIVE_Y, \
           OCR_WIDTH, OCR_HEIGHT, OCR_UPSCALE_FACTOR, OCR_THRESHOLD_VALUE, \
           MOUSE_MOVE_DURATION, DRAG_DURATION, POST_ACTION_DELAY

    if not os.path.exists(CONFIG_FILE):
        log_message(f"WARNING: {CONFIG_FILE} not found. Writing default config.")
        default_cfg_obj = configparser.ConfigParser(interpolation=None)
        for section, options in DEFAULT_CONFIG.items(): default_cfg_obj[section.lower()] = options # Write sections in lowercase
        with open(CONFIG_FILE, 'w') as configfile: default_cfg_obj.write(configfile)
    
    config.read(CONFIG_FILE) # configparser reads sections/options case-insensitively

    def get_cfg_val(section_name, option_name, type_func=str, is_int=False, is_float=False):
        # Section and option names for DEFAULT_CONFIG are all lowercase
        default_section_name = section_name.lower()
        default_option_name = option_name.lower()
        try:
            # Read from config file (case-insensitive for configparser)
            # Fallback directly to DEFAULT_CONFIG value if not found in file
            val_str = config.get(section_name, option_name, 
                                 fallback=DEFAULT_CONFIG[default_section_name][default_option_name])
            if is_int: return int(val_str)
            if is_float: return float(val_str)
            return type_func(val_str)
        except (ValueError) as e:
            log_message(f"Config ERROR: Invalid value for '{option_name}' in '[{section_name}]': '{config.get(section_name,option_name, fallback='ERROR_NO_FALLBACK')}' (Error: {e}). Using hardcoded default.")
            val_str = DEFAULT_CONFIG[default_section_name][default_option_name] # Get default again
            if is_int: return int(val_str)
            if is_float: return float(val_str)
            return type_func(val_str)
        except Exception as e:
            log_message(f"Config CRITICAL ERROR for '{option_name}' in '[{section_name}]' (Error: {e}).")
            if is_int: return 0
            if is_float: return 0.0
            return ""

    GAME_WINDOW_TITLE = get_cfg_val('General', 'GameWindowTitle')
    TESSERACT_CMD_PATH = get_cfg_val('General', 'TesseractCmdPath')
    CALCULATE_HOTKEY = get_cfg_val('Hotkeys', 'CalculateAndPlan') # Keys from INI can be mixed case
    EXECUTE_SORT_HOTKEY = get_cfg_val('Hotkeys', 'ExecuteSort')
    FULL_UI_CALIBRATION_HOTKEY = get_cfg_val('Hotkeys', 'StartFullUICalibration')
    SAVE_CONFIG_HOTKEY = get_cfg_val('Hotkeys', 'SaveCalibratedValues')
    EXIT_SCRIPT_HOTKEY = get_cfg_val('Hotkeys', 'ExitScript')
    SET_GRID_ORIGIN_IND_HOTKEY = get_cfg_val('Hotkeys', 'SetGridOrigin')
    CALIBRATE_SLOT_DIM_IND_HOTKEY = get_cfg_val('Hotkeys', 'CalibrateSlotDimensions')
    CALIBRATE_SLOT_X_GAP_IND_HOTKEY = get_cfg_val('Hotkeys', 'CalibrateSlotXGap')
    CALIBRATE_SLOT_Y_GAP_IND_HOTKEY = get_cfg_val('Hotkeys', 'CalibrateSlotYGap')
    CALIBRATE_TIER_COLOR_IND_HOTKEY = get_cfg_val('Hotkeys', 'CalibrateTierColorPoint')
    CALIBRATE_OCR_REGION_IND_HOTKEY = get_cfg_val('Hotkeys', 'CalibrateOCRRegion')

    GRID_OFFSET_X = get_cfg_val('GridStructure', 'GridOffsetX', is_int=True)
    GRID_OFFSET_Y = get_cfg_val('GridStructure', 'GridOffsetY', is_int=True)
    NUM_COLS = get_cfg_val('GridStructure', 'NumCols', is_int=True); MAX_NUM_ROWS = get_cfg_val('GridStructure', 'MaxNumRows', is_int=True)
    SLOT_WIDTH = get_cfg_val('GridStructure', 'SlotWidth', is_int=True); SLOT_HEIGHT = get_cfg_val('GridStructure', 'SlotHeight', is_int=True)
    SLOT_GAP_X = get_cfg_val('GridStructure', 'SlotGapX', is_int=True); SLOT_GAP_Y = get_cfg_val('GridStructure', 'SlotGapY', is_int=True)
    COLOR_PATCH_RELATIVE_X = get_cfg_val('ColorPatch', 'RelativeX', is_int=True)
    COLOR_PATCH_RELATIVE_Y = get_cfg_val('ColorPatch', 'RelativeY', is_int=True)
    COLOR_PATCH_SIZE = get_cfg_val('ColorPatch', 'Size', is_int=True); COLOR_TOLERANCE = get_cfg_val('ColorPatch', 'Tolerance', is_int=True)
    TIER_COLORS = {}
    for i in range(1, 6):
        key_in_ini = f'Tier{i}'; key_in_default = f'tier{i}' # default keys are lowercase
        rgb_str = get_cfg_val('TierColors', key_in_ini)
        try: TIER_COLORS[i] = tuple(map(int, rgb_str.split(',')))
        except: TIER_COLORS[i] = tuple(map(int, DEFAULT_CONFIG['tiercolors'][key_in_default].split(','))) # Access DEFAULT_CONFIG with lowercase section
    OCR_RELATIVE_X = get_cfg_val('OCR', 'RelativeX', is_int=True); OCR_RELATIVE_Y = get_cfg_val('OCR', 'RelativeY', is_int=True)
    OCR_WIDTH = get_cfg_val('OCR', 'Width', is_int=True); OCR_HEIGHT = get_cfg_val('OCR', 'Height', is_int=True)
    OCR_UPSCALE_FACTOR = get_cfg_val('OCR', 'UpscaleFactor', is_int=True); OCR_THRESHOLD_VALUE = get_cfg_val('OCR', 'ThresholdValue', is_int=True)
    MOUSE_MOVE_DURATION = get_cfg_val('MouseMovement', 'MoveDuration', is_float=True)
    DRAG_DURATION = get_cfg_val('MouseMovement', 'DragDuration', is_float=True)
    POST_ACTION_DELAY = get_cfg_val('MouseMovement', 'PostActionDelay', is_float=True)

pytesseract_available = False
def initialize_tesseract(): # Unchanged from previous working version
    global pytesseract_available, TESSERACT_CMD_PATH, pytesseract
    if pytesseract is None: log_message("Pytesseract module not imported. OCR disabled."); pytesseract_available=False; return
    try:
        if TESSERACT_CMD_PATH and os.path.exists(TESSERACT_CMD_PATH): pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD_PATH
        version = pytesseract.get_tesseract_version()
        log_message(f"Tesseract {version} (Cmd: '{pytesseract.pytesseract.tesseract_cmd}').")
        pytesseract_available = True
    except Exception as e: log_message(f"WARN: Init Tesseract (Path:'{TESSERACT_CMD_PATH}'): {e}"); pytesseract_available=False

is_processing = False; last_calculated_plan = None; script_running = True
DEBUG_IMAGE_FOLDER = "execution_debug_images"; full_ui_calibration_state = {"active": False, "step": 0, "points": []}
individual_calibration_tool_state = {} # For individual calibration tools

def ensure_debug_folder(): # Unchanged
    if not os.path.exists(DEBUG_IMAGE_FOLDER): os.makedirs(DEBUG_IMAGE_FOLDER); log_message(f"Created: {DEBUG_IMAGE_FOLDER}")

# --- UNIFIED & INDIVIDUAL UI CALIBRATION FUNCTIONS ---
# (start_or_advance_full_ui_calibration, set_grid_origin_individually, calibrate_slot_dimensions_individually, etc. from previous message)
# Ensure they use the correct global hotkey variables loaded from config for their log messages.
# Example for one individual function:
def set_grid_origin_individually():
    global GRID_OFFSET_X, GRID_OFFSET_Y, is_processing
    if is_processing: log_message("Busy."); return
    is_processing = True
    log_message(f"INDIV. CALIB: GRID ORIGIN. Hover TOP-LEFT of FIRST slot. Press {SET_GRID_ORIGIN_IND_HOTKEY} again.") # Use loaded hotkey
    
    tool_name = "IndividualGridOrigin"
    # ... (rest of the logic for this function as provided before) ...
    if tool_name not in individual_calibration_tool_state:
        individual_calibration_tool_state[tool_name] = {"step":1}
        is_processing = False; return
    game_rect = get_game_window_rect(); # ... (rest of function)
    if not game_rect: log_message("ERR: Game window not found."); is_processing=False; del individual_calibration_tool_state[tool_name]; return
    game_x_abs, game_y_abs, _, _ = game_rect; mx, my = pyautogui.position()
    GRID_OFFSET_X = mx - game_x_abs; GRID_OFFSET_Y = my - game_y_abs
    log_message(f"SUCCESS: Indiv. Grid Origin: X={GRID_OFFSET_X}, Y={GRID_OFFSET_Y}. Press {SAVE_CONFIG_HOTKEY} to save.")
    del individual_calibration_tool_state[tool_name]; is_processing = False


def start_or_advance_full_ui_calibration():
    global full_ui_calibration_state, is_processing, GRID_OFFSET_X, GRID_OFFSET_Y, \
           SLOT_WIDTH, SLOT_HEIGHT, SLOT_GAP_X, SLOT_GAP_Y, \
           COLOR_PATCH_RELATIVE_X, COLOR_PATCH_RELATIVE_Y, \
           OCR_RELATIVE_X, OCR_RELATIVE_Y, OCR_WIDTH, OCR_HEIGHT

    if is_processing and not full_ui_calibration_state["active"]:
        log_message("Busy with another process. Try calibration again shortly.")
        return
    is_processing = True # Mark as busy for other tools, unless it's this tool itself

    state = full_ui_calibration_state
    calib_hotkey = FULL_UI_CALIBRATION_HOTKEY # From global config

    if not state["active"]: # First press to start calibration
        state["active"] = True
        state["step"] = 1
        state["points"] = []
        log_message(f"--- Full UI Calibration Started (Hotkey: {calib_hotkey}) ---")
    
    mx, my = pyautogui.position() # Get current mouse position for each step

    if state["step"] == 1:
        log_message(f"STEP 1/8: Hover over TOP-LEFT of FIRST slot (row 0, col 0). Press {calib_hotkey}.")
        if len(state["points"]) == 0: state["points"].append(None) # Placeholder for P1
        else: # This is the click for P1
            game_rect = get_game_window_rect()
            if not game_rect: log_message("ERR: Game window not found. Calibration aborted."); state["active"]=False; is_processing=False; return
            game_x_abs, game_y_abs, _, _ = game_rect
            state["points"][0] = (mx, my)
            GRID_OFFSET_X = mx - game_x_abs
            GRID_OFFSET_Y = my - game_y_abs
            log_message(f"  P1 (Slot0 TL / Grid Origin) captured at ({mx},{my}). Relative Grid Offset: X={GRID_OFFSET_X}, Y={GRID_OFFSET_Y}")
            state["step"] = 2
        is_processing = False; return # Wait for next click

    if state["step"] == 2:
        log_message(f"STEP 2/8: Hover over TOP-RIGHT of SAME FIRST slot. Press {calib_hotkey}.")
        if len(state["points"]) == 1: state["points"].append(None)
        else:
            state["points"][1] = (mx, my)
            p1x, _ = state["points"][0]
            SLOT_WIDTH = abs(mx - p1x)
            log_message(f"  P2 (Slot0 TR) captured at ({mx},{my}). Calculated SLOT_WIDTH: {SLOT_WIDTH}")
            state["step"] = 3
        is_processing = False; return

    if state["step"] == 3:
        log_message(f"STEP 3/8: Hover over BOTTOM-LEFT of SAME FIRST slot. Press {calib_hotkey}.")
        if len(state["points"]) == 2: state["points"].append(None)
        else:
            state["points"][2] = (mx, my)
            _, p1y = state["points"][0]
            SLOT_HEIGHT = abs(my - p1y)
            log_message(f"  P3 (Slot0 BL) captured at ({mx},{my}). Calculated SLOT_HEIGHT: {SLOT_HEIGHT}")
            state["step"] = 4
        is_processing = False; return
        
    if state["step"] == 4:
        log_message(f"STEP 4/8: Hover over TOP-LEFT of slot TO THE RIGHT (row 0, col 1). Press {calib_hotkey}.")
        if len(state["points"]) == 3: state["points"].append(None)
        else:
            state["points"][3] = (mx,my)
            p2x, _ = state["points"][1] # Top-right of first slot
            SLOT_GAP_X = abs(mx - p2x)
            log_message(f"  P4 (Slot1 TL) captured at ({mx},{my}). Calculated SLOT_GAP_X: {SLOT_GAP_X}")
            state["step"] = 5
        is_processing = False; return

    if state["step"] == 5:
        log_message(f"STEP 5/8: Hover over TOP-LEFT of slot BELOW (row 1, col 0). Press {calib_hotkey}.")
        if len(state["points"]) == 4: state["points"].append(None)
        else:
            state["points"][4] = (mx,my)
            p1x, p1y = state["points"][0] # Top-left of first slot
            # SLOT_GAP_Y = abs(my - p3y) if using p3 (bottom-left of first slot)
            SLOT_GAP_Y = abs(my - (p1y + SLOT_HEIGHT)) # Gap = Slot1_TopY - (Slot0_TopY + Slot0_Height)
            log_message(f"  P5 (Slot_Below TL) captured at ({mx},{my}). Calculated SLOT_GAP_Y: {SLOT_GAP_Y}")
            state["step"] = 6
        is_processing = False; return

    if state["step"] == 6: # Tier Color Patch CENTER relative to Slot0 TL
        log_message(f"STEP 6/8 (Tier Color): Ref: Slot0 TL ({state['points'][0]}). Hover over CENTER of TIER COLOR sample area on book. Press {calib_hotkey}.")
        if len(state["points"]) == 5: state["points"].append(None)
        else:
            state["points"][5] = (mx,my)
            p1x, p1y = state["points"][0]
            COLOR_PATCH_RELATIVE_X = mx - p1x
            COLOR_PATCH_RELATIVE_Y = my - p1y
            log_message(f"  P6 (TierColor Center) at ({mx},{my}). Relative to Slot0 TL: X={COLOR_PATCH_RELATIVE_X}, Y={COLOR_PATCH_RELATIVE_Y}")
            log_message(f"    (COLOR_PATCH_SIZE={COLOR_PATCH_SIZE} from config defines sample area around this point)")
            state["step"] = 7
        is_processing = False; return

    if state["step"] == 7: # OCR TL relative to Slot0 TL
        log_message(f"STEP 7/8 (Stack# TL): Ref: Slot0 TL ({state['points'][0]}). Hover over TOP-LEFT of STACK COUNT number. Press {calib_hotkey}.")
        if len(state["points"]) == 6: state["points"].append(None)
        else:
            state["points"][6] = (mx,my) # This is P7 (OCR TL Absolute)
            p1x, p1y = state["points"][0]
            OCR_RELATIVE_X = mx - p1x
            OCR_RELATIVE_Y = my - p1y
            log_message(f"  P7 (Stack# TL) at ({mx},{my}). Relative to Slot0 TL: X={OCR_RELATIVE_X}, Y={OCR_RELATIVE_Y}")
            state["step"] = 8
        is_processing = False; return

    if state["step"] == 8: # OCR BR relative to OCR TL (P7)
        log_message(f"STEP 8/8 (Stack# BR): Ref: Stack# TL ({state['points'][6]}). Hover over BOTTOM-RIGHT of STACK COUNT number. Press {calib_hotkey}.")
        if len(state["points"]) == 7: state["points"].append(None) # This will be P8 (OCR BR Absolute)
        else:
            p7x_abs, p7y_abs = state["points"][6] # P7 (OCR TL Absolute)
            ocr_br_x_abs, ocr_br_y_abs = mx, my   # P8 (OCR BR Absolute)
            OCR_WIDTH = abs(ocr_br_x_abs - p7x_abs)
            OCR_HEIGHT = abs(ocr_br_y_abs - p7y_abs)
            log_message(f"  P8 (Stack# BR) at ({mx},{my}). Calculated OCR_WIDTH: {OCR_WIDTH}, OCR_HEIGHT: {OCR_HEIGHT}")
            log_message(f"--- Full UI Calibration Complete! ---")
            log_message(f"  GridOffset:({GRID_OFFSET_X},{GRID_OFFSET_Y}), Slot:({SLOT_WIDTH}x{SLOT_HEIGHT}), Gap:({SLOT_GAP_X}x{SLOT_GAP_Y})")
            log_message(f"  ColorPatch Rel:({COLOR_PATCH_RELATIVE_X},{COLOR_PATCH_RELATIVE_Y}), Size:{COLOR_PATCH_SIZE}")
            log_message(f"  OCR Rel:({OCR_RELATIVE_X},{OCR_RELATIVE_Y}), Size:({OCR_WIDTH}x{OCR_HEIGHT})")
            log_message(f"Press '{SAVE_CONFIG_HOTKEY}' to save these values to config.ini.")
            state["active"] = False # Reset calibration state
            state["step"] = 0
            state["points"] = []
        is_processing = False; return
 
def calibrate_slot_dimensions_individually():
    global SLOT_WIDTH, SLOT_HEIGHT, individual_calibration_tool_state, is_processing
    if is_processing: log_message("Busy."); return
    is_processing = True
    
    tool_name = "IndividualSlotDimensions"
    hotkey = config.get('Hotkeys', 'CalibrateSlotDimensions')

    if tool_name not in individual_calibration_tool_state:
        log_message(f"INDIV. CALIB: SLOT DIM. STEP 1: Hover TOP-LEFT of ANY slot. Press {hotkey}.")
        individual_calibration_tool_state[tool_name] = {"step": 1}
        is_processing = False; return

    if individual_calibration_tool_state[tool_name]["step"] == 1:
        mx, my = pyautogui.position()
        individual_calibration_tool_state[tool_name]["p1"] = (mx, my)
        individual_calibration_tool_state[tool_name]["step"] = 2
        log_message(f"  P1 (Slot TL) captured at ({mx},{my}). STEP 2: Hover BOTTOM-RIGHT of SAME slot. Press {hotkey}.")
        is_processing = False; return

    if individual_calibration_tool_state[tool_name]["step"] == 2:
        mx, my = pyautogui.position()
        p1x, p1y = individual_calibration_tool_state[tool_name]["p1"]
        SLOT_WIDTH, SLOT_HEIGHT = abs(mx - p1x), abs(my - p1y)
        log_message(f"  P2 (Slot BR) at ({mx},{my}). SUCCESS: SLOT_WIDTH={SLOT_WIDTH}, SLOT_HEIGHT={SLOT_HEIGHT}")
        log_message(f"  Press {SAVE_CONFIG_HOTKEY} to save.")
        del individual_calibration_tool_state[tool_name]
        is_processing = False; return

def calibrate_slot_x_gap_individually():
    global SLOT_GAP_X, individual_calibration_tool_state, is_processing
    # ... (similar 2-step logic: click Top-Right of slot A, then Top-Left of slot B (to its right))
    if is_processing: log_message("Busy."); return; is_processing=True
    tool="IndXGap"; hotkey=config.get('Hotkeys','CalibrateSlotXGap')
    if tool not in individual_calibration_tool_state:
        log_message(f"INDIV. CALIB: X-GAP. S1: Hover TR of Slot A. Press {hotkey}.")
        individual_calibration_tool_state[tool]={"step":1}; is_processing=False; return
    if individual_calibration_tool_state[tool]["step"]==1:
        mx,my=pyautogui.position(); individual_calibration_tool_state[tool]["p1"]=(mx,my); individual_calibration_tool_state[tool]["step"]=2
        log_message(f"  P1 (SlotA TR) at ({mx},{my}). S2: Hover TL of Slot B (right of A). Press {hotkey}.")
        is_processing=False; return
    if individual_calibration_tool_state[tool]["step"]==2:
        mx,my=pyautogui.position(); p1x,_=individual_calibration_tool_state[tool]["p1"]
        SLOT_GAP_X=abs(mx-p1x)
        log_message(f"  P2 (SlotB TL) at ({mx},{my}). SUCCESS: SLOT_GAP_X={SLOT_GAP_X}. Save with {SAVE_CONFIG_HOTKEY}.")
        del individual_calibration_tool_state[tool]; is_processing=False; return

def calibrate_slot_y_gap_individually():
    global SLOT_GAP_Y, individual_calibration_tool_state, is_processing
    # ... (similar 2-step logic: click Bottom-Left of slot A, then Top-Left of slot B (below it))
    if is_processing: log_message("Busy."); return; is_processing=True
    tool="IndYGap"; hotkey=config.get('Hotkeys','CalibrateSlotYGap')
    if tool not in individual_calibration_tool_state:
        log_message(f"INDIV. CALIB: Y-GAP. S1: Hover BL of Slot A. Press {hotkey}.")
        individual_calibration_tool_state[tool]={"step":1}; is_processing=False; return
    if individual_calibration_tool_state[tool]["step"]==1:
        mx,my=pyautogui.position(); individual_calibration_tool_state[tool]["p1"]=(mx,my); individual_calibration_tool_state[tool]["step"]=2
        log_message(f"  P1 (SlotA BL) at ({mx},{my}). S2: Hover TL of Slot B (below A). Press {hotkey}.")
        is_processing=False; return
    if individual_calibration_tool_state[tool]["step"]==2:
        mx,my=pyautogui.position(); _,p1y=individual_calibration_tool_state[tool]["p1"]
        SLOT_GAP_Y=abs(my-p1y)
        log_message(f"  P2 (SlotB TL) at ({mx},{my}). SUCCESS: SLOT_GAP_Y={SLOT_GAP_Y}. Save with {SAVE_CONFIG_HOTKEY}.")
        del individual_calibration_tool_state[tool]; is_processing=False; return

def calibrate_tier_color_point_individually(): # Calibrates COLOR_PATCH_RELATIVE_X/Y
    global COLOR_PATCH_RELATIVE_X, COLOR_PATCH_RELATIVE_Y, individual_calibration_tool_state, is_processing
    # ... (2-step: click Top-Left of a reference slot, then click center of color patch WITHIN that slot)
    if is_processing: log_message("Busy."); return; is_processing=True
    tool="IndColorPatch"; hotkey=config.get('Hotkeys','CalibrateTierColorPoint')
    if tool not in individual_calibration_tool_state:
        log_message(f"INDIV. CALIB: TIER COLOR POINT. S1: Hover TL of ANY slot (ref). Press {hotkey}.")
        individual_calibration_tool_state[tool]={"step":1}; is_processing=False; return
    if individual_calibration_tool_state[tool]["step"]==1:
        mx,my=pyautogui.position(); individual_calibration_tool_state[tool]["ref_slot_tl"]=(mx,my); individual_calibration_tool_state[tool]["step"]=2
        log_message(f"  Ref Slot TL at ({mx},{my}). S2: Hover CENTER of Tier Color area IN THIS SLOT. Press {hotkey}.")
        is_processing=False; return
    if individual_calibration_tool_state[tool]["step"]==2:
        mx,my=pyautogui.position(); r_sx,r_sy=individual_calibration_tool_state[tool]["ref_slot_tl"]
        COLOR_PATCH_RELATIVE_X, COLOR_PATCH_RELATIVE_Y = mx-r_sx, my-r_sy
        log_message(f"  Color Patch Center at ({mx},{my}). SUCCESS: COLOR_PATCH_RELATIVE_X={COLOR_PATCH_RELATIVE_X}, Y={COLOR_PATCH_RELATIVE_Y}.")
        log_message(f"    (COLOR_PATCH_SIZE={COLOR_PATCH_SIZE} from config is used around this). Save with {SAVE_CONFIG_HOTKEY}.")
        del individual_calibration_tool_state[tool]; is_processing=False; return

def calibrate_ocr_region_individually(): # Calibrates OCR_RELATIVE_X/Y and OCR_WIDTH/HEIGHT
    global OCR_RELATIVE_X, OCR_RELATIVE_Y, OCR_WIDTH, OCR_HEIGHT, individual_calibration_tool_state, is_processing
    # ... (3-step: click Top-Left of ref slot, then TL of number in that slot, then BR of number in that slot)
    if is_processing: log_message("Busy."); return; is_processing=True
    tool="IndOCR"; hotkey=config.get('Hotkeys','CalibrateOCRRegion')
    if tool not in individual_calibration_tool_state:
        log_message(f"INDIV. CALIB: OCR REGION. S1: Hover TL of ANY slot (ref). Press {hotkey}.")
        individual_calibration_tool_state[tool]={"step":1}; is_processing=False; return
    if individual_calibration_tool_state[tool]["step"]==1: # Slot TL
        mx,my=pyautogui.position(); individual_calibration_tool_state[tool]["ref_slot_tl"]=(mx,my); individual_calibration_tool_state[tool]["step"]=2
        log_message(f"  Ref Slot TL at ({mx},{my}). S2: Hover TL of STACK NUMBER in this slot. Press {hotkey}.")
        is_processing=False; return
    if individual_calibration_tool_state[tool]["step"]==2: # OCR TL
        mx,my=pyautogui.position(); individual_calibration_tool_state[tool]["ocr_tl_abs"]=(mx,my); individual_calibration_tool_state[tool]["step"]=3
        log_message(f"  Stack# TL at ({mx},{my}). S3: Hover BR of SAME STACK NUMBER. Press {hotkey}.")
        is_processing=False; return
    if individual_calibration_tool_state[tool]["step"]==3: # OCR BR
        mx,my=pyautogui.position(); r_sx,r_sy=individual_calibration_tool_state[tool]["ref_slot_tl"]; ocr_tlx,ocr_tly=individual_calibration_tool_state[tool]["ocr_tl_abs"]
        OCR_RELATIVE_X, OCR_RELATIVE_Y = ocr_tlx-r_sx, ocr_tly-r_sy
        OCR_WIDTH, OCR_HEIGHT = abs(mx-ocr_tlx), abs(my-ocr_tly)
        if OCR_WIDTH<=0 or OCR_HEIGHT<=0: log_message("ERR: OCR Width/Height is 0/neg. Try again.")
        else: log_message(f"  Stack# BR at ({mx},{my}). SUCCESS: OCR Rel.X={OCR_RELATIVE_X}, Rel.Y={OCR_RELATIVE_Y}, W={OCR_WIDTH}, H={OCR_HEIGHT}. Save with {SAVE_CONFIG_HOTKEY}.")
        del individual_calibration_tool_state[tool]; is_processing=False; return

def save_calibrated_values_to_config(): # Updated to save new calibrated globals
    global GRID_OFFSET_X, GRID_OFFSET_Y, SLOT_WIDTH, SLOT_HEIGHT, SLOT_GAP_X, SLOT_GAP_Y, \
           COLOR_PATCH_RELATIVE_X, COLOR_PATCH_RELATIVE_Y, \
           OCR_RELATIVE_X, OCR_RELATIVE_Y, OCR_WIDTH, OCR_HEIGHT, \
           NUM_COLS, MAX_NUM_ROWS, COLOR_PATCH_SIZE, COLOR_TOLERANCE, TIER_COLORS # Add others if needed
    
    log_message("Saving current calibrated values to config.ini...")
    cfg_to_save = configparser.ConfigParser(interpolation=None) # Use interpolation=None
    
    # Read existing config FIRST to preserve comments and structure as much as possible
    # However, configparser itself will strip comments on write.
    # The main goal here is to preserve sections and options not touched by calibration.
    if os.path.exists(CONFIG_FILE):
        cfg_to_save.read(CONFIG_FILE)

    # Ensure sections exist using lowercase names (consistent with DEFAULT_CONFIG)
    sections_to_update_directly = ['gridstructure', 'colorpatch', 'ocr', 'tiercolors'] # lowercase
    all_expected_sections = list(DEFAULT_CONFIG.keys()) # These are already lowercase

    for section_key_lc in all_expected_sections: # Iterate with lowercase keys
        if not cfg_to_save.has_section(section_key_lc):
            # configparser often creates sections with the case provided.
            # If we want to ensure lowercase sections in the output file,
            # we should add them with lowercase names if they don't exist.
            # However, if config.ini had [General], has_section('general') might be false.
            # Let's find the actual section name if it exists with different casing.
            actual_section_name = next((s for s in cfg_to_save.sections() if s.lower() == section_key_lc), None)
            if not actual_section_name:
                cfg_to_save.add_section(section_key_lc) # Add with lowercase if truly missing
            # If we want to force all sections to be lowercase on save, more work is needed
            # to rebuild the config object with lowercase section names.
            # For now, let's assume we primarily care about options within existing (possibly mixed-case) sections.

    # Helper to set value, ensuring section exists (using lowercase for new sections)
    def safe_set(parser, section_lc, option_lc, value_str):
        # Find the actual section name (could be mixed case in existing INI)
        actual_section = next((s for s in parser.sections() if s.lower() == section_lc), section_lc)
        if not parser.has_section(actual_section):
            parser.add_section(actual_section) # Will add as section_lc if it was truly new
        parser.set(actual_section, option_lc, value_str)


    # Update GridStructure (using lowercase keys for set)
    safe_set(cfg_to_save, 'gridstructure', 'gridoffsetx', str(GRID_OFFSET_X))
    safe_set(cfg_to_save, 'gridstructure', 'gridoffsety', str(GRID_OFFSET_Y))
    safe_set(cfg_to_save, 'gridstructure', 'slotwidth', str(SLOT_WIDTH))
    safe_set(cfg_to_save, 'gridstructure', 'slotheight', str(SLOT_HEIGHT))
    safe_set(cfg_to_save, 'gridstructure', 'slotgapx', str(SLOT_GAP_X))
    safe_set(cfg_to_save, 'gridstructure', 'slotgapy', str(SLOT_GAP_Y))
    safe_set(cfg_to_save, 'gridstructure', 'numcols', str(NUM_COLS))
    safe_set(cfg_to_save, 'gridstructure', 'maxnumrows', str(MAX_NUM_ROWS))

    # Update ColorPatch
    safe_set(cfg_to_save, 'colorpatch', 'relativex', str(COLOR_PATCH_RELATIVE_X))
    safe_set(cfg_to_save, 'colorpatch', 'relativey', str(COLOR_PATCH_RELATIVE_Y))
    safe_set(cfg_to_save, 'colorpatch', 'size', str(COLOR_PATCH_SIZE))
    safe_set(cfg_to_save, 'colorpatch', 'tolerance', str(COLOR_TOLERANCE))
    
    # Update OCR
    safe_set(cfg_to_save, 'ocr', 'relativex', str(OCR_RELATIVE_X))
    safe_set(cfg_to_save, 'ocr', 'relativey', str(OCR_RELATIVE_Y))
    safe_set(cfg_to_save, 'ocr', 'width', str(OCR_WIDTH))
    safe_set(cfg_to_save, 'ocr', 'height', str(OCR_HEIGHT))
    safe_set(cfg_to_save, 'ocr', 'upscalefactor', str(OCR_UPSCALE_FACTOR))
    safe_set(cfg_to_save, 'ocr', 'thresholdvalue', str(OCR_THRESHOLD_VALUE))

    # TierColors
    for tier, rgb in TIER_COLORS.items():
         safe_set(cfg_to_save, 'tiercolors', f'tier{tier}', f"{rgb[0]},{rgb[1]},{rgb[2]}") # key is already tierX
    
    # Preserve other sections and their options that were not directly calibrated
    # This ensures we don't lose [General], [Hotkeys], [MouseMovement]
    for section_lc in DEFAULT_CONFIG: # section_lc is e.g., 'general', 'hotkeys'
        if section_lc not in sections_to_update_directly: 
            actual_section_in_cfg = next((s for s in cfg_to_save.sections() if s.lower() == section_lc), section_lc)
            if not cfg_to_save.has_section(actual_section_in_cfg) and section_lc in DEFAULT_CONFIG:
                 cfg_to_save.add_section(actual_section_in_cfg) # Add if missing, using original case if found, else lowercase

            for key_lc, default_value_str in DEFAULT_CONFIG[section_lc].items():
                # Only write if this specific option wasn't set by direct calibration
                # This check is a bit redundant given the sections_to_update_directly list,
                # but safe.
                # We need to ensure we don't overwrite already existing values in these other sections
                # unless they were missing entirely.
                
                # Find the key with potentially different casing in the actual config
                actual_key_in_cfg_section = None
                if cfg_to_save.has_section(actual_section_in_cfg):
                    actual_key_in_cfg_section = next((k for k in cfg_to_save.options(actual_section_in_cfg) if k.lower() == key_lc), None)

                if actual_key_in_cfg_section:
                    # Key exists, preserve its value if it's from these "other" sections
                    # This part is tricky: we don't want to overwrite existing values in [General] etc.
                    # The current logic implicitly does this by only explicitly setting calibrated values.
                    # Let's ensure we just don't touch them unless we are adding a missing default.
                    pass # Already handled by reading the file first.
                elif default_value_str is not None : 
                    # If key is missing in config but present in defaults for this "other" section, add it.
                    # This ensures that if we add new default config options later, they get written.
                    safe_set(cfg_to_save, section_lc, key_lc, default_value_str)


    with open(CONFIG_FILE, 'w') as configfile:
        cfg_to_save.write(configfile)
    log_message(f"Calibrated values (and preserved others) saved to {CONFIG_FILE}.")


# --- Helper Functions (get_slot_center_coords, etc.) ---
# (Unchanged from the last "full script" version, they use the global config variables)
def get_slot_center_coords(game_x_abs, game_y_abs, num_rows_to_scan): # ... uses global GRID_OFFSET_X/Y etc.
    centers = []
    base_x = game_x_abs + GRID_OFFSET_X
    base_y = game_y_abs + GRID_OFFSET_Y
    for r in range(num_rows_to_scan):
        for c in range(NUM_COLS):
            centers.append((base_x + c*(SLOT_WIDTH+SLOT_GAP_X) + SLOT_WIDTH//2,
                            base_y + r*(SLOT_HEIGHT+SLOT_GAP_Y) + SLOT_HEIGHT//2))
    return centers

def get_color_patch_coords_for_slot(slot_idx, game_x_abs, game_y_abs): # ... uses global COLOR_PATCH_RELATIVE_X/Y etc.
    r, c = slot_idx // NUM_COLS, slot_idx % NUM_COLS
    slot_tl_x_abs = game_x_abs + GRID_OFFSET_X + c*(SLOT_WIDTH+SLOT_GAP_X)
    slot_tl_y_abs = game_y_abs + GRID_OFFSET_Y + r*(SLOT_HEIGHT+SLOT_GAP_Y)
    patch_center_x_abs = slot_tl_x_abs + COLOR_PATCH_RELATIVE_X
    patch_center_y_abs = slot_tl_y_abs + COLOR_PATCH_RELATIVE_Y
    return patch_center_x_abs, patch_center_y_abs

def get_average_color_from_patch(img, patch_cx_abs, patch_cy_abs, game_x_abs, game_y_abs, slot_idx_str=""): # ... uses COLOR_PATCH_SIZE
    patch_cx_rel, patch_cy_rel = patch_cx_abs - game_x_abs, patch_cy_abs - game_y_abs
    half = COLOR_PATCH_SIZE // 2
    crop_box = (max(0, patch_cx_rel - half), max(0, patch_cy_rel - half),
                min(img.width, patch_cx_rel + half), min(img.height, patch_cy_rel + half))
    if crop_box[0] >= crop_box[2] or crop_box[1] >= crop_box[3]: return None
    patch_img = img.crop(crop_box)
    if patch_img.size[0] == 0 or patch_img.size[1] == 0: return None
    if patch_img.mode != 'RGB': patch_img = patch_img.convert('RGB')
    return tuple(np.mean(np.array(patch_img).reshape(-1, 3), axis=0).astype(int))

def identify_tier_from_color(rgb_tuple, slot_idx_str=""): # ... uses TIER_COLORS, COLOR_TOLERANCE
    if rgb_tuple is None: return None
    best_match, min_d = None, float('inf')
    for tier, t_color in TIER_COLORS.items():
        d = sum(abs(rgb_tuple[i] - t_color[i]) for i in range(3))
        if d < min_d: min_d, best_match = d, tier
    return best_match if min_d <= COLOR_TOLERANCE else None

def get_color_under_mouse_periodic(interval=1): # For TIER_COLOR manual calibration
    log_message("Calibrating TIER_COLORS. Press Ctrl+C in console to stop.") # ... rest of function
    try:
        while True:
            x,y = pyautogui.position(); img = ImageGrab.grab(bbox=(x-2,y-2,x+2,y+2))
            if img.mode != 'RGB': img = img.convert('RGB')
            avg_c = tuple(np.mean(np.array(img).reshape(-1,3), axis=0).astype(int))
            log_message(f"Mouse ({x},{y}) - Avg 5x5 Color: {avg_c}")
            time.sleep(interval)
    except KeyboardInterrupt: log_message("Tier Color calibration stopped.")
    except Exception as e: log_message(f"Tier Color sampling error: {e}")


def get_game_window_rect(): # ... uses GAME_WINDOW_TITLE
    try:
        wins = gw.getWindowsWithTitle(GAME_WINDOW_TITLE)
        if not wins: log_message(f"ERR: Window '{GAME_WINDOW_TITLE}' not found."); return None
        win = wins[0]
        if not win.isActive: log_message(f"WARN: '{GAME_WINDOW_TITLE}' not active.")
        if win.isMinimized: log_message(f"ERR: '{GAME_WINDOW_TITLE}' minimized."); return None
        return (win.left, win.top, win.width, win.height)
    except Exception as e: log_message(f"ERR getting game window: {e}"); return None

def get_stack_count_from_image_region(slot_img_crop, slot_idx_str=""): # ... uses OCR_* globals
    global pytesseract_available, OCR_UPSCALE_FACTOR, OCR_THRESHOLD_VALUE, pytesseract
    if not pytesseract_available: return 1
    try:
        img = slot_img_crop.convert('L')
        w, h = img.size; img = img.resize((w*OCR_UPSCALE_FACTOR, h*OCR_UPSCALE_FACTOR), Image.LANCZOS)
        img = img.point(lambda p: 255 if p > OCR_THRESHOLD_VALUE else 0); img = ImageOps.invert(img)
        img.save(os.path.join(DEBUG_IMAGE_FOLDER, f"Step_OCR_Slot_{slot_idx_str}_Processed.png"))
        cfg = r'--oem 3 --psm 7 -c tessedit_char_whitelist=0123456789'
        txt = pytesseract.image_to_string(img, config=cfg).strip()
        return int(txt) if txt.isdigit() and int(txt) > 0 else 1
    except Exception as e: log_message(f"Slot {slot_idx_str} OCR error: {e}"); return 1

def smooth_drag(sx, sy, ex, ey): # ... uses MouseMovement globals
    log_message(f"Dragging from ({sx},{sy}) to ({ex},{ey})")
    pyautogui.moveTo(sx, sy, duration=MOUSE_MOVE_DURATION, tween=pyautogui.easeInOutQuad); time.sleep(0.05)
    pyautogui.mouseDown(); time.sleep(0.1)
    pyautogui.moveTo(ex, ey, duration=DRAG_DURATION, tween=pyautogui.easeInOutQuad); time.sleep(0.05)
    pyautogui.mouseUp(); log_message(f"Drag done. Pause {POST_ACTION_DELAY}s"); time.sleep(POST_ACTION_DELAY)


# --- MAIN LOGIC (calculate_sort_plan, execute_sort_plan) ---
# These functions need to be complete and use the global config variables.
# calculate_sort_plan needs the TypeError fix for draw.rectangle
def calculate_sort_plan():
    global is_processing, last_calculated_plan
    # Uses global config variables like GRID_OFFSET_X, SLOT_WIDTH, NUM_COLS, etc.

    if is_processing: log_message("Busy."); return
    is_processing = True; log_message("Calculating sort plan..."); last_calculated_plan = None
    log_message(f"Grid Offset: X={GRID_OFFSET_X}, Y={GRID_OFFSET_Y}")
    log_message(f"Grid: {NUM_COLS}x{MAX_NUM_ROWS}(max), Slot:{SLOT_WIDTH}x{SLOT_HEIGHT}, Gap:{SLOT_GAP_X}x{SLOT_GAP_Y}")

    game_rect = get_game_window_rect()
    if not game_rect: is_processing=False; return
    game_x, game_y, game_w, game_h = game_rect # game_w, game_h for screenshot boundary checks

    try:
        win=gw.getWindowsWithTitle(GAME_WINDOW_TITLE)[0]
        if not win.isActive:
            log_message("Activating game window..."); win.activate(); time.sleep(0.5)
            if not win.isActive: log_message("WARN: Failed to activate game window. Ensure it's focused."); # Continue if activation fails but don't block
        screenshot = ImageGrab.grab(bbox=game_rect)
        ensure_debug_folder(); screenshot.save(os.path.join(DEBUG_IMAGE_FOLDER, "Step_0_FullScan_Screenshot.png"))
    except Exception as e: log_message(f"Screenshot error: {e}"); is_processing=False; return

    log_message("Scanning slots..."); scanned_items_initial_state=[] # Stores items with their initial physical slot data
    eff_rows=0
    debug_ss_slots=screenshot.copy(); draw=ImageDraw.Draw(debug_ss_slots)

    # --- Stage 1: Scan the screen and identify all items and their properties ---
    for r in range(MAX_NUM_ROWS):
        row_items_found_this_scan=False
        for c in range(NUM_COLS):
            s_idx = r*NUM_COLS+c
            
            # Slot coordinates relative to the screenshot for drawing
            s_rel_x = GRID_OFFSET_X + c*(SLOT_WIDTH+SLOT_GAP_X) 
            s_rel_y = GRID_OFFSET_Y + r*(SLOT_HEIGHT+SLOT_GAP_Y)
            draw.rectangle([s_rel_x, s_rel_y, s_rel_x+SLOT_WIDTH, s_rel_y+SLOT_HEIGHT], outline="blue", width=1)
            draw.text((s_rel_x+2,s_rel_y+2), str(s_idx), fill="yellow")

            # Absolute screen coordinates for color patch sampling
            pcx_abs, pcy_abs = get_color_patch_coords_for_slot(s_idx, game_x, game_y)
            avg_c = get_average_color_from_patch(screenshot, pcx_abs, pcy_abs, game_x, game_y, str(s_idx))
            tier = identify_tier_from_color(avg_c, str(s_idx))
            
            # Draw color patch sample area (relative to screenshot)
            cp_rel_x = pcx_abs - game_x - COLOR_PATCH_SIZE//2; cp_rel_y = pcy_abs - game_y - COLOR_PATCH_SIZE//2
            draw.rectangle([cp_rel_x, cp_rel_y, cp_rel_x+COLOR_PATCH_SIZE, cp_rel_y+COLOR_PATCH_SIZE], outline="red", width=1)

            if tier is not None:
                row_items_found_this_scan=True
                if r + 1 > eff_rows: eff_rows = r + 1 # Track the max row we've found an item in

                # OCR Region Calculation (absolute screen coordinates, then relative for cropping)
                slot_tl_abs_x = game_x + GRID_OFFSET_X + c*(SLOT_WIDTH+SLOT_GAP_X)
                slot_tl_abs_y = game_y + GRID_OFFSET_Y + r*(SLOT_HEIGHT+SLOT_GAP_Y)
                
                ocr_tl_abs_x = slot_tl_abs_x + OCR_RELATIVE_X
                ocr_tl_abs_y = slot_tl_abs_y + OCR_RELATIVE_Y
                ocr_br_abs_x = ocr_tl_abs_x + OCR_WIDTH
                ocr_br_abs_y = ocr_tl_abs_y + OCR_HEIGHT

                # OCR crop coordinates relative to the screenshot
                ocr_l_rel, ocr_t_rel = ocr_tl_abs_x-game_x, ocr_tl_abs_y-game_y
                ocr_r_rel, ocr_b_rel = ocr_br_abs_x-game_x, ocr_br_abs_y-game_y
                draw.rectangle([ocr_l_rel, ocr_t_rel, ocr_r_rel, ocr_b_rel], outline="lime", width=1)
                
                s_count=1
                if ocr_l_rel < ocr_r_rel and ocr_t_rel < ocr_b_rel and \
                   ocr_r_rel <= screenshot.width and ocr_b_rel <= screenshot.height and \
                   ocr_l_rel >=0 and ocr_t_rel >=0: # Ensure crop is valid
                    sc_crop = screenshot.crop((ocr_l_rel,ocr_t_rel,ocr_r_rel,ocr_b_rel))
                    sc_crop.save(os.path.join(DEBUG_IMAGE_FOLDER,f"Step_OCR_Slot_{s_idx}_Raw.png"))
                    s_count = get_stack_count_from_image_region(sc_crop, str(s_idx))
                else:
                    log_message(f"WARN: Invalid OCR crop coordinates for slot {s_idx}. Defaulting count to 1.")

                # Physical center of this slot on screen
                s_cx_abs, s_cy_abs = slot_tl_abs_x+SLOT_WIDTH//2, slot_tl_abs_y+SLOT_HEIGHT//2
                scanned_items_initial_state.append({
                    'tier':tier, 'count':s_count, 
                    'original_slot_index':s_idx, # This item was found at physical slot s_idx
                    'id': f"item_orig_{s_idx}",    # A unique ID based on original position
                    'current_physical_coords': (s_cx_abs,s_cy_abs) # Its current screen center
                })
                log_message(f"Slot {s_idx}(R{r}C{c}): T{tier},C{s_count},Clr{avg_c}")
        
        if not row_items_found_this_scan and r>=1 and eff_rows > 0 and r >= eff_rows : 
            # If this row is empty, AND we've already found items in a previous row (eff_rows > 0),
            # AND this empty row is at or after the last known effective row.
            log_message(f"Stop scan: Row {r} (0-idx) empty after items found up to row {eff_rows-1}.");break
        if not row_items_found_this_scan and r > 3 and eff_rows == 0 : 
            # If first ~4 rows are completely empty
            log_message(f"Stop scan: Initial {r+1} rows appear empty.");break
    
    debug_ss_slots.save(os.path.join(DEBUG_IMAGE_FOLDER, "Step_1_ScannedSlots_Layout.png"))
    if not scanned_items_initial_state:log_message("No items found.");is_processing=False;return
    log_message(f"Scan done. Max row with items: {eff_rows-1 if eff_rows > 0 else 'None'}. Items found: {len(scanned_items_initial_state)}")

    # --- Stage 2: Determine target order and generate moves ---
    target_sorted_items_by_properties = sorted(
        scanned_items_initial_state, 
        key=lambda x: (x['tier'], -x['count'], x['original_slot_index']) # Tier asc, Count desc, OriginalPos asc
    )
    
    log_message(f"--- Target Sorted Order (Properties of items that should be in these final slots) ---")
    for i,item_props in enumerate(target_sorted_items_by_properties[:10]): # Log first 10
        log_message(f"Target Slot {i} should contain: Item (ID {item_props['id']}) with T{item_props['tier']},C{item_props['count']}")

    # `current_simulated_layout` maps a physical slot index to the item `id` currently in it.
    current_simulated_layout = [None] * (MAX_NUM_ROWS * NUM_COLS)
    # `item_details_map` maps item `id` to its full details (tier, count, id, original_slot_index)
    item_details_map = {}

    for item_data in scanned_items_initial_state:
        current_simulated_layout[item_data['original_slot_index']] = item_data['id']
        item_details_map[item_data['id']] = item_data

    # Get the absolute screen coordinates for all physical slot centers we might use
    # Use eff_rows if known, otherwise MAX_NUM_ROWS for safety, though moves only go up to len(target_items)
    num_rows_for_centers = eff_rows if eff_rows > 0 else MAX_NUM_ROWS 
    all_physical_slot_centers = get_slot_center_coords(game_x, game_y, num_rows_for_centers)

    moves_to_make = []
    # Iterate through each target physical slot (0, 1, 2, ...) up to the number of items we have
    for target_physical_slot_idx in range(len(target_sorted_items_by_properties)):
        # Get the properties of the item that *should* be in this target_physical_slot_idx
        desired_item_id_for_target_slot = target_sorted_items_by_properties[target_physical_slot_idx]['id']

        # Find where this desired_item_id_for_target_slot *currently is* in our simulation
        current_physical_location_of_desired_item = -1
        try:
            current_physical_location_of_desired_item = current_simulated_layout.index(desired_item_id_for_target_slot)
        except ValueError:
            log_message(f"CRIT ERR: Desired item ID {desired_item_id_for_target_slot} not found in simulation. Abort.");is_processing=False;return

        # If the desired item is not already in its correct target physical slot
        if current_physical_location_of_desired_item != target_physical_slot_idx:
            # This item needs to move.
            # Source: current_physical_location_of_desired_item
            # Destination: target_physical_slot_idx

            # Ensure coordinates are valid before adding to move list
            if not (0 <= current_physical_location_of_desired_item < len(all_physical_slot_centers) and \
                    0 <= target_physical_slot_idx < len(all_physical_slot_centers)):
                log_message(f"WARN: Skipping move due to invalid slot index. From: {current_physical_location_of_desired_item}, To: {target_physical_slot_idx}")
                continue

            moves_to_make.append({
                "from_slot_idx": current_physical_location_of_desired_item,
                "to_slot_idx": target_physical_slot_idx,
                "item_id_being_moved": desired_item_id_for_target_slot, # Item that belongs in target_idx
                "from_coords": all_physical_slot_centers[current_physical_location_of_desired_item],
                "to_coords": all_physical_slot_centers[target_physical_slot_idx]
            })
            
            # Simulate the swap in `current_simulated_layout`
            item_id_being_moved = current_simulated_layout[current_physical_location_of_desired_item]
            item_id_displaced = current_simulated_layout[target_physical_slot_idx] # Can be None if target is empty

            current_simulated_layout[target_physical_slot_idx] = item_id_being_moved
            current_simulated_layout[current_physical_location_of_desired_item] = item_id_displaced
            # No need to update 'current_physical_coords' in item_details_map as we use all_physical_slot_centers

    if moves_to_make:
        log_message("--- Calculated Action Plan (0-indexed physical slots) ---")
        for i,m in enumerate(moves_to_make):
            item_props = item_details_map[m['item_id_being_moved']]
            log_message(f"Move {i+1}: Drag item (ID {m['item_id_being_moved']}, T{item_props['tier']}C{item_props['count']}) "
                        f"from current physical_slot {m['from_slot_idx']} to target physical_slot {m['to_slot_idx']}")
        last_calculated_plan={"moves":moves_to_make}
    else:log_message("Inventory already sorted or no moves needed based on scan.")
    log_message("Sort plan calculation finished.");is_processing=False
    
def execute_sort_plan(): # Logic unchanged
    # ... (same as previous full script)
    global is_processing, last_calculated_plan
    if is_processing: log_message("Busy."); return
    if not last_calculated_plan or not last_calculated_plan["moves"]: log_message("No plan. Numpad1 first."); return
    is_processing=True; log_message("Executing sort plan..."); time.sleep(2)
    if not get_game_window_rect(): is_processing=False; log_message("Game window lost."); return
    win = gw.getWindowsWithTitle(GAME_WINDOW_TITLE)[0]
    if not win.isActive: log_message("Game window not active."); is_processing=False; return

    for i, move in enumerate(last_calculated_plan["moves"]):
        sx,sy = move.get("from_coords", (None,None)); ex,ey = move.get("to_coords", (None,None)) # Safer access
        if not (sx and sy and ex and ey): log_message(f"ERR: Bad coords move {i+1}. Skip."); continue
        log_message(f"Move {i+1}: Drag ({sx},{sy}) to ({ex},{ey})")
        smooth_drag(sx,sy,ex,ey)
        if i < len(last_calculated_plan["moves"])-1: log_message("Pause..."); time.sleep(0.2)
    log_message("Execution finished."); last_calculated_plan=None; is_processing=False


def request_exit(): # Unhookall fix applied
    global script_running; log_message("Exit requested by hotkey.")
    try: keyboard.unhook_all(); log_message("Hotkeys unhooked by exit request.")
    except Exception as e: log_message(f"Note: Error unhook_all in request_exit: {e}")
    script_running = False; log_message("Script will now terminate. Close console if needed.")

# --- Helper for Hotkey Registration ---
def register_hotkeys(hotkey_definitions_list): # Logic unchanged
    log_message("--- Registering Hotkeys ---") # ... (rest of function)
    for key_config_name, func_to_call, desc_text in hotkey_definitions_list:
        try:
            hotkey_str = config.get('Hotkeys', key_config_name.lower(), fallback=None) # Read lowercase key
            if hotkey_str:
                keyboard.add_hotkey(hotkey_str, func_to_call)
                log_message(f"  '{hotkey_str}' -> {desc_text}")
            else:
                log_message(f"  WARN: Hotkey for '{key_config_name}' not defined in config.ini [Hotkeys].")
        except Exception as e: log_message(f"  ERR registering '{key_config_name}': {e}.")
    log_message("--- Hotkeys Active ---")


# --- SCRIPT EXECUTION ---
if __name__ == "__main__":
    log_message("Inventory Sorter Script Loading...")
    load_config()       
    initialize_tesseract() 
    ensure_debug_folder()  

    log_message(f"--- Script Configuration Summary ---")
    log_message(f"  Game Window: '{GAME_WINDOW_TITLE}'")
    log_message(f"  Config File: '{CONFIG_FILE}' (Review values if issues persist!)")
    
    # Use lowercase for key_config_name to match DEFAULT_CONFIG and INI convention
    hotkey_actions_list = [
        ('calculateandplan', calculate_sort_plan, "Scan & Plan Sort"),
        ('executesort', execute_sort_plan, "Execute Sort Plan"),
        ('setgridorigin', set_grid_origin_individually, "Indiv: Set Grid Origin (Top-Left of 1st Slot)"),
        ('startfulluicalibration', start_or_advance_full_ui_calibration, "Full UI Calibration Cycle (All Geometry)"),
        ('calibrateslotdimensions', calibrate_slot_dimensions_individually, "Indiv: Set Slot Width/Height (2 clicks)"),
        ('calibrateslotxgap', calibrate_slot_x_gap_individually, "Indiv: Set Slot X-Gap (2 clicks)"),
        ('calibrateslotygap', calibrate_slot_y_gap_individually, "Indiv: Set Slot Y-Gap (2 clicks)"),
        ('calibratetiercolorpoint', calibrate_tier_color_point_individually, "Indiv: Set Tier Color Sample Point (2 clicks)"),
        ('calibrateocrregion', calibrate_ocr_region_individually, "Indiv: Set OCR Stack Count Region (3 clicks)"),
        ('savecalibratedvalues', save_calibrated_values_to_config, "Save ALL Current Calibrated Values to config.ini"),
        ('exitscript', request_exit, "Unhook Keys & Prepare for Exit")
    ]
    register_hotkeys(hotkey_actions_list)
    
    log_message("Console active. Enter 'calibratecolors' (for tier colors) or 'exit' (console input loop).")

    while script_running:
        try:
            if not script_running: break # Moved redundant check earlier
            cmd = input("> ").strip().lower()
            # if not script_running: break # Redundant check removed

            if cmd == 'calibratecolors': get_color_under_mouse_periodic()
            elif cmd == 'exit': log_message("Exiting console loop. Hotkeys still active."); break
            elif cmd: log_message(f"Unknown cmd: '{cmd}'. Use 'calibratecolors' or 'exit'.")
        except EOFError: log_message("EOFError. Non-interactive mode."); break
        except KeyboardInterrupt: log_message("\nCtrl+C: Exiting."); script_running=False
    
    try: keyboard.unhook_all(); log_message("All hotkeys unhooked on final exit.") # Fixed unhook logic
    except Exception as e: log_message(f"Note: Error during final unhook_all: {e}")
    log_message("Script terminated.")