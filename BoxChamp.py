# -*- coding: utf-8 -*-
"""
BoxChamp – Multibox Manager (Windows)

What's new in this cleaned-up build
-----------------------------------
- All UI & code comments are now in **English**
- **Repeater Regions** feature removed entirely (overlay, editor, settings)
- Modern **dark theme** (Fusion) + polished styling
- Simplified, friendlier English UI labels & groups
- Mouse broadcast kept (hold modifier to broadcast clicks simultaneously)
- Borderless game windows; main window can cover the taskbar (real "borderless fullscreen")
- Topmost mode: "always" / "active_only" (default) / "never"
- Per-start dialog: "Run Auto-Login now?"
- Memory Reading for HP/MP stats in Dashboard (requires pymem)
- **Advanced Combat Rotations** with conditional logic engine, per-client assignment,
  and a global toggle hotkey.

Install:
    python -m pip install PySide6 pywin32 keyboard mouse screeninfo psutil pymem

Start:
    python boxMain.py --debug
"""

import os, sys, json, time, threading, subprocess, logging, argparse, faulthandler, traceback, atexit, platform, ctypes, string
from dataclasses import dataclass, field, asdict
from typing import List, Tuple, Dict, Optional, Any, Callable, Literal
import copy

# ---------- Paths next to the executable ------------------------------------
def _get_base_dir() -> str:
    exe_or_file = sys.executable if getattr(sys, "frozen", False) else __file__
    return os.path.dirname(os.path.abspath(exe_or_file))

BASE_DIR = _get_base_dir()
APP_DIR = os.path.join(BASE_DIR, ".boxchamp")
SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")
LOG_PATH = os.path.join(APP_DIR, "boxchamp.log")

def ensure_dirs():
    os.makedirs(APP_DIR, exist_ok=True)

# ---------- Logging ----------------------------------------------------------
def setup_logging(debug: bool):
    ensure_dirs()
    log = logging.getLogger("boxchamp")
    log.setLevel(logging.DEBUG if debug else logging.INFO)
    log.propagate = False
    if getattr(log, "_boxchamp_handlers_set", False):
        return log
    fmt = logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s")
    handlers: List[logging.Handler] = []
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    ch.setLevel(logging.DEBUG if debug else logging.INFO)
    handlers.append(ch)
    try:
        fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
        fh.setFormatter(fmt); fh.setLevel(logging.DEBUG); handlers.append(fh)
    except Exception:
        pass
    for h in handlers: log.addHandler(h)
    log._boxchamp_handlers_set = True
    return log

# ---------- OS prerequisites -------------------------------------------------
if platform.system().lower() != "windows":
    print("BoxChamp is intended for Windows only.")
    sys.exit(1)

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except Exception: pass

# ---------- Dependencies -----------------------------------------------------
try:
    from PySide6 import QtCore, QtWidgets, QtGui
except Exception as e:
    print("PySide6 could not be loaded:", e)
    print("Fix: python -m pip install PySide6")
    sys.exit(1)

try:
    import win32gui, win32con, win32api, win32process
except Exception as e:
    QtWidgets.QMessageBox.critical(None, "Missing dependency",
                                     f"pywin32 is required:\n{e}\n\nFix:\npython -m pip install pywin32")
    sys.exit(1)

try:
    import keyboard, mouse
except Exception as e:
    QtWidgets.QMessageBox.critical(None, "Missing dependency",
                                     f"'keyboard' and 'mouse' modules are required:\n{e}\n\nFix:\npython -m pip install keyboard mouse")
    sys.exit(1)

try:
    from screeninfo import get_monitors
except Exception as e:
    QtWidgets.QMessageBox.critical(None, "Missing dependency",
                                     f"'screeninfo' is required:\n{e}\n\nFix:\npython -m pip install screeninfo")
    sys.exit(1)

try:
    import psutil
except Exception as e:
    QtWidgets.QMessageBox.critical(None, "Missing dependency",
                                     f"'psutil' is required:\n{e}\n\nFix:\npython -m pip install psutil")
    sys.exit(1)

try:
    from pymem import Pymem
    from pymem.process import module_from_name
except Exception as e:
    QtWidgets.QMessageBox.critical(None, "Missing dependency",
                                     f"'pymem' is required for HP/MP reading:\n{e}\n\nFix:\npython -m pip install pymem")
    sys.exit(1)


# ---------- Globals ----------------------------------------------------------
APP_NAME = "BoxChamp"
DEFAULT_TITLES = ["world of warcraft", "wow"]

# ---------- Memory Reading Constants (WoW 3.3.5a 12340) ---------------------
CUR_MGR_POINTER = 0x00C79CE0
CUR_MGR_OFFSET = 0x2ED0
FIRST_OBJECT_OFFSET = 0xAC
NEXT_OBJECT_OFFSET = 0x3C
LOCAL_GUID_OFFSET = 0xC0
OBJECT_GUID_OFFSET = 0x30
M_STORAGE_OFFSET = 0x8
TARGET_GUID_STATIC = 0x00BD07B0
COMBO_POINTS_STATIC = 0x00BD084D

# Descriptor Offsets (multiplied by 4 for dword size)
UNIT_FIELD_HEALTH = 0x18 * 4
UNIT_FIELD_MAXHEALTH = 0x20 * 4
UNIT_FIELD_POWER1 = 0x19 * 4
UNIT_FIELD_MAXPOWER1 = 0x21 * 4
UNIT_FIELD_LEVEL = 0x36 * 4

# Target Name Chain Offsets
UNIT_NAME_1 = 0x964
UNIT_NAME_2 = 0x5C

# ---------- Data Model / Config ---------------------------------------------
STAT_TYPES = Literal[
    "player_hp_percent", "player_mp_percent", "target_hp_percent", "has_target",
    "target_name", "combo_points", "player_level", "target_level",
    "target_is_own_character"
]
OPERATOR_TYPES = Literal["<", ">", "==", "!=", "<=", ">="]
VALID_STATS: List[STAT_TYPES] = [
    "player_hp_percent", "player_mp_percent", "target_hp_percent", "has_target",
    "target_name", "combo_points", "player_level", "target_level",
    "target_is_own_character"
]
VALID_OPERATORS: List[OPERATOR_TYPES] = ["<", ">", "==", "!=", "<=", ">="]

@dataclass
class Condition:
    stat: STAT_TYPES = "player_hp_percent"
    operator: OPERATOR_TYPES = "<"
    value: Any = 50

@dataclass
class RotationRule:
    keys_to_press: List[str] = field(default_factory=list)
    conditions: List[Condition] = field(default_factory=list)

@dataclass
class CombatRotation:
    name: str = "New Rotation"
    hotkey: Optional[str] = None
    rules: List[RotationRule] = field(default_factory=list)
    loop_interval: float = 0.5

@dataclass
class SlotConfig:
    name: str="Profile1"
    args: str=""
    account: str=""
    password: str=""
    character_name: str=""
    realm: Optional[str]=None
    character_steps: int=0
    login_delays: Dict[str,float]=field(default_factory=lambda:{
        "after_start": 8.0, "after_user": 0.2, "after_pass": 0.2,
        "after_enter": 5.0, "after_char": 4.0
    })
    cpu_affinity: List[int]=field(default_factory=list)
    priority: str="below_normal"
    assigned_rotation: Optional[str] = None

@dataclass
class CharacterSet:
    name: str="Team1"
    exe_path: str = r"C:\Games\WoW\wow.exe"
    workdir: str = r"C:\Games\WoW"
    slots: List[str]=field(default_factory=lambda:["Profile1"])
    auto_login: bool=True
    start_interval: float=2.0
    stop_grace: float=3.0
    layout_mode: str = "main_left_slaves_right"
    grid_cols: int=1
    grid_rows: int=4

@dataclass
class MacroStep:
    type: str="key"
    value: str=""
    delay: float=0.0
    mouse: Dict[str,Any]=field(default_factory=dict)

@dataclass
class Macro:
    name: str=""
    hotkey: Optional[str]=None
    target: str="all"
    slots: List[str]=field(default_factory=list)
    steps: List[MacroStep]=field(default_factory=list)
    rr_index: int=0
    loop: bool=False
    loop_interval: float=0.5
    loop_count: Optional[int]=None

@dataclass
class ClickBarButton:
    label: str
    macro: str

@dataclass
class ClickBar:
    enabled: bool=False
    x: int=100
    y: int=100
    width: int=200
    height: int=40
    buttons: List[ClickBarButton]=field(default_factory=lambda:[ClickBarButton("Burst","Burst")])

@dataclass
class KeyMap:
    whitelist: List[str]=field(default_factory=lambda:list("123456789"))
    mouse_hold: str="ctrl"
    toggle_broadcast_hotkey: str="ctrl+alt+b"
    swap_hotkey: str="ctrl+alt+s"
    cycle_main_hotkey: Optional[str]=None
    toggle_rotations_hotkey: str = "f12"
    broadcast_only_when_client_focused: bool=True
    mouse_broadcast_enabled: bool=True
    auto_assist_enabled: bool=False
    auto_assist_prefix_keys: List[str] = field(default_factory=lambda: ["f2", "h"])

@dataclass
class GroupTargetingConfig:
    enabled: bool = True
    target_master_hotkey: str = "alt+f1"
    target_slave_hotkeys: List[str] = field(default_factory=lambda: ["f2", "f3", "f4", "f5"])
    key_for_self: str = "f1"
    key_for_master: str = "f2"

@dataclass
class Settings:
    window_title_filters: List[str]=field(default_factory=lambda:DEFAULT_TITLES)
    game_executable_names: List[str] = field(default_factory=lambda: ["Wow.exe"])
    window_rename_scheme: str = "By Slot Name"
    tile_padding: int=0
    broadcast_enabled: bool=True
    broadcast_all_keys: bool=False
    clickbar: ClickBar=field(default_factory=ClickBar)
    keymap: KeyMap=field(default_factory=KeyMap)
    group_targeting: GroupTargetingConfig=field(default_factory=GroupTargetingConfig)
    borderless: bool=True
    main_fullscreen_taskbar: bool=True
    main_topmost_mode: str="active_only"

@dataclass
class AppConfig:
    slots: Dict[str,SlotConfig]=field(default_factory=lambda:{ "Profile1": SlotConfig(name="Profile1") })
    sets: List[CharacterSet]=field(default_factory=lambda:[ CharacterSet(name="Team1", slots=["Profile1"]) ])
    macros: Dict[str,Macro]=field(default_factory=lambda:{
        "FollowMain": Macro(name="FollowMain", hotkey="f6", target="all_except_main",
                            steps=[MacroStep(type="key", value="f11")]),
        "Burst": Macro(name="Burst", hotkey="f7", target="all",
                       steps=[MacroStep(type="key", value="1"),
                              MacroStep(type="delay", delay=0.10),
                              MacroStep(type="key", value="2"),
                              MacroStep(type="delay", delay=0.10),
                              MacroStep(type="key", value="3")]),
    })
    rotations: Dict[str, CombatRotation] = field(default_factory=dict)
    settings: Settings=field(default_factory=Settings)

# ---------- Config IO --------------------------------------------------------
def load_config() -> AppConfig:
    ensure_dirs()
    if not os.path.exists(SETTINGS_PATH):
        cfg=AppConfig()
        with open(SETTINGS_PATH,"w",encoding="utf-8") as f:
            json.dump(asdict(cfg), f, ensure_ascii=False, indent=2)
        return cfg
    with open(SETTINGS_PATH,"r",encoding="utf-8") as f:
        data=json.load(f)

    def to_dc(cls, d):
        if isinstance(d, dict):
            fieldtypes={f.name:f.type for f in cls.__dataclass_fields__.values()}
            kwargs={}
            valid_keys = {f.name for f in cls.__dataclass_fields__.values()}
            for k,t in fieldtypes.items():
                if k not in d: continue
                v=d[k]
                if hasattr(t,"__dataclass_fields__"):
                    kwargs[k]=to_dc(t,v)
                elif getattr(t,"__origin__",None) is list and hasattr(t.__args__[0], "__dataclass_fields__"):
                    kwargs[k]=[to_dc(t.__args__[0], i) for i in v]
                elif getattr(t,"__origin__",None) is dict:
                    vt=t.__args__[1]
                    if hasattr(vt,"__dataclass_fields__"):
                        kwargs[k]={kk:to_dc(vt,vv) for kk,vv in v.items()}
                    else:
                        kwargs[k]=v
                else:
                    kwargs[k]=v
            
            final_kwargs = {k: v for k, v in kwargs.items() if k in valid_keys}
            return cls(**final_kwargs)
        return d

    slots={k:to_dc(SlotConfig,v) for k,v in data.get("slots",{}).items()}
    sets=[to_dc(CharacterSet,s) for s in data.get("sets",[])]
    macs={}
    for name,m in data.get("macros",{}).items():
        steps=[to_dc(MacroStep, st) for st in m.get("steps",[])]
        macs[name]=Macro(name=name, hotkey=m.get("hotkey"), target=m.get("target","all"),
                             slots=m.get("slots",[]), steps=steps,
                             rr_index=int(m.get("rr_index",0)),
                             loop=bool(m.get("loop", False)),
                             loop_interval=float(m.get("loop_interval",0.5)),
                             loop_count=m.get("loop_count", None))
    rots = {}
    for name, r_data in data.get("rotations", {}).items():
        rules = []
        for rule_data in r_data.get("rules", []):
            # Compatibility for old key_to_press
            if "key_to_press" in rule_data and "keys_to_press" not in rule_data:
                rule_data["keys_to_press"] = [rule_data["key_to_press"]]
            rules.append(to_dc(RotationRule, rule_data))
        rots[name] = CombatRotation(
            name=name,
            hotkey=r_data.get("hotkey"), # Hotkey pro Rotation wird aktuell nicht verwendet
            rules=rules,
            loop_interval=float(r_data.get("loop_interval", 0.5))
        )
    settings=to_dc(Settings, data.get("settings",{}))
    return AppConfig(slots=slots, sets=sets, macros=macs, rotations=rots, settings=settings)

def save_config(cfg:AppConfig):
    ensure_dirs()
    with open(SETTINGS_PATH,"w",encoding="utf-8") as f:
        json.dump(asdict(cfg), f, ensure_ascii=False, indent=2)

# ---------- Memory Reader Class ----------------------------------------------
class MemoryReader:
    def __init__(self, pid: int, log: logging.Logger):
        self.log = log
        self.pm = Pymem(pid)
        self.object_manager_base = 0

    def _get_object_manager_base(self):
        if self.object_manager_base:
            return self.object_manager_base
        try:
            client_connection = self.pm.read_uint(CUR_MGR_POINTER)
            if not client_connection: return 0
            self.object_manager_base = self.pm.read_uint(client_connection + CUR_MGR_OFFSET)
            return self.object_manager_base
        except Exception:
            self.object_manager_base = 0
            return 0
            
    def _find_object_base_by_guid(self, guid_to_find: int) -> int:
        obj_mgr = self._get_object_manager_base()
        if not obj_mgr or not guid_to_find:
            return 0
        
        current_obj = self.pm.read_uint(obj_mgr + FIRST_OBJECT_OFFSET)
        for _ in range(4096):
            if not current_obj: break
            try:
                obj_guid = self.pm.read_ulonglong(current_obj + OBJECT_GUID_OFFSET)
                if obj_guid == guid_to_find:
                    return current_obj
                current_obj = self.pm.read_uint(current_obj + NEXT_OBJECT_OFFSET)
            except Exception:
                break
        return 0

    def get_combat_stats(self) -> Optional[Dict[str, Any]]:
        stats = {
            'player_hp': 0, 'player_max_hp': 0, 'player_mp': 0, 'player_max_mp': 0,
            'player_hp_percent': 0, 'player_mp_percent': 0, 'player_level': 0,
            'target_hp': 0, 'target_max_hp': 0, 'target_hp_percent': 0, 'target_level': 0,
            'has_target': False, 'target_name': '', 'combo_points': 0
        }
        try:
            obj_mgr = self._get_object_manager_base()
            if not obj_mgr: return None
            
            stats['combo_points'] = self.pm.read_uchar(COMBO_POINTS_STATIC)
            
            player_guid = self.pm.read_ulonglong(obj_mgr + LOCAL_GUID_OFFSET)
            if not player_guid: return None
            
            player_base = self._find_object_base_by_guid(player_guid)
            if not player_base: return None

            descriptor_ptr = self.pm.read_uint(player_base + M_STORAGE_OFFSET)
            if not descriptor_ptr: return None

            stats['player_hp'] = self.pm.read_int(descriptor_ptr + UNIT_FIELD_HEALTH)
            stats['player_max_hp'] = self.pm.read_int(descriptor_ptr + UNIT_FIELD_MAXHEALTH)
            stats['player_mp'] = self.pm.read_int(descriptor_ptr + UNIT_FIELD_POWER1)
            stats['player_max_mp'] = self.pm.read_int(descriptor_ptr + UNIT_FIELD_MAXPOWER1)
            stats['player_level'] = self.pm.read_int(descriptor_ptr + UNIT_FIELD_LEVEL)
            
            if stats['player_max_hp'] > 0:
                stats['player_hp_percent'] = round((stats['player_hp'] / stats['player_max_hp']) * 100)
            if stats['player_max_mp'] > 0:
                stats['player_mp_percent'] = round((stats['player_mp'] / stats['player_max_mp']) * 100)
            
            target_guid = self.pm.read_ulonglong(TARGET_GUID_STATIC)
            if target_guid > 0:
                stats['has_target'] = True
                target_base = self._find_object_base_by_guid(target_guid)
                if target_base:
                    target_desc_ptr = self.pm.read_uint(target_base + M_STORAGE_OFFSET)
                    if target_desc_ptr:
                        stats['target_hp'] = self.pm.read_int(target_desc_ptr + UNIT_FIELD_HEALTH)
                        stats['target_max_hp'] = self.pm.read_int(target_desc_ptr + UNIT_FIELD_MAXHEALTH)
                        stats['target_level'] = self.pm.read_int(target_desc_ptr + UNIT_FIELD_LEVEL)
                        if stats['target_max_hp'] > 0:
                           stats['target_hp_percent'] = round((stats['target_hp'] / stats['target_max_hp']) * 100)
                    try:
                        name_ptr_1 = self.pm.read_uint(target_base + UNIT_NAME_1)
                        if name_ptr_1:
                            name_ptr_2 = self.pm.read_uint(name_ptr_1 + UNIT_NAME_2)
                            if name_ptr_2:
                                stats['target_name'] = self.pm.read_string(name_ptr_2)
                    except Exception:
                        stats['target_name'] = ''
            else:
                stats['has_target'] = False

            return stats
        except Exception as e:
            self.log.debug(f"get_combat_stats failed for PID {self.pm.process_id}: {e}")
            return None

# ---------- Win32 helpers ----------------------------------------------------
def enum_windows() -> List[int]:
    hwnds=[]
    def _enum(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            hwnds.append(hwnd)
    win32gui.EnumWindows(_enum, None)
    return hwnds

def hwnd_pid(hwnd:int) -> int:
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return pid

def enum_windows_by_title(title_filters: List[str], exe_names: List[str]) -> List[int]:
    hwnds=[]; title_lowers=[s.lower() for s in title_filters]
    exe_lowers = [name.lower() for name in exe_names]

    def _enum(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd): return
        
        title = (win32gui.GetWindowText(hwnd) or "").lower()
        if not any(s in title for s in title_lowers):
            return

        try:
            pid = hwnd_pid(hwnd)
            if not pid: return
            p = psutil.Process(pid)
            if p.name().lower() not in exe_lowers:
                return
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            return

        style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        if style & win32con.WS_EX_TOOLWINDOW: return
        hwnds.append(hwnd)

    win32gui.EnumWindows(_enum, None)
    unique = sorted(set(hwnds), key=lambda h: ((win32gui.GetWindowText(h) or "").lower(), h))
    return unique

def windows_for_pid(pid:int) -> List[int]:
    return [h for h in enum_windows() if hwnd_pid(h)==pid and win32gui.IsWindowVisible(h)]

def bring_to_front(hwnd:int):
    try:
        flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, flags)
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, flags)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


def set_window_pos(hwnd:int, x:int,y:int,w:int,h:int, topmost=False):
    flags = win32con.SWP_NOACTIVATE | win32con.SWP_SHOWWINDOW
    insert_after = win32con.HWND_TOPMOST if topmost else win32con.HWND_NOTOPMOST
    try:
        win32gui.SetWindowPos(hwnd, insert_after, x,y,w,h, flags)
    except Exception:
        pass

def client_rect(hwnd:int) -> Tuple[int,int,int,int]:
    r = win32gui.GetClientRect(hwnd)
    pt= win32gui.ClientToScreen(hwnd, (0,0))
    return pt[0], pt[1], r[2]-r[0], r[3]-r[0]

def screen_rect(hwnd:int) -> Tuple[int,int,int,int]:
    r = win32gui.GetWindowRect(hwnd)
    return r[0], r[1], r[2]-r[0], r[3]-r[0]

def _mk_lparam(x:int,y:int)->int:
    return (y<<16) | (x & 0xFFFF)

# ---------- Layout engine ----------------------------------------------------
class ClientLayout:
    def __init__(self, padding: int):
        self.pad = padding

    def monitors(self):
        mons = get_monitors()
        mons.sort(key=lambda m: (m.x, m.y))
        return mons

    def positions(self, hwnds: List[int], mode: str, cols: int, rows: int) -> Dict[int, Tuple[int, int, int, int]]:
        pos = {}
        if not hwnds: return pos
        mons = self.monitors()
        main_hwnd = hwnds[0]
        slaves = hwnds[1:]

        # Determine main and slave monitor rectangles
        if mode == "main_on_monitor_1" and len(mons) > 1:
            main_mon, slave_mon = mons[0], mons[1]
        elif mode == "main_on_monitor_2" and len(mons) > 1:
            main_mon, slave_mon = mons[1], mons[0]
        else:
            main_mon, slave_mon = mons[0], mons[0]
        
        # Assign main window position (always flush to its monitor edges)
        pos[main_hwnd] = (main_mon.x, main_mon.y, main_mon.width, main_mon.height)

        # Determine the region for slave windows
        slave_rect = (slave_mon.x, slave_mon.y, slave_mon.width, slave_mon.height)
        if main_mon == slave_mon: # If on the same monitor, split the space
            main_width = int(main_mon.width * 0.6)
            if mode == "main_left_slaves_right":
                pos[main_hwnd] = (main_mon.x, main_mon.y, main_width, main_mon.height)
                slave_rect = (main_mon.x + main_width, main_mon.y, main_mon.width - main_width, main_mon.height)
            elif mode == "main_right_slaves_left":
                slave_width = main_mon.width - main_width
                slave_rect = (main_mon.x, main_mon.y, slave_width, main_mon.height)
                pos[main_hwnd] = (main_mon.x + slave_width, main_mon.y, main_width, main_mon.height)

        if not slaves: return pos

        # Tile slaves within the slave_rect
        sx, sy, sw, sh = slave_rect
        cols, rows = max(1, cols), max(1, rows)
        
        total_pad_x = (cols - 1) * self.pad
        total_pad_y = (rows - 1) * self.pad
        
        cell_w = (sw - total_pad_x) / cols
        cell_h = (sh - total_pad_y) / rows
        
        idx = 0
        for r in range(rows):
            for c in range(cols):
                if idx >= len(slaves): break
                x = sx + c * (cell_w + self.pad)
                y = sy + r * (cell_h + self.pad)
                pos[slaves[idx]] = (round(x), round(y), round(cell_w), round(cell_h))
                idx += 1
        return pos

# ---------- Virtual Key Code Map ---------------------------------------------
VK_CODE_MAP = {
    'backspace': 0x08, 'tab': 0x09, 'enter': 0x0D, 'shift': 0x10, 'ctrl': 0x11,
    'alt': 0x12, 'pause': 0x13, 'caps_lock': 0x14, 'esc': 0x1B, 'space': 0x20,
    'page_up': 0x21, 'page_down': 0x22, 'end': 0x23, 'home': 0x24,
    'left': 0x25, 'up': 0x26, 'right': 0x27, 'down': 0x28, 'print_screen': 0x2C,
    'insert': 0x2D, 'delete': 0x2E,
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34, '5': 0x35, '6': 0x36,
    '7': 0x37, '8': 0x38, '9': 0x39,
    'a': 0x41, 'b': 0x42, 'c': 0x43, 'd': 0x44, 'e': 0x45, 'f': 0x46, 'g': 0x47,
    'h': 0x48, 'i': 0x49, 'j': 0x4A, 'k': 0x4B, 'l': 0x4C, 'm': 0x4D, 'n': 0x4E,
    'o': 0x4F, 'p': 0x50, 'q': 0x51, 'r': 0x52, 's': 0x53, 't': 0x54, 'u': 0x55,
    'v': 0x56, 'w': 0x57, 'x': 0x58, 'y': 0x59, 'z': 0x5A,
    'f1': 0x70, 'f2': 0x71, 'f3': 0x72, 'f4': 0x73, 'f5': 0x74, 'f6': 0x75,
    'f7': 0x76, 'f8': 0x77, 'f9': 0x78, 'f10': 0x79, 'f11': 0x7A, 'f12': 0x7B,
    'numpad_0': 0x60, 'numpad_1': 0x61, 'numpad_2': 0x62, 'numpad_3': 0x63,
    'numpad_4': 0x64, 'numpad_5': 0x65, 'numpad_6': 0x66, 'numpad_7': 0x67,
    'numpad_8': 0x68, 'numpad_9': 0x69,
    'multiply': 0x6A, 'add': 0x6B, 'separator': 0x6C, 'subtract': 0x6D,
    'decimal': 0x6E, 'divide': 0x6F
}


# ---------- Controller -------------------------------------------------------
class BoxChampController(QtCore.QObject):
    status_changed = QtCore.Signal(str)
    clients_changed = QtCore.Signal(list)
    set_state_changed = QtCore.Signal(str)
    broadcast_state_changed = QtCore.Signal()
    launch_error = QtCore.Signal(str, str)

    def __init__(self, cfg:AppConfig, log:logging.Logger):
        super().__init__()
        self.cfg=cfg
        self.log=log
        self.hwnds: List[int]=[]
        self.main_hwnd: Optional[int]=None
        self.processes: Dict[str, psutil.Process]={}
        self.proc_hwnd: Dict[str,int]={}
        self.memory_readers: Dict[int, MemoryReader] = {}
        self.running_set: Optional[CharacterSet]=None
        self._broadcast_lock=threading.RLock()
        self._reserved_hotkeys:set[str]=set()
        self._rr_state: Dict[str,int]={}
        self._kb_hook: Optional[Callable]=None
        self._mouse_hook: Optional[Callable]=None
        self._hotkey_handles: List[tuple[str,Callable]]=[]
        self._macro_threads: "set[threading.Thread]" = set()
        self._mouse_hotkeys: List[Tuple[str,set,Callable]] = []
        self._loop_stop: Dict[str, threading.Event] = {}
        self._loop_threads: Dict[str, threading.Thread] = {}
        self._running_rotations: Dict[str, threading.Event] = {}
        self._main_is_topmost: Optional[bool]=None
        self._broadcast_keys_down = set()
        self._topmost_timer = QtCore.QTimer(self)
        self._topmost_timer.setInterval(400)
        self._topmost_timer.timeout.connect(self._enforce_topmost)
        self._topmost_timer.start()
        self._rotation_engine_thread: Optional[threading.Thread] = None
        
        self._health_check_timer = QtCore.QTimer(self)
        self._health_check_timer.setInterval(20000) # Increased to 20 seconds
        self._health_check_timer.timeout.connect(self._check_running_set_health)
        self._health_check_timer.start()

    def _perform_soft_stop(self):
        """Cleans up the internal state of a running set without killing processes."""
        self.processes.clear()
        self.proc_hwnd.clear()
        self.memory_readers.clear()
        self.running_set = None
        self.set_state_changed.emit("stopped")
        self.refresh_clients()

    def _check_running_set_health(self):
        if not self.running_set:
            return

        # This check is for when the launch fails completely
        if not self.processes:
            self.log.warning(f"Resetting stale state for set '{self.running_set.name}' (no processes tracked).")
            self._perform_soft_stop()
            return

        # This check is for when the processes were launched but have since closed
        any_running = False
        for proc in self.processes.values():
            try:
                if proc.is_running():
                    any_running = True
                    break
            except psutil.NoSuchProcess:
                continue

        if not any_running:
            self.log.info(f"Detected that all processes for set '{self.running_set.name}' have been closed. Resetting state.")
            self.status_changed.emit("Set stopped (windows closed).")
            self._perform_soft_stop()

    def refresh_clients(self):
        managed_hwnds = set()
        sorted_managed_hwnds = []
        
        if self.running_set:
            for slot_name in self.running_set.slots:
                hwnd = self.proc_hwnd.get(slot_name)
                if hwnd and win32gui.IsWindow(hwnd):
                    if hwnd not in managed_hwnds:
                        sorted_managed_hwnds.append(hwnd)
                        managed_hwnds.add(hwnd)

        title_filters = self.cfg.settings.window_title_filters
        valid_exes = self.cfg.settings.game_executable_names
        unmanaged_hwnds = [h for h in enum_windows_by_title(title_filters, valid_exes) if h not in managed_hwnds]
        hwnds = sorted_managed_hwnds + unmanaged_hwnds

        if self.main_hwnd in hwnds:
            hwnds.remove(self.main_hwnd)
            hwnds.insert(0, self.main_hwnd)

        self.hwnds = hwnds
        
        if hwnds and (self.main_hwnd is None or self.main_hwnd not in hwnds):
            self.main_hwnd = hwnds[0]
        elif not hwnds:
            self.main_hwnd = None

        items = []
        for h in hwnds:
            try:
                if win32gui.IsWindow(h):
                    items.append((h, win32gui.GetWindowText(h)))
            except Exception:
                pass
        
        self.clients_changed.emit(items)
        self.log.debug(f"Refreshed clients: {[(hex(h), t) for h, t in items]}")

    def _rename_window(self, slot: SlotConfig, hwnd: int):
        scheme = self.cfg.settings.window_rename_scheme
        new_title = ""
        if scheme == "By Slot Name":
            new_title = slot.name
        elif scheme == "By Account Name":
            new_title = slot.account
        elif scheme == "By Character Name":
            new_title = slot.character_name
        
        if new_title:
            for i in range(5):
                try:
                    win32gui.SetWindowText(hwnd, new_title)
                    current_title = win32gui.GetWindowText(hwnd)
                    if current_title == new_title:
                        self.log.info(f"Renamed window for slot {slot.name} to '{new_title}'")
                        return
                    time.sleep(0.2)
                except Exception as e:
                    self.log.error(f"Attempt {i+1} to rename window for slot {slot.name} failed: {e}")
                    time.sleep(0.2)
            self.log.warning(f"Failed to rename window for slot {slot.name} to '{new_title}' after several attempts.")

    def _make_borderless(self, hwnd:int):
        try:
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            style &= ~(win32con.WS_CAPTION | win32con.WS_THICKFRAME | win32con.WS_MINIMIZEBOX | win32con.WS_MAXIMIZEBOX | win32con.WS_SYSMENU)
            style |= win32con.WS_POPUP
            win32gui.SetWindowLong(hwnd, win32con.GWL_STYLE, style)
            ex = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            ex &= ~(win32con.WS_EX_DLGMODALFRAME | win32con.WS_EX_CLIENTEDGE | win32con.WS_EX_STATICEDGE)
            win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, ex)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0,0,0,0,
                                  win32con.SWP_NOMOVE|win32con.SWP_NOSIZE|win32con.SWP_FRAMECHANGED|win32con.SWP_NOOWNERZORDER|win32con.SWP_SHOWWINDOW)
        except Exception as e:
            self.log.warning(f"Borderless failed: {e}")

    def _fullscreen_monitor(self, hwnd:int):
        try:
            hmon = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
            info = win32api.GetMonitorInfo(hmon)
            l,t,r,b = info.get("Monitor", info.get("Work", (0,0,0,0)))
            w = max(1, r - l); h = max(1, b - t)
            if self.cfg.settings.borderless: self._make_borderless(hwnd)
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, l, t, w, h,
                                  win32con.SWP_FRAMECHANGED|win32con.SWP_NOOWNERZORDER|win32con.SWP_SHOWWINDOW)
        except Exception as e:
            self.log.warning(f"Fullscreen failed: {e}")

    def _set_topmost_state(self, hwnd:int, top:bool):
        try:
            if top:
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0,0,0,0,
                                      win32con.SWP_NOMOVE|win32con.SWP_NOSIZE|win32con.SWP_NOACTIVATE)
            else:
                win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0,0,0,0,
                                      win32con.SWP_NOMOVE|win32con.SWP_NOSIZE|win32con.SWP_NOACTIVATE)
            self._main_is_topmost=top
        except Exception: pass

    def _enforce_topmost(self):
        S=self.cfg.settings
        if not (S.main_fullscreen_taskbar and self.hwnds):
            return
        main=self.hwnds[0]
        mode=S.main_topmost_mode.lower()
        fg = win32gui.GetForegroundWindow()
        desired = True if mode=="always" else False if mode=="never" else (fg in self.hwnds or fg==main)
        if self._main_is_topmost is None or desired != self._main_is_topmost:
            self._set_topmost_state(main, desired)

    def apply_layout(self):
        if not self.hwnds:
            self.status_changed.emit("No WoW clients found."); return
        S=self.cfg.settings
        cs=self.running_set
        if not cs:
            self.status_changed.emit("No set is running. Select a set to apply its layout."); return
        
        pos = ClientLayout(S.tile_padding).positions(self.hwnds, cs.layout_mode, cs.grid_cols, cs.grid_rows)
        for hwnd,(x,y,w,h) in pos.items():
            if S.borderless: self._make_borderless(hwnd)
            set_window_pos(hwnd, int(x), int(y), int(w), int(h), topmost=False)

        self._enforce_topmost()
        self.status_changed.emit("Layout applied.")
        self.log.info("Layout applied.")

    def make_main(self, hwnd:int):
        if hwnd not in self.hwnds: return
        self.hwnds.remove(hwnd); self.hwnds.insert(0,hwnd); self.main_hwnd=hwnd
        self.apply_layout(); bring_to_front(hwnd); self.refresh_clients()

    def cycle_main(self):
        if len(self.hwnds)<=1: return
        self.make_main(self.hwnds[1])

    def start_set(self, cs:CharacterSet, auto_login_override: Optional[bool]=None):
        if self.running_set:
            self.status_changed.emit("A set is already running."); return
        self.set_state_changed.emit("starting")
        threading.Thread(target=self._thread_start_set, args=(cs, auto_login_override), daemon=True, name="BC-StartSet").start()

    def _thread_start_set(self, cs:CharacterSet, auto_login_override: Optional[bool]):
        try:
            self.processes.clear(); self.proc_hwnd.clear(); self.memory_readers.clear()
            for slot_name in sorted(list(set(cs.slots))):
                slot=self.cfg.slots.get(slot_name)
                if not slot: continue
                try:
                    if not os.path.isfile(cs.exe_path):
                        self.log.error(f"Set '{cs.name}': exe_path does not exist: {cs.exe_path}")
                        continue
                    if cs.workdir and not os.path.isdir(cs.workdir):
                        self.log.warning(f"Set '{cs.name}': workdir does not exist: {cs.workdir}")
                    args=[cs.exe_path]+([a for a in slot.args.split(" ") if a] if slot.args else [])
                    proc = subprocess.Popen(args, cwd=cs.workdir or None)
                    self.processes[slot_name]=psutil.Process(proc.pid)
                    self._apply_proc_tuning(slot_name, slot)
                    hwnd=self._wait_for_window(proc.pid, timeout=max(20.0, slot.login_delays.get("after_start",8.0)+10))
                    if hwnd:
                        self.proc_hwnd[slot_name]=hwnd
                        self._rename_window(slot, hwnd)
                        try:
                            self.memory_readers[hwnd] = MemoryReader(proc.pid, self.log)
                            self.log.info(f"MemoryReader attached to PID {proc.pid} for slot {slot.name}")
                        except Exception as e:
                            self.log.error(f"Failed to attach MemoryReader to PID {proc.pid}: {e}")
                            
                    self.log.info(f"{slot.name} started (PID {proc.pid})")
                except Exception:
                    error_message = f"Failed to start the process for slot '{slot.name}'.\n\n"
                    error_message += f"Path: {cs.exe_path}\n"
                    error_message += f"Workdir: {cs.workdir}\n\n"
                    error_message += f"Error:\n{traceback.format_exc()}"
                    self.log.error(error_message)
                    self.launch_error.emit(slot.name, error_message)
                time.sleep(max(0.2, cs.start_interval))

            if not self.processes:
                self.log.error(f"Failed to start any clients for set '{cs.name}'. Aborting.")
                self.status_changed.emit("Failed to start any clients.")
                self.set_state_changed.emit("stopped")
                return
            
            self.running_set = cs

            self.refresh_clients(); self.apply_layout()
            if self.hwnds: self.main_hwnd=self.hwnds[0]
            do_login = auto_login_override if auto_login_override is not None else cs.auto_login
            if do_login:
                for slot_name in cs.slots:
                    slot=self.cfg.slots.get(slot_name); hwnd=self.proc_hwnd.get(slot_name)
                    if slot and hwnd: self._auto_login(hwnd, slot)
            self.set_state_changed.emit("running")
            self.status_changed.emit("Set started.")
        except Exception:
            self.log.exception("Start error")
            self.set_state_changed.emit("stopped")
            self.running_set = None

    def stop_set(self):
        if not self.running_set:
            self.status_changed.emit("No active set."); return
        self.stop_all_rotations()
        cs=self.running_set
        self.set_state_changed.emit("stopping")
        threading.Thread(target=self._thread_stop_set, args=(cs,), daemon=True, name="BC-StopSet").start()

    def _thread_stop_set(self, cs:CharacterSet):
        try:
            for slot, proc in list(self.processes.items()):
                try: proc.terminate()
                except Exception: pass
            time.sleep(max(0.5, cs.stop_grace))
            for slot, proc in list(self.processes.items()):
                try:
                    if proc.is_running():
                        proc.kill()
                except Exception: pass
            self.processes.clear(); self.proc_hwnd.clear(); self.memory_readers.clear()
            self.status_changed.emit("Set stopped.")
        finally:
            self.running_set=None
            self.set_state_changed.emit("stopped")
            self.refresh_clients()

    def _wait_for_window(self, pid:int, timeout:float=25.0) -> Optional[int]:
        t0=time.time()
        while time.time()-t0<timeout:
            wins=windows_for_pid(pid)
            if wins:
                wins.sort(key=lambda h: screen_rect(h)[2]*screen_rect(h)[3], reverse=True)
                return wins[0]
            time.sleep(0.3)
        return None

    def _apply_proc_tuning(self, slot_name:str, slot:SlotConfig):
        try:
            p=self.processes.get(slot_name)
            if not p: return
            if slot.cpu_affinity:
                try:
                    max_cpu = psutil.cpu_count(logical=True) or 0
                    aff = [c for c in slot.cpu_affinity if 0 <= c < max_cpu]
                    if aff: p.cpu_affinity(aff)
                except Exception: pass
            pr_map={
                "idle": psutil.IDLE_PRIORITY_CLASS,
                "below_normal": psutil.BELOW_NORMAL_PRIORITY_CLASS,
                "normal": psutil.NORMAL_PRIORITY_CLASS,
                "above_normal": psutil.ABOVE_NORMAL_PRIORITY_CLASS,
                "high": psutil.HIGH_PRIORITY_CLASS
            }
            p.nice(pr_map.get((slot.priority or "below_normal").lower(), psutil.BELOW_NORMAL_PRIORITY_CLASS))
        except Exception:
            pass

    def _type_text(self, text:str): keyboard.write(text, delay=0)
    def _press(self, key:str): keyboard.send(key)

    def _auto_login(self, hwnd:int, slot:SlotConfig):
        d=slot.login_delays
        bring_to_front(hwnd)
        time.sleep(max(0.0, d.get("after_start",8.0)))
        if slot.account:
            self._type_text(slot.account); time.sleep(d.get("after_user",0.2))
        self._press("tab"); time.sleep(0.1)
        if slot.password:
            self._type_text(slot.password); time.sleep(d.get("after_pass",0.2))
        self._press("enter"); time.sleep(d.get("after_enter",5.0))
        for _ in range(max(0,int(slot.character_steps))):
            self._press("down"); time.sleep(0.05)
        self._press("enter"); time.sleep(d.get("after_char",4.0))

    def toggle_broadcast(self):
        s=self.cfg.settings
        s.broadcast_enabled=not s.broadcast_enabled
        self.status_changed.emit(f"Keyboard broadcast: {'ON' if s.broadcast_enabled else 'OFF'}")
        self.broadcast_state_changed.emit()

    def toggle_broadcast_all(self):
        s=self.cfg.settings
        s.broadcast_all_keys=not s.broadcast_all_keys
        self.status_changed.emit(f"Broadcast ALL keys: {'ON' if s.broadcast_all_keys else 'OFF'}")
        self.broadcast_state_changed.emit()

    def toggle_mouse_broadcast(self):
        s=self.cfg.settings.keymap
        s.mouse_broadcast_enabled=not s.mouse_broadcast_enabled
        self.status_changed.emit(f"Mouse broadcast: {'ON' if s.mouse_broadcast_enabled else 'OFF'}")

    def _post_key_event(self, hwnd: int, key: str, event_type: str):
        key_lower = key.lower()
        vk_code = VK_CODE_MAP.get(key_lower)
        if vk_code is None:
            self.log.warning(f"No VK_CODE mapping for key '{key}'")
            return
        
        scan_code = win32api.MapVirtualKey(vk_code, 0)
        
        if event_type == 'down':
            msg = win32con.WM_KEYDOWN
            lParam = 1 | (scan_code << 16)
        else: # 'up'
            msg = win32con.WM_KEYUP
            lParam = 1 | (scan_code << 16) | (1 << 30) | (1 << 31)

        try:
            win32gui.PostMessage(hwnd, msg, vk_code, lParam)
        except Exception as e:
            self.log.debug(f"Post key event failed for key '{key}' ({hex(vk_code)}) on {hex(hwnd)}: {e}")

    def _broadcast_key_event(self, key: str, targets: List[int], event_type: str):
        with self._broadcast_lock:
            for hwnd in targets:
                self._post_key_event(hwnd, key, event_type)
    
    def _broadcast_char_event(self, char: str, targets: List[int]):
        with self._broadcast_lock:
            for hwnd in targets:
                try:
                    win32gui.PostMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                except Exception as e:
                    self.log.debug(f"Post char event failed for '{char}' on {hwnd}: {e}")

    def _post_mouse_click(self, hwnd:int, rx:float, ry:float, button:str):
        try:
            x,y,w,h = client_rect(hwnd)
            cx=max(0, min(int(rx*w), w-1))
            cy=max(0, min(int(ry*h), h-1))
            lp=_mk_lparam(cx, cy)
            if button=='left':
                win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lp)
                win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP,   0, lp)
            elif button=='right':
                win32gui.PostMessage(hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lp)
                win32gui.PostMessage(hwnd, win32con.WM_RBUTTONUP,   0, lp)
            else:
                win32gui.PostMessage(hwnd, win32con.WM_MBUTTONDOWN, win32con.MK_MBUTTON, lp)
                win32gui.PostMessage(hwnd, win32con.WM_MBUTTONUP,   0, lp)
        except Exception as e:
            self.log.debug(f"Post click failed on {hwnd}: {e}")

    def _broadcast_click(self, rx:float, ry:float, button:str, targets:List[int]):
        threads=[]
        for hwnd in targets:
            t=threading.Thread(target=self._post_mouse_click, args=(hwnd,rx,ry,button), daemon=True)
            threads.append(t); t.start()
        for t in threads: t.join(timeout=0.02)

    def _targets_for(self, target:str, named_slots:List[str]=[]) -> List[int]:
        if not self.hwnds: return []
        if target=="main": return [self.main_hwnd] if self.main_hwnd else []
        if target=="all": return self.hwnds[:]
        if target=="all_except_main": return [h for h in self.hwnds if h!=self.main_hwnd]
        if target=="slots":
            hwnds=[]
            for s in named_slots:
                hwnd = self.proc_hwnd.get(s)
                if hwnd: hwnds.append(hwnd)
            return hwnds
        if target=="round_robin":
            arr=[h for h in self.hwnds if h!=self.main_hwnd] or self.hwnds[:1]
            if not arr: return []
            key="__rr_default__"
            idx=self._rr_state.get(key,0)%len(arr)
            self._rr_state[key]=idx+1
            return [arr[idx]]
        return self.hwnds[:]

    def _run_macro_body(self, m:Macro):
        targets=self._targets_for(m.target, m.slots)
        if not targets: return
        if m.target=="round_robin":
            arr=[h for h in self.hwnds if h!=self.main_hwnd] or self.hwnds[:1]
            if not arr: return
            idx=m.rr_index % len(arr)
            targets=[arr[idx]]
            m.rr_index=(m.rr_index+1)%len(arr); save_config(self.cfg)
        
        for st in m.steps:
            if st.type=="delay":
                time.sleep(max(0.0, st.delay or 0)); continue
            if st.delay: time.sleep(max(0.0, st.delay))
            
            if st.type=="key":
                key_value = (st.value or "").lower().strip()
                if key_value in VK_CODE_MAP:
                    self._broadcast_key_event(key_value, targets, 'down')
                    time.sleep(0.02) # Small delay is good practice
                    self._broadcast_key_event(key_value, targets, 'up')
                elif len(key_value) == 1:
                    self._broadcast_char_event(key_value, targets)
            elif st.type=="text":
                original=self.main_hwnd
                for hwnd in targets:
                    bring_to_front(hwnd); time.sleep(0.005)
                    keyboard.write(st.value, delay=0)
                if original: bring_to_front(original)
            elif st.type=="mouse":
                btn=str(st.mouse.get("button","left")).lower()
                rx=float(st.mouse.get("x",0.5)); ry=float(st.mouse.get("y",0.5))
                self._broadcast_click(rx,ry,btn,targets)

    def run_macro(self, m:Macro):
        t=threading.Thread(target=self._run_macro_body, args=(m,), daemon=True, name=f"BC-Macro-{m.name}")
        self._macro_threads.add(t)
        def _cleanup(th=t):
            try: th.join()
            except Exception: pass
            finally: self._macro_threads.discard(th)
        threading.Thread(target=_cleanup, daemon=True).start()
        t.start()

    def start_macro_loop(self, m:Macro):
        if m.name in self._loop_stop:
            self.status_changed.emit(f"Loop '{m.name}' already running."); return
        stop_evt=threading.Event()
        self._loop_stop[m.name]=stop_evt
        
        def _loop():
            i = 0
            while not stop_evt.is_set():
                self.run_macro(m)
                
                i += 1
                if m.loop_count and i >= m.loop_count:
                    break
                
                if stop_evt.wait(max(0.0, m.loop_interval)):
                    break
            
            self._loop_stop.pop(m.name, None)
            self._loop_threads.pop(m.name, None)
            self.status_changed.emit(f"Loop '{m.name}' finished.")
            
        th=threading.Thread(target=_loop, daemon=True, name=f"BC-Loop-{m.name}")
        self._loop_threads[m.name]=th; th.start()
        self.status_changed.emit(f"Loop '{m.name}' started.")

    def stop_macro_loop(self, name:str):
        ev=self._loop_stop.get(name)
        if ev: ev.set()

    def stop_all_loops(self):
        for ev in list(self._loop_stop.values()): ev.set()

    def _check_condition(self, condition: Condition, stats: Dict[str, Any]) -> bool:
        stat_value = stats.get(condition.stat)
        if stat_value is None: return False

        if condition.stat in ('has_target', 'target_is_own_character'):
            expected_bool = bool(condition.value > 0)
            is_true = False
            if condition.stat == 'has_target':
                is_true = stat_value
            elif condition.stat == 'target_is_own_character':
                own_character_names = {slot.character_name.lower() for slot in self.cfg.slots.values() if slot.character_name}
                target_name = stats.get('target_name', '').lower()
                if target_name:
                    is_true = target_name in own_character_names

            if condition.operator == '==':
                return is_true == expected_bool
            elif condition.operator == '!=':
                return is_true != expected_bool
            return False
            
        elif condition.stat == 'target_name':
            stat_str = str(stat_value).lower()
            cond_str = str(condition.value).lower()
            if condition.operator == '==':
                return stat_str == cond_str
            if condition.operator == '!=':
                return stat_str != cond_str
            return False
            
        op_map = {
            "<": lambda a, b: a < b, ">": lambda a, b: a > b,
            "==": lambda a, b: a == b, "!=": lambda a, b: a != b,
            "<=": lambda a, b: a <= b, ">=": lambda a, b: a >= b,
        }
        if condition.operator not in op_map: return False
        
        # Ensure value is numeric for numeric comparisons
        try:
            numeric_value = float(condition.value)
        except (ValueError, TypeError):
            return False # Cannot compare non-numeric value

        return op_map[condition.operator](stat_value, numeric_value)

    def _thread_rotation_engine(self, stop_event: threading.Event):
        self.status_changed.emit("Rotation Engine STARTED.")
        self.log.info("Rotation engine thread started.")
        client_rotation_map: Dict[int, Tuple[CombatRotation, float]] = {}

        while not stop_event.is_set():
            # Dynamically find clients with assigned rotations
            active_clients = {}
            for slot_name, hwnd in self.proc_hwnd.items():
                if not win32gui.IsWindow(hwnd): continue
                slot_cfg = self.cfg.slots.get(slot_name)
                if slot_cfg and slot_cfg.assigned_rotation:
                    rotation = self.cfg.rotations.get(slot_cfg.assigned_rotation)
                    if rotation:
                        active_clients[hwnd] = rotation

            for hwnd, rotation in active_clients.items():
                last_exec_time = client_rotation_map.get(hwnd, (None, 0.0))[1]
                if time.time() - last_exec_time < rotation.loop_interval:
                    continue

                reader = self.memory_readers.get(hwnd)
                stats = reader.get_combat_stats() if reader else None
                if not stats: continue

                for rule in rotation.rules:
                    all_conditions_met = all(self._check_condition(cond, stats) for cond in rule.conditions)
                    if all_conditions_met:
                        self.log.debug(f"Rot '{rotation.name}' on {hwnd}: Firing '{rule.keys_to_press}'")
                        for key_to_press in rule.keys_to_press:
                            self._broadcast_key_event(key_to_press, [hwnd], 'down')
                            time.sleep(0.05)
                            self._broadcast_key_event(key_to_press, [hwnd], 'up')
                            time.sleep(0.05) # Delay before next key in sequence
                        client_rotation_map[hwnd] = (rotation, time.time())
                        break 
            
            if stop_event.wait(0.1): # Master loop check interval
                break
        
        self._rotation_engine_thread = None
        self.status_changed.emit("Rotation Engine STOPPED.")
        self.log.info("Rotation engine thread stopped.")

    def toggle_rotation_engine(self):
        if getattr(self, "_rotation_engine_thread", None) and self._rotation_engine_thread.is_alive():
            self._rotation_stop_event.set()
        else:
            if not self.running_set:
                self.status_changed.emit("Cannot start rotation engine: No set is running.")
                return
            self._rotation_stop_event = threading.Event()
            self._rotation_engine_thread = threading.Thread(target=self._thread_rotation_engine, args=(self._rotation_stop_event,), daemon=True)
            self._rotation_engine_thread.start()

    def stop_all_rotations(self):
        if getattr(self, "_rotation_engine_thread", None) and self._rotation_engine_thread.is_alive():
            self._rotation_stop_event.set()

    def _is_client_fg(self) -> bool:
        if not self.cfg.settings.keymap.broadcast_only_when_client_focused: return True
        return win32gui.GetForegroundWindow() in self.hwnds

    def _collect_reserved_hotkeys(self):
        km=self.cfg.settings.keymap
        r=set([km.toggle_broadcast_hotkey, km.swap_hotkey])
        if km.cycle_main_hotkey: r.add(km.cycle_main_hotkey)
        for _,m in self.cfg.macros.items():
            if m.hotkey: r.add(m.hotkey)
        if km.toggle_rotations_hotkey:
            r.add(km.toggle_rotations_hotkey)
        self._reserved_hotkeys=set(hk for hk in r if hk)

    @staticmethod
    def _parse_hotkey(hk:str)->Tuple[set, Optional[str], Optional[str]]:
        hk=(hk or "").strip().lower()
        if not hk: return set(), None, None
        parts=[p for p in hk.replace(" ", "").split("+") if p]
        mods=set(); key=None; mouse_btn=None
        for p in parts:
            if p in ("ctrl","alt","shift","win","windows"):
                mods.add("win" if p=="windows" else p)
            elif p.startswith("mouse:"):
                mouse_btn=p.split(":",1)[1]
                if mouse_btn in ("left","right","middle","x","x1","x2"):
                    mouse_btn = {"x1":"x"}.get(mouse_btn, mouse_btn)
            else:
                key=p
        return mods, key, mouse_btn

    def start_hooks(self):
        self.stop_hooks()
        self._collect_reserved_hotkeys()
        km=self.cfg.settings.keymap

        def _add_kb_hotkey(combo:str, func:Callable):
            if not combo: return
            try:
                keyboard.add_hotkey(combo, func); self._hotkey_handles.append((combo, func))
            except Exception as e:
                self.log.error(f"Failed to register hotkey '{combo}': {e}")

        _add_kb_hotkey(km.toggle_broadcast_hotkey, self.toggle_broadcast)
        _add_kb_hotkey(km.swap_hotkey, self.cycle_main)
        if km.cycle_main_hotkey: _add_kb_hotkey(km.cycle_main_hotkey, self.cycle_main)
        _add_kb_hotkey(km.toggle_rotations_hotkey, self.toggle_rotation_engine)

        self._mouse_hotkeys.clear()
        
        # Register Macro Hotkeys
        for _,m in self.cfg.macros.items():
            if not m.hotkey: continue
            mods, key, mouse_btn = self._parse_hotkey(m.hotkey)
            def make_cb(mm=m):
                def _cb():
                    if mm.loop:
                        if mm.name in self._loop_stop: self.stop_macro_loop(mm.name)
                        else: self.start_macro_loop(mm)
                    else:
                        self.run_macro(mm)
                return _cb
            cb=make_cb()
            if mouse_btn: self._mouse_hotkeys.append((mouse_btn, mods, cb))
            else: _add_kb_hotkey(m.hotkey, cb)
            
        def on_key(e):
            if not self._is_client_fg(): return
            
            fg_hwnd = win32gui.GetForegroundWindow()
            targets_to_broadcast = [h for h in self.hwnds if h != fg_hwnd]
            name=(e.name or "").lower()
            event_type = e.event_type

            if name in self._broadcast_keys_down and event_type == 'down':
                return
            
            if event_type == 'down':
                self._broadcast_keys_down.add(name)
            elif event_type == 'up':
                self._broadcast_keys_down.discard(name)

            MODIFIER_KEYS = {
                "shift", "ctrl", "alt", "left windows", "right windows", "windows", "win",
                "umschalt", "strg", "linke windows"
            }

            if self.cfg.settings.broadcast_all_keys:
                if name not in MODIFIER_KEYS:
                    self._broadcast_key_event(name, targets_to_broadcast, event_type)
                return

            if not self.cfg.settings.broadcast_enabled: return
            if name in MODIFIER_KEYS: return
            
            wl=set(k.lower() for k in self.cfg.settings.keymap.whitelist)
            
            try:
                for hk in self._reserved_hotkeys:
                    if hk and keyboard.is_pressed(hk): return
            except Exception: pass
            
            if name in wl:
                if targets_to_broadcast:
                    if self.cfg.settings.keymap.auto_assist_enabled and event_type == 'down':
                        for prefix_key in self.cfg.settings.keymap.auto_assist_prefix_keys:
                            self._broadcast_key_event(prefix_key, targets_to_broadcast, 'down')
                            time.sleep(0.05)
                            self._broadcast_key_event(prefix_key, targets_to_broadcast, 'up')
                            time.sleep(0.05)
                        self._broadcast_key_event(name, targets_to_broadcast, 'down')
                        time.sleep(0.05)
                        self._broadcast_key_event(name, targets_to_broadcast, 'up')
                    elif not self.cfg.settings.keymap.auto_assist_enabled:
                        self._broadcast_key_event(name, targets_to_broadcast, event_type)

        def on_mouse(ev):
            if isinstance(ev, mouse.ButtonEvent) and ev.event_type=="down":
                btn = ev.button
                mods_now=set(n for n,probe in (("ctrl","ctrl"),("alt","alt"),("shift","shift"),("win","windows")) if keyboard.is_pressed(probe))
                for mb, mods, cb in self._mouse_hotkeys:
                    if btn==mb and mods.issubset(mods_now):
                        cb(); return
            
            if not self.cfg.settings.keymap.mouse_broadcast_enabled: return
            if isinstance(ev, mouse.ButtonEvent) and ev.event_type=="up":
                btn=ev.button
                if btn not in ("left","right","middle"): return
                if not self.hwnds: return
                
                fg_hwnd = win32gui.GetForegroundWindow()
                if fg_hwnd not in self.hwnds: return
                
                mx,my,mw,mh = client_rect(fg_hwnd)
                cur=win32api.GetCursorPos()
                mw=max(1,mw); mh=max(1,mh)
                rx=(cur[0]-mx)/mw; ry=(cur[1]-my)/mh
                
                if not (0 <= rx <= 1 and 0 <= ry <= 1): return

                hold=self.cfg.settings.keymap.mouse_hold
                hold_pressed = keyboard.is_pressed(hold) if hold else False
                if hold_pressed:
                    targets_to_broadcast = [h for h in self.hwnds if h != fg_hwnd]
                    if targets_to_broadcast:
                        self._broadcast_click(rx,ry,btn,targets_to_broadcast)

        self._kb_hook = on_key; self._mouse_hook = on_mouse
        keyboard.hook(on_key); mouse.hook(on_mouse)

    def stop_hooks(self):
        self._broadcast_keys_down.clear()
        for combo, func in self._hotkey_handles:
            try: keyboard.remove_hotkey(combo)
            except Exception: pass
        self._hotkey_handles.clear()
        if self._kb_hook:
            try: keyboard.unhook(self._kb_hook)
            except Exception: pass
            self._kb_hook=None
        if self._mouse_hook:
            try: mouse.unhook(self._mouse_hook)
            except Exception: pass
            self._mouse_hook=None
        self._mouse_hotkeys.clear()
        self.stop_all_loops()
        self.stop_all_rotations()

# ---------- Small helpers ----------------------------------------------------
def affinity_to_text(aff: List[int]) -> str:
    return ",".join(str(i) for i in aff)

def parse_affinity(text: str) -> List[int]:
    out=[]
    for part in text.replace(" ", "").split(","):
        if not part: continue
        try:
            v=int(part)
            if v>=0: out.append(v)
        except Exception: pass
    return out

# ---------- Hotkey Capture Dialog -------------------------------------------
class HotkeyCaptureDialog(QtWidgets.QDialog):
    hotkey_captured = QtCore.Signal(str)
    dialog_accept = QtCore.Signal()
    dialog_reject = QtCore.Signal()

    def __init__(self, parent=None, capture_mouse=True):
        super().__init__(parent)
        self.setWindowTitle("Capture Hotkey")
        self.resize(380,130)
        v=QtWidgets.QVBoxLayout(self)
        self._capture_mouse = capture_mouse
        msg = "Press the desired key or mouse button." if capture_mouse else "Press the desired key."
        msg += "\nCombine with Ctrl/Alt/Shift. Enter to confirm, Esc to cancel."
        self.lbl=QtWidgets.QLabel(msg)
        self.lbl.setAlignment(QtCore.Qt.AlignCenter)
        v.addWidget(self.lbl,1)
        self.captured: Optional[str]=None
        self._kb_hook=None; self._mouse_hook=None
        
        self.hotkey_captured.connect(self._update_label_text, QtCore.Qt.QueuedConnection)
        self.dialog_accept.connect(self.accept, QtCore.Qt.QueuedConnection)
        self.dialog_reject.connect(self.reject, QtCore.Qt.QueuedConnection)

    @QtCore.Slot(str)
    def _update_label_text(self, text):
        self.lbl.setText(f"Captured: {text}\nEnter to accept, Esc to cancel")

    def _compose(self, main:str, is_mouse=False)->str:
        parts=[]
        if keyboard.is_pressed("ctrl"): parts.append("ctrl")
        if keyboard.is_pressed("alt"): parts.append("alt")
        if keyboard.is_pressed("shift"): parts.append("shift")
        if keyboard.is_pressed("windows"): parts.append("win")
        parts.append(f"mouse:{main}" if is_mouse else main)
        return "+".join(parts)

    def _on_key(self, e):
        # This function runs in a background thread
        if e.event_type!="down": return
        name=(e.name or "").lower()
        
        if name in ("esc","escape"):
            self.dialog_reject.emit()
            return
        if name in ("enter","return"):
            if self.captured:
                self.dialog_accept.emit()
            return
        if name in ("ctrl","alt","shift","windows","left windows","right windows"): return
        
        self.captured=self._compose(name, is_mouse=False)
        self.hotkey_captured.emit(self.captured)

    def _on_mouse(self, ev):
        if not self._capture_mouse: return
        # This function runs in a background thread
        if isinstance(ev, mouse.ButtonEvent) and ev.event_type=="down":
            btn = ev.button
            self.captured=self._compose(btn, is_mouse=True)
            self.hotkey_captured.emit(self.captured)

    def exec(self) -> int:
        # Deactivate main hooks to prevent interference
        self.parent().ctrl.stop_hooks()
        self._kb_hook=keyboard.hook(self._on_key)
        self._mouse_hook=mouse.hook(self._on_mouse)
        try:
            return super().exec()
        finally:
            try:
                if self._kb_hook: keyboard.unhook(self._kb_hook)
                if self._mouse_hook: mouse.unhook(self._mouse_hook)
            except Exception: pass
            self._kb_hook=None; self._mouse_hook=None
            # Reactivate main hooks
            self.parent().ctrl.start_hooks()

# ---------- Editors & Preview -----------------------------------------------
class SlotEditor(QtWidgets.QDialog):
    def __init__(self, cfg:AppConfig, slot: Optional[SlotConfig]=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Client (Slot)")
        self.cfg=cfg
        self.slot = slot or SlotConfig()
        self.resize(580, 420)
        form=QtWidgets.QFormLayout(self)
        form.setContentsMargins(12,12,12,12)
        form.setVerticalSpacing(8)

        self.ed_name=QtWidgets.QLineEdit(self.slot.name)
        self.ed_args=QtWidgets.QLineEdit(self.slot.args)
        self.ed_acc=QtWidgets.QLineEdit(self.slot.account)
        self.ed_pass=QtWidgets.QLineEdit(self.slot.password); self.ed_pass.setEchoMode(QtWidgets.QLineEdit.Password)
        self.ed_char_name=QtWidgets.QLineEdit(self.slot.character_name)
        self.ed_realm=QtWidgets.QLineEdit(self.slot.realm or "")
        self.spin_steps=QtWidgets.QSpinBox(); self.spin_steps.setRange(0,50); self.spin_steps.setValue(self.slot.character_steps)

        d=self.slot.login_delays
        self.d_after_start=QtWidgets.QDoubleSpinBox(); self.d_after_start.setRange(0,120); self.d_after_start.setSingleStep(0.1); self.d_after_start.setValue(d.get("after_start",8.0))
        self.d_after_user=QtWidgets.QDoubleSpinBox(); self.d_after_user.setRange(0,10); self.d_after_user.setSingleStep(0.1); self.d_after_user.setValue(d.get("after_user",0.2))
        self.d_after_pass=QtWidgets.QDoubleSpinBox(); self.d_after_pass.setRange(0,10); self.d_after_pass.setSingleStep(0.1); self.d_after_pass.setValue(d.get("after_pass",0.2))
        self.d_after_enter=QtWidgets.QDoubleSpinBox(); self.d_after_enter.setRange(0,60); self.d_after_enter.setSingleStep(0.1); self.d_after_enter.setValue(d.get("after_enter",5.0))
        self.d_after_char=QtWidgets.QDoubleSpinBox(); self.d_after_char.setRange(0,60); self.d_after_char.setSingleStep(0.1); self.d_after_char.setValue(d.get("after_char",4.0))

        self.ed_aff=QtWidgets.QLineEdit(affinity_to_text(self.slot.cpu_affinity))
        self.cb_prio=QtWidgets.QComboBox(); self.cb_prio.addItems(["idle","below_normal","normal","above_normal","high"])
        self.cb_prio.setCurrentText(self.slot.priority)

        self.cb_rot = QtWidgets.QComboBox()
        self.cb_rot.addItem("None")
        self.cb_rot.addItems(sorted(self.cfg.rotations.keys()))
        if self.slot.assigned_rotation:
            self.cb_rot.setCurrentText(self.slot.assigned_rotation)

        form.addRow("Slot Name:", self.ed_name)
        form.addRow("Arguments:", self.ed_args)
        form.addRow("Account:", self.ed_acc)
        form.addRow("Password:", self.ed_pass)
        form.addRow("Character Name:", self.ed_char_name)
        form.addRow("Realm:", self.ed_realm)
        form.addRow("Character Steps (↓):", self.spin_steps)

        grp_del=QtWidgets.QGroupBox("Login Delays (seconds)")
        g=QtWidgets.QFormLayout(grp_del)
        g.setContentsMargins(10,10,10,10); g.setVerticalSpacing(6)
        g.addRow("After Start:", self.d_after_start)
        g.addRow("After Username:", self.d_after_user)
        g.addRow("After Password:", self.d_after_pass)
        g.addRow("After Enter:", self.d_after_enter)
        g.addRow("After Character:", self.d_after_char)
        form.addRow(grp_del)

        form.addRow("CPU Affinity (e.g. 0,2,4):", self.ed_aff)
        form.addRow("Priority:", self.cb_prio)
        form.addRow("Assigned Rotation:", self.cb_rot)

        btns=QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        form.addRow(btns)
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject)

    def accept(self):
        s=self.slot
        s.name=self.ed_name.text().strip() or s.name
        s.args=self.ed_args.text()
        s.account=self.ed_acc.text()
        s.password=self.ed_pass.text()
        s.character_name=self.ed_char_name.text().strip()
        s.realm=self.ed_realm.text().strip() or None
        s.character_steps=self.spin_steps.value()
        s.login_delays={
            "after_start": self.d_after_start.value(),
            "after_user": self.d_after_user.value(),
            "after_pass": self.d_after_pass.value(),
            "after_enter": self.d_after_enter.value(),
            "after_char": self.d_after_char.value(),
        }
        s.cpu_affinity=parse_affinity(self.ed_aff.text())
        s.priority=self.cb_prio.currentText()
        s.assigned_rotation = self.cb_rot.currentText() if self.cb_rot.currentText() != "None" else None
        super().accept()

class SetEditor(QtWidgets.QDialog):
    def __init__(self, cfg:AppConfig, cs: Optional[CharacterSet]=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Character Set")
        self.cfg=cfg
        self.cs = cs or CharacterSet(name="NewSet", slots=[])
        self.resize(600, 680)
        v=QtWidgets.QVBoxLayout(self)
        v.setContentsMargins(12,12,12,12)
        f=QtWidgets.QFormLayout()
        f.setVerticalSpacing(8)

        self.ed_name=QtWidgets.QLineEdit(self.cs.name)
        self.ed_exe=QtWidgets.QLineEdit(self.cs.exe_path)
        self.btn_exe=QtWidgets.QPushButton("Browse…")
        self.ed_work=QtWidgets.QLineEdit(self.cs.workdir)
        self.btn_work=QtWidgets.QPushButton("Browse…")
        
        row_exe=QtWidgets.QHBoxLayout(); row_exe.addWidget(self.ed_exe,1); row_exe.addWidget(self.btn_exe)
        row_work=QtWidgets.QHBoxLayout(); row_work.addWidget(self.ed_work,1); row_work.addWidget(self.btn_work)
        
        self.cb_auto=QtWidgets.QCheckBox(); self.cb_auto.setChecked(self.cs.auto_login)
        self.d_start=QtWidgets.QDoubleSpinBox(); self.d_start.setRange(0,30); self.d_start.setSingleStep(0.1); self.d_start.setValue(self.cs.start_interval)
        self.d_stop=QtWidgets.QDoubleSpinBox(); self.d_stop.setRange(0,30); self.d_stop.setSingleStep(0.1); self.d_stop.setValue(self.cs.stop_grace)

        f.addRow("Set Name:", self.ed_name)
        f.addRow("Executable:", row_exe)
        f.addRow("Working Folder:", row_work)
        f.addRow("Auto-Login:", self.cb_auto)
        f.addRow("Start Interval (s):", self.d_start)
        f.addRow("Stop Grace (s):", self.d_stop)
        v.addLayout(f)
        
        self.btn_exe.clicked.connect(self._pick_exe)
        self.btn_work.clicked.connect(self._pick_work)

        grp=QtWidgets.QGroupBox("Slots in Set")
        gl=QtWidgets.QVBoxLayout(grp); gl.setContentsMargins(10,10,10,10); gl.setSpacing(6)
        self.list=QtWidgets.QListWidget()
        self.list.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        for k,slot in self.cfg.slots.items():
            it=QtWidgets.QListWidgetItem(f"{slot.name}  ({k})")
            it.setFlags(it.flags() | QtCore.Qt.ItemIsUserCheckable)
            it.setCheckState(QtCore.Qt.Checked if k in self.cs.slots else QtCore.Qt.Unchecked)
            it.setData(QtCore.Qt.UserRole, k)
            self.list.addItem(it)
        gl.addWidget(self.list)

        lay_grp=QtWidgets.QGroupBox("Layout & Display")
        l=QtWidgets.QGridLayout(lay_grp)
        self.cb_layout = QtWidgets.QComboBox()
        self.layouts = {
            "Main Left, Slaves Right": "main_left_slaves_right",
            "Main Right, Slaves Left": "main_right_slaves_left",
            "Main on Monitor 1, Slaves on 2": "main_on_monitor_1",
            "Main on Monitor 2, Slaves on 1": "main_on_monitor_2"
        }
        self.cb_layout.addItems(self.layouts.keys())
        current_layout_key = next((key for key, value in self.layouts.items() if value == self.cs.layout_mode), "Main Left, Slaves Right")
        self.cb_layout.setCurrentText(current_layout_key)

        self.spin_cols=QtWidgets.QSpinBox(); self.spin_cols.setRange(1,8); self.spin_cols.setValue(self.cs.grid_cols)
        self.spin_rows=QtWidgets.QSpinBox(); self.spin_rows.setRange(1,8); self.spin_rows.setValue(self.cs.grid_rows)
        
        self.preview = LayoutPreview()
        self.preview.setMinimumHeight(150)

        presets_bar = QtWidgets.QHBoxLayout()
        for c, r in [(1,4), (2,2), (3,2)]:
            btn = QtWidgets.QPushButton(f"{c}x{r}")
            btn.clicked.connect(lambda _, co=c, ro=r: (self.spin_cols.setValue(co), self.spin_rows.setValue(ro)))
            presets_bar.addWidget(btn)

        l.addWidget(QtWidgets.QLabel("Layout Mode:"), 0, 0)
        l.addLayout(presets_bar, 2, 1)
        l.addWidget(self.cb_layout, 0, 1)
        l.addWidget(QtWidgets.QLabel("Slave Grid:"), 1, 0)
        grid_box=QtWidgets.QHBoxLayout(); grid_box.addWidget(self.spin_cols); grid_box.addWidget(QtWidgets.QLabel("x")); grid_box.addWidget(self.spin_rows)
        l.addLayout(grid_box, 1, 1)
        l.addWidget(self.preview, 3, 0, 1, 2)
        
        v.addWidget(grp)
        v.addWidget(lay_grp)

        btns=QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        v.addWidget(btns)
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject)

        self.cb_layout.currentTextChanged.connect(self._update_preview)
        self.spin_cols.valueChanged.connect(self._update_preview)
        self.spin_rows.valueChanged.connect(self._update_preview)
        self.list.itemChanged.connect(self._update_preview)
        self._update_preview()

    def _update_preview(self):
        mode = self.layouts[self.cb_layout.currentText()]
        cols = self.spin_cols.value()
        rows = self.spin_rows.value()
        num_slots = sum(1 for i in range(self.list.count()) if self.list.item(i).checkState() == QtCore.Qt.Checked)
        self.preview.configure(mode, cols, rows, num_slots)

    def _pick_exe(self):
        p,_=QtWidgets.QFileDialog.getOpenFileName(self, "Select EXE", self.ed_exe.text() or BASE_DIR, "Programs (*.exe);;All files (*.*)")
        if p: self.ed_exe.setText(p)

    def _pick_work(self):
        d=QtWidgets.QFileDialog.getExistingDirectory(self, "Select Working Folder", self.ed_work.text() or BASE_DIR)
        if d: self.ed_work.setText(d)

    def accept(self):
        cs=self.cs
        cs.name=self.ed_name.text().strip() or cs.name
        cs.exe_path=self.ed_exe.text().strip()
        cs.workdir=self.ed_work.text().strip()
        cs.auto_login=self.cb_auto.isChecked()
        cs.start_interval=self.d_start.value()
        cs.stop_grace=self.d_stop.value()
        cs.layout_mode = self.layouts[self.cb_layout.currentText()]
        cs.grid_cols=self.spin_cols.value()
        cs.grid_rows=self.spin_rows.value()
        slots=[]
        for i in range(self.list.count()):
            it=self.list.item(i)
            if it.checkState()==QtCore.Qt.Checked:
                slots.append(it.data(QtCore.Qt.UserRole))
        cs.slots=slots
        super().accept()
        
class LayoutPreview(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.monitors = get_monitors()
        self.mode = "main_left_slaves_right"
        self.cols = 1
        self.rows = 4
        self.num_slots = 5

    def configure(self, mode: str, cols: int, rows: int, num_slots: int):
        self.mode = mode
        self.cols = cols
        self.rows = rows
        self.num_slots = num_slots
        self.update()

    def paintEvent(self, ev: QtGui.QPaintEvent):
        p = QtGui.QPainter(self)
        p.setRenderHint(QtGui.QPainter.Antialiasing)
        
        total_w = sum(m.width for m in self.monitors)
        total_h = max(m.height for m in self.monitors)
        scale = min(self.width() / total_w, self.height() / total_h) * 0.95
        
        font = p.font(); font.setPointSize(8); p.setFont(font)

        for i, mon in enumerate(self.monitors):
            mon_rect = QtCore.QRectF(mon.x * scale, mon.y * scale, mon.width * scale, mon.height * scale)
            p.setPen(QtGui.QColor("#555"))
            p.setBrush(QtGui.QColor("#333"))
            p.drawRect(mon_rect)
            p.setPen(QtGui.QColor("#888"))
            p.drawText(mon_rect.adjusted(5, 5, 0, 0), f"Monitor {i+1}")
        
        layout = ClientLayout(padding=2)
        hwnds = list(range(self.num_slots))
        positions = layout.positions(hwnds, self.mode, self.cols, self.rows)

        for i, hwnd in enumerate(hwnds):
            if hwnd not in positions: continue
            x, y, w, h = positions[hwnd]
            win_rect = QtCore.QRectF(x * scale, y * scale, w * scale, h * scale)
            if i == 0:
                p.setBrush(QtGui.QColor(42, 130, 218, 150)) # Main
                p.setPen(QtGui.QColor(42, 130, 218))
                label = "Main"
            else:
                p.setBrush(QtGui.QColor(200, 200, 200, 80)) # Slave
                p.setPen(QtGui.QColor(200, 200, 200))
                label = f"S{i}"
            
            p.drawRect(win_rect)
            p.drawText(win_rect, QtCore.Qt.AlignCenter, label)

# ---------- Macro Editor & Helpers -----------------------------------------------------
class MouseStepDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, step: Optional[MacroStep] = None):
        super().__init__(parent)
        self.setWindowTitle("Mouse Click Step")
        self.step_data = step.mouse if step and step.mouse else {"button": "left", "x": 0.5, "y": 0.5}

        layout = QtWidgets.QFormLayout(self)
        self.button_combo = QtWidgets.QComboBox()
        self.button_combo.addItems(["left", "right", "middle"])
        self.button_combo.setCurrentText(self.step_data.get("button", "left"))

        self.x_spin = QtWidgets.QDoubleSpinBox()
        self.x_spin.setRange(0.0, 1.0)
        self.x_spin.setDecimals(3)
        self.x_spin.setSingleStep(0.01)
        self.x_spin.setValue(self.step_data.get("x", 0.5))

        self.y_spin = QtWidgets.QDoubleSpinBox()
        self.y_spin.setRange(0.0, 1.0)
        self.y_spin.setDecimals(3)
        self.y_spin.setSingleStep(0.01)
        self.y_spin.setValue(self.step_data.get("y", 0.5))
        
        layout.addRow("Button:", self.button_combo)
        layout.addRow("X (relative):", self.x_spin)
        layout.addRow("Y (relative):", self.y_spin)

        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)

    def accept(self):
        self.step_data["button"] = self.button_combo.currentText()
        self.step_data["x"] = self.x_spin.value()
        self.step_data["y"] = self.y_spin.value()
        super().accept()

    def get_data(self):
        return self.step_data

class RecordingDialog(QtWidgets.QDialog):
    stopped = QtCore.Signal()
    def __init__(self, parent=None):
        super().__init__(parent, QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool)
        self.setWindowTitle("Recording...")
        layout = QtWidgets.QVBoxLayout(self)
        self.label = QtWidgets.QLabel("Recording actions...\nPress 'Stop' to finish.")
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.stop_button = QtWidgets.QPushButton("Stop Recording")
        self.stop_button.clicked.connect(self.stop)
        layout.addWidget(self.label)
        layout.addWidget(self.stop_button)
        self.setFixedSize(200, 100)
    
    def stop(self):
        self.stopped.emit()
        self.close()

class MacroEditor(QtWidgets.QDialog):
    key_was_captured = QtCore.Signal(int, str)

    def __init__(self, cfg:AppConfig, ctrl: BoxChampController, name: Optional[str]=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Macro")
        self.cfg=cfg
        self.ctrl = ctrl
        self.m = cfg.macros.get(name) if name else Macro(name="NewMacro")
        if not self.m.steps: self.m.steps = []
        self._is_new = name is None
        self.resize(580, 610)
        v=QtWidgets.QVBoxLayout(self); v.setContentsMargins(12,12,12,12); v.setSpacing(8)
        f=QtWidgets.QFormLayout(); f.setVerticalSpacing(8)

        self.ed_name=QtWidgets.QLineEdit(self.m.name)
        row_hot=QtWidgets.QHBoxLayout()
        self.ed_hotkey=QtWidgets.QLineEdit(self.m.hotkey or "")
        btn_pick=QtWidgets.QPushButton("Capture…")
        row_hot.addWidget(self.ed_hotkey, 1); row_hot.addWidget(btn_pick)
        btn_pick.clicked.connect(self._pick_hotkey)

        self.cb_target=QtWidgets.QComboBox(); self.cb_target.addItems(["all","main","slots","all_except_main","round_robin"])
        self.cb_target.setCurrentText(self.m.target)
        self.ed_slots=QtWidgets.QLineEdit(" ".join(self.m.slots))

        f.addRow("Name:", self.ed_name)
        hotw=QtWidgets.QWidget(); hotw.setLayout(row_hot); f.addRow("Hotkey:", hotw)
        f.addRow("Target:", self.cb_target)
        f.addRow("Slots (if 'slots'):", self.ed_slots)

        grp_loop=QtWidgets.QGroupBox("Loop")
        fl=QtWidgets.QFormLayout(grp_loop); fl.setContentsMargins(10,10,10,10); fl.setVerticalSpacing(6)
        self.cb_loop=QtWidgets.QCheckBox("Enable loop"); self.cb_loop.setChecked(self.m.loop)
        self.sp_interval=QtWidgets.QDoubleSpinBox(); self.sp_interval.setRange(0.01, 10.0); self.sp_interval.setSingleStep(0.05); self.sp_interval.setValue(self.m.loop_interval)
        self.sp_count=QtWidgets.QSpinBox(); self.sp_count.setRange(0, 1000000); self.sp_count.setValue(self.m.loop_count or 0)
        fl.addRow(self.cb_loop)
        fl.addRow("Interval (s):", self.sp_interval)
        fl.addRow("Count (0 = infinite):", self.sp_count)

        v.addLayout(f); v.addWidget(grp_loop)

        self.steps=QtWidgets.QTableWidget(0,3)
        self.steps.setHorizontalHeaderLabels(["Type","Value / Info","Delay (s)"])
        self.steps.horizontalHeader().setStretchLastSection(False)
        self.steps.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        v.addWidget(self.steps)

        for st in self.m.steps: self._add_step_row(st)

        btns_line=QtWidgets.QHBoxLayout()
        btns_line.setSpacing(8)
        btn_add=QtWidgets.QPushButton("Add Step"); btn_del=QtWidgets.QPushButton("Remove Selected"); btn_rec=QtWidgets.QPushButton("Record Macro")
        btns_line.addWidget(btn_add); btns_line.addWidget(btn_del); btns_line.addStretch(); btns_line.addWidget(btn_rec)
        v.addLayout(btns_line)
        btn_add.clicked.connect(self._add_step_manually)
        btn_del.clicked.connect(lambda: self.steps.removeRow(self.steps.currentRow()) if self.steps.currentRow()>=0 else None)
        btn_rec.clicked.connect(self._start_recording)

        btns=QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        v.addWidget(btns)
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject)

        self._recording_hook = None
        self._recorded_steps = []
        self._last_event_time = 0
        
        self.key_was_captured.connect(self._on_key_captured_update_gui)

    def _pick_hotkey(self):
        dlg=HotkeyCaptureDialog(self)
        if dlg.exec():
            if dlg.captured: self.ed_hotkey.setText(dlg.captured)
            
    def _add_step_manually(self):
        self._add_step_row(MacroStep())

    def _add_step_row(self, st:MacroStep):
        r=self.steps.rowCount(); self.steps.insertRow(r)
        
        cb=QtWidgets.QComboBox(); cb.addItems(["key","text","delay","mouse"]); cb.setCurrentText(st.type)
        self.steps.setCellWidget(r,0,cb)
        
        dly=QtWidgets.QDoubleSpinBox(); dly.setRange(0,60); dly.setSingleStep(0.01); dly.setDecimals(3); dly.setValue(st.delay)
        self.steps.setCellWidget(r,2,dly)
        
        item = QtWidgets.QTableWidgetItem()
        item.setData(QtCore.Qt.UserRole, st)
        self.steps.setItem(r, 1, item)

        self._update_step_row_widgets(r)
        cb.currentIndexChanged.connect(lambda _, row=r: self._update_step_row_widgets(row))
    
    def _update_step_row_widgets(self, row):
        new_type = self.steps.cellWidget(row, 0).currentText()
        item = self.steps.item(row, 1)
        step = item.data(QtCore.Qt.UserRole)
        step.type = new_type

        if new_type == "key":
            btn = QtWidgets.QPushButton(step.value or "Set Key")
            btn.clicked.connect(lambda _, r=row: self._capture_key_for_step(r))
            self.steps.setCellWidget(row, 1, btn)
        elif new_type == "text":
            ed = QtWidgets.QLineEdit(step.value)
            self.steps.setCellWidget(row, 1, ed)
        elif new_type == "mouse":
            btn = QtWidgets.QPushButton(f"Click: {step.mouse.get('button','left')} @ ({step.mouse.get('x',0.5):.2f}, {step.mouse.get('y',0.5):.2f})")
            btn.clicked.connect(lambda _, r=row: self._edit_mouse_step(r))
            self.steps.setCellWidget(row, 1, btn)
        else:
            self.steps.setCellWidget(row, 1, QtWidgets.QWidget())

    def _edit_mouse_step(self, row):
        item = self.steps.item(row, 1)
        step = item.data(QtCore.Qt.UserRole)
        
        dlg = MouseStepDialog(self, step)
        if dlg.exec():
            step.mouse = dlg.get_data()
            btn = self.steps.cellWidget(row, 1)
            btn.setText(f"Click: {step.mouse.get('button','left')} @ ({step.mouse.get('x',0.5):.2f}, {step.mouse.get('y',0.5):.2f})")

    def _start_recording(self):
        self.ctrl.stop_hooks()
        self._recorded_steps = []
        self._last_event_time = time.time()
        
        self.rec_dialog = RecordingDialog(self)
        self.rec_dialog.stopped.connect(self._stop_recording)
        
        self._recording_hook = keyboard.hook(self._on_record_event, suppress=True)
        self.rec_dialog.show()

    def _on_record_event(self, e):
        if e.event_type != "down": return
        
        current_time = time.time()
        delay = current_time - self._last_event_time
        self._last_event_time = current_time
        
        if e.name in ("shift","ctrl","alt","left windows","right windows","windows","win"): return
        
        step = MacroStep(type="key", value=e.name, delay=round(delay, 3))
        self._recorded_steps.append(step)

    def _stop_recording(self):
        if self._recording_hook:
            keyboard.unhook(self._recording_hook)
            self._recording_hook = None
        
        self.steps.setRowCount(0)
        for step in self._recorded_steps:
            self._add_step_row(step)
        
        self.ctrl.start_hooks()

    def _capture_key_for_step(self, row):
        btn = self.steps.cellWidget(row, 1)
        if not btn: return
        btn.setText("Press a key...")
        btn.setEnabled(False)
        self.ctrl.stop_hooks()
        threading.Thread(target=self._thread_read_key, args=(row,), daemon=True).start()

    def _thread_read_key(self, row):
        try:
            key_name = keyboard.read_key(suppress=True)
            self.key_was_captured.emit(row, key_name)
        except Exception as e:
            self.ctrl.log.error(f"Error capturing key: {e}")
            self.key_was_captured.emit(row, "")

    @QtCore.Slot(int, str)
    def _on_key_captured_update_gui(self, row, key_name):
        try:
            btn = self.steps.cellWidget(row, 1)
            item = self.steps.item(row, 1)
            if not btn or not item: return

            step_data = item.data(QtCore.Qt.UserRole)
            original_value = step_data.value
            step_data.value = key_name
            
            btn.setText(key_name or original_value or "Set Key")
            btn.setEnabled(True)
        finally:
            self.ctrl.start_hooks()

    def accept(self):
        name = self.ed_name.text().strip() or "Macro"
        hot = self.ed_hotkey.text().strip() or None
        tgt = self.cb_target.currentText()
        slots = [s for s in self.ed_slots.text().split() if s]
        steps = []
        for r in range(self.steps.rowCount()):
            step_data = self.steps.item(r, 1).data(QtCore.Qt.UserRole)
            step_data.type = self.steps.cellWidget(r, 0).currentText()
            step_data.delay = self.steps.cellWidget(r, 2).value()
            
            if step_data.type == "text":
                widget = self.steps.cellWidget(r, 1)
                if isinstance(widget, QtWidgets.QLineEdit):
                    step_data.value = widget.text()

            steps.append(step_data)
        
        if self._is_new and name in self.cfg.macros:
            QtWidgets.QMessageBox.warning(self, "Name exists", "There is already a macro with this name.")
            return
            
        self.m.name = name
        self.m.hotkey = hot
        self.m.target = tgt
        self.m.slots = slots
        self.m.steps = steps
        self.m.loop = self.cb_loop.isChecked()
        self.m.loop_interval = self.sp_interval.value()
        self.m.loop_count = (self.sp_count.value() or None)
        self.cfg.macros[name] = self.m
        super().accept()

# ---------- Rotation Editor --------------------------------------------------
class ConditionEditor(QtWidgets.QDialog):
    def __init__(self, cfg: AppConfig, parent=None, condition: Optional[Condition]=None):
        super().__init__(parent)
        self.cfg = cfg
        self.cond = condition or Condition()
        self.setWindowTitle("Edit Condition")
        layout = QtWidgets.QFormLayout(self)
        
        self.cb_stat = QtWidgets.QComboBox()
        self.cb_stat.addItems(VALID_STATS)
        
        self.cb_op = QtWidgets.QComboBox()
        
        self.value_stack = QtWidgets.QStackedWidget()
        self.sp_val = QtWidgets.QSpinBox(); self.sp_val.setRange(0, 200) # Increased range for level etc.
        self.ed_val = QtWidgets.QLineEdit()
        self.value_stack.addWidget(self.sp_val)
        self.value_stack.addWidget(self.ed_val)
        
        self.cb_stat.setCurrentText(self.cond.stat)
        
        layout.addRow("Statistic:", self.cb_stat)
        layout.addRow("Operator:", self.cb_op)
        layout.addRow("Value:", self.value_stack)
        self.cb_stat.currentTextChanged.connect(self._update_ui)
        
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept); btns.rejected.connect(self.reject)
        layout.addRow(btns)
        self._update_ui(self.cond.stat)
        self.cb_op.setCurrentText(self.cond.operator)


    def _update_ui(self, stat_name: str):
        self.cb_op.clear()
        
        is_bool_stat = stat_name in ('has_target', 'target_is_own_character')
        
        if stat_name == 'target_name':
            self.value_stack.setCurrentWidget(self.ed_val)
            self.ed_val.setText(str(self.cond.value) if isinstance(self.cond.value, str) else "")
            self.cb_op.addItems(["==", "!="])
        else: # Default numeric type
            self.value_stack.setCurrentWidget(self.sp_val)
            self.sp_val.setValue(int(self.cond.value) if isinstance(self.cond.value, (int, float)) else 0)
            self.cb_op.addItems(VALID_OPERATORS)

        if is_bool_stat:
            self.sp_val.setToolTip("Use 1 for TRUE and 0 for FALSE.")
            self.sp_val.setRange(0, 1)
            self.cb_op.clear()
            self.cb_op.addItems(["==", "!="])
        else:
            self.sp_val.setToolTip("")
            self.sp_val.setRange(0, 99999)
        
        # Ensure the current operator is valid for the new list of operators
        if self.cb_op.findText(self.cond.operator) > -1:
            self.cb_op.setCurrentText(self.cond.operator)

        
    def accept(self):
        self.cond.stat = self.cb_stat.currentText()
        self.cond.operator = self.cb_op.currentText()
        
        current_widget = self.value_stack.currentWidget()
        if current_widget == self.ed_val:
            self.cond.value = self.ed_val.text()
        else: # self.sp_val
            self.cond.value = self.sp_val.value()
            
        super().accept()

class RotationEditor(QtWidgets.QDialog):
    def __init__(self, cfg: AppConfig, ctrl: BoxChampController, name: Optional[str] = None, parent=None):
        super().__init__(parent)
        self.cfg = cfg
        self.ctrl = ctrl
        self.rot = copy.deepcopy(cfg.rotations.get(name)) if name and name in cfg.rotations else CombatRotation()
        self._original_name = name
        self._is_new = name is None
        self.setWindowTitle("Edit Rotation")
        self.resize(700, 650)

        main_layout = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()
        self.ed_name = QtWidgets.QLineEdit(self.rot.name)
        self.sp_interval = QtWidgets.QDoubleSpinBox()
        self.sp_interval.setRange(0.1, 10.0); self.sp_interval.setSingleStep(0.1)
        self.sp_interval.setValue(self.rot.loop_interval)
        form.addRow("Name:", self.ed_name)
        form.addRow("Check Interval (s):", self.sp_interval)
        main_layout.addLayout(form)
        
        # --- Client Assignment Group ---
        assign_group = QtWidgets.QGroupBox("Assigned Clients")
        assign_layout = QtWidgets.QVBoxLayout(assign_group)
        self.slot_list = QtWidgets.QListWidget()
        assign_layout.addWidget(self.slot_list)
        main_layout.addWidget(assign_group)
        
        self.all_slots = sorted(self.cfg.slots.items())
        for slot_key, slot_obj in self.all_slots:
            item = QtWidgets.QListWidgetItem(slot_obj.name)
            item.setData(QtCore.Qt.UserRole, slot_key)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
            
            if slot_obj.assigned_rotation == self.rot.name:
                item.setCheckState(QtCore.Qt.Checked)
            else:
                item.setCheckState(QtCore.Qt.Unchecked)
            self.slot_list.addItem(item)
        
        # --- Keys and Conditions ---
        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        
        keys_widget = QtWidgets.QWidget()
        keys_layout = QtWidgets.QVBoxLayout(keys_widget)
        keys_layout.addWidget(QtWidgets.QLabel("Keys/Skills (evaluated top-down):"))
        self.keys_list = QtWidgets.QListWidget()
        keys_layout.addWidget(self.keys_list)
        key_btns = QtWidgets.QHBoxLayout()
        btn_add_key = QtWidgets.QPushButton("Add Key")
        btn_edit_key = QtWidgets.QPushButton("Edit Key")
        btn_rem_key = QtWidgets.QPushButton("Remove Key")
        key_btns.addWidget(btn_add_key); key_btns.addWidget(btn_edit_key); key_btns.addWidget(btn_rem_key)
        keys_layout.addLayout(key_btns)
        
        cond_widget = QtWidgets.QWidget()
        cond_layout = QtWidgets.QVBoxLayout(cond_widget)
        cond_layout.addWidget(QtWidgets.QLabel("Conditions for selected key (ALL must be true):"))
        self.cond_list = QtWidgets.QListWidget()
        cond_layout.addWidget(self.cond_list)
        cond_btns = QtWidgets.QHBoxLayout()
        btn_add_cond = QtWidgets.QPushButton("Add Condition")
        btn_edit_cond = QtWidgets.QPushButton("Edit Condition")
        btn_rem_cond = QtWidgets.QPushButton("Remove Condition")
        cond_btns.addWidget(btn_add_cond); cond_btns.addWidget(btn_edit_cond); cond_btns.addWidget(btn_rem_cond)
        cond_layout.addLayout(cond_btns)

        splitter.addWidget(keys_widget); splitter.addWidget(cond_widget)
        main_layout.addWidget(splitter)
        
        dialog_btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Save | QtWidgets.QDialogButtonBox.Cancel)
        main_layout.addWidget(dialog_btns)

        btn_add_key.clicked.connect(self._add_key_step)
        btn_edit_key.clicked.connect(self._edit_key_step)
        btn_rem_key.clicked.connect(self._remove_key_step)
        btn_add_cond.clicked.connect(self._add_condition)
        btn_edit_cond.clicked.connect(self._edit_condition)
        btn_rem_cond.clicked.connect(self._remove_condition)
        self.keys_list.currentItemChanged.connect(self._populate_conditions)
        dialog_btns.accepted.connect(self.accept); dialog_btns.rejected.connect(self.reject)
        self._populate_keys_list()

    def _populate_keys_list(self):
        current_row = self.keys_list.currentRow()
        self.keys_list.clear()
        for i, rule in enumerate(self.rot.rules):
            item = QtWidgets.QListWidgetItem(f"Press '{', '.join(rule.keys_to_press)}'")
            item.setData(QtCore.Qt.UserRole, i)
            self.keys_list.addItem(item)
        
        if 0 <= current_row < self.keys_list.count():
            self.keys_list.setCurrentRow(current_row)
        
        self.cond_list.clear()

    def _populate_conditions(self, current_item, _):
        self.cond_list.clear()
        if not current_item: return
        rule_index = current_item.data(QtCore.Qt.UserRole)
        rule = self.rot.rules[rule_index]
        for i, cond in enumerate(rule.conditions):
            c_item = QtWidgets.QListWidgetItem(f"IF {cond.stat} {cond.operator} {cond.value}")
            c_item.setData(QtCore.Qt.UserRole, i)
            self.cond_list.addItem(c_item)

    def _add_key_step(self):
        keys_str, ok = QtWidgets.QInputDialog.getText(self, "Add Key/Skill Step", "Keys to press (comma-separated):")
        if ok and keys_str:
            keys = [k.strip() for k in keys_str.split(',') if k.strip()]
            if keys:
                self.rot.rules.append(RotationRule(keys_to_press=keys))
                self._populate_keys_list()
                self.keys_list.setCurrentRow(self.keys_list.count() - 1)
    
    def _edit_key_step(self):
        item = self.keys_list.currentItem()
        if not item: return
        
        rule_index = item.data(QtCore.Qt.UserRole)
        rule = self.rot.rules[rule_index]
        current_keys_str = ", ".join(rule.keys_to_press)
        
        new_keys_str, ok = QtWidgets.QInputDialog.getText(self, "Edit Key/Skill Step", "Keys to press (comma-separated):", text=current_keys_str)
        
        if ok and new_keys_str:
            keys = [k.strip() for k in new_keys_str.split(',') if k.strip()]
            if keys:
                rule.keys_to_press = keys
                self._populate_keys_list()

    def _remove_key_step(self):
        item = self.keys_list.currentItem()
        if not item: return
        rule_index = item.data(QtCore.Qt.UserRole)
        del self.rot.rules[rule_index]
        self._populate_keys_list()

    def _add_condition(self):
        key_item = self.keys_list.currentItem()
        if not key_item: return
        dlg = ConditionEditor(self.cfg, self)
        if dlg.exec():
            rule = self.rot.rules[key_item.data(QtCore.Qt.UserRole)]
            rule.conditions.append(dlg.cond)
            self._populate_conditions(key_item, None)

    def _edit_condition(self):
        key_item = self.keys_list.currentItem()
        cond_item = self.cond_list.currentItem()
        if not key_item or not cond_item: return
        rule = self.rot.rules[key_item.data(QtCore.Qt.UserRole)]
        cond_index = cond_item.data(QtCore.Qt.UserRole)
        dlg = ConditionEditor(self.cfg, self, rule.conditions[cond_index])
        if dlg.exec():
            self._populate_conditions(key_item, None)

    def _remove_condition(self):
        key_item = self.keys_list.currentItem()
        cond_item = self.cond_list.currentItem()
        if not key_item or not cond_item: return
        rule = self.rot.rules[key_item.data(QtCore.Qt.UserRole)]
        del rule.conditions[cond_item.data(QtCore.Qt.UserRole)]
        self._populate_conditions(key_item, None)

    def accept(self):
        name = self.ed_name.text().strip()
        if not name: return
        if self._is_new and name in self.cfg.rotations:
            QtWidgets.QMessageBox.warning(self, "Name Exists", "A rotation with this name already exists.")
            return

        # Handle renaming first: update all slots that were pointing to the old name
        if not self._is_new and name != self._original_name:
            for slot in self.cfg.slots.values():
                if slot.assigned_rotation == self._original_name:
                    slot.assigned_rotation = name
        
        # Handle checkbox changes
        for i in range(self.slot_list.count()):
            item = self.slot_list.item(i)
            slot_key = item.data(QtCore.Qt.UserRole)
            slot_to_modify = self.cfg.slots.get(slot_key)
            if not slot_to_modify: continue

            is_checked = (item.checkState() == QtCore.Qt.Checked)
            
            if is_checked:
                slot_to_modify.assigned_rotation = name
            else:
                # Only un-assign if it was assigned to THIS rotation
                if slot_to_modify.assigned_rotation == name:
                    slot_to_modify.assigned_rotation = None

        # Save the rotation object itself
        if not self._is_new and name != self._original_name:
            self.cfg.rotations.pop(self._original_name, None)
        
        self.rot.name = name
        self.rot.loop_interval = self.sp_interval.value()
        self.cfg.rotations[name] = self.rot
        super().accept()

# ---------- Click Bar --------------------------------------------------------
class ClickBarButtonsDialog(QtWidgets.QDialog):
    def __init__(self, cfg:AppConfig, parent=None):
        super().__init__(parent); self.cfg=cfg
        self.setWindowTitle("Click Bar Buttons")
        self.resize(430,300)
        v=QtWidgets.QVBoxLayout(self); self.list=QtWidgets.QListWidget(); v.addWidget(self.list)
        for b in self.cfg.settings.clickbar.buttons: self.list.addItem(f"{b.label} → {b.macro}")
        hb=QtWidgets.QHBoxLayout(); add=QtWidgets.QPushButton("Add"); rem=QtWidgets.QPushButton("Remove")
        hb.addWidget(add); hb.addWidget(rem); v.addLayout(hb)
        btns=QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close); v.addWidget(btns)
        add.clicked.connect(self._add); rem.clicked.connect(self._rem); btns.rejected.connect(self.reject)
    def _add(self):
        label,ok=QtWidgets.QInputDialog.getText(self,"Label","Button Title:")
        if not ok or not label: return
        macro,ok=QtWidgets.QInputDialog.getText(self,"Macro","Macro Name:")
        if not ok or not macro: return
        self.cfg.settings.clickbar.buttons.append(ClickBarButton(label=label, macro=macro))
        self.list.addItem(f"{label} → {macro}"); save_config(self.cfg)
    def _rem(self):
        row=self.list.currentRow()
        if row<0: return
        self.cfg.settings.clickbar.buttons.pop(row)
        self.list.takeItem(row); save_config(self.cfg)

class ClickBarWindow(QtWidgets.QWidget):
    def __init__(self, ctrl:BoxChampController, cfg:AppConfig):
        super().__init__(None, QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool)
        self.ctrl=ctrl; self.cfg=cfg
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self.setWindowTitle("Click Bar")
        self._drag=False; self._drag_pos=QtCore.QPoint()
        
        main_layout = QtWidgets.QHBoxLayout(self)
        main_layout.setContentsMargins(4,4,4,4)
        
        self.frame=QtWidgets.QFrame()
        self.frame.setStyleSheet("QFrame{background:rgba(30,30,30,160);border:1px solid #555;border-radius:4px;}")
        self.layout = QtWidgets.QHBoxLayout(self.frame)
        self.layout.setContentsMargins(4,4,4,4)
        self.layout.setSpacing(4)
        
        grip_layout = QtWidgets.QHBoxLayout()
        grip_layout.setContentsMargins(0,0,0,0)
        grip_layout.addWidget(self.frame)
        grip_layout.addWidget(QtWidgets.QSizeGrip(self), 0, QtCore.Qt.AlignBottom | QtCore.Qt.AlignRight)
        main_layout.addLayout(grip_layout)
        
        self.rebuild()

    def rebuild(self):
        while self.layout.count():
            item = self.layout.takeAt(0)
            widget = item.widget()
            if widget: widget.deleteLater()
            
        cb=self.cfg.settings.clickbar
        self.setGeometry(cb.x, cb.y, cb.width, cb.height)
        
        for b in cb.buttons:
            btn=QtWidgets.QPushButton(b.label); btn.setCursor(QtCore.Qt.PointingHandCursor)
            btn.clicked.connect(lambda _, m=b.macro: self._fire(m))
            self.layout.addWidget(btn)

    def _fire(self, macro_name:str):
        m=self.cfg.macros.get(macro_name)
        if m: self.ctrl.run_macro(m)

    def mousePressEvent(self, e:QtGui.QMouseEvent):
        if e.button() == QtCore.Qt.LeftButton:
            self._drag=True
            self._drag_pos=e.globalPosition().toPoint()-self.frameGeometry().topLeft()
    
    def mouseMoveEvent(self, e:QtGui.QMouseEvent):
        if self._drag: self.move(e.globalPosition().toPoint()-self._drag_pos)

    def mouseReleaseEvent(self, e:QtGui.QMouseEvent):
        self._drag=False
        cb=self.cfg.settings.clickbar
        cb.x=self.x(); cb.y=self.y()
        save_config(self.cfg)

    def resizeEvent(self, event:QtGui.QResizeEvent):
        super().resizeEvent(event)
        if not self.isMaximized() and not self.isMinimized():
            cb=self.cfg.settings.clickbar
            cb.width=self.width(); cb.height=self.height()
            save_config(self.cfg)

# ---------- Broadcast Overlay ------------------------------------------------
class BroadcastOverlayWindow(QtWidgets.QWidget):
    def __init__(self, ctrl: BoxChampController, cfg: AppConfig):
        super().__init__(None, QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint | QtCore.Qt.Tool)
        self.ctrl = ctrl
        self.cfg = cfg
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self.setWindowTitle("Broadcast Overlay")
        self._drag = False
        self._drag_pos = QtCore.QPoint()

        outer = QtWidgets.QVBoxLayout(self)
        outer.setContentsMargins(4, 4, 4, 4)
        frame = QtWidgets.QFrame()
        frame.setStyleSheet("QFrame{background:rgba(30,30,30,200);border:1px solid #555;border-radius:4px;}")
        layout = QtWidgets.QFormLayout(frame)
        layout.setContentsMargins(8, 8, 8, 8)

        self.btn_broadcast = QtWidgets.QPushButton()
        self.btn_broadcast.setCheckable(True)
        self.btn_broadcast_all = QtWidgets.QPushButton()
        self.btn_broadcast_all.setCheckable(True)

        layout.addRow("Broadcast:", self.btn_broadcast)
        layout.addRow("Broadcast ALL:", self.btn_broadcast_all)
        outer.addWidget(frame)

        self.btn_broadcast.clicked.connect(self.ctrl.toggle_broadcast)
        self.btn_broadcast_all.clicked.connect(self.ctrl.toggle_broadcast_all)
        self.ctrl.broadcast_state_changed.connect(self.update_states, QtCore.Qt.QueuedConnection)
        self.update_states()

    def update_states(self):
        s = self.cfg.settings
        self.btn_broadcast.setChecked(s.broadcast_enabled)
        self.btn_broadcast.setText("ON" if s.broadcast_enabled else "OFF")
        self.btn_broadcast.setStyleSheet("background-color: #2a82da;" if s.broadcast_enabled else "background-color: #555;")
        
        self.btn_broadcast_all.setChecked(s.broadcast_all_keys)
        self.btn_broadcast_all.setText("ON" if s.broadcast_all_keys else "OFF")
        self.btn_broadcast_all.setStyleSheet("background-color: #c23b22;" if s.broadcast_all_keys else "background-color: #555;")

    def mousePressEvent(self, e:QtGui.QMouseEvent):
        self._drag=True; self._drag_pos=e.globalPosition().toPoint()-self.frameGeometry().topLeft()
    def mouseMoveEvent(self, e:QtGui.QMouseEvent):
        if self._drag: self.move(e.globalPosition().toPoint()-self._drag_pos)
    def mouseReleaseEvent(self, e:QtGui.QMouseEvent):
        self._drag=False
# ---------- Management Tabs --------------------------------------------------

# DIESEN GANZEN BLOCK AUSSCHNEIDEN

class GroupTargetingTab(QtWidgets.QWidget):
    def __init__(self, main:'MainWindow'):
        super().__init__()
        self.main = main
        v_layout = QtWidgets.QVBoxLayout(self)
        v_layout.setContentsMargins(12, 12, 12, 12)
        v_layout.setSpacing(8)

        gt_group = QtWidgets.QGroupBox("Group Targeting Settings")
        form_layout = QtWidgets.QFormLayout(gt_group)

        self.cb_enabled = QtWidgets.QCheckBox("Enable Group Targeting Hotkeys")
        self.cb_enabled.setToolTip("Enables the hotkeys defined below to coordinate party targeting.")
        form_layout.addRow(self.cb_enabled)

        self.ed_target_slaves = QtWidgets.QLineEdit()
        self.ed_target_slaves.setToolTip("Comma-separated hotkeys the master presses to target slaves.\n(e.g., f2,f3,f4,f5)")
        form_layout.addRow("Target Slave Hotkeys:", self.ed_target_slaves)

        self.ed_target_master = QtWidgets.QLineEdit()
        self.ed_target_master.setToolTip("Hotkey to make all slaves target the master.")
        form_layout.addRow("Target Master Hotkey:", self.ed_target_master)
        
        form_layout.addRow(QtWidgets.QLabel("--- Key Definitions ---"))

        self.ed_key_self = QtWidgets.QLineEdit()
        self.ed_key_self.setToolTip("The key a character presses to target itself (e.g., f1).")
        form_layout.addRow("Key for Self-Target:", self.ed_key_self)
        
        self.ed_key_master = QtWidgets.QLineEdit()
        self.ed_key_master.setToolTip("The key a slave presses to target the master (e.g., f2).")
        form_layout.addRow("Key for Targeting Master:", self.ed_key_master)
        
        v_layout.addWidget(gt_group)
        v_layout.addStretch()

        save_button = QtWidgets.QPushButton("Save Group Targeting Settings")
        save_button.clicked.connect(self._save)
        v_layout.addWidget(save_button)
        
        self.load_settings()

    def load_settings(self):
        gt_cfg = self.main.cfg.settings.group_targeting
        self.cb_enabled.setChecked(gt_cfg.enabled)
        self.ed_target_master.setText(gt_cfg.target_master_hotkey)
        self.ed_target_slaves.setText(", ".join(gt_cfg.target_slave_hotkeys))
        self.ed_key_self.setText(gt_cfg.key_for_self)
        self.ed_key_master.setText(gt_cfg.key_for_master)
        
    def _save(self):
        gt_cfg = self.main.cfg.settings.group_targeting
        gt_cfg.enabled = self.cb_enabled.isChecked()
        gt_cfg.target_master_hotkey = self.ed_target_master.text().strip()
        gt_cfg.target_slave_hotkeys = [k.strip() for k in self.ed_target_slaves.text().split(",") if k.strip()]
        gt_cfg.key_for_self = self.ed_key_self.text().strip()
        gt_cfg.key_for_master = self.ed_key_master.text().strip()
        
        save_config(self.main.cfg)
        self.main.ctrl.start_hooks() # Re-register all hotkeys with the new settings
        self.main.status.showMessage("Group Targeting settings saved.", 3000)

class ClientsTab(QtWidgets.QWidget):
    def __init__(self, main:'MainWindow'):
        super().__init__()
        self.main=main
        v=QtWidgets.QVBoxLayout(self)
        v.setContentsMargins(12,12,12,12)
        v.setSpacing(8)
        self.table=QtWidgets.QTableWidget(0,5)
        self.table.setHorizontalHeaderLabels(["Slot Name","Account","Character Name","Args","Priority"])
        self.table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.reload_table()
        v.addWidget(self.table,1)

        btns=QtWidgets.QHBoxLayout()
        btns.setSpacing(8)
        add=QtWidgets.QPushButton("Add")
        edit=QtWidgets.QPushButton("Edit")
        rem=QtWidgets.QPushButton("Remove")
        dup=QtWidgets.QPushButton("Duplicate")
        btns.addWidget(add); btns.addWidget(edit); btns.addWidget(rem); btns.addWidget(dup); btns.addStretch(1)
        v.addLayout(btns)

        add.clicked.connect(self.add_slot)
        edit.clicked.connect(self.edit_slot)
        rem.clicked.connect(self.remove_slot)
        dup.clicked.connect(self.dup_slot)

    def reload_table(self):
        cfg=self.main.cfg
        self.table.setRowCount(0)
        for key, slot in sorted(cfg.slots.items()):
            r=self.table.rowCount(); self.table.insertRow(r)
            vals=[slot.name, slot.account, slot.character_name, slot.args, slot.priority]
            for c,val in enumerate(vals):
                it=QtWidgets.QTableWidgetItem(str(val))
                if c==0: it.setData(QtCore.Qt.UserRole, key)
                self.table.setItem(r,c,it)

    def _selected_key(self) -> Optional[str]:
        r=self.table.currentRow()
        if r<0: return None
        it=self.table.item(r,0)
        return it.data(QtCore.Qt.UserRole) if it else None

    def add_slot(self):
        dlg=SlotEditor(self.main.cfg, None, self)
        if dlg.exec():
            s=dlg.slot
            key=s.name if s.name and s.name not in self.main.cfg.slots else f"Slot{len(self.main.cfg.slots)+1}"
            s.name=key if s.name=="" else s.name
            self.main.cfg.slots[key]=s
            save_config(self.main.cfg)
            self.reload_table()
            self.main.status.showMessage("Client added.", 3000)

    def edit_slot(self):
        key=self._selected_key()
        if not key: return
        slot=self.main.cfg.slots[key]
        dlg=SlotEditor(self.main.cfg, slot, self)
        if dlg.exec():
            self.main.cfg.slots[key]=dlg.slot
            if key != dlg.slot.name:
                self.main.cfg.slots[dlg.slot.name] = self.main.cfg.slots.pop(key)
                for cs in self.main.cfg.sets:
                    if key in cs.slots:
                        cs.slots = [dlg.slot.name if s == key else s for s in cs.slots]

            save_config(self.main.cfg)
            self.reload_table()
            self.main.status.showMessage("Client saved.", 3000)

    def remove_slot(self):
        key=self._selected_key()
        if not key: return
        if QtWidgets.QMessageBox.question(self,"Remove", f"Delete {key}?")==QtWidgets.QMessageBox.Yes:
            for cs in self.main.cfg.sets:
                if key in cs.slots: cs.slots.remove(key)
            self.main.cfg.slots.pop(key, None)
            save_config(self.main.cfg)
            self.reload_table()
            self.main.status.showMessage("Client removed.", 3000)

    def dup_slot(self):
        key=self._selected_key()
        if not key: return
        base=self.main.cfg.slots[key]
        new=SlotConfig(**json.loads(json.dumps(asdict(base))))
        new.name=base.name+"_copy"
        idx=1
        new_key=new.name
        while new_key in self.main.cfg.slots:
            idx+=1; new_key=f"{new.name}{idx}"
        self.main.cfg.slots[new_key]=new
        save_config(self.main.cfg)
        self.reload_table()
        self.main.status.showMessage("Client duplicated.", 3000)

class SetsTab(QtWidgets.QWidget):
    def __init__(self, main:'MainWindow'):
        super().__init__()
        self.main=main
        v=QtWidgets.QVBoxLayout(self)
        v.setContentsMargins(12,12,12,12)
        v.setSpacing(8)
        self.table=QtWidgets.QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Set Name", "# Slots", "Layout", "Game Path"])
        self.table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        v.addWidget(self.table,1)
        btns=QtWidgets.QHBoxLayout()
        btns.setSpacing(8)
        add=QtWidgets.QPushButton("Add")
        edit=QtWidgets.QPushButton("Edit")
        rem=QtWidgets.QPushButton("Remove")
        btns.addWidget(add); btns.addWidget(edit); btns.addWidget(rem); btns.addStretch(1)
        v.addLayout(btns)
        add.clicked.connect(self.add_set)
        edit.clicked.connect(self.edit_set)
        rem.clicked.connect(self.remove_set)
        self.reload()

    def reload(self):
        self.table.setRowCount(0)
        for i, cs in enumerate(self.main.cfg.sets):
            r = self.table.rowCount()
            self.table.insertRow(r)
            
            name_item = QtWidgets.QTableWidgetItem(cs.name)
            name_item.setData(QtCore.Qt.UserRole, i)
            
            self.table.setItem(r, 0, name_item)
            self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(str(len(cs.slots))))
            self.table.setItem(r, 2, QtWidgets.QTableWidgetItem(cs.layout_mode))
            self.table.setItem(r, 3, QtWidgets.QTableWidgetItem(cs.exe_path))

    def add_set(self):
        dlg=SetEditor(self.main.cfg, None, self)
        if dlg.exec():
            self.main.cfg.sets.append(dlg.cs)
            save_config(self.main.cfg)
            self.reload()
            self.main.status.showMessage("Set added.", 3000)

    def edit_set(self):
        r = self.table.currentRow()
        if r < 0: return
        idx=self.table.item(r, 0).data(QtCore.Qt.UserRole)
        cs=self.main.cfg.sets[idx]
        dlg=SetEditor(self.main.cfg, cs, self)
        if dlg.exec():
            save_config(self.main.cfg)
            self.reload()
            self.main.status.showMessage("Set saved.", 3000)

    def remove_set(self):
        r = self.table.currentRow()
        if r < 0: return
        idx=self.table.item(r, 0).data(QtCore.Qt.UserRole)
        if QtWidgets.QMessageBox.question(self,"Remove", "Delete this set?")==QtWidgets.QMessageBox.Yes:
            self.main.cfg.sets.pop(idx)
            save_config(self.main.cfg)
            self.reload()
            self.main.status.showMessage("Set removed.", 3000)

class GeneralTab(QtWidgets.QWidget):
    def __init__(self, main:'MainWindow'):
        super().__init__()
        self.main = main
        
        outer_layout = QtWidgets.QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        
        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QtWidgets.QFrame.NoFrame)
        
        container = QtWidgets.QWidget()
        v = QtWidgets.QVBoxLayout(container)
        v.setContentsMargins(12, 12, 12, 12)
        v.setSpacing(8)

        # --- KEY BROADCASTING GROUP ---
        kb_group=QtWidgets.QGroupBox("Key Broadcasting")
        kb_layout = QtWidgets.QVBoxLayout(kb_group)
        
        self.wl_list = QtWidgets.QListWidget()
        self.wl_list.setToolTip("Only keys in this list will be broadcast to slave windows.")
        wl_box = QtWidgets.QVBoxLayout()
        wl_box.addWidget(QtWidgets.QLabel("Broadcast Whitelist:"))
        wl_box.addWidget(self.wl_list)
        wl_btns = QtWidgets.QHBoxLayout()
        wl_add = QtWidgets.QPushButton("Add Key")
        wl_rem = QtWidgets.QPushButton("Remove Selected")
        wl_btns.addWidget(wl_add); wl_btns.addWidget(wl_rem)
        wl_box.addLayout(wl_btns)
        kb_layout.addLayout(wl_box)
        
        self.cb_auto_assist = QtWidgets.QCheckBox("Auto Assist Enable")
        self.cb_auto_assist.setToolTip("If checked, pressing a whitelisted key sends the 'Assist Keys' sequence to slaves first.")
        kb_layout.addWidget(self.cb_auto_assist)

        assist_layout = QtWidgets.QHBoxLayout()
        self.ed_auto_assist_keys = QtWidgets.QLineEdit()
        self.ed_auto_assist_keys.setToolTip("Comma-separated keys to press before the actual key (e.g., 'f2, h').")
        assist_layout.addWidget(QtWidgets.QLabel("Assist Keys:"))
        assist_layout.addWidget(self.ed_auto_assist_keys)
        kb_layout.addLayout(assist_layout)

        v.addWidget(kb_group)
        
        wl_add.clicked.connect(lambda: self._add_key_to_list(self.wl_list))
        wl_rem.clicked.connect(lambda: self._rem_key_from_list(self.wl_list))

        # --- WINDOW MANAGEMENT GROUP ---
        wm_group = QtWidgets.QGroupBox("Window Management")
        f_wm = QtWidgets.QFormLayout(wm_group)
        self.cb_rename = QtWidgets.QComboBox()
        self.cb_rename.addItems(["No Rename", "By Slot Name", "By Account Name", "By Character Name"])
        self.cb_rename.setToolTip("Automatically rename game windows after they launch.")
        
        self.ed_titles=QtWidgets.QLineEdit()
        self.ed_titles.setToolTip("Comma-separated list of keywords to identify game windows (not case-sensitive).")

        self.ed_exe_names = QtWidgets.QLineEdit()
        self.ed_exe_names.setToolTip("Comma-separated list of process names to identify as game clients (e.g., Wow.exe, WowClassic.exe).")

        self.cb_borderless=QtWidgets.QCheckBox("Borderless Windows");
        self.cb_borderless.setToolTip("Remove the title bar and borders from all game windows.")
        
        self.cb_full=QtWidgets.QCheckBox("Main Window covers Taskbar");
        self.cb_full.setToolTip("Allows the main window to cover the entire screen, including the taskbar.")

        self.cb_topmost_mode=QtWidgets.QComboBox(); self.cb_topmost_mode.addItems(["always","active_only","never"])
        self.cb_topmost_mode.setToolTip(
            "<b>Always:</b> Main window always on top.<br>"
            "<b>Active Only:</b> Main window on top only when a game window is focused.<br>"
            "<b>Never:</b> Default window behavior."
        )

        self.spin_pad=QtWidgets.QSpinBox(); self.spin_pad.setRange(0,40)
        self.spin_pad.setToolTip("The amount of empty space (in pixels) between tiled windows.")

        f_wm.addRow("Rename Scheme:", self.cb_rename)
        f_wm.addRow("Window Title Filters:", self.ed_titles)
        f_wm.addRow("Game Executable Names:", self.ed_exe_names)
        f_wm.addRow(self.cb_borderless)
        f_wm.addRow(self.cb_full)
        f_wm.addRow("Topmost Mode:", self.cb_topmost_mode)
        f_wm.addRow("Tiling Padding:", self.spin_pad)
        v.addWidget(wm_group)

        # --- MISC GROUP ---
        misc_group = QtWidgets.QGroupBox("Hotkeys & Overlays")
        f_misc = QtWidgets.QFormLayout(misc_group)
        self.ed_tgl=QtWidgets.QLineEdit(); self.ed_tgl.setToolTip("Global hotkey to enable/disable keyboard broadcasting via the overlay.")
        self.ed_swap=QtWidgets.QLineEdit(); self.ed_swap.setToolTip("Hotkey to swap main window with the first slave.")
        self.ed_mouse_hold=QtWidgets.QLineEdit(); self.ed_mouse_hold.setToolTip("Modifier key to hold to broadcast mouse clicks (e.g., 'ctrl', 'alt').")
        self.ed_toggle_rotations_hotkey = QtWidgets.QLineEdit(); self.ed_toggle_rotations_hotkey.setToolTip("Global hotkey to enable/disable the combat rotation engine.")
        self.cb_mouse_broadcast_enabled=QtWidgets.QCheckBox("Enable Mouse Broadcast");
        self.cb_mouse_broadcast_enabled.setToolTip("Globally enable or disable the mouse broadcasting feature.")
        self.btn_edit_clickbar = QtWidgets.QPushButton("Edit Click Bar Position/Size")
        self.btn_edit_clickbar.setToolTip("Shows the Click Bar so you can drag and resize it visually.")
        self.btn_edit_clickbar.clicked.connect(lambda: self.main.toggle_clickbar(force_on=True))
        f_misc.addRow("Toggle Broadcast Hotkey:", self.ed_tgl)
        f_misc.addRow("Swap Main Hotkey:", self.ed_swap)
        f_misc.addRow("Mouse Broadcast Hold Key:", self.ed_mouse_hold)
        f_misc.addRow("Toggle Rotations Hotkey:", self.ed_toggle_rotations_hotkey)
        f_misc.addRow(self.cb_mouse_broadcast_enabled)
        f_misc.addRow(self.btn_edit_clickbar)
        v.addWidget(misc_group)
        
        hb=QtWidgets.QHBoxLayout()
        hb.setSpacing(8)
        save=QtWidgets.QPushButton("Save Settings")
        reload=QtWidgets.QPushButton("Reload Settings")
        openf=QtWidgets.QPushButton("Open Folder")
        hb.addWidget(save); hb.addWidget(reload); hb.addWidget(openf)
        v.addLayout(hb)
        v.addStretch(1)
        
        scroll_area.setWidget(container)
        outer_layout.addWidget(scroll_area)

        for combo in container.findChildren(QtWidgets.QComboBox):
            combo.installEventFilter(self)

        save.clicked.connect(self._save)
        reload.clicked.connect(self._reload)
        openf.clicked.connect(lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(APP_DIR)))
        
        self.load_settings()

    def eventFilter(self, obj, event):
        if event.type() == QtCore.QEvent.Type.Wheel and isinstance(obj, QtWidgets.QComboBox):
            return True # Eat the event
        return super().eventFilter(obj, event)

    def _add_key_to_list(self, list_widget):
        text, ok = QtWidgets.QInputDialog.getText(self, "Add Key", "Enter key name (e.g., 'f1', 'g', 'numpad_5'):")
        if ok and text:
            list_widget.addItem(text.strip().lower())

    def _rem_key_from_list(self, list_widget):
        for item in list_widget.selectedItems():
            list_widget.takeItem(list_widget.row(item))
            
    def load_settings(self):
        s = self.main.cfg.settings
        km = s.keymap
        self.cb_rename.setCurrentText(s.window_rename_scheme)
        self.ed_titles.setText(", ".join(s.window_title_filters))
        self.ed_exe_names.setText(", ".join(s.game_executable_names))
        self.cb_borderless.setChecked(s.borderless)
        self.cb_full.setChecked(s.main_fullscreen_taskbar)
        self.cb_topmost_mode.setCurrentText(s.main_topmost_mode)
        self.spin_pad.setValue(s.tile_padding)
        self.wl_list.clear(); self.wl_list.addItems(km.whitelist)
        self.cb_auto_assist.setChecked(km.auto_assist_enabled)
        self.ed_auto_assist_keys.setText(", ".join(km.auto_assist_prefix_keys))
        self.ed_tgl.setText(km.toggle_broadcast_hotkey)
        self.ed_swap.setText(km.swap_hotkey)
        self.ed_mouse_hold.setText(km.mouse_hold)
        self.ed_toggle_rotations_hotkey.setText(km.toggle_rotations_hotkey)
        self.cb_mouse_broadcast_enabled.setChecked(km.mouse_broadcast_enabled)

    def _save(self):
        s=self.main.cfg.settings
        km=s.keymap
        s.window_rename_scheme = self.cb_rename.currentText()
        s.window_title_filters=[x.strip() for x in self.ed_titles.text().split(",") if x.strip()]
        s.game_executable_names=[x.strip() for x in self.ed_exe_names.text().split(",") if x.strip()]
        s.borderless=self.cb_borderless.isChecked()
        s.main_fullscreen_taskbar=self.cb_full.isChecked()
        s.main_topmost_mode=self.cb_topmost_mode.currentText()
        s.tile_padding=self.spin_pad.value()
        km.whitelist = [self.wl_list.item(i).text() for i in range(self.wl_list.count())]
        km.auto_assist_enabled = self.cb_auto_assist.isChecked()
        prefix_keys_text = self.ed_auto_assist_keys.text().strip()
        km.auto_assist_prefix_keys = [key.strip().lower() for key in prefix_keys_text.replace(" ", "").split(",") if key.strip()]
        km.toggle_broadcast_hotkey=self.ed_tgl.text().strip() or km.toggle_broadcast_hotkey
        km.swap_hotkey=self.ed_swap.text().strip() or km.swap_hotkey
        km.mouse_hold=self.ed_mouse_hold.text().strip()
        km.toggle_rotations_hotkey = self.ed_toggle_rotations_hotkey.text().strip() or km.toggle_rotations_hotkey
        km.mouse_broadcast_enabled=self.cb_mouse_broadcast_enabled.isChecked()
        
        save_config(self.main.cfg)
        self.main.ctrl.start_hooks()
        self.main.ctrl._enforce_topmost()
        self.main.status.showMessage("Settings saved.", 3000)

    def _reload(self):
        self.main.cfg = load_config()
        self.load_settings()
        self.main.status.showMessage("Settings reloaded.", 3000)

class MacrosTab(QtWidgets.QWidget):
    def __init__(self, main:'MainWindow'):
        super().__init__()
        self.main=main
        v=QtWidgets.QVBoxLayout(self)
        v.setContentsMargins(12,12,12,12)
        v.setSpacing(8)
        self.table=QtWidgets.QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["Macro Name", "Hotkey", "Target", "Steps", "Looping"])
        self.table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)

        v.addWidget(self.table,1)
        hb=QtWidgets.QHBoxLayout()
        hb.setSpacing(8)
        add=QtWidgets.QPushButton("Add")
        edit=QtWidgets.QPushButton("Edit")
        rem=QtWidgets.QPushButton("Remove")
        stop=QtWidgets.QPushButton("Stop All Loops")
        hb.addWidget(add); hb.addWidget(edit); hb.addWidget(rem); hb.addStretch(); hb.addWidget(stop)
        v.addLayout(hb)
        add.clicked.connect(self.add_macro)
        edit.clicked.connect(self.edit_macro)
        rem.clicked.connect(self.remove_macro)
        stop.clicked.connect(self.main.ctrl.stop_all_loops)
        self.reload()

    def reload(self):
        self.table.setRowCount(0)
        for name, m in sorted(self.main.cfg.macros.items()):
            r = self.table.rowCount()
            self.table.insertRow(r)
            
            name_item = QtWidgets.QTableWidgetItem(name)
            name_item.setData(QtCore.Qt.UserRole, name)

            self.table.setItem(r, 0, name_item)
            self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(m.hotkey or "-"))
            self.table.setItem(r, 2, QtWidgets.QTableWidgetItem(m.target))
            self.table.setItem(r, 3, QtWidgets.QTableWidgetItem(str(len(m.steps))))
            self.table.setItem(r, 4, QtWidgets.QTableWidgetItem("Yes" if m.loop else "No"))


    def add_macro(self):
        dlg=MacroEditor(self.main.cfg, self.main.ctrl, None, self)
        if dlg.exec():
            save_config(self.main.cfg)
            self.main.ctrl.start_hooks()
            self.reload()
            self.main.status.showMessage("Macro added.", 3000)

    def edit_macro(self):
        r = self.table.currentRow()
        if r < 0: return
        name=self.table.item(r, 0).data(QtCore.Qt.UserRole)
        dlg=MacroEditor(self.main.cfg, self.main.ctrl, name, self)
        if dlg.exec():
            save_config(self.main.cfg)
            self.main.ctrl.start_hooks()
            self.reload()
            self.main.status.showMessage("Macro saved.", 3000)

    def remove_macro(self):
        r = self.table.currentRow()
        if r < 0: return
        name=self.table.item(r, 0).data(QtCore.Qt.UserRole)
        if QtWidgets.QMessageBox.question(self,"Remove", f"Delete macro {name}?")==QtWidgets.QMessageBox.Yes:
            self.main.cfg.macros.pop(name, None)
            save_config(self.main.cfg)
            self.main.ctrl.start_hooks()
            self.reload()
            self.main.status.showMessage("Macro removed.", 3000)

class RotationsTab(QtWidgets.QWidget):
    def __init__(self, main: 'MainWindow'):
        super().__init__()
        self.main = main
        v_layout = QtWidgets.QVBoxLayout(self)
        v_layout.setContentsMargins(12, 12, 12, 12)
        v_layout.setSpacing(8)

        self.table = QtWidgets.QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Rotation Name", "# Rules", "Interval (s)"])
        self.table.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        v_layout.addWidget(self.table, 1)

        h_layout = QtWidgets.QHBoxLayout()
        h_layout.setSpacing(8)
        btn_add = QtWidgets.QPushButton("Add")
        btn_edit = QtWidgets.QPushButton("Edit")
        btn_rem = QtWidgets.QPushButton("Remove")
        btn_stop_all = QtWidgets.QPushButton("Stop All Rotations")
        
        h_layout.addWidget(btn_add)
        h_layout.addWidget(btn_edit)
        h_layout.addWidget(btn_rem)
        h_layout.addStretch()
        h_layout.addWidget(btn_stop_all)
        v_layout.addLayout(h_layout)

        btn_add.clicked.connect(self.add_rotation)
        btn_edit.clicked.connect(self.edit_rotation)
        btn_rem.clicked.connect(self.remove_rotation)
        btn_stop_all.clicked.connect(self.main.ctrl.stop_all_rotations)
        
        self.reload()

    def reload(self):
        self.table.setRowCount(0)
        for name, rot in sorted(self.main.cfg.rotations.items()):
            r = self.table.rowCount()
            self.table.insertRow(r)
            
            name_item = QtWidgets.QTableWidgetItem(name)
            name_item.setData(QtCore.Qt.UserRole, name)
            
            self.table.setItem(r, 0, name_item)
            self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(str(len(rot.rules))))
            self.table.setItem(r, 2, QtWidgets.QTableWidgetItem(f"{rot.loop_interval:.2f}"))

    def add_rotation(self):
        dlg = RotationEditor(self.main.cfg, self.main.ctrl, None, self)
        if dlg.exec():
            save_config(self.main.cfg)
            self.main.ctrl.start_hooks()
            self.reload()
            self.main.tab_clients.reload_table() # Reload to reflect assignment changes
            self.main.status.showMessage("Rotation added.", 3000)

    def edit_rotation(self):
        r = self.table.currentRow()
        if r < 0: return
        name = self.table.item(r, 0).data(QtCore.Qt.UserRole)
        dlg = RotationEditor(self.main.cfg, self.main.ctrl, name, self)
        if dlg.exec():
            save_config(self.main.cfg)
            self.main.ctrl.start_hooks()
            self.reload()
            self.main.tab_clients.reload_table() # Reload to reflect assignment changes
            self.main.status.showMessage("Rotation saved.", 3000)

    def remove_rotation(self):
        r = self.table.currentRow()
        if r < 0: return
        name = self.table.item(r, 0).data(QtCore.Qt.UserRole)
        if QtWidgets.QMessageBox.question(self, "Remove", f"Delete rotation '{name}'?") == QtWidgets.QMessageBox.Yes:
            self.main.cfg.rotations.pop(name, None)
            # Update slots that might still have this rotation assigned
            for slot in self.main.cfg.slots.values():
                if slot.assigned_rotation == name:
                    slot.assigned_rotation = None
            save_config(self.main.cfg)
            self.main.ctrl.start_hooks()
            self.reload()
            self.main.tab_clients.reload_table() # Reload to reflect assignment changes
            self.main.status.showMessage("Rotation removed.", 3000)

class DashboardTab(QtWidgets.QWidget):
    client_clicked = QtCore.Signal(int)
    def __init__(self, main:'MainWindow'):
        super().__init__()
        self.main=main
        self.ctrl = main.ctrl
        self.client_cards: Dict[int, ClientCard] = {}

        v=QtWidgets.QVBoxLayout(self)
        v.setContentsMargins(12,12,12,12)
        v.setSpacing(8)
        self.grid_container=QtWidgets.QWidget(); self.grid=QtWidgets.QGridLayout(self.grid_container)
        self.grid.setContentsMargins(0,0,0,0); self.grid.setSpacing(8)
        scroll=QtWidgets.QScrollArea(); scroll.setWidgetResizable(True); scroll.setWidget(self.grid_container)
        v.addWidget(scroll,1)

        self.update_timer = QtCore.QTimer(self)
        self.update_timer.timeout.connect(self.update_card_stats)
        self.update_timer.start(2000)

    def update_card_stats(self):
        for hwnd, card in self.client_cards.items():
            card.update_stats()
    
    def render_clients(self, items: List[Tuple[int,str]]):
        current_hwnds = {item[0] for item in items}
        
        # Remove cards for closed windows
        for hwnd in list(self.client_cards.keys()):
            if hwnd not in current_hwnds:
                card = self.client_cards.pop(hwnd)
                card.deleteLater()
        
        # Add or update cards
        col=row=0
        num_cols = max(1, self.width() // (ClientCard.CARD_WIDTH + 10))
        hwnd_to_slot = {v: k for k, v in self.ctrl.proc_hwnd.items()}

        for hwnd,title in items:
            if hwnd not in self.client_cards:
                slot_name = hwnd_to_slot.get(hwnd)
                slot_config = self.ctrl.cfg.slots.get(slot_name) if slot_name else None
                
                card=ClientCard(hwnd, self.ctrl, slot_config)
                card.clicked.connect(self.client_clicked.emit)
                self.client_cards[hwnd] = card
                self.grid.addWidget(card, row, col)
                col+=1
                if col>=num_cols: col=0; row+=1
        
        self.highlight_main_window()

    def highlight_main_window(self):
        for hwnd, card in self.client_cards.items():
            card.set_is_main(hwnd == self.ctrl.main_hwnd)

class ClientCard(QtWidgets.QFrame):
    clicked = QtCore.Signal(int)
    CARD_WIDTH = 260
    CARD_HEIGHT = 160
    
    def __init__(self, hwnd:int, ctrl: BoxChampController, slot_config: Optional[SlotConfig]):
        super().__init__()
        self.hwnd = hwnd
        self.ctrl = ctrl
        self.slot_config = slot_config
        self.process = None
        
        try:
            pid = hwnd_pid(self.hwnd)
            self.process = psutil.Process(pid)
            self.process.cpu_percent(interval=None) # Initialize for subsequent calls
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            self.process = None

        self.setFixedSize(self.CARD_WIDTH, self.CARD_HEIGHT)
        self.setFrameShape(QtWidgets.QFrame.StyledPanel)
        
        self.main_layout = QtWidgets.QVBoxLayout(self)
        self.main_layout.setContentsMargins(8,8,8,8)
        
        title_bar = QtWidgets.QHBoxLayout()
        self.title_label = QtWidgets.QLabel(f"<b>{slot_config.name if slot_config else 'Unknown'}</b>")
        self.pid_label = QtWidgets.QLabel(f"PID: {self.process.pid if self.process else 'N/A'}")
        self.pid_label.setStyleSheet("color: #aaa;")
        title_bar.addWidget(self.title_label)
        title_bar.addStretch()
        title_bar.addWidget(self.pid_label)
        
        self.main_layout.addLayout(title_bar)
        
        self.char_label = QtWidgets.QLabel(slot_config.character_name if slot_config and slot_config.character_name else "<i>No Character</i>")
        self.main_layout.addWidget(self.char_label)
        self.main_layout.addStretch(1)

        # HP and MP Bars
        self.hp_bar = QtWidgets.QProgressBar()
        self.hp_bar.setTextVisible(True)
        self.hp_bar.setStyleSheet("QProgressBar { border: 1px solid #555; border-radius: 4px; text-align: center; background-color: #333; }"
                                  "QProgressBar::chunk { background-color: #d63031; border-radius: 3px; }")
        
        self.mp_bar = QtWidgets.QProgressBar()
        self.mp_bar.setTextVisible(True)
        self.mp_bar.setStyleSheet("QProgressBar { border: 1px solid #555; border-radius: 4px; text-align: center; background-color: #333; }"
                                  "QProgressBar::chunk { background-color: #0984e3; border-radius: 3px; }")
                                  
        self.main_layout.addWidget(self.hp_bar)
        self.main_layout.addWidget(self.mp_bar)

        # CPU and Memory stats
        stats_layout = QtWidgets.QFormLayout()
        stats_layout.setContentsMargins(0, 8, 0, 0)
        self.cpu_label = QtWidgets.QLabel("...")
        self.mem_label = QtWidgets.QLabel("...")
        stats_layout.addRow("CPU:", self.cpu_label)
        stats_layout.addRow("Mem:", self.mem_label)
        
        self.main_layout.addLayout(stats_layout)
        self.set_is_main(False)
        self.update_stats()

    def set_is_main(self, is_main):
        if is_main:
            self.setStyleSheet("QFrame { border: 2px solid #2a82da; border-radius: 4px; background: rgba(42, 130, 218, 0.1); }")
        else:
            self.setStyleSheet("QFrame { border: 1px solid #444; border-radius: 4px; background: rgba(60, 60, 70, 0.5); }")

    def update_stats(self):
        # Update CPU/Mem
        if self.process and self.process.is_running():
            try:
                cpu = self.process.cpu_percent(interval=None)
                mem = self.process.memory_info().rss / (1024 * 1024)
                self.cpu_label.setText(f"{cpu:.1f} %")
                self.mem_label.setText(f"{mem:.0f} MB")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                self.cpu_label.setText("N/A")
                self.mem_label.setText("N/A")
        else:
            self.cpu_label.setText("N/A")
            self.mem_label.setText("N/A")
        
        # Update HP/MP from combat_stats for consistency
        reader = self.ctrl.memory_readers.get(self.hwnd)
        stats = reader.get_combat_stats() if reader else None

        if stats and stats.get('player_max_hp', 0) > 0:
            self.hp_bar.setMaximum(stats['player_max_hp'])
            self.hp_bar.setValue(stats['player_hp'])
            self.hp_bar.setFormat(f"HP: {stats['player_hp']} / {stats['player_max_hp']}")
        else:
            self.hp_bar.setValue(0)
            self.hp_bar.setFormat("HP: N/A")
        
        if stats and stats.get('player_max_mp', 0) > 0:
            self.mp_bar.setMaximum(stats['player_max_mp'])
            self.mp_bar.setValue(stats['player_mp'])
            self.mp_bar.setFormat(f"MP: {stats['player_mp']} / {stats['player_max_mp']}")
        else:
            self.mp_bar.setValue(0)
            self.mp_bar.setFormat("MP: N/A")
            
    def contextMenuEvent(self, event: QtGui.QContextMenuEvent):
        menu = QtWidgets.QMenu(self)
        set_main_action = menu.addAction("Set as Main")
        bring_front_action = menu.addAction("Bring to Front")
        
        action = menu.exec(self.mapToGlobal(event.pos()))
        
        if action == set_main_action:
            self.ctrl.make_main(self.hwnd)
        elif action == bring_front_action:
            bring_to_front(self.hwnd)
    
    def mousePressEvent(self, e:QtGui.QMouseEvent):
        if e.button() == QtCore.Qt.LeftButton:
            self.clicked.emit(self.hwnd)
        super().mousePressEvent(e)

# ---------- The Main Window --------------------------------------------------
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, cfg:AppConfig, log:logging.Logger, debug:bool):
        super().__init__()
        self.cfg=cfg; self.log=log; self.debug=debug
        self.ctrl=BoxChampController(cfg, log)
        self.setWindowTitle("BoxChamp – Multibox Manager")
        self.setWindowIcon(QtGui.QIcon("logo.png"))
        self.resize(1200, 700)
        self.setMinimumSize(800, 550)

        tb=self.addToolBar("Actions")
        style=self.style()
        def act(text, icon, slot, checkable=False):
            a=QtGui.QAction(style.standardIcon(icon), text, self)
            a.setCheckable(checkable)
            a.triggered.connect(slot); tb.addAction(a); return a
        tb.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
        act("Find Clients", QtWidgets.QStyle.SP_BrowserReload, self.ctrl.refresh_clients)
        act("Apply Layout", QtWidgets.QStyle.SP_DialogApplyButton, self.ctrl.apply_layout)
        tb.addSeparator()
        act("Start Set", QtWidgets.QStyle.SP_MediaPlay, self.start_set_dialog)
        act("Stop Set", QtWidgets.QStyle.SP_MediaStop, self.ctrl.stop_set)
        tb.addSeparator()
        act("Toggle Overlay", QtWidgets.QStyle.SP_ComputerIcon, self.toggle_broadcast_overlay)
        act("Toggle Click Bar", QtWidgets.QStyle.SP_DialogYesButton, self.toggle_clickbar)
        tb.addSeparator()
        act("Save", QtWidgets.QStyle.SP_DialogSaveButton, lambda: (save_config(self.cfg), self.status.showMessage("Settings saved.", 3000)))
        act("Open Folder", QtWidgets.QStyle.SP_DirOpenIcon, lambda: QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(APP_DIR)))

        self.status=QtWidgets.QStatusBar(); self.setStatusBar(self.status)

        central_widget = QtWidgets.QWidget()
        self.setCentralWidget(central_widget)
        
        self.stacked_layout = QtWidgets.QStackedLayout(central_widget)
        
        self.logo_label = QtWidgets.QLabel()
        self.logo_label.setAlignment(QtCore.Qt.AlignCenter)
        self.logo_pixmap = QtGui.QPixmap("logo.png")
        
        opacity_effect = QtWidgets.QGraphicsOpacityEffect(self.logo_label)
        opacity_effect.setOpacity(0.08)
        self.logo_label.setGraphicsEffect(opacity_effect)

        self.tabs=QtWidgets.QTabWidget()
        
        self.stacked_layout.addWidget(self.logo_label)
        self.stacked_layout.addWidget(self.tabs)
        self.stacked_layout.setCurrentWidget(self.tabs)

        self.tab_dashboard=DashboardTab(self)
        self.tab_clients=ClientsTab(self)
        self.tab_sets=SetsTab(self)
        self.tab_macros=MacrosTab(self)
        self.tab_rotations = RotationsTab(self)
        self.tab_group_targeting = GroupTargetingTab(self)
        self.tab_general=GeneralTab(self)

        self.tabs.addTab(self.tab_dashboard, "Dashboard")
        self.tabs.addTab(self.tab_clients, "Clients")
        self.tabs.addTab(self.tab_sets, "Character Sets")
        self.tabs.addTab(self.tab_macros, "Macros")
        self.tabs.addTab(self.tab_rotations, "Rotations")
        self.tabs.addTab(self.tab_group_targeting, "Group Targeting")
        self.tabs.addTab(self.tab_general, "General")
        self.tab_dashboard.client_clicked.connect(self.ctrl.make_main)

        self.ctrl.status_changed.connect(self.status.showMessage)
        self.ctrl.clients_changed.connect(self.tab_dashboard.render_clients)
        self.ctrl.clients_changed.connect(self.tab_dashboard.highlight_main_window)
        self.ctrl.set_state_changed.connect(self.on_set_state)
        self.ctrl.launch_error.connect(self.on_launch_error)

        self.ctrl.refresh_clients()
        self.ctrl.start_hooks()
        
        self.clickbar_win: Optional[ClickBarWindow]=None
        if self.cfg.settings.clickbar.enabled: self.toggle_clickbar(force_on=True)
        self.broadcast_overlay: Optional[BroadcastOverlayWindow] = None
        
        self.overlay_timer = QtCore.QTimer(self)
        self.overlay_timer.timeout.connect(self._keep_overlays_on_top)
        self.overlay_timer.start(500)

        self.status.showMessage(f"App started • Settings: {SETTINGS_PATH}", 7000)
        self.log.info(f"{APP_NAME} ready • Settings: {SETTINGS_PATH}")
        if debug:
            self.status.showMessage("Debug mode is ON. Check console for detailed logs.", 5000)
            
        QtWidgets.QApplication.instance().aboutToQuit.connect(self._cleanup)
    
    @QtCore.Slot(str, str)
    def on_launch_error(self, slot_name, message):
        QtWidgets.QMessageBox.critical(self, f"Launch Error ({slot_name})", message)

    def _keep_overlays_on_top(self):
        for overlay in [self.clickbar_win, self.broadcast_overlay]:
            if overlay and overlay.isVisible():
                try:
                    win32gui.SetWindowPos(int(overlay.winId()), win32con.HWND_TOPMOST, 0, 0, 0, 0,
                                          win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
                except Exception:
                    pass

    def resizeEvent(self, event: QtGui.QResizeEvent):
        super().resizeEvent(event)
        if hasattr(self, "logo_pixmap") and not self.logo_pixmap.isNull():
            scaled_pixmap = self.logo_pixmap.scaled(self.size(), QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
            self.logo_label.setPixmap(scaled_pixmap)
        self.tab_dashboard.render_clients([(h, win32gui.GetWindowText(h)) for h in self.ctrl.hwnds])

    def _reload_settings(self):
        new_cfg=load_config()
        self.apply_new_config(new_cfg)
        self.status.showMessage("Settings reloaded.", 3000)

    def apply_new_config(self, new_cfg:AppConfig):
        self.cfg=new_cfg
        self.ctrl.cfg=new_cfg
        self.tab_clients.reload_table()
        self.tab_sets.reload()
        self.tab_macros.reload()
        self.tab_rotations.reload()
        
        old_general_tab = self.findChild(GeneralTab)
        if old_general_tab:
            idx = self.tabs.indexOf(old_general_tab)
            self.tabs.removeTab(idx)
            old_general_tab.deleteLater()
        
        self.tab_general=GeneralTab(self)
        self.tabs.insertTab(idx if 'idx' in locals() else self.tabs.count(), self.tab_general, "General")

        self.ctrl.start_hooks()
        self.ctrl.refresh_clients()
        self.ctrl._enforce_topmost()

    def _cleanup(self):
        try:
            self.ctrl.stop_hooks()
        except Exception:
            pass

    def start_set_dialog(self):
        if self.ctrl.running_set:
            self.status.showMessage("A set is already running.")
            return

        if not self.cfg.sets:
            QtWidgets.QMessageBox.information(self, "No set", "Create a set first in 'Character Sets'.")
            return
            
        items=[cs.name for cs in self.cfg.sets]
        name,ok=QtWidgets.QInputDialog.getItem(self,"Start Set","Set:", items, editable=False)
        if not ok: return
        cs=next((c for c in self.cfg.sets if c.name==name), None)
        if not cs: return
        
        self.status.showMessage(f"Selected set: {cs.name}")

        default_btn = QtWidgets.QMessageBox.Yes if cs.auto_login else QtWidgets.QMessageBox.No
        res = QtWidgets.QMessageBox.question(self, "Auto-Login", "Run Auto-Login now?",
                                             QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Cancel,
                                             defaultButton=default_btn)
        if res==QtWidgets.QMessageBox.Cancel:
            return
            
        auto_now = (res==QtWidgets.QMessageBox.Yes)
        self.ctrl.start_set(cs, auto_login_override=auto_now)

    def on_set_state(self, state:str):
        self.status.showMessage({"running":"Set running.","stopped":"Set stopped.","starting":"Starting set…","stopping":"Stopping set…"}\
                                .get(state, state), 3000)
        if state in ["running", "stopped"]:
            self.tab_dashboard.highlight_main_window()

    def toggle_clickbar(self, force_on:bool=False):
        if self.clickbar_win and self.clickbar_win.isVisible():
            if force_on:
                bring_to_front(int(self.clickbar_win.winId()))
                return
            self.clickbar_win.close(); self.clickbar_win=None
            self.cfg.settings.clickbar.enabled=False; save_config(self.cfg); return
            
        self.clickbar_win=ClickBarWindow(self.ctrl, self.cfg)
        self.clickbar_win.show()
        self.cfg.settings.clickbar.enabled=True; save_config(self.cfg)
        
    def toggle_broadcast_overlay(self):
        if self.broadcast_overlay and self.broadcast_overlay.isVisible():
            self.broadcast_overlay.close()
            self.broadcast_overlay = None
        else:
            self.broadcast_overlay = BroadcastOverlayWindow(self.ctrl, self.cfg)
            self.broadcast_overlay.show()
# ---------- Theming ----------------------------------------------------------
def apply_dark_theme(app: QtWidgets.QApplication):
    app.setStyle("Fusion")
    palette = QtGui.QPalette()
    base = QtGui.QColor(45,45,48)
    alt  = QtGui.QColor(37,37,38)
    text = QtGui.QColor(220,220,220)
    dis  = QtGui.QColor(127,127,127)
    btn  = QtGui.QColor(53,53,53)
    high = QtGui.QColor(42,130,218)

    palette.setColor(QtGui.QPalette.Window, base)
    palette.setColor(QtGui.QPalette.WindowText, text)
    palette.setColor(QtGui.QPalette.Base, alt)
    palette.setColor(QtGui.QPalette.AlternateBase, base)
    palette.setColor(QtGui.QPalette.ToolTipBase, text)
    palette.setColor(QtGui.QPalette.ToolTipText, text)
    palette.setColor(QtGui.QPalette.Text, text)
    palette.setColor(QtGui.QPalette.Disabled, QtGui.QPalette.Text, dis)
    palette.setColor(QtGui.QPalette.Button, btn)
    palette.setColor(QtGui.QPalette.ButtonText, text)
    palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
    palette.setColor(QtGui.QPalette.Highlight, high)
    palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.white)
    app.setPalette(palette)

    app.setStyleSheet("""
        QGroupBox {
            border: 1px solid #555; border-radius: 8px; margin-top: 10px; padding: 8px 8px 8px 8px;
        }
        QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
        QTabWidget::pane { border: 1px solid #444; border-radius: 6px; background: transparent; }
        QTabWidget > QWidget { background-color: #2d2d30; }
        QTabBar::tab { padding: 6px 10px; margin: 2px; border: 1px solid #444; border-radius: 6px; background-color: #353535; }
        QTabBar::tab:selected { background: #2a82da; }
        QTableView, QTableWidget { gridline-color: #444; selection-background-color: #2a82da; }
        QPushButton { padding: 4px 8px; border: 1px solid #666; border-radius: 4px; }
        QPushButton:hover { border-color: #8aa; }
        QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox { padding: 3px; }
        QStatusBar { border-top: 1px solid #444; }
        QToolTip { border: 1px solid #666; background-color: #2d2d30; color: #ddd; padding: 4px; }
    """)

# ---------- Exception Hook / Entry ------------------------------------------
def excepthook(exc_type, exc, tb):
    msg = "".join(traceback.format_exception(exc_type, exc, tb))
    try: logging.getLogger("boxchamp").error("Unhandled exception:\n%s", msg)
    except Exception: pass
    QtWidgets.QMessageBox.critical(None, "Unhandled Exception", msg)
    sys.__excepthook__(exc_type, exc, tb)

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument("--debug", action="store_true", help="Debug logging & info dialog")
    args=parser.parse_args()

    faulthandler.enable(); log=setup_logging(args.debug); sys.excepthook = excepthook

    try: ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("BoxChamp.App")
    except Exception: pass

    app=QtWidgets.QApplication(sys.argv)
    apply_dark_theme(app)
    cfg=load_config()
    win=MainWindow(cfg, log, args.debug)
    win.show()

    def _atexit():
        try: win._cleanup()
        except Exception: pass
    atexit.register(_atexit)

    sys.exit(app.exec())

if __name__=="__main__":
    main()