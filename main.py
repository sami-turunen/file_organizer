import os
import shutil
import sys
import win32com.client
import winreg
import configparser
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

ides = [
    "visual studio code", "pycharm", "sublime text", "atom", "notepad++", "vim", 
    "emacs", "geany", "eclipse", "intellij", "visual studio", "xcode", "android studio", 
    "appcode", "webstorm", "phpstorm", "rubymine", "clion", "datagrip", "goland", 
    "rider", "mps", "teamcity", "tmcbeans"
]

browsers = [
    "google chrome", "microsoft edge", "mozilla firefox", "safari", "opera", 
    "brave", "vivaldi", "chromium", "zen", "school - edge", "personal - brave"
]

ai_softwares = ["chatgpt", "perplexity", "copilot", "microsoft 365 copilot"]

dev_tools = [
    "mongodb compass", "mongodb", "studio 3t", "winscp", "everything"
]

utilities = [
    "wiztree", "windirstat", "wise memory optimizer", "wise memory...", 
    "hp support assistant", "chrome remote desktop"
]

games = ["epic games launcher", "epic games", "roblox", "roblox studio"]

script_extensions = [".py", ".ps1", ".bat", ".sh", ".cmd", ".vbs"]

document_extensions = [".txt", ".doc", ".docx", ".pdf"]

image_extensions = [".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp"]

def get_clean_name(filename):
    name, _ = os.path.splitext(filename)
    return name.strip().lower()

def is_ide(filename): return any(ide in get_clean_name(filename) for ide in ides)
def is_browser(filename): return any(browser in get_clean_name(filename) for browser in browsers)
def is_ai_software(filename): return any(ai in get_clean_name(filename) for ai in ai_softwares)
def is_dev_tool(filename): return any(tool in get_clean_name(filename) for tool in dev_tools)
def is_utility(filename): return any(u in get_clean_name(filename) for u in utilities)
def is_game(filename): return any(g in get_clean_name(filename) for g in games)
def is_script(filename): return any(filename.lower().endswith(ext) for ext in script_extensions)
def is_document(filename): return any(filename.lower().endswith(ext) for ext in document_extensions)
def is_image(filename): 
    return any(filename.lower().endswith(ext) for ext in image_extensions)

def is_shortcut_broken(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == ".lnk":
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(file_path)
            target_path = shortcut.TargetPath
            if not target_path:
                return False
            
            target_path = os.path.expandvars(target_path.strip('"'))
            return not os.path.exists(target_path)
        except Exception:
            return False

    elif ext == ".url":
        try:
            config = configparser.ConfigParser()
            config.read(file_path, encoding="utf-8")
            url = config.get("InternetShortcut", "URL", fallback="")
            
            if url.startswith("file:///"):
                local_path = url.replace("file:///", "").replace("/", "\\")
                return not os.path.exists(local_path)
            return False
        except Exception:
            return False
            
    return False

def get_windows_desktop_path():
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        )
        path, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        return os.path.expandvars(path)
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


def organize_file(file_path):
    if not os.path.exists(file_path):
        return

    item = os.path.basename(file_path)

    if item.lower() == "desktop.ini" or os.path.isdir(file_path):
        return
    if item.endswith((".tmp", ".crdownload", ".part")):
        return

    if item.lower().endswith((".lnk", ".url")) and is_shortcut_broken(file_path):
        try:
            os.remove(file_path)
            print(f"Removed broken shortcut: {item}")
            return
        except Exception as e:
            print(f"Failed to remove broken shortcut {item}: {e}")
            return

    time.sleep(1)

    folder_name = None
    if is_browser(item):
        folder_name = "Browsers"
    elif is_ide(item):
        folder_name = "IDEs"
    elif is_ai_software(item):
        folder_name = "AI Software"
    elif is_dev_tool(item):
        folder_name = "Dev Tools"
    elif is_utility(item):
        folder_name = "Utilities"
    elif is_game(item):
        folder_name = "Games"
    elif is_script(item):
        folder_name = "Scripts"
    elif is_document(item):
        folder_name = "Text_Documents"
    elif is_image(item):
        folder_name = "Images"

    if folder_name:
        user_desktop = get_windows_desktop_path()
        destination_folder = os.path.join(user_desktop, folder_name)
        os.makedirs(destination_folder, exist_ok=True)
        destination_path = os.path.join(destination_folder, item)

        max_retries = 5
        for attempt in range(max_retries):
            try:
                shutil.move(file_path, destination_path)
                print(f"Moved {item} -> {folder_name}")
                break
            except PermissionError:
                if attempt < max_retries - 1:
                    time.sleep(1.5)
                else:
                    print(f"File locked by another process: {item}")
            except Exception as e:
                print(f"Failed to move {item}: {e}")
                break

class DesktopHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            organize_file(event.src_path)

    def on_moved(self, event):
        if not event.is_directory:
            organize_file(event.dest_path)

def initial_sweep(desktop_paths):
    for path in desktop_paths:
        if os.path.exists(path):
            for item in os.listdir(path):
                organize_file(os.path.join(path, item))

if __name__ == "__main__":
    user_desktop = get_windows_desktop_path()
    public_desktop = os.path.join(os.environ.get("PUBLIC", r"C:\Users\Public"), "Desktop")

    initial_sweep([user_desktop, public_desktop])

    event_handler = DesktopHandler()
    observer = Observer()

    if os.path.exists(user_desktop):
        observer.schedule(event_handler, user_desktop, recursive=False)
    if os.path.exists(public_desktop):
        observer.schedule(event_handler, public_desktop, recursive=False)

    observer.start()
    print("Monitoring desktop for changes...")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
