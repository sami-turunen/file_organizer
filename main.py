import os
import shutil
import sys
import win32com.client
import winreg
import configparser

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

games = ["epic games launcher", "epic games"]

script_extensions = [".py", ".ps1", ".bat", ".sh", ".cmd", ".vbs"]

def get_clean_name(filename):
    name, _ = os.path.splitext(filename)
    return name.strip().lower()

def is_ide(filename):
    clean_name = get_clean_name(filename)
    return any(ide in clean_name for ide in ides)

def is_browser(filename):
    clean_name = get_clean_name(filename)
    return any(browser in clean_name for browser in browsers)

def is_ai_software(filename):
    clean_name = get_clean_name(filename)
    return any(ai in clean_name for ai in ai_softwares)

def is_dev_tool(filename):
    clean_name = get_clean_name(filename)
    return any(tool in clean_name for tool in dev_tools)

def is_utility(filename):
    clean_name = get_clean_name(filename)
    return any(u in clean_name for u in utilities)

def is_game(filename):
    clean_name = get_clean_name(filename)
    return any(g in clean_name for g in games)

def is_script(filename):
    return any(filename.lower().endswith(ext) for ext in script_extensions)

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

def main():
    print("Welcome to the desktop organizer!")
    user_desktop = get_windows_desktop_path()
    public_desktop = os.path.join(os.environ.get("PUBLIC", r"C:\Users\Public"), "Desktop")
    
    desktops = [user_desktop, public_desktop]

    for desktop_path in desktops:
        if not os.path.exists(desktop_path):
            continue

        for item in os.listdir(desktop_path):
            file_path = os.path.join(desktop_path, item)

            if item.lower() == "desktop.ini" or os.path.isdir(file_path):
                continue

            if item.lower().endswith((".lnk", ".url")) and is_shortcut_broken(file_path):
                print(f"Shortcut broken: {item}")
                try:
                    os.remove(file_path)
                    print(f"Removed {item}")
                except PermissionError:
                    print(f"Permission denied to remove {item}. Run terminal as Administrator.")
                continue

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

            if folder_name:
                destination_folder = os.path.join(user_desktop, folder_name)
                os.makedirs(destination_folder, exist_ok=True)
                destination_path = os.path.join(destination_folder, item)
                
                try:
                    shutil.move(file_path, destination_path)
                    print(f"Moved {item} -> {folder_name}")
                except PermissionError:
                    print(f"Permission denied to move {item}. Run terminal as Administrator.")
                except Exception as e:
                    print(f"Failed to move {item}: {e}")

    print("Desktop organization complete.")

if __name__ == "__main__":
    main()