import sys
import os

# このアンインストーラーはWindows専用です。
if sys.platform != "win32":
    print("エラー: このアンインストーラーはWindows専用です。", file=sys.stderr)
    from tkinter import messagebox, Tk
    root = Tk()
    root.withdraw()
    messagebox.showerror("エラー", "このアンインストーラーはWindows専用です。")
    sys.exit(1)

import shutil
import winreg
import subprocess
from tkinter import messagebox, Tk

# このスクリプトは、インストーラーによって設定されたレジストリから
# アプリケーションのGUIDとインストールパスを読み取り、アンインストールを実行します。

# --- 定数 (equa.py/pyraxis.pyと一致させる) ---
APP_PUBLISHER = "StudioNosa"
APP_NAME = "EQUA"
REG_UNINSTALL_PATH = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

def get_uninstall_info(app_guid):
    """レジストリからインストール情報を取得する"""
    # HKEY_LOCAL_MACHINE (管理者権限) と HKEY_CURRENT_USER (ユーザー単位) の両方を検索
    for hkey_root in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
        try:
            key_path = os.path.join(REG_UNINSTALL_PATH, app_guid)
            with winreg.OpenKey(hkey_root, key_path, 0, winreg.KEY_READ) as key:
                install_path, _ = winreg.QueryValueEx(key, "InstallLocation")
                display_name, _ = winreg.QueryValueEx(key, "DisplayName")
                publisher, _ = winreg.QueryValueEx(key, "Publisher")
                return install_path, display_name, publisher, hkey_root
        except FileNotFoundError:
            continue  # 見つからなければ次のHKEYを試す
        except Exception as e:
            messagebox.showerror("エラー", f"インストール情報の取得中に予期せぬエラーが発生しました。\n{e}")
            return None, None, None, None
    
    # どちらのHKEYにも見つからなかった場合
    messagebox.showerror("エラー", "アンインストール情報が見つかりません。アプリケーションが正しくインストールされていない可能性があります。")
    return None, None, None, None

def remove_user_data():
    """ユーザーデータフォルダを削除する"""
    try:
        local_app_data = os.getenv('LOCALAPPDATA')
        if not local_app_data:
            print("環境変数 'LOCALAPPDATA' が見つかりません。")
            return

        user_data_path = os.path.join(local_app_data, APP_PUBLISHER, APP_NAME)
        if os.path.isdir(user_data_path):
            print(f"ユーザーデータフォルダを削除しています: {user_data_path}")
            shutil.rmtree(user_data_path, ignore_errors=True)
        else:
            print(f"ユーザーデータフォルダが見つかりません: {user_data_path}")
    except Exception as e:
        # 失敗してもアンインストール処理は続行する
        print(f"ユーザーデータの削除中にエラーが発生しました: {e}")

def remove_files(install_path):
    """インストール先のファイルとフォルダを削除する"""
    if os.path.isdir(install_path):
        print(f"Deleting folder: {install_path}")
        shutil.rmtree(install_path, ignore_errors=True)

def remove_desktop_shortcut(display_name):
    """デスクトップのショートカットを削除する"""
    try:
        # WScript.Shellを使ってデスクトップパスを確実に取得
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        desktop = shell.SpecialFolders("Desktop")
        shortcut_path = os.path.join(desktop, f"{display_name}.lnk")
        if os.path.exists(shortcut_path):
            print(f"Deleting shortcut: {shortcut_path}")
            os.remove(shortcut_path)
    except Exception as e:
        print(f"Could not remove shortcut: {e}") # 失敗しても処理は続行

def remove_start_menu_shortcut(display_name, publisher):
    """スタートメニューのショートカットとフォルダを削除する"""
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        start_menu_folder = shell.SpecialFolders("AllUsersPrograms")
        publisher_folder = os.path.join(start_menu_folder, publisher)
        shortcut_path = os.path.join(publisher_folder, f"{display_name}.lnk")

        # ショートカットを削除
        if os.path.exists(shortcut_path):
            print(f"Deleting Start Menu shortcut: {shortcut_path}")
            os.remove(shortcut_path)
        
        # 発行元のフォルダが空なら削除
        if os.path.isdir(publisher_folder) and not os.listdir(publisher_folder):
            print(f"Deleting empty Start Menu folder: {publisher_folder}")
            os.rmdir(publisher_folder)
    except Exception as e:
        print(f"Could not remove Start Menu shortcut: {e}") # 失敗しても処理は続行

def remove_registry_entry(app_guid, hkey):
    """レジストリからアンインストール情報を削除する"""
    try:
        key_path = os.path.join(REG_UNINSTALL_PATH, app_guid)
        winreg.DeleteKey(hkey, key_path)
        hkey_name = "HKEY_LOCAL_MACHINE" if hkey == winreg.HKEY_LOCAL_MACHINE else "HKEY_CURRENT_USER"
        print(f"Deleted registry key: {hkey_name}\\{key_path}")
    except FileNotFoundError:
        print(f"Registry key not found, skipping deletion: {key_path}")
    except PermissionError:
        print(f"Permission denied to delete registry key: {key_path}")
    except Exception as e:
        print(f"Could not delete registry key: {e}") # 失敗しても処理は続行

def self_delete():
    """アンインストーラー自身を削除するバッチファイルを作成して実行する"""
    uninstaller_path = sys.executable
    batch_script = f"""
@echo off
chcp 65001 > nul
echo アンインストーラーのクリーンアップ...
timeout /t 2 /nobreak > nul
del "{uninstaller_path}"
del "%~f0"
"""
    script_path = os.path.join(os.path.dirname(uninstaller_path), 'cleanup.bat')
    with open(script_path, 'w', encoding='utf-8') as f:
        f.write(batch_script)
    
    subprocess.Popen(f'"{script_path}"', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)

def main():
    # GUIのルートウィンドウを非表示にする
    root = Tk()
    root.withdraw()

    if len(sys.argv) < 2:
        messagebox.showerror("エラー", "アンインストール情報が不足しています。")
        sys.exit(1)
    
    app_guid = sys.argv[1]

    install_path, display_name, publisher, hkey = get_uninstall_info(app_guid)
    if not all([install_path, display_name, publisher, hkey]):
        sys.exit(1)

    if messagebox.askyesno("アンインストール", f"「{display_name}」をコンピューターからアンインストールしますか？"):
        # ユーザーデータ削除の確認
        if messagebox.askyesno("ユーザーデータの削除", "ブックマーク、履歴、設定などのユーザーデータをすべて削除しますか？\n\nこの操作は元に戻せません。"):
            remove_user_data()

        remove_files(install_path)
        remove_desktop_shortcut(display_name)
        remove_start_menu_shortcut(display_name, publisher)
        remove_registry_entry(app_guid, hkey)
        messagebox.showinfo("完了", f"「{display_name}」のアンインストールが完了しました。")
        self_delete()
    
if __name__ == '__main__':
    main()