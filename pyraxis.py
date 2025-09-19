import sys
import os

# このインストーラーはWindows専用です。
if sys.platform != "win32":
    print("エラー: このインストーラーはWindows専用です。", file=sys.stderr)
    # GUI環境かもしれないので、GUIでのメッセージボックス表示を試みる
    try:
        from PyQt6.QtWidgets import QApplication, QMessageBox
        app = QApplication([])
        QMessageBox.critical(None, "エラー", "このインストーラーはWindows専用です。")
    except ImportError:
        pass # PyQt6がなければコンソール出力のみ
    sys.exit(1)

import shutil
import time
import uuid
import winreg

from PyQt6.QtWidgets import (
    QApplication, QWizard, QWizardPage, QVBoxLayout, QLabel, QLineEdit,
    QPushButton, QFileDialog, QProgressBar, QCheckBox, QTextEdit, QMessageBox
)
from PyQt6.QtCore import Qt, pyqtSignal, QThread

# win32comをインポート（ショートカット作成用）
# pip install pywin32
try:
    import win32com.client
    from win32com.client import Dispatch
    from win32com.client import pywintypes # COMエラーを捕捉するために追加
except ImportError:
    win32com = None
    pywintypes = None # pywintypesもNoneに設定

# --- 定数 ---
APP_NAME = "EQUA"
APP_VERSION = "0.1.0"  # equa.pyのバージョンと合わせる
APP_PUBLISHER = "StudioNosa" # equa.pyの発行元と合わせる
INSTALLER_NAME = "PyRAXIS Installer"
EXECUTABLE_NAME = "equa.exe"
ICON_NAME = "equa.ico"
UNINSTALLER_NAME = "uninstaller.exe"
# アプリケーション固有の固定GUID (名前ベースで生成)
# これにより、上書きインストール時に「プログラムと機能」に重複して登録されるのを防ぎます。
APP_GUID = str(uuid.uuid5(uuid.NAMESPACE_DNS, f'{APP_PUBLISHER}.{APP_NAME}'))
APP_GUID = f"{{{APP_GUID}}}"

def get_resource_path(relative_path):
    """リソースへのパスを取得する。PyInstallerでバンドルされた場合に対応。"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstallerの一時フォルダ
        base_path = os.path.join(sys._MEIPASS, "app")
    else:
        # 通常のPython環境（デバッグ用）
        base_path = os.path.join(os.path.abspath("."), "app")
    return os.path.join(base_path, relative_path)

def get_windows_theme():
    """Windowsの個人設定からアプリのテーマ（ライト/ダーク）を取得する"""
    if sys.platform != 'win32' or not winreg:
        return "ダーク"
    try:
        # HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize')
        # 'AppsUseLightTheme' の値を取得 (1: ライト, 0: ダーク)
        value, _ = winreg.QueryValueEx(key, 'AppsUseLightTheme')
        winreg.CloseKey(key)
        return "ライト" if value > 0 else "ダーク"
    except (FileNotFoundError, OSError):
        # レジストリキーが存在しない/読めない場合
        return "ダーク"

def cleanup_old_uninstall_entries():
    """古いバージョンのアンインストール情報をレジストリから検索して削除する"""
    uninstall_reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    current_fixed_guid = APP_GUID.strip('{}').lower()
    keys_to_delete = []

    # HKLMとHKCUの両方を検索
    for hkey_root in [winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER]:
        try:
            # 読み取り専用で開く
            with winreg.OpenKey(hkey_root, uninstall_reg_path, 0, winreg.KEY_READ) as uninstall_key:
                i = 0
                while True:
                    try:
                        guid = winreg.EnumKey(uninstall_key, i)
                        i += 1
                        
                        # 現在のインストーラーが使用する固定GUIDは削除対象外
                        if guid.lower() == current_fixed_guid:
                            continue

                        # 各GUIDキーを開いてDisplayNameを確認
                        with winreg.OpenKey(uninstall_key, guid, 0, winreg.KEY_READ) as app_key:
                            try:
                                display_name, _ = winreg.QueryValueEx(app_key, "DisplayName")
                                # DisplayNameが一致したら削除候補に追加
                                if display_name == APP_NAME:
                                    keys_to_delete.append((hkey_root, guid))
                            except FileNotFoundError:
                                # DisplayNameがないキーは無視
                                continue
                    except OSError:
                        # これ以上キーがない場合はループを抜ける
                        break
        except FileNotFoundError:
            # Uninstallキー自体が存在しない場合はスキップ
            continue

    # 収集した古いキーを削除
    if keys_to_delete:
        for hkey_root, guid in keys_to_delete:
            try:
                full_key_path = os.path.join(uninstall_reg_path, guid)
                winreg.DeleteKey(hkey_root, full_key_path)
            except (PermissionError, OSError) as e:
                # 権限エラーなどが発生しても、インストール処理は続行させる
                print(f"警告: 古いレジストリキーの削除に失敗しました: {guid}, {e}")

# --- スタイルシート ---
DARK_INSTALLER_STYLESHEET = """
    QWizard, QWizardPage {
        background-color: #2E3440; /* Nord Polar Night */
    }

    QLabel {
        color: #D8DEE9; /* Nord Snow Storm */
        font-size: 10pt;
    }

    QWizard::title {
        color: #88C0D0; /* Nord Frost */
        font-size: 18pt;
        font-weight: bold;
        padding: 15px 10px;
    }

    QWizard::subTitle {
        color: #E5E9F0;
        font-size: 10pt;
        padding-left: 10px;
        padding-bottom: 10px;
    }

    QCheckBox {
        color: #D8DEE9;
        font-size: 10pt;
        spacing: 8px;
    }

    QCheckBox::indicator {
        width: 16px;
        height: 16px;
        border: 1px solid #4C566A;
        border-radius: 4px; /* 丸み */
        background-color: #3B4252;
    }

    QCheckBox::indicator:hover {
        border: 1px solid #81A1C1;
    }

    QCheckBox::indicator:checked {
        background-color: #88C0D0;
        border: 1px solid #88C0D0;
    }

    QLineEdit, QTextEdit {
        background-color: #3B4252;
        color: #ECEFF4;
        border: 1px solid #4C566A;
        border-radius: 4px; /* 丸み */
        padding: 5px;
        font-size: 10pt;
    }

    #LicenseText {
        font-family: "Consolas", "Courier New", monospace;
    }

    QPushButton {
        background-color: #5E81AC; /* Nord Frost */
        color: #ECEFF4;
        border: none;
        padding: 8px 16px; /* 少し小さく */
        border-radius: 4px; /* 丸み */
        font-size: 9pt; /* 少し小さく */
        font-weight: bold;
    }

    QPushButton:hover {
        background-color: #81A1C1;
    }

    QPushButton:pressed {
        background-color: #88C0D0;
    }

    QPushButton:disabled {
        background-color: #4C566A;
        color: #6E7889;
    }

    QProgressBar {
        border: 1px solid #4C566A;
        border-radius: 5px; /* 丸み */
        text-align: center;
        color: #ECEFF4;
        background-color: #434C5E;
        max-height: 16px; /* 細くする */
        font-weight: bold;
        font-size: 9pt;
    }

    QProgressBar::chunk {
        background-color: #88C0D0; /* Nord Frost */
        border-radius: 4px; /* 丸み */
    }
"""

LIGHT_INSTALLER_STYLESHEET = """
    QWizard, QWizardPage {
        background-color: #ECEFF4; /* Nord Snow Storm */
    }

    QLabel {
        color: #2E3440; /* Nord Polar Night */
        font-size: 10pt;
    }

    QWizard::title {
        color: #5E81AC; /* Nord Frost */
        font-size: 18pt;
        font-weight: bold;
        padding: 15px 10px;
    }

    QWizard::subTitle {
        color: #4C566A;
        font-size: 10pt;
        padding-left: 10px;
        padding-bottom: 10px;
    }

    QCheckBox {
        color: #2E3440;
        font-size: 10pt;
        spacing: 8px;
    }

    QCheckBox::indicator {
        width: 16px;
        height: 16px;
        border: 1px solid #D8DEE9;
        border-radius: 4px;
        background-color: #E5E9F0;
    }

    QCheckBox::indicator:hover {
        border: 1px solid #81A1C1;
    }

    QCheckBox::indicator:checked {
        background-color: #5E81AC;
        border: 1px solid #5E81AC;
    }

    QLineEdit, QTextEdit {
        background-color: #FFFFFF;
        color: #2E3440;
        border: 1px solid #D8DEE9;
        border-radius: 4px;
        padding: 5px;
        font-size: 10pt;
    }

    #LicenseText {
        font-family: "Consolas", "Courier New", monospace;
    }

    QPushButton {
        background-color: #5E81AC; /* Nord Frost */
        color: #ECEFF4;
        border: none;
        padding: 8px 16px;
        border-radius: 4px;
        font-size: 9pt;
        font-weight: bold;
    }

    QPushButton:hover {
        background-color: #81A1C1;
    }

    QPushButton:pressed {
        background-color: #88C0D0;
    }

    QPushButton:disabled {
        background-color: #D8DEE9;
        color: #848E9F;
    }

    QProgressBar {
        border: 1px solid #D8DEE9;
        border-radius: 5px;
        text-align: center;
        color: #2E3440;
        background-color: #E5E9F0;
        max-height: 16px;
        font-weight: bold;
        font-size: 9pt;
    }

    QProgressBar::chunk {
        background-color: #88C0D0; /* Nord Frost */
        border-radius: 4px;
    }
"""

class InstallThread(QThread):
    """インストール処理を別スレッドで実行するためのクラス"""
    progress = pyqtSignal(int)
    finished = pyqtSignal(bool, str) # 成功/失敗、メッセージ

    def __init__(self, install_path, create_shortcut_enabled, create_start_menu_shortcut_enabled):
        super().__init__()
        self.install_path = install_path
        self.create_shortcut_enabled = create_shortcut_enabled
        self.create_start_menu_shortcut_enabled = create_start_menu_shortcut_enabled

    def run(self):
        try:
            # 古いアンインストール情報をクリーンアップ
            cleanup_old_uninstall_entries()
            self.progress.emit(5)

            # インストール先ディレクトリを作成
            os.makedirs(self.install_path, exist_ok=True)
            self.progress.emit(10)

            # インストールするファイルリスト
            files_to_install = [EXECUTABLE_NAME, ICON_NAME, UNINSTALLER_NAME]

            # ファイルをコピー
            for i, filename in enumerate(files_to_install):
                source_path = get_resource_path(filename)
                dest_path = os.path.join(self.install_path, filename)
                
                if not os.path.exists(source_path):
                    raise FileNotFoundError(f"ソースファイルが見つかりません: {source_path}")

                shutil.copy(source_path, dest_path)
                # 進捗更新: 10% (ベース) + 70% (ファイルコピー用) * (現在のファイルインデックス + 1) / 全ファイル数
                self.progress.emit(10 + int(70 * (i + 1) / len(files_to_install)))
                time.sleep(0.1) # 進捗を視覚的に見せるためのウェイトを短縮

            # レジストリにアンインストール情報を書き込む
            # この処理は管理者権限を要求します
            self.write_uninstall_info()
            self.progress.emit(90) # レジストリ書き込み完了

            # ショートカットを作成
            all_shortcuts_successful = True
            shortcut_error_messages = []

            if self.create_shortcut_enabled:
                if win32com:
                    success, msg = self.create_desktop_shortcut()
                    if not success:
                        all_shortcuts_successful = False
                        shortcut_error_messages.append(f"デスクトップ: {msg}")
                else:
                    all_shortcuts_successful = False
                    shortcut_error_messages.append("デスクトップ: 'pywin32' が見つかりません。")
            
            if self.create_start_menu_shortcut_enabled:
                if win32com:
                    success, msg = self.create_start_menu_shortcut()
                    if not success:
                        all_shortcuts_successful = False
                        shortcut_error_messages.append(f"スタートメニュー: {msg}")
                else:
                    all_shortcuts_successful = False
                    shortcut_error_messages.append("スタートメニュー: 'pywin32' が見つかりません。")
            
            self.progress.emit(100)

            final_installation_message = "インストールが正常に完了しました。"
            if not all_shortcuts_successful:
                errors = "\n".join(shortcut_error_messages)
                final_installation_message += f"\n\nただし、ショートカットの作成に一部失敗しました:\n{errors}"
            
            self.finished.emit(True, final_installation_message)

        except Exception as e:
            # ファイルコピーやディレクトリ作成など、主要なインストール処理でエラーが発生した場合
            self.finished.emit(False, f"インストール中に致命的なエラーが発生しました:\n{e}")

    def write_uninstall_info(self):
        """レジストリにアンインストール情報を書き込む"""
        try:
            # HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{GUID}
            reg_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            
            # Uninstallキーを開く (なければ作成)
            key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
            
            # アプリケーション用のキーを作成
            app_key = winreg.CreateKey(key, APP_GUID)

            # 各値を設定
            winreg.SetValueEx(app_key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
            winreg.SetValueEx(app_key, "DisplayVersion", 0, winreg.REG_SZ, APP_VERSION)
            winreg.SetValueEx(app_key, "Publisher", 0, winreg.REG_SZ, APP_PUBLISHER)
            winreg.SetValueEx(app_key, "InstallLocation", 0, winreg.REG_SZ, self.install_path)
            winreg.SetValueEx(app_key, "DisplayIcon", 0, winreg.REG_SZ, os.path.join(self.install_path, ICON_NAME))
            uninstall_command = f'"{os.path.join(self.install_path, UNINSTALLER_NAME)}" "{APP_GUID}"'
            winreg.SetValueEx(app_key, "UninstallString", 0, winreg.REG_SZ, uninstall_command)
            winreg.SetValueEx(app_key, "NoModify", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(app_key, "NoRepair", 0, winreg.REG_DWORD, 1)

        except PermissionError:
            raise PermissionError("レジストリへの書き込みに失敗しました。管理者としてインストーラーを実行してください。")

    def create_desktop_shortcut(self):
        """デスクトップにショートカットを作成する"""
        try:
            shell = Dispatch('WScript.Shell')
            # 環境変数ではなく、WScript.Shellから直接デスクトップパスを取得する方が、
            # 管理者権限での実行時などでも安定して動作します。
            desktop = shell.SpecialFolders("Desktop")
            shortcut_path = os.path.join(desktop, f"{APP_NAME}.lnk")
            target_path = os.path.join(self.install_path, EXECUTABLE_NAME)
            icon_path = os.path.join(self.install_path, ICON_NAME)

            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = target_path
            shortcut.IconLocation = icon_path
            shortcut.WorkingDirectory = self.install_path
            shortcut.save()
            return True, "" # 成功
        except pywintypes.com_error as e:
            # COMエラーの詳細を抽出
            error_message = e.args[2][2] if len(e.args) > 2 and len(e.args[2]) > 2 else str(e)
            return False, f"ショートカットの保存に失敗しました: {error_message}"
        except Exception as e:
            return False, f"予期せぬエラーによりショートカットの作成に失敗しました: {e}"

    def create_start_menu_shortcut(self):
        """スタートメニューにショートカットを作成する"""
        try:
            shell = Dispatch('WScript.Shell')
            # AllUsersPrograms は全ユーザーのスタートメニュー
            start_menu_folder = shell.SpecialFolders("AllUsersPrograms")
            publisher_folder = os.path.join(start_menu_folder, APP_PUBLISHER)
            
            # 発行元名のフォルダを作成
            os.makedirs(publisher_folder, exist_ok=True)

            shortcut_path = os.path.join(publisher_folder, f"{APP_NAME}.lnk")
            target_path = os.path.join(self.install_path, EXECUTABLE_NAME)
            icon_path = os.path.join(self.install_path, ICON_NAME)

            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = target_path
            shortcut.IconLocation = icon_path
            shortcut.WorkingDirectory = self.install_path
            shortcut.save()
            return True, "" # 成功
        except pywintypes.com_error as e:
            error_message = e.args[2][2] if len(e.args) > 2 and len(e.args[2]) > 2 else str(e)
            return False, f"ショートカットの保存に失敗しました: {error_message}"
        except Exception as e:
            return False, f"予期せぬエラーによりショートカットの作成に失敗しました: {e}"


class IntroPage(QWizardPage):
    """ウェルカムページ"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTitle(f"ようこそ {APP_NAME} のインストーラーへ")
        self.setSubTitle(f"このウィザードは {APP_NAME} をお使いのコンピューターにインストールします。")

        layout = QVBoxLayout()
        label = QLabel(f"{APP_NAME} のインストールを開始するには「Next」をクリックしてください。")
        label.setWordWrap(True)
        layout.addWidget(label)
        self.setLayout(layout)


class LicensePage(QWizardPage):
    """ライセンス同意ページ"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTitle("ライセンス契約")
        self.setSubTitle("インストールを続行する前に、ライセンス契約の条項をお読みください。")

        layout = QVBoxLayout()
        
        self.license_text = QTextEdit()
        self.license_text.setObjectName("LicenseText")
        self.license_text.setReadOnly(True)
        layout.addWidget(self.license_text)

        self.agree_checkbox = QCheckBox("ライセンス契約の条項に同意します。")
        # isCompleteを再評価するためにシグナルを接続
        self.agree_checkbox.toggled.connect(self.completeChanged.emit)
        layout.addWidget(self.agree_checkbox)

        self.setLayout(layout)

    def initializePage(self):
        """ページが表示されるときにライセンスファイルを読み込む"""
        license_path = get_resource_path('license.txt')
        try:
            with open(license_path, 'r', encoding='utf-8') as f:
                self.license_text.setText(f.read())
        except FileNotFoundError:
            self.license_text.setText("エラー: app/license.txt が見つかりませんでした。")
        
        # 初期状態ではチェックボックスはオフ
        self.agree_checkbox.setChecked(False)

    def isComplete(self):
        """「次へ」ボタンを有効にする条件"""
        return self.agree_checkbox.isChecked()


class InstallPathPage(QWizardPage):
    """インストール先選択ページ"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTitle("インストール先の選択")
        self.setSubTitle("プログラムをインストールするフォルダーを選択してください。")

        layout = QVBoxLayout()
        
        self.path_edit = QLineEdit()
        browse_button = QPushButton("参照...")
        browse_button.clicked.connect(self.browse)
        
        self.shortcut_checkbox = QCheckBox("デスクトップにショートカットを作成する")
        self.shortcut_checkbox.setChecked(True)

        self.start_menu_checkbox = QCheckBox("スタートメニューにショートカットを作成する")
        self.start_menu_checkbox.setChecked(True)

        path_layout = QVBoxLayout()
        path_layout.addWidget(QLabel("インストール先フォルダー:"))
        path_layout.addWidget(self.path_edit)
        path_layout.addWidget(browse_button, alignment=Qt.AlignmentFlag.AlignLeft)
        path_layout.addSpacing(20)
        path_layout.addWidget(self.shortcut_checkbox)
        path_layout.addWidget(self.start_menu_checkbox)

        layout.addLayout(path_layout)
        self.setLayout(layout)

        self.registerField("installPath*", self.path_edit)
        self.registerField("createShortcut", self.shortcut_checkbox)
        self.registerField("createStartMenuShortcut", self.start_menu_checkbox)

    def initializePage(self):
        default_path = os.path.join(os.environ.get("ProgramFiles", "C:\\Program Files"), APP_NAME)
        self.path_edit.setText(default_path)

    def browse(self):
        directory = QFileDialog.getExistingDirectory(self, "フォルダーを選択", self.path_edit.text())
        if directory:
            self.path_edit.setText(directory)

    def validatePage(self):
        """ページを離れる前に検証し、既存のインストールを確認する"""
        install_path = self.field("installPath")
        executable_path = os.path.join(install_path, EXECUTABLE_NAME)

        # 既存のインストールをチェック
        if os.path.exists(executable_path):
            reply = QMessageBox.question(
                self,
                "インストールの確認",
                f"指定されたフォルダーには既に '{APP_NAME}' がインストールされているようです。\n\n"
                "既存のバージョンを上書きしてインストールを続行しますか？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return False  # インストールを中断し、このページに留まる
        
        return True # インストールを続行


class InstallPage(QWizardPage):
    """インストール実行ページ"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTitle("インストール中")
        self.setSubTitle(f"{APP_NAME} をインストールしています。しばらくお待ちください...")

        layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        self.status_label = QLabel("準備中...")
        self.status_label.setWordWrap(True)
        
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.status_label)
        self.setLayout(layout)

        # setField/fieldでプログラム的に値を設定するフィールドを登録します。
        # PyQtではウィジェットとの関連付けが必須なため、ダミーのウィジェットを作成して登録します。
        self._success_holder = QCheckBox()
        self._message_holder = QLineEdit()

        self.registerField("installSuccess", self._success_holder)
        self.registerField("installMessage", self._message_holder)

    def initializePage(self):
        self.setCommitPage(True)
        self.wizard().setButtonLayout([]) # 次へ/戻るボタンを非表示

        self.progress_bar.setValue(0)
        install_path = self.field("installPath")
        create_shortcut = self.field("createShortcut")
        create_start_menu_shortcut = self.field("createStartMenuShortcut") 

        self.thread = InstallThread(install_path, create_shortcut, create_start_menu_shortcut)
        self.thread.progress.connect(self.progress_bar.setValue)
        self.thread.finished.connect(self.on_finished)
        self.thread.start()

    def on_finished(self, success, message):
        self.setField("installSuccess", success)
        self.setField("installMessage", message)
        self.status_label.setText(message)
        self.wizard().next()


class FinishPage(QWizardPage):
    """完了ページ"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setTitle("インストール完了")
        
        layout = QVBoxLayout()
        self.message_label = QLabel()
        self.message_label.setWordWrap(True)
        layout.addWidget(self.message_label)

        # アプリケーション起動チェックボックスを追加
        self.launch_checkbox = QCheckBox(f"{APP_NAME} を起動する")
        self.launch_checkbox.setChecked(True)
        layout.addSpacing(20)
        layout.addWidget(self.launch_checkbox)

        self.setLayout(layout)

        # フィールドを登録
        self.registerField("launchAppOnFinish", self.launch_checkbox)

    def initializePage(self):
        # InstallPageで非表示にしたボタンを再表示する
        self.wizard().setButtonLayout([
            QWizard.WizardButton.Stretch,
            QWizard.WizardButton.FinishButton
        ])

        if self.field("installSuccess"):
            self.setSubTitle(f"{APP_NAME} のインストールが完了しました。")
        else:
            self.setSubTitle("インストールに失敗しました。")
        self.message_label.setText(self.field("installMessage"))


class InstallerWizard(QWizard):
    """インストーラーウィザード本体"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(INSTALLER_NAME)

        self.addPage(IntroPage())
        self.addPage(LicensePage())
        self.addPage(InstallPathPage())
        self.addPage(InstallPage())
        self.addPage(FinishPage())

        # ボタンのテキストを英語 ("Next" など) に統一し、表記揺れをなくす
        self.setButtonText(QWizard.WizardButton.BackButton, "< Back")
        self.setButtonText(QWizard.WizardButton.NextButton, "Next >")
        self.setButtonText(QWizard.WizardButton.CommitButton, "Install")
        self.setButtonText(QWizard.WizardButton.FinishButton, "Finish")
        self.setButtonText(QWizard.WizardButton.CancelButton, "Cancel")

        self.setWizardStyle(QWizard.WizardStyle.ModernStyle)
        self.setOption(QWizard.WizardOption.NoBackButtonOnLastPage)

    def accept(self):
        """完了ボタンが押されたときの処理"""
        # チェックボックスがオンの場合、アプリケーションを起動する
        if self.field("launchAppOnFinish"):
            try:
                install_path = self.field("installPath")
                executable_path = os.path.join(install_path, EXECUTABLE_NAME)
                if os.path.exists(executable_path):
                    import subprocess
                    subprocess.Popen([executable_path], cwd=install_path)
            except Exception as e:
                print(f"アプリケーションの起動に失敗しました: {e}")
        
        super().accept() # ウィザードを閉じる


if __name__ == '__main__':
    app = QApplication(sys.argv)
    wizard = InstallerWizard()

    # OSのテーマを判定してスタイルシートを適用
    theme = get_windows_theme()
    if theme == "ライト":
        wizard.setStyleSheet(LIGHT_INSTALLER_STYLESHEET)
    else:
        wizard.setStyleSheet(DARK_INSTALLER_STYLESHEET)

    wizard.show()
    sys.exit(app.exec())