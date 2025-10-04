# -*- coding: utf-8 -*-

# 必要なモジュールをインポート
import sys
import os
import json
import urllib.request
import urllib.parse
import re
import sqlite3
try: # winregはWindows専用モジュールなので、他のOSでエラーにならないようにする
    import winreg
except ImportError:
    winreg = None # Windows以外のOS用のフォールバック
from packaging.version import parse as parse_version, InvalidVersion # バージョン番号の比較に使用
from datetime import datetime # 日時情報の扱いに使用
from html.parser import HTMLParser # HTMLの解析に使用 (ブックマークインポート)
import qtawesome as qta # Font Awesomeアイコンを使用するためのライブラリ
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QLineEdit, QTabWidget, QInputDialog, QGroupBox, QStyle, QProxyStyle, QStyleOptionTab, QStyleFactory,
    QWidget, QSizePolicy, QHBoxLayout, QPushButton, QListWidget, QDialog, QVBoxLayout, QMessageBox,
    QLabel, QListWidgetItem, QMenu, QFileDialog, QProgressBar, QScrollArea, QColorDialog, QComboBox,
    QStackedWidget, QCheckBox
)
from PyQt6.QtWebEngineWidgets import QWebEngineView # ウェブページを表示するためのウィジェット
from PyQt6.QtWebEngineCore import QWebEngineSettings, QWebEngineProfile, QWebEngineUrlRequestInterceptor, QWebEnginePage
from PyQt6.QtCore import QUrl, QSettings, Qt, QStandardPaths, QSize, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QIcon, QCloseEvent, QAction, QDesktopServices, QPixmap, QColor

# アプリケーションのバージョンとGitHubリポジトリ情報
__version__ = "0.2.0"
GITHUB_REPO_OWNER = "Keychrom"
GITHUB_REPO_NAME = "Project-EQUA-Portable"

# --- ポータブル化対応 ---
def get_portable_base_path():
    """ポータブルモード用のベースパス（実行ファイルのあるディレクトリ）を取得する。"""
    if hasattr(sys, 'frozen'):
        # PyInstallerでexe化された場合
        return os.path.dirname(sys.executable)
    # スクリプトとして実行された場合
    return os.path.dirname(os.path.abspath(__file__))

PORTABLE_BASE_PATH = get_portable_base_path()

# --- 定数 ---
SETTINGS_FILE_NAME = "settings.ini"
DATA_DIR_NAME = "data"
DEFAULT_ADBLOCK_LIST_URL = "https://easylist.to/easylist/easylist.txt"

# PyInstallerで作成されたexeファイル内でリソースファイル（アイコンなど）のパスを解決するためのヘルパー関数
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller/Nuitka """
    # PyInstaller/Nuitkaでバンドルされているかチェック
    if getattr(sys, 'frozen', False):
        if hasattr(sys, '_MEIPASS'):
            # PyInstallerは一時フォルダを作成し、そのパスを_MEIPASSに格納します
            base_path = os.path.join(sys._MEIPASS, 'app')
        else:
            base_path = os.path.join(sys.executable, 'app')
    else:
        # 開発環境（.pyを直接実行）の場合
        base_path = os.path.join(os.path.abspath("."), 'app')
    return os.path.join(base_path, relative_path)

# 開いているウィンドウの参照を保持するグローバルリスト
# これにより、ウィンドウがスコープ外に出てもガベージコレクションされなくなる
windows = []
# 永続プロファイルを保持するためのグローバル変数
persistent_profile = None # Cookieやキャッシュなどを保持するプロファイル

# 広告ブロック用リクエストインターセプター
class AdBlockInterceptor(QWebEngineUrlRequestInterceptor):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ad_domains = set()
        # ad_block_list.txtのパスを決定 (ポータブル化)
        data_path = os.path.join(PORTABLE_BASE_PATH, DATA_DIR_NAME)
        self.block_list_path = os.path.join(data_path, 'ad_block_list.txt') # ブロックリストのファイルパス
        self.load_domains() # ドメインリストを読み込む

    def load_domains(self):
        """ブロックリストファイルからドメインを読み込む。ファイルがなければデフォルト値で作成する"""
        self.ad_domains.clear()
        try:
            if not os.path.exists(self.block_list_path):
                # ファイルが存在しない場合、デフォルトのリストで作成
                with open(self.block_list_path, 'w', encoding='utf-8') as f:
                    default_domains = [
                        "! EasyList系の基本的な広告ドメイン",
                        "doubleclick.net",
                        "googlesyndication.com",
                        "googleadservices.com",
                        "adservice.google.com",
                        "adnxs.com",
                        "scorecardresearch.com",
                        "criteo.com",
                        "pubmatic.com",
                        "ad.yieldmanager.com",
                        "adform.net"
                    ]
                    f.write('\n'.join(default_domains) + '\n')

            with open(self.block_list_path, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    # コメント行、空行、セクションマーカー、例外ルール(@@)を無視
                    if not line or line.startswith(('!', '[', '#', '@@')):
                        continue

                    # EasyListのドメイン指定ルール (例: ||example.com^) を簡易的にパース
                    if line.startswith('||'): # ドメイン指定ルール
                        # '||' を取り除き、オプション部分('^'以降)を分離
                        domain_part = line[2:].split('^', 1)[0]

                        # この簡易ブロッカーはパス指定ルール('/'を含む)をサポートしないため、無視する。
                        # これにより、"||google.com/ads/" のようなルールが "google.com" として誤って解釈されるのを防ぐ。
                        if '/' in domain_part:
                            continue
                        
                        # ドメイン名として妥当か簡単なチェック（先頭のワイルドカードとドットは除去）
                        domain = domain_part.lstrip('*.')
                        if '.' in domain and ' ' not in domain:
                            self.ad_domains.add(domain)
                    # その他の単純なドメイン指定も考慮 (後方互換性のため)
                    # '/'を含まない、' 'を含まない、'.'を含むものをドメインとみなす
                    elif '/' not in line and ' ' not in line and '.' in line:
                        self.ad_domains.add(line)
        except Exception as e:
            print(f"広告ブロックリストの読み込みに失敗しました: {e}")

    def interceptRequest(self, info):
        """ウェブページからのリクエストをインターセプト(傍受)する"""
        url_host = info.requestUrl().host()
        # ドメインリストと後方一致でマッチングを行う
        for domain in self.ad_domains:
            # url_hostがブロック対象ドメインそのものであるか、
            # またはブロック対象ドメインのサブドメインであるかをチェック
            # 例: domain="example.com" の場合、"example.com" と "sub.example.com" にマッチ
            if url_host == domain or url_host.endswith('.' + domain):
                info.block(True)
                return # 一致したらブロックして終了

# 広告ブロックリストを非同期で更新するためのワーカースレッド
class UpdateBlocklistThread(QThread):
    # シグナル: 処理完了時に (成功/失敗, コンテンツ, エラーメッセージ) を送信
    finished = pyqtSignal(bool, str, str)  # success, content, error_message

    def __init__(self, url, parent=None):
        super().__init__(parent)
        self.url = url

    # スレッドのメイン処理
    def run(self):
        try:
            req = urllib.request.Request(
                self.url,
                headers={'User-Agent': 'Mozilla/5.0'}  # 403 Forbiddenを避けるため
            )
            with urllib.request.urlopen(req, timeout=30) as response:
                if response.status == 200:
                    content = response.read().decode('utf-8', errors='ignore')
                    self.finished.emit(True, content, "") # 成功シグナルを送信
                else:
                    self.finished.emit(False, "", f"サーバーエラー: {response.status}") # 失敗シグナルを送信
        except Exception as e:
            self.finished.emit(False, "", str(e)) # 例外発生時に失敗シグナルを送信

# GitHubリリースを非同期でチェックするためのワーカースレッド
class UpdateCheckThread(QThread):
    # シグナル: 処理完了時に (成功/失敗, 最新バージョン, リリースURL, アセットURL, エラーメッセージ) を送信
    finished = pyqtSignal(bool, str, str, str, str)  # success, latest_version, release_url, asset_url, error_message

    def __init__(self, owner, repo, parent=None):
        super().__init__(parent)
        self.owner = owner
        self.repo = repo

    # スレッドのメイン処理
    def run(self):
        try:
            # /releases/latest はプレリリースを無視するため、/releases を使用して最新のリリース（プレリリースを含む）を取得します
            url = f"https://api.github.com/repos/{self.owner}/{self.repo}/releases"
            req = urllib.request.Request(
                url,
                headers={'User-Agent': 'EQUA-Update-Checker', 'Accept': 'application/vnd.github.v3+json'}
            )
            with urllib.request.urlopen(req, timeout=10) as response:
                if response.status == 200:
                    releases = json.loads(response.read().decode('utf-8'))
                    if releases:  # リリースのリストが空でないことを確認
                        latest_release = releases[0]  # 最初の要素が最新のリリース
                        latest_version = latest_release.get('tag_name', '')
                        release_url = latest_release.get('html_url', '')
                        asset_url = ""
                        # OSに応じたアセットを探す
                        asset_pattern = None
                        if sys.platform == 'win32':
                            # ポータブル版のzipファイルを探す
                            asset_pattern = re.compile(r'EQUA-Portable\.exe$', re.IGNORECASE)
                        # 他のOS用のアセットも必要ならここに追加

                        for asset in latest_release.get('assets', []):
                            if asset_pattern and asset_pattern.search(asset.get('name', '')):
                                asset_url = asset.get('browser_download_url')
                                break # 最初に見つかったものを採用
                        self.finished.emit(True, latest_version, release_url, asset_url, "")
                    else: # リリースが一つもない場合
                        self.finished.emit(True, "", "", "", "") # 成功として扱うが、バージョン情報なし
                else:
                    self.finished.emit(False, "", "", "", f"サーバーエラー: {response.status}\nURL: {url}")
        except Exception as e:
            # エラーメッセージに、どのURLで失敗したかを含める
            url_for_error = f"https://api.github.com/repos/{self.owner}/{self.repo}/releases"
            self.finished.emit(False, "", "", "", f"{e}\nURL: {url_for_error}")

# --- テーマ関連の定数 ---
THEME_COLORS = {
    "ダーク": {
        "icon_color": "#D8DEE9",
        "icon_active_color": "#88C0D0",
        "progress_bar_color": "#88C0D0",
    },
    "ライト": {
        "icon_color": "#4C566A",
        "icon_active_color": "#5E81AC",
        "progress_bar_color": "#81A1C1",
    }
}

# ダークテーマのスタイルシート (Nord)
DARK_STYLESHEET = """ 
/* 全体のフォントと基本カラー */
QWidget {
    background-color: #2E3440; /* Nord Polar Night */
    color: #D8DEE9; /* Nord Snow Storm */
    font-family: "Segoe UI", "Meiryo", "sans-serif";
    font-size: 10pt;
    border: none;
}

/* メインウィンドウとダイアログ */
QMainWindow, QDialog {
    background-color: #3B4252; /* Nord Polar Night */
}

/* ナビゲーションバー */
#NavigationBar {
    background-color: #3B4252;
    border-bottom: 1px solid #4C566A;
    padding: 2px;
}

#NavigationBar QLineEdit {
    background-color: #2E3440;
    color: #ECEFF4;
    border: 1px solid #4C566A;
    border-radius: 13px;
    padding: 3px 12px; /* 上下のパディングをさらに縮小 */
    font-size: 10pt;
}

#NavigationBar QLineEdit:focus {
    border: 1px solid #88C0D0; /* Nord Frost */
}

#NavigationBar QPushButton {
    background-color: transparent;
    border: none;
    padding: 4px;
    border-radius: 13px;
    width: 26px;
    height: 26px;
}

#NavigationBar QPushButton:hover {
    background-color: #4C566A;
}

#NavigationBar QPushButton:pressed {
    background-color: #434C5E;
}

/* タブウィジェット */
QTabWidget::pane:top {
    /* コンテンツエリアの上部に境界線を引きます。これが「謎の線」の正体です。 */
    border-top: 1px solid #4C566A;
}
QTabBar:top {
    /* タブバー自体の下の境界線は不要です。 */
    border-bottom: none;
}
QTabWidget::pane:left {
    border-left: 1px solid #4C566A;
}
QTabWidget::pane:right {
    border-right: 1px solid #4C566A;
}

QTabBar:left {
    border-right: 1px solid #434C5E;
}
QTabBar:right {
    border-left: 1px solid #434C5E;
}
QTabBar:bottom {
    border-top: none;
}

QTabBar::tab:top {
    background: #3B4252;
    color: #D8DEE9;
    padding: 7px 20px; /* 高さを調整して、下の境界線がはみ出さないようにする */
    border: 1px solid transparent;
    border-top-left-radius: 5px;
    border-top-right-radius: 5px;
    margin-right: 1px; /* タブ間の区切り */
    min-width: 180px; /* タブの最小幅 */
    max-width: 220px; /* タブの最大幅 */
}

QTabBar::tab:bottom {
    background: #3B4252;
    color: #D8DEE9;
    padding: 7px 20px;
    border: 1px solid transparent;
    border-bottom-left-radius: 5px;
    border-bottom-right-radius: 5px;
    margin-right: 1px;
    min-width: 180px;
    max-width: 220px;
}

QTabBar::tab:left, QTabBar::tab:right {
    background: #3B4252;
    color: #D8DEE9;
    padding: 8px 12px;
    border: 1px solid transparent;
    margin: 0 1px 1px 1px; /* 上下左右の余白 */
    border-radius: 5px;
    min-width: 40px; /* タブの高さ */
}

QTabBar::tab:hover { background: #434C5E; }

/* 共通の選択タブスタイル */
QTabBar::tab:selected {
    background: #2E3440; /* ウィジェット背景と同じにして一体感を出す */
    color: #ECEFF4;
    border: 1px solid #4C566A; /* 境界線を基本として定義 */
}

/* 上タブが選択された場合: 下の境界線を消す */
QTabBar::tab:selected:top {
    /* 選択されたタブを1pxだけ下にずらし、paneの上部境界線の上に重ねて隠します。 */
    margin-bottom: -1px; 
    /* 下の境界線を、スタイルも含めて背景色と完全に一致させることで、線を消します。 */
    border-bottom: 1px solid #2E3440;
}

/* 左タブが選択された場合: 右の境界線を消す */
QTabBar::tab:selected:left {
    margin-right: -1px;
    border-right-color: #2E3440;
}

/* 右タブが選択された場合: 左の境界線を消す */
QTabBar::tab:selected:right {
    margin-left: -1px;
    border-left-color: #2E3440;
}

/* 下タブが選択された場合: 上の境界線を消す */
QTabBar::tab:selected:bottom {
    /* 選択されたタブを1pxだけ上にずらし、paneの下部境界線の上に重ねて隠します。 */
    margin-top: -1px;
    /* 上の境界線を、スタイルも含めて背景色と完全に一致させることで、線を消します。 */
    border-top: 1px solid #2E3440;
}

QTabBar::close-button {
    margin: 2px;
}
QTabBar::close-button:hover {
    background: #4C566A;
    border-radius: 8px;
}

/* メニュー */
QMenu {
    background-color: #3B4252;
    border: 1px solid #4C566A;
    padding: 5px;
}

QMenu::item {
    padding: 8px 25px;
    border-radius: 4px;
}

QMenu::item:selected {
    background-color: #81A1C1; /* Nord Frost */
}

QMenu::separator {
    height: 1px;
    background: #4C566A;
    margin: 5px 0;
}

/* メッセージボックス */
QMessageBox {
    background-color: #3B4252; /* Nord Polar Night */
}

QMessageBox QLabel {
    background-color: transparent;
    color: #D8DEE9;
    font-size: 10pt;
}

/* ボタン */
QPushButton {
    background-color: #5E81AC; /* Nord Frost */
    color: #ECEFF4;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #81A1C1;
}

QPushButton:pressed {
    background-color: #88C0D0;
}

/* リストウィジェット */
QListWidget {
    background-color: #2E3440;
    border: 1px solid #4C566A;
    border-radius: 4px;
    padding: 2px;
}

QListWidget::item {
    padding: 10px;
    border-radius: 3px;
}

QListWidget::item:hover {
    background-color: #434C5E;
}

QListWidget::item:selected {
    background-color: #5E81AC;
    color: #ECEFF4;
}

/* グループボックス */
QGroupBox {
    border: 1px solid #434C5E; /* 少し明るい色に */
    border-radius: 8px;
    margin-top: 15px;
    padding: 20px 15px 15px 15px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 10px;
    left: 10px;
    color: #88C0D0;
    background-color: #2E3440; /* 背景色で線を途切れさせる */
    font-weight: bold;
}

/* テキスト入力 */
QLineEdit {
    background-color: #434C5E;
    border: 1px solid #4C566A;
    padding: 6px;
    border-radius: 4px;
}

/* プログレスバー */
QProgressBar {
    border: 1px solid #4C566A;
    border-radius: 5px;
    text-align: center;
    color: #ECEFF4;
    background-color: #434C5E;
}

QProgressBar::chunk {
    background-color: #88C0D0; /* Nord Frost */
    border-radius: 4px;
}

/* スクロールバー */
QScrollBar:vertical { border: none; background: #2E3440; width: 12px; margin: 0; }
QScrollBar::handle:vertical { background: #4C566A; min-height: 25px; border-radius: 6px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; }
QScrollBar:horizontal { border: none; background: #2E3440; height: 12px; margin: 0; }
QScrollBar::handle:horizontal { background: #4C566A; min-width: 25px; border-radius: 6px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: none; }

/* 設定ダイアログのカテゴリリスト */
#SettingsCategories {
    background-color: #3B4252;
    border-right: 1px solid #434C5E;
}

#SettingsCategories::item {
    padding: 12px 20px;
    border-left: 3px solid transparent; /* 未選択時は透明 */
}

#SettingsCategories::item:hover {
    background-color: #434C5E;
}

#SettingsCategories::item:selected {
    background-color: #2E3440; /* 右ペインと一体化 */
    color: #ECEFF4; /* テキスト色を白に */
    border-left: 3px solid #88C0D0;
}
"""

# ライトテーマのスタイルシート (Nord)
LIGHT_STYLESHEET = """
/* 全体のフォントと基本カラー */
QWidget {
    background-color: #ECEFF4; /* Nord Snow Storm */
    color: #2E3440; /* Nord Polar Night */
    font-family: "Segoe UI", "Meiryo", "sans-serif";
    font-size: 10pt;
    border: none;
}

/* メインウィンドウとダイアログ */
QMainWindow, QDialog {
    background-color: #E5E9F0;
}

/* ナビゲーションバー */
#NavigationBar {
    background-color: #E5E9F0;
    border-bottom: 1px solid #D8DEE9;
    padding: 2px;
}

#NavigationBar QLineEdit {
    background-color: #FFFFFF;
    color: #2E3440;
    border: 1px solid #D8DEE9;
    border-radius: 13px;
    padding: 3px 12px;
    font-size: 10pt;
}

#NavigationBar QLineEdit:focus {
    border: 1px solid #5E81AC; /* Nord Frost */
}

#NavigationBar QPushButton {
    background-color: transparent;
    border: none;
    padding: 4px;
    border-radius: 13px;
    width: 26px;
    height: 26px;
}

#NavigationBar QPushButton:hover {
    background-color: #D8DEE9;
}

#NavigationBar QPushButton:pressed {
    background-color: #BFC9DB;
}

/* タブウィジェット */
QTabWidget::pane:top {
    /* コンテンツエリアの上部に境界線を引きます。これが「謎の線」の正体です。 */
    border-top: 1px solid #D8DEE9;
}
QTabBar:top {
    /* タブバー自体の下の境界線は不要です。 */
    border-bottom: none;
}
QTabWidget::pane:left {
    border-left: 1px solid #D8DEE9;
}
QTabWidget::pane:right {
    border-right: 1px solid #D8DEE9;
}

QTabBar:left {
    border-right: 1px solid #E5E9F0;
}
QTabBar:right {
    border-left: 1px solid #E5E9F0;
}
QTabBar:bottom {
    border-top: none;
}

QTabBar::tab:top {
    background: #E5E9F0;
    color: #4C566A;
    padding: 7px 20px; /* 高さを調整して、下の境界線がはみ出さないようにする */
    border: 1px solid transparent;
    border-top-left-radius: 5px;
    border-top-right-radius: 5px;
    margin-right: 1px; /* タブ間の区切り */
    min-width: 180px; /* タブの最小幅 */
    max-width: 220px; /* タブの最大幅 */
}

QTabBar::tab:bottom {
    background: #E5E9F0;
    color: #4C566A;
    padding: 7px 20px;
    border: 1px solid transparent;
    border-bottom-left-radius: 5px;
    border-bottom-right-radius: 5px;
    margin-right: 1px;
    min-width: 180px;
    max-width: 220px;
}

QTabBar::tab:left, QTabBar::tab:right {
    background: #E5E9F0;
    color: #4C566A;
    padding: 8px 12px;
    border: 1px solid transparent;
    margin: 0 1px 1px 1px; /* 上下左右の余白 */
    border-radius: 5px;
    min-width: 40px; /* タブの高さ */
}

QTabBar::tab:hover { background: #D8DEE9; }

/* 共通の選択タブスタイル */
QTabBar::tab:selected {
    background: #ECEFF4; /* ウィジェット背景と同じ */
    color: #2E3440;
    border: 1px solid #D8DEE9; /* 境界線を基本として定義 */
}

/* 上タブが選択された場合: 下の境界線を消す */
QTabBar::tab:selected:top {
    /* 選択されたタブを1pxだけ下にずらし、paneの上部境界線の上に重ねて隠します。 */
    margin-bottom: -1px;
    /* 下の境界線を、スタイルも含めて背景色と完全に一致させることで、線を消します。 */
    border-bottom: 1px solid #ECEFF4;
}

/* 左タブが選択された場合: 右の境界線を消す */
QTabBar::tab:selected:left {
    margin-right: -1px;
    border-right-color: #ECEFF4;
}

/* 右タブが選択された場合: 左の境界線を消す */
QTabBar::tab:selected:right {
    margin-left: -1px;
    border-left-color: #ECEFF4;
}

/* 下タブが選択された場合: 上の境界線を消す */
QTabBar::tab:selected:bottom {
    /* 選択されたタブを1pxだけ上にずらし、paneの下部境界線の上に重ねて隠します。 */
    margin-top: -1px;
    /* 上の境界線を、スタイルも含めて背景色と完全に一致させることで、線を消します。 */
    border-top: 1px solid #ECEFF4;
}

QTabBar::close-button {
    margin: 2px;
}
QTabBar::close-button:hover {
    background: #BFC9DB;
    border-radius: 8px;
}

/* メニュー */
QMenu {
    background-color: #E5E9F0;
    border: 1px solid #D8DEE9;
    padding: 5px;
}

QMenu::item {
    padding: 8px 25px;
    border-radius: 4px;
}

QMenu::item:selected {
    background-color: #81A1C1; /* Nord Frost */
    color: #ECEFF4;
}

QMenu::separator {
    height: 1px;
    background: #D8DEE9;
    margin: 5px 0;
}

/* メッセージボックス */
QMessageBox {
    background-color: #E5E9F0;
}

QMessageBox QLabel {
    background-color: transparent;
    color: #2E3440;
    font-size: 10pt;
}

/* ボタン */
QPushButton {
    background-color: #5E81AC; /* Nord Frost */
    color: #ECEFF4;
    border: none;
    padding: 8px 16px;
    border-radius: 4px;
    font-weight: bold;
}

QPushButton:hover {
    background-color: #81A1C1;
}

QPushButton:pressed {
    background-color: #88C0D0;
}

/* リストウィジェット */
QListWidget {
    background-color: #ECEFF4;
    border: 1px solid #D8DEE9;
    border-radius: 4px;
    padding: 2px;
}

QListWidget::item {
    padding: 10px;
    border-radius: 3px;
}

QListWidget::item:hover {
    background-color: #E5E9F0;
}

QListWidget::item:selected {
    background-color: #5E81AC;
    color: #ECEFF4;
}

/* グループボックス */
QGroupBox {
    border: 1px solid #D8DEE9;
    border-radius: 8px;
    margin-top: 15px;
    padding: 20px 15px 15px 15px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 0 10px;
    left: 10px;
    color: #5E81AC;
    background-color: #ECEFF4; /* 背景色で線を途切れさせる */
    font-weight: bold;
}

/* テキスト入力 */
QLineEdit {
    background-color: #FFFFFF;
    border: 1px solid #D8DEE9;
    padding: 6px;
    border-radius: 4px;
}

/* プログレスバー */
QProgressBar {
    border: 1px solid #D8DEE9;
    border-radius: 5px;
    text-align: center;
    color: #2E3440;
    background-color: #E5E9F0;
}

QProgressBar::chunk {
    background-color: #88C0D0; /* Nord Frost */
    border-radius: 4px;
}

/* スクロールバー */
QScrollBar:vertical { border: none; background: #ECEFF4; width: 12px; margin: 0; }
QScrollBar::handle:vertical { background: #BFC9DB; min-height: 25px; border-radius: 6px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; }
QScrollBar:horizontal { border: none; background: #ECEFF4; height: 12px; margin: 0; }
QScrollBar::handle:horizontal { background: #BFC9DB; min-width: 25px; border-radius: 6px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; }
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: none; }

/* 設定ダイアログのカテゴリリスト */
#SettingsCategories {
    background-color: #E5E9F0;
    border-right: 1px solid #D8DEE9;
}

#SettingsCategories::item {
    padding: 12px 20px;
    border-left: 3px solid transparent; /* 未選択時は透明 */
}

#SettingsCategories::item:hover {
    background-color: #D8DEE9;
}

#SettingsCategories::item:selected {
    background-color: #ECEFF4; /* 右ペインと一体化 */
    color: #2E3440; /* テキスト色を濃く */
    border-left: 3px solid #5E81AC;
}
"""

# テーマ名とスタイルシートの辞書
THEMES = {
    "ダーク": DARK_STYLESHEET,
    "ライト": LIGHT_STYLESHEET,
}

SEARCH_ENGINES = {
    "Google": "https://www.google.com/search?q={}",
    "Bing": "https://www.bing.com/search?q={}",
    "DuckDuckGo": "https://duckduckgo.com/?q={}",
    "Yahoo Japan": "https://search.yahoo.co.jp/search?p={}",
    "Yandex": "https://yandex.com/search/?text={}",
    "Perplexity": "https://www.perplexity.ai/search?q={}",
    "Baidu": "https://www.baidu.com/s?wd={}",
    "Startpage": "https://www.startpage.com/do/dsearch?query={}",
    "Naver": "https://search.naver.com/search.naver?query={}",
    "Ecosia": "https://www.ecosia.org/search?q={}",
    "Seznam": "https://search.seznam.cz/?q={}",
    "Lilo": "https://search.lilo.org/searchweb.php?query={}",
    "Mojeek": "https://www.mojeek.com/search?q={}",
}

class HorizontalTextTabStyle(QProxyStyle):
    """垂直タブのテキストを水平に描画するためのカスタムスタイル"""
    def __init__(self, style=None):
        super().__init__(style)

    def sizeFromContents(self, contentsType, option, size, widget):
        if contentsType == QStyle.ContentsType.CT_TabBarTab and option.shape in (QTabBar.Shape.RoundedWest, QTabBar.Shape.RoundedEast):
            # 垂直タブの場合、水平タブとしてサイズを計算し、幅と高さを入れ替えて返す
            # これにより、QTabBarがサイズを転置して解釈した結果、意図した横長のタブサイズになる
            fm = option.fontMetrics
            text_size = fm.size(Qt.TextFlag.TextShowMnemonic, option.text)
            
            icon_width = 0
            if not option.icon.isNull():
                icon_width = option.iconSize.width() + 4 # アイコンとテキストの間のスペース

            # スタイルシートのpadding (QTabBar::tab:left,right の padding: 8px 12px) を考慮
            width = text_size.width() + icon_width + 24 # 左右パディング (12*2)
            height = max(text_size.height(), option.iconSize.height()) + 16 # 上下パディング (8*2)

            return QSize(height, width) # QTabBarが転置するので、(高さ, 幅) を返す

        return super().sizeFromContents(contentsType, option, size, widget)

    def drawControl(self, element, option, painter, widget=None):
        if element == QStyle.ControlElement.CE_TabBarTab and isinstance(option, QStyleOptionTab):
            if option.shape in (QTabBar.Shape.RoundedWest, QTabBar.Shape.RoundedEast):
                # 垂直タブの描画を乗っ取り、水平タブとして描画させる
                h_option = QStyleOptionTab(option)
                h_option.shape = QTabBar.Shape.RoundedNorth
                # 1. 水平タブとして背景と枠を描画
                self.proxy().drawControl(QStyle.ControlElement.CE_TabBarTabShape, h_option, painter, widget)
                # 2. アイコンとテキストを水平に描画
                self.proxy().drawControl(QStyle.ControlElement.CE_TabBarTabLabel, h_option, painter, widget)
                return # 描画を完了したのでここで終了

        super().drawControl(element, option, painter, widget)

def get_windows_theme(): # Windowsのみ対応
    """Windowsの個人設定からアプリのテーマ（ライト/ダーク）を取得する"""
    # ポータブル化のため、レジストリの読み取りを無効化し、常にダークをデフォルトとする。
    # テーマはsettings.iniで管理される。
    return "ダーク"

# ブックマークHTML解析用クラス (HTMLParserを継承)
class BookmarkHTMLParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_a_tag = False
        self.bookmarks = []
        self.current_href = ""
        self.current_title = ""

    def handle_starttag(self, tag, attrs):
        """開始タグを処理する (<a> タグを検出)"""
        if tag == 'a':
            attrs_dict = dict(attrs)
            if 'href' in attrs_dict:
                self.in_a_tag = True
                self.current_href = attrs_dict['href']
                self.current_title = ""

    def handle_endtag(self, tag):
        """終了タグを処理する (</a> タグを検出)"""
        if tag == 'a' and self.in_a_tag:
            self.in_a_tag = False
            if self.current_href.startswith(('http://', 'https://')) and self.current_title:
                self.bookmarks.append({'title': self.current_title.strip(), 'url': self.current_href})

    def handle_data(self, data):
        if self.in_a_tag:
            # <a> タグ内のテキストをタイトルとして取得
            self.current_title += data

# JavaScriptのコンソールエラーを抑制するためのカスタムWebEnginePage
class SilentWebEnginePage(QWebEnginePage):
    def javaScriptConsoleMessage(self, level, message, lineNumber, sourceID):
        # ウェブサイト側から出力されるJavaScriptのコンソールメッセージを
        # ターミナルに表示しないようにします。
        # デバッグ時に確認したい場合は、以下のコメントを解除してください。
        # print(f"JS Console ({sourceID}:{lineNumber}): {message}")
        pass

    def createWindow(self, window_type):
        """ウェブページからの新しいタブ/ウィンドウの作成要求 (window.openなど) をハンドルする"""
        # self.parent() はこのページがセットされている QWebEngineView を返す
        view = self.parent()
        if not view:
            return None

        # 親ウィジェットを辿ってBrowserWindowインスタンスを見つける
        main_window = view
        while main_window and not isinstance(main_window, BrowserWindow):
            main_window = main_window.parent()

        if not main_window:
            return None

        # 「新しいウィンドウで開く」の場合
        if window_type == QWebEnginePage.WebWindowType.WebBrowserWindow:
            # 現在のウィンドウがプライベートかどうかでプロファイルを決定
            if main_window.is_private:
                profile = QWebEngineProfile()  # 新しい一時的なプライベートプロファイル
            else:
                profile = persistent_profile  # グローバルな永続プロファイル

            if profile is None:  # フォールバック
                profile = QWebEngineProfile()

            new_window = BrowserWindow(profile=profile)
            windows.append(new_window)
            new_window.show()

            # 新しいウィンドウにはコンストラクタによって既にタブが1つ作成されている。
            # そのタブのページオブジェクトを返すことで、Qtが要求されたURLをロードする。
            browser = new_window.tabs.currentWidget()
            if browser:
                return browser.page()
            else:
                _browser, new_page = new_window._create_new_browser(set_as_current=True, label="読み込み中...")
                return new_page

        # 「新しいタブで開く」またはその他の場合
        if window_type == QWebEnginePage.WebWindowType.WebBrowserBackgroundTab:
            set_as_current = False
        else:
            set_as_current = True

        _browser, new_page = main_window._create_new_browser(set_as_current=set_as_current, label="読み込み中...")
        return new_page

# ブックマークウィンドウクラス
class BookmarkWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ブックマーク")
        self.setGeometry(300, 300, 400, 600)
        self.parent = parent
        # UIの構築
        layout = QVBoxLayout()

        # 検索バー
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("ブックマークを検索...")
        self.search_bar.textChanged.connect(self.filter_items)
        layout.addWidget(self.search_bar)

        # ブックマークリストウィジェット
        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)

        # ボタン用の水平レイアウト
        button_layout = QHBoxLayout()

        # インポート/エクスポートボタン
        import_button = QPushButton("インポート")
        import_button.clicked.connect(self.handle_import)
        button_layout.addWidget(import_button)

        export_button = QPushButton("エクスポート")
        export_button.clicked.connect(self.handle_export)
        button_layout.addWidget(export_button)

        # ブックマークの追加ボタン
        add_button = QPushButton("現在のページをブックマーク")
        add_button.clicked.connect(self.add_bookmark)
        button_layout.addWidget(add_button)

        layout.addLayout(button_layout)
        self.setLayout(layout)
        self.filter_items() # 初回ロード

        # ダブルクリックでブックマークを開く
        self.list_widget.itemDoubleClicked.connect(self.open_bookmark)
        # 右クリックメニューの作成
        self.list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_widget.customContextMenuRequested.connect(self.show_context_menu)

    def filter_items(self):
        """検索バーのテキストに基づいてブックマークをフィルタリング"""
        search_text = self.search_bar.text().lower()
        self.list_widget.clear()
        for bookmark in self.parent.bookmarks:
            if search_text in bookmark['title'].lower() or search_text in bookmark['url'].lower():
                item = QListWidgetItem(f"{bookmark['title']} ({bookmark['url']})")
                item.setData(Qt.ItemDataRole.UserRole, bookmark)
                self.list_widget.addItem(item)

    def load_bookmarks(self):
        """ブックマークリストを再読み込みし、表示を更新する"""
        # 検索バーをクリアする。テキストが入っていればtextChangedシグナルが発火し、リストが更新される。
        self.search_bar.clear()
        # 検索バーがもともと空だった場合、textChangedが発火しないため、
        # 明示的にfilter_itemsを呼び出してリストの更新を保証する。
        self.filter_items()

    def add_bookmark(self):
        """現在のタブのURLとタイトルをブックマークに追加"""
        current_browser = self.parent.tabs.currentWidget()
        if current_browser:
            url = current_browser.url().toString()
            title = current_browser.title()
            # ブックマークが存在しないかチェック
            if not any(b['url'] == url for b in self.parent.bookmarks):
                self.parent.bookmarks.insert(0, {"title": title, "url": url})
                self.parent.save_bookmarks()
                self.load_bookmarks()
                QMessageBox.information(self, "完了", "ブックマークに追加しました。")
            else:
                QMessageBox.warning(self, "警告", "このページはすでにブックマークされています。")

    def open_bookmark(self, item):
        """ブックマークをクリックしてページを開く"""
        bookmark = item.data(Qt.ItemDataRole.UserRole)
        self.parent.add_new_tab(QUrl(bookmark['url'])) # 親ウィンドウに新しいタブの作成を依頼
        self.accept()

    def show_context_menu(self, pos):
        """右クリックメニューを表示"""
        item = self.list_widget.itemAt(pos)
        if item:
            menu = QMenu(self)
            delete_action = QAction("削除", self)
            delete_action.triggered.connect(lambda: self.delete_bookmark(item))
            menu.addAction(delete_action)
            menu.exec(self.list_widget.mapToGlobal(pos))
            
    def delete_bookmark(self, item):
        """ブックマークを削除"""
        bookmark_to_delete = item.data(Qt.ItemDataRole.UserRole)
        self.parent.bookmarks = [b for b in self.parent.bookmarks if b['url'] != bookmark_to_delete['url']]
        self.parent.save_bookmarks()
        self.load_bookmarks()

    def handle_import(self):
        """インポート処理を親ウィンドウに依頼し、完了後にリストを更新する"""
        self.parent.import_bookmarks()
        self.load_bookmarks()

    def handle_export(self):
        """エクスポート処理を親ウィンドウに依頼する"""
        self.parent.export_bookmarks()

# 履歴ウィンドウクラス
class HistoryWindow(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("履歴")
        self.setGeometry(200, 200, 500, 600)
        self.parent = parent
        # UIの構築
        layout = QVBoxLayout()

        # 検索バー
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("履歴を検索...")
        self.search_bar.textChanged.connect(self.filter_items)
        layout.addWidget(self.search_bar)

        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)
        self.setLayout(layout)

        self.filter_items() # 初回ロード（検索バーは空なので全件表示）

        # ダブルクリックで履歴のページを開く
        self.list_widget.itemDoubleClicked.connect(self.open_history_url)

    def filter_items(self):
        """検索バーのテキストに基づいて履歴をフィルタリング"""
        search_text = self.search_bar.text().strip()
        self.list_widget.clear()

        try:
            conn = sqlite3.connect(self.parent.history_db_path)
            cursor = conn.cursor()
            
            if search_text:
                # 検索語でタイトルとURLを検索
                query = """
                    SELECT title, url, last_visit_time FROM history
                    WHERE title LIKE ? OR url LIKE ?
                    ORDER BY last_visit_time DESC
                    LIMIT 200
                """
                search_pattern = f"%{search_text}%"
                cursor.execute(query, (search_pattern, search_pattern))
            else:
                # 検索語がなければ全件を降順で取得
                query = """
                    SELECT title, url, last_visit_time FROM history
                    ORDER BY last_visit_time DESC
                    LIMIT 200
                """
                cursor.execute(query)

            for title, url, last_visit_time in cursor.fetchall():
                try:
                    dt = datetime.fromisoformat(last_visit_time)
                    time_str = dt.strftime('%Y/%m/%d %H:%M')
                except (ValueError, TypeError):
                    time_str = "不明な日時"
                
                display_text = f"{title}\n{url}\n{time_str}"
                item = QListWidgetItem(display_text)
                item.setData(Qt.ItemDataRole.UserRole, url)
                self.list_widget.addItem(item)

        except sqlite3.Error as e:
            print(f"履歴データベースの読み込みに失敗しました: {e}")
        finally:
            if conn:
                conn.close()

    def open_history_url(self, item):
        """リストの項目をダブルクリックしたときにURLを新しいタブで開く"""
        url = item.data(Qt.ItemDataRole.UserRole)
        if url:
            self.parent.add_new_tab(QUrl(url))
            self.accept()

# 設定ダイアログクラス
class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("設定")
        self.setMinimumSize(700, 580)

        # ad_block_list.txtのパスを取得
        self.ad_block_list_path = self.parent.ad_blocker.block_list_path

        # メインレイアウト (左にカテゴリ、右に設定ページ)
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 左側のカテゴリリスト
        self.categories_widget = QListWidget()
        self.categories_widget.setFixedWidth(180)
        self.categories_widget.setObjectName("SettingsCategories")

        # 右側の設定ページ (QStackedWidget)
        self.pages_widget = QStackedWidget()

        main_layout.addWidget(self.categories_widget)
        main_layout.addWidget(self.pages_widget)

        # カテゴリとページを作成
        self.create_categories_and_pages()

        # テーマに基づいてアイコンの色を初期設定
        self.update_theme_elements()

        # カテゴリリストの選択が変更されたらページを切り替える
        self.categories_widget.currentRowChanged.connect(self.pages_widget.setCurrentIndex)
        self.categories_widget.setCurrentRow(0)

    def update_theme_elements(self):
        """テーマ変更時にアイコンの色などを更新する"""
        # 現在のテーマに合わせたアイコンの色を取得
        icon_color = self.parent.theme_colors['icon_color']
        self.categories_widget.item(0).setIcon(qta.icon('fa5s.cog', color=icon_color))
        self.categories_widget.item(1).setIcon(qta.icon('fa5s.layer-group', color=icon_color))
        self.categories_widget.item(2).setIcon(qta.icon('fa5s.user-shield', color=icon_color))
        self.categories_widget.item(3).setIcon(qta.icon('fa5s.bookmark', color=icon_color))
        self.categories_widget.item(4).setIcon(qta.icon('fa5s.shield-alt', color=icon_color))
        self.categories_widget.item(5).setIcon(qta.icon('fa5s.info-circle', color=icon_color))

        # 「このEQUAについて」ページを再生成して差し替える
        about_page_widget = self.pages_widget.widget(5)
        self.pages_widget.removeWidget(about_page_widget) # 古いページを削除
        self.pages_widget.insertWidget(5, self.create_about_page())
        about_page_widget.deleteLater()

    def create_categories_and_pages(self):
        """カテゴリリストと対応するページを作成し、ウィジェットに追加する"""
        # 一般
        self.categories_widget.addItem(QListWidgetItem(qta.icon('fa5s.cog'), "一般"))
        self.pages_widget.addWidget(self.create_general_page())

        # タブグループ
        self.categories_widget.addItem(QListWidgetItem(qta.icon('fa5s.layer-group'), "タブグループ"))
        self.pages_widget.addWidget(self.create_group_management_page())

        # プライバシー
        self.categories_widget.addItem(QListWidgetItem(qta.icon('fa5s.user-shield'), "プライバシー"))
        self.pages_widget.addWidget(self.create_privacy_page())

        # ブックマーク
        self.categories_widget.addItem(QListWidgetItem(qta.icon('fa5s.bookmark'), "ブックマーク"))
        self.pages_widget.addWidget(self.create_bookmark_page())

        # 広告ブロック
        self.categories_widget.addItem(QListWidgetItem(qta.icon('fa5s.shield-alt'), "広告ブロック"))
        self.pages_widget.addWidget(self.create_ad_block_page())

        # このEQUAについて
        self.categories_widget.addItem(QListWidgetItem(qta.icon('fa5s.info-circle'), "このEQUAについて"))
        self.pages_widget.addWidget(self.create_about_page())

    def create_about_page(self):
        """「このアプリについて」ページを作成する"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setSpacing(15)
        layout.setContentsMargins(40, 40, 40, 40) # ページ内の余白を調整
        layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        link_color = self.parent.theme_colors['icon_active_color']
        
        # アプリアイコン
        app_icon_label = QLabel()
        icon_pixmap = QPixmap(resource_path('equa.ico')).scaled(QSize(80, 80), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        app_icon_label.setPixmap(icon_pixmap)
        app_icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(app_icon_label)

        # アプリ名とバージョン
        title_label = QLabel("EQUA")
        title_font = title_label.font()
        title_font.setPointSize(22)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)

        version_label = QLabel(f"バージョン: {__version__}")
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(version_label)

        layout.addSpacing(20)

        # 開発者情報
        dev_label = QLabel("開発者: StudioNosa (猫星　吹恋)")
        dev_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(dev_label)

        # GitHubリンク
        github_url = f"https://github.com/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}"
        github_link_label = QLabel(f'<a href="{github_url}" style="color: {link_color};">GitHubリポジトリ</a>')
        github_link_label.setOpenExternalLinks(True)
        github_link_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(github_link_label)

        layout.addSpacing(20)

        # 使用ライブラリ情報
        libs_group = QGroupBox("使用ライブラリ")
        libs_layout = QVBoxLayout(libs_group)
        libs_layout.setSpacing(10)

        qt_label = QLabel(f'<b>Qt Framework</b> - <a href="https://www.qt.io/" style="color: {link_color};">The Qt Company Ltd.</a>')
        qt_label.setOpenExternalLinks(True)
        libs_layout.addWidget(qt_label)

        pyqt6_label = QLabel(f'<b>PyQt6</b> - <a href="https://www.riverbankcomputing.com/software/pyqt/" style="color: {link_color};">The Qt Company Ltd. and Riverbank Computing Ltd.</a>')
        pyqt6_label.setOpenExternalLinks(True)
        libs_layout.addWidget(pyqt6_label)

        qtawesome_label = QLabel(f'<b>QtAwesome</b> (Font Awesome 5) - <a href="https://github.com/spyder-ide/qtawesome" style="color: {link_color};">The Spyder Development Team</a>')
        qtawesome_label.setOpenExternalLinks(True)
        libs_layout.addWidget(qtawesome_label)
        layout.addWidget(libs_group)

        # ライセンス情報
        license_group = QGroupBox("ライセンス")
        license_layout = QVBoxLayout(license_group)
        license_label = QLabel(
            'このアプリケーションは <b>GNU General Public License v3.0</b> の下で配布されています。<br>'
            f'<a href="https://www.gnu.org/licenses/gpl-3.0.html" style="color: {link_color};">ライセンス条文はこちら</a>'
        )
        license_label.setOpenExternalLinks(True)
        license_label.setWordWrap(True)
        license_layout.addWidget(license_label)
        layout.addWidget(license_group)

        return page

    def create_general_page(self):
        """「一般」設定ページを作成する"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 20, 30, 20)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.setSpacing(25)

        # 外観設定
        appearance_group = QGroupBox("外観")
        appearance_layout = QVBoxLayout(appearance_group)
        appearance_layout.setSpacing(15)
        
        theme_layout = QHBoxLayout()
        theme_label = QLabel("テーマ:")
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(list(THEMES.keys())) # 「自動」を削除
        current_theme = self.parent.settings.value("theme", "ダーク")
        self.theme_combo.setCurrentText(current_theme)
        self.theme_combo.currentTextChanged.connect(self.parent.change_theme)
        theme_layout.addWidget(theme_label)
        theme_layout.addWidget(self.theme_combo)
        appearance_layout.addLayout(theme_layout)

        # タブの表示位置設定
        tab_pos_layout = QHBoxLayout()
        tab_pos_label = QLabel("タブの表示位置:")
        self.tab_pos_combo = QComboBox()
        self.tab_pos_combo.addItems(["上", "下", "左", "右"])
        current_tab_pos = self.parent.settings.value("tab_position", "上")
        self.tab_pos_combo.setCurrentText(current_tab_pos)
        self.tab_pos_combo.currentTextChanged.connect(self.parent.change_tab_position)
        tab_pos_layout.addWidget(tab_pos_label)
        tab_pos_layout.addWidget(self.tab_pos_combo)
        appearance_layout.addLayout(tab_pos_layout)
        layout.addWidget(appearance_group)
        
        # 起動設定
        startup_group = QGroupBox("起動設定")
        startup_layout = QVBoxLayout(startup_group)
        startup_layout.setSpacing(15)
        
        # 検索エンジン設定
        search_engine_layout = QHBoxLayout()
        search_engine_label = QLabel("デフォルトの検索エンジン:")
        self.search_engine_combo = QComboBox()
        self.search_engine_combo.addItems(SEARCH_ENGINES.keys())
        current_search_engine = self.parent.settings.value("search_engine", "Google")
        self.search_engine_combo.setCurrentText(current_search_engine)
        self.search_engine_combo.currentTextChanged.connect(self.parent.change_search_engine)
        search_engine_layout.addWidget(search_engine_label)
        search_engine_layout.addWidget(self.search_engine_combo)
        startup_layout.addLayout(search_engine_layout)
        
        url_layout = QHBoxLayout()
        url_label = QLabel("新しいタブのデフォルトURL:")
        self.url_input = QLineEdit(self.parent.default_new_tab_url)
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.url_input)
        startup_layout.addLayout(url_layout)

        # アップデートチェック設定
        self.update_check_checkbox = QCheckBox("起動時にアップデートを確認する")
        self.update_check_checkbox.setChecked(self.parent.settings.value("update_check_enabled", True, type=bool))
        self.update_check_checkbox.toggled.connect(self.toggle_update_check)
        startup_layout.addWidget(self.update_check_checkbox)

        # セッション復元設定
        self.session_restore_checkbox = QCheckBox("起動時に前回のセッションを復元する")
        self.session_restore_checkbox.setChecked(self.parent.settings.value("session_restore_enabled", True, type=bool))
        self.session_restore_checkbox.toggled.connect(self.toggle_session_restore)
        startup_layout.addWidget(self.session_restore_checkbox)

        # ウィンドウサイズ復元設定
        self.window_geometry_restore_checkbox = QCheckBox("終了時のウィンドウサイズと位置を記憶する")
        self.window_geometry_restore_checkbox.setChecked(self.parent.settings.value("window_geometry_restore_enabled", True, type=bool))
        self.window_geometry_restore_checkbox.toggled.connect(self.toggle_window_geometry_restore)
        startup_layout.addWidget(self.window_geometry_restore_checkbox)

        # ハードウェアアクセラレーション設定
        self.hw_accel_checkbox = QCheckBox("ハードウェアアクセラレーションを有効にする (要再起動)")
        self.hw_accel_checkbox.setChecked(self.parent.settings.value("hw_accel_enabled", True, type=bool))
        self.hw_accel_checkbox.toggled.connect(self.toggle_hw_accel)
        startup_layout.addWidget(self.hw_accel_checkbox)

        save_button = QPushButton("URLを保存")
        save_button.clicked.connect(self.save_general_settings)
        startup_layout.addWidget(save_button, alignment=Qt.AlignmentFlag.AlignRight)
        
        layout.addWidget(startup_group)

        # ダウンロード設定
        download_group = QGroupBox("ダウンロード")
        download_layout = QVBoxLayout(download_group)
        download_layout.setSpacing(15)
        
        path_layout = QHBoxLayout()
        path_label = QLabel("ファイルの保存先:")
        
        default_download_dir = os.path.join(PORTABLE_BASE_PATH, "downloads")
        self.download_path_edit = QLineEdit(self.parent.settings.value("download_path", default_download_dir))
        self.download_path_edit.setReadOnly(True) # ユーザーに直接編集させず、ダイアログ経由にする
        
        browse_button = QPushButton("参照...")
        browse_button.clicked.connect(self.select_download_folder)
        
        path_layout.addWidget(path_label)
        path_layout.addWidget(self.download_path_edit)
        path_layout.addWidget(browse_button)
        download_layout.addLayout(path_layout)
        layout.addWidget(download_group)
        return page

    def create_privacy_page(self):
        """「プライバシー」設定ページを作成する"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 20, 30, 20)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.setSpacing(25)

        # Cookie設定
        cookie_group = QGroupBox("Cookie設定")
        cookie_layout = QVBoxLayout(cookie_group)
        cookie_layout.setSpacing(15)

        cookie_policy_layout = QHBoxLayout()
        cookie_policy_label = QLabel("Cookieの取り扱い:")
        self.cookie_policy_combo = QComboBox()

        # 利用可能なCookieポリシーを動的に構築
        self.cookie_policies = [
            ("すべてのCookieを許可する", QWebEngineProfile.PersistentCookiesPolicy.AllowPersistentCookies)
        ]
        # Qt 6.2以降で利用可能なサードパーティCookieブロックオプションを確認
        if hasattr(QWebEngineProfile.PersistentCookiesPolicy, 'BlockThirdPartyCookies'):
            self.cookie_policies.append(
                ("サードパーティのCookieをブロックする", QWebEngineProfile.PersistentCookiesPolicy.BlockThirdPartyCookies)
            )
        self.cookie_policies.append(
            ("すべての永続Cookieをブロックする", QWebEngineProfile.PersistentCookiesPolicy.NoPersistentCookies)
        )

        self.cookie_policy_combo.addItems([item[0] for item in self.cookie_policies])

        # 保存されたポリシーのenum値に基づいて、コンボボックスの選択状態を復元
        default_policy_value = QWebEngineProfile.PersistentCookiesPolicy.AllowPersistentCookies.value
        current_policy_value = self.parent.settings.value("privacy/cookie_policy_value", default_policy_value, type=int)
        
        current_index = next((i for i, (_, policy_enum) in enumerate(self.cookie_policies) if policy_enum.value == current_policy_value), 0)
        self.cookie_policy_combo.setCurrentIndex(current_index)
        self.cookie_policy_combo.currentIndexChanged.connect(self.change_cookie_policy)

        cookie_policy_layout.addWidget(cookie_policy_label)
        cookie_policy_layout.addWidget(self.cookie_policy_combo)
        cookie_layout.addLayout(cookie_policy_layout)
        layout.addWidget(cookie_group)

        # 閲覧データ設定
        data_group = QGroupBox("閲覧データ")
        data_layout = QVBoxLayout(data_group)
        data_layout.setSpacing(15)

        self.clear_on_exit_checkbox = QCheckBox("ブラウザ終了時に閲覧データを削除する")
        self.clear_on_exit_checkbox.setChecked(self.parent.settings.value("privacy/clear_on_exit", False, type=bool))
        self.clear_on_exit_checkbox.toggled.connect(
            lambda checked: self.parent.settings.setValue("privacy/clear_on_exit", checked)
        )
        data_layout.addWidget(self.clear_on_exit_checkbox)

        clear_data_button = QPushButton("閲覧データを今すぐ削除...")
        clear_data_button.clicked.connect(self.handle_clear_browsing_data)
        data_layout.addWidget(clear_data_button, alignment=Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(data_group)
        return page

    def create_group_management_page(self):
        """「タブグループ」設定ページを作成する"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 20, 30, 20)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.setSpacing(25)

        group = QGroupBox("タブグループ管理")
        group_layout = QVBoxLayout(group)
        group_layout.setSpacing(15)

        self.group_list = QListWidget() # グループ一覧
        self.load_groups()
        group_layout.addWidget(self.group_list)

        group_button_layout = QHBoxLayout()
        rename_button = QPushButton("名前の変更")
        rename_button.clicked.connect(self.rename_selected_group)
        group_button_layout.addWidget(rename_button)

        color_button = QPushButton("色の変更")
        color_button.clicked.connect(self.change_selected_group_color)
        group_button_layout.addWidget(color_button)
        group_layout.addLayout(group_button_layout)

        group_button_layout2 = QHBoxLayout()
        ungroup_button = QPushButton("グループを解散")
        ungroup_button.clicked.connect(self.ungroup_selected_group)
        group_button_layout2.addWidget(ungroup_button)

        close_group_button = QPushButton("グループを閉じる")
        close_group_button.clicked.connect(self.close_selected_group)
        group_button_layout2.addWidget(close_group_button)
        group_layout.addLayout(group_button_layout2)

        layout.addWidget(group)
        return page

    def create_bookmark_page(self):
        """「ブックマーク」設定ページを作成する"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 20, 30, 20)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.setSpacing(25)

        group = QGroupBox("ブックマークデータの管理")
        group_layout = QHBoxLayout(group)
        group_layout.setSpacing(15)

        import_button = QPushButton("HTMLからインポート")
        import_button.clicked.connect(self.parent.import_bookmarks)
        group_layout.addWidget(import_button)

        export_button = QPushButton("HTMLへエクスポート")
        export_button.clicked.connect(self.parent.export_bookmarks)
        group_layout.addWidget(export_button)

        layout.addWidget(group)
        return page

    def create_ad_block_page(self):
        """「広告ブロック」設定ページを作成する"""
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 20, 30, 20)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        layout.setSpacing(25)

        # ON/OFF設定
        enable_group = QGroupBox("広告ブロック設定")
        enable_layout = QVBoxLayout(enable_group)
        enable_layout.setSpacing(15)
        self.ad_block_checkbox = QCheckBox("広告ブロックを有効にする")
        self.ad_block_checkbox.setChecked(self.parent.settings.value("ad_block_enabled", True, type=bool))
        self.ad_block_checkbox.stateChanged.connect(self.toggle_ad_blocking)
        enable_layout.addWidget(self.ad_block_checkbox)
        layout.addWidget(enable_group)

        # 起動時更新チェックボックスを追加
        self.autoupdate_checkbox = QCheckBox("起動時にリストを自動で更新する")
        self.autoupdate_checkbox.setChecked(self.parent.settings.value("ad_block_autoupdate_enabled", True, type=bool))
        self.autoupdate_checkbox.toggled.connect(
            lambda checked: self.parent.settings.setValue("ad_block_autoupdate_enabled", checked)
        )
        enable_layout.addWidget(self.autoupdate_checkbox)

        # 自動更新設定
        update_group = QGroupBox("リストの自動更新")
        update_layout = QVBoxLayout(update_group)
        update_layout.setSpacing(15)

        url_layout = QHBoxLayout()
        url_label = QLabel("更新元URL:")
        self.update_url_input = QLineEdit(self.parent.settings.value("ad_block_update_url", DEFAULT_ADBLOCK_LIST_URL))
        self.update_url_input.setToolTip("更新に使用する広告ブロックリスト（ABP形式など）のURL")
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.update_url_input)
        update_layout.addLayout(url_layout)

        update_button_layout = QHBoxLayout()
        self.last_updated_label = QLabel()
        self.update_last_updated_label()
        self.update_list_button = QPushButton("今すぐ更新")
        self.update_list_button.clicked.connect(self.start_manual_blocklist_update)
        update_button_layout.addWidget(self.last_updated_label, alignment=Qt.AlignmentFlag.AlignLeft)
        update_button_layout.addWidget(self.update_list_button, alignment=Qt.AlignmentFlag.AlignRight)
        update_layout.addLayout(update_button_layout)
        layout.addWidget(update_group)

        # リスト管理
        list_group = QGroupBox("ブロックするドメインのリスト")
        list_layout = QVBoxLayout(list_group)
        list_layout.setSpacing(10)
        self.block_list_widget = QListWidget()
        self.block_list_widget.setToolTip("ここにブロックしたいドメイン名（例: example.com）を一行ずつ追加します。")
        self.block_list_widget.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        list_layout.addWidget(self.block_list_widget)
        
        button_layout = QHBoxLayout()
        add_button = QPushButton("追加")
        add_button.clicked.connect(self.add_block_domain)
        remove_button = QPushButton("削除")
        remove_button.clicked.connect(self.remove_block_domain)
        button_layout.addStretch()
        button_layout.addWidget(add_button)
        button_layout.addWidget(remove_button)
        list_layout.addLayout(button_layout)
        
        layout.addWidget(list_group)
        
        self.load_block_list()
        return page

    def load_groups(self):
        """タブグループの一覧をリストウィジェットに読み込む"""
        self.group_list.clear()
        for name, color in self.parent.groups.items():
            pixmap = QPixmap(16, 16)
            pixmap.fill(color)
            icon = QIcon(pixmap)
            item = QListWidgetItem(icon, name)
            self.group_list.addItem(item)

    def rename_selected_group(self):
        """選択されているグループの名前を変更する"""
        currentItem = self.group_list.currentItem()
        if not currentItem: return
        old_name = currentItem.text()
        new_name, ok = QInputDialog.getText(self, "グループ名の変更", f"「{old_name}」の新しい名前:", text=old_name)
        if ok and new_name and new_name != old_name and new_name not in self.parent.groups:
            self.parent.rename_group(old_name, new_name)
            self.load_groups()

    def change_selected_group_color(self):
        """選択されているグループの色を変更する"""
        currentItem = self.group_list.currentItem()
        if not currentItem: return
        group_name = currentItem.text()
        color = QColorDialog.getColor(self.parent.groups[group_name], self, "グループの色を選択")
        if color.isValid():
            self.parent.change_group_color(group_name, color)
            self.load_groups()

    def ungroup_selected_group(self):
        """選択されているグループを解散する"""
        currentItem = self.group_list.currentItem()
        if not currentItem: return
        group_name = currentItem.text()
        reply = QMessageBox.question(self, "確認", f"グループ「{group_name}」を解散しますか？\n（タブは閉じられません）",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.parent.ungroup_tabs(group_name)
            self.load_groups()

    def close_selected_group(self):
        """選択されているグループに属するタブをすべて閉じる"""
        currentItem = self.group_list.currentItem()
        if not currentItem: return
        group_name = currentItem.text()
        reply = QMessageBox.question(self, "確認", f"グループ「{group_name}」に属するすべてのタブを閉じますか？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.parent.close_group(group_name)
            self.load_groups()

    def save_general_settings(self):
        """一般設定を保存する"""
        url_text = self.url_input.text().strip() # 前後の空白を削除
        if url_text:
            qurl = QUrl(url_text)
            # スキーム(http, httpsなど)がなければ、httpsを補完する
            if not qurl.scheme():
                qurl.setScheme("https")
            
            valid_url = qurl.toString()
            self.url_input.setText(valid_url) # UIの表示も更新する
            self.parent.default_new_tab_url = valid_url
            self.parent.settings.setValue("default_new_tab_url", valid_url)
            QMessageBox.information(self, "成功", "設定を保存しました。")
        else:
            QMessageBox.warning(self, "警告", "URLを入力してください。")

    def toggle_update_check(self, checked):
        """アップデートチェックの有効/無効をQSettingsに保存する"""
        self.parent.settings.setValue("update_check_enabled", checked)

    def toggle_session_restore(self, checked):
        """セッション復元機能の有効/無効をQSettingsに保存する"""
        self.parent.settings.setValue("session_restore_enabled", checked)

    def toggle_window_geometry_restore(self, checked):
        """ウィンドウサイズ復元機能の有効/無効をQSettingsに保存する"""
        self.parent.settings.setValue("window_geometry_restore_enabled", checked)

    def toggle_hw_accel(self, checked):
        """ハードウェアアクセラレーションの有効/無効を保存し、再起動を促す"""
        self.parent.settings.setValue("hw_accel_enabled", checked)
        QMessageBox.information(self, "設定の変更", "この設定はアプリケーションの再起動後に有効になります。")

    def select_download_folder(self):
        """ダウンロードフォルダを選択するダイアログを開き、設定を保存する"""
        current_path = self.download_path_edit.text()
        directory = QFileDialog.getExistingDirectory(self, "ダウンロード先フォルダを選択", current_path)
        if directory:
            self.download_path_edit.setText(directory)
            self.parent.settings.setValue("download_path", directory)

    def change_cookie_policy(self, index):
        """Cookieポリシー設定を保存し、適用する"""
        if 0 <= index < len(self.cookie_policies):
            # コンボボックスのインデックスではなく、対応するenumの値を保存する
            policy_value = self.cookie_policies[index][1].value
            self.parent.settings.setValue("privacy/cookie_policy_value", policy_value)
            apply_cookie_policy() # グローバル関数を呼んで全ウィンドウに適用

    def handle_clear_browsing_data(self):
        """「閲覧データを削除」ボタンが押されたときの処理"""
        reply = QMessageBox.question(self, "閲覧データの削除",
                                     "すべての閲覧データ（履歴、Cookie、キャッシュ、ダウンロード履歴）を削除しますか？\n\nこの操作は元に戻せません。",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            for window in windows: # 全ウィンドウに削除を指示
                window.clear_browsing_data()
            QMessageBox.information(self, "完了", "閲覧データを削除しました。")

    def toggle_ad_blocking(self, state):
        """広告ブロックのON/OFFを切り替える"""
        enabled = (state == Qt.CheckState.Checked.value) # チェックボックスの状態を取得
        # すべてのウィンドウに設定を適用
        for window in windows:
            window.set_ad_blocking(enabled)

    def load_block_list(self):
        """ブロックリストのユーザー定義部分を表示に読み込む"""
        self.block_list_widget.clear()
        try:
            if os.path.exists(self.ad_block_list_path):
                with open(self.ad_block_list_path, 'r', encoding='utf-8') as f:
                    in_user_section = False
                    user_rules_marker = "[User Defined Rules]"
                    for line in f:
                        if line.strip() == user_rules_marker:
                            in_user_section = True
                            continue
                        if in_user_section:
                            domain = line.strip()
                            if domain and not domain.startswith('!'):
                                self.block_list_widget.addItem(domain)
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"ブロックリストの読み込みに失敗しました:\n{e}")

    def save_block_list_and_reload(self):
        """リストウィジェットの内容をユーザー定義ルールとしてファイルに保存し、全ウィンドウのブロッカーを更新する"""
        try:
            auto_updated_rules = []
            user_rules_marker = "[User Defined Rules]"
            # 既存のファイルから自動更新部分を読み込む
            if os.path.exists(self.ad_block_list_path):
                 with open(self.ad_block_list_path, 'r', encoding='utf-8') as f:
                     for line in f:
                         if line.strip() == user_rules_marker:
                             break # ユーザー定義セクションに到達したら読み込みを停止
                         auto_updated_rules.append(line)
            
            with open(self.ad_block_list_path, 'w', encoding='utf-8') as f:
                # 自動更新部分を書き戻す
                f.writelines(auto_updated_rules)
                # 末尾に改行がない場合に備える
                if auto_updated_rules and not auto_updated_rules[-1].endswith('\n'):
                    f.write('\n')
                
                # ユーザー定義セクションを書き込む
                f.write(f'\n{user_rules_marker}\n')
                for i in range(self.block_list_widget.count()):
                    f.write(self.block_list_widget.item(i).text() + '\n')
            
            for window in windows: # 開いているすべてのウィンドウの広告ブロッカーを更新
                if window.ad_block_enabled:
                    window.ad_blocker.load_domains()
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"ブロックリストの保存に失敗しました:\n{e}")

    def add_block_domain(self):
        """ブロックリストに新しいドメインを追加する"""
        domain, ok = QInputDialog.getText(self, "ドメイン追加", "ブロックするドメインを入力してください:", QLineEdit.EchoMode.Normal)
        if ok and domain:
            domain = domain.strip().lower()
            if not domain: return
            if not self.block_list_widget.findItems(domain, Qt.MatchFlag.MatchExactly):
                self.block_list_widget.addItem(domain)
                self.save_block_list_and_reload()
            else:
                QMessageBox.warning(self, "警告", "そのドメインは既に追加されています。")

    def remove_block_domain(self):
        """ブロックリストから選択したドメインを削除する"""
        selected_items = self.block_list_widget.selectedItems()
        if not selected_items: return
        reply = QMessageBox.question(self, "確認", f"{len(selected_items)}個のドメインを削除しますか？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            for item in selected_items:
                self.block_list_widget.takeItem(self.block_list_widget.row(item))
            self.save_block_list_and_reload()

    def update_last_updated_label(self):
        """最終更新日時のラベルを更新する"""
        last_updated_iso = self.parent.settings.value("ad_block_last_updated", "")
        if last_updated_iso:
            try:
                dt = datetime.fromisoformat(last_updated_iso)
                self.last_updated_label.setText(f"最終更新: {dt.strftime('%Y/%m/%d %H:%M')}") # 見やすい形式に変換
            except ValueError:
                self.last_updated_label.setText("最終更新: 不明")
        else:
            self.last_updated_label.setText("最終更新: なし")

    def start_manual_blocklist_update(self):
        """「今すぐ更新」ボタンが押されたときの処理"""
        url = self.update_url_input.text().strip()
        if not url:
            QMessageBox.warning(self, "警告", "更新元URLを入力してください。")
            return
        self.parent.settings.setValue("ad_block_update_url", url) # 入力されたURLを保存
        self.update_list_button.setEnabled(False)
        self.update_list_button.setText("更新中...")
        self.parent.start_blocklist_update(silent=False) # 親ウィンドウに更新開始を依頼 (silent=Falseで結果をダイアログ表示)

    def on_update_finished(self, success, message):
        """親ウィンドウからのシグナルを受けてUIを更新するスロット"""
        self.update_list_button.setEnabled(True) # ボタンを再度有効化
        self.update_list_button.setText("今すぐ更新")
        if success:
            self.update_last_updated_label()
            self.load_block_list()
        if message:
            if success: QMessageBox.information(self, "成功", message)
            else: QMessageBox.critical(self, "更新失敗", message)

# アップデートファイルダウンロード用スレッド
class UpdateDownloadThread(QThread):
    # シグナル: (受信済みバイト数, 全体バイト数), (成功/失敗, 保存パス, エラーメッセージ)
    progress = pyqtSignal(int, int)
    finished = pyqtSignal(bool, str, str)  # success, downloaded_path, error_message

    def __init__(self, url, save_path, parent=None):
        super().__init__(parent)
        self.url = url
        self.save_path = save_path
        self._is_cancelled = False

    # スレッドのメイン処理
    def run(self):
        try:
            req = urllib.request.Request(self.url, headers={'User-Agent': 'EQUA-Update-Downloader'})
            with urllib.request.urlopen(req, timeout=10) as response:
                total_size = int(response.getheader('Content-Length', 0))
                bytes_received = 0
                chunk_size = 8192
                with open(self.save_path, 'wb') as f:
                    while True:
                        if self._is_cancelled:
                            self.finished.emit(False, "", "ダウンロードがキャンセルされました。")
                            return
                        chunk = response.read(chunk_size)
                        if not chunk:
                            break
                        f.write(chunk)
                        bytes_received += len(chunk)
                        self.progress.emit(bytes_received, total_size)
            self.finished.emit(True, self.save_path, "")
        except Exception as e:
            self.finished.emit(False, "", str(e))

    def cancel(self):
        """ダウンロードをキャンセルする"""
        self._is_cancelled = True

# アップデートダウンロードダイアログ
class UpdateDownloadDialog(QDialog):
    # シグナル: (成功/失敗, 保存パス)
    download_finished = pyqtSignal(bool, str) # success, downloaded_path

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("アップデートをダウンロード中")
        self.setModal(True)
        self.setFixedSize(400, 120)
        # UIの構築
        layout = QVBoxLayout(self)
        self.label = QLabel("アップデートファイルをダウンロードしています...")
        self.progress_bar = QProgressBar()
        self.progress_label = QLabel("0 MB / 0 MB")
        self.cancel_button = QPushButton("キャンセル")
        self.cancel_button.clicked.connect(self.cancel_download)

        layout.addWidget(self.label)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.progress_label, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.cancel_button, alignment=Qt.AlignmentFlag.AlignRight)

    def start_download(self, url, save_path):
        """ダウンロードスレッドを開始する"""
        self.download_thread = UpdateDownloadThread(url, save_path)
        self.download_thread.progress.connect(self.update_progress)
        self.download_thread.finished.connect(self.on_finished)
        self.download_thread.start()

    def update_progress(self, received, total):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(received)
        self.progress_label.setText(f"{received/1024/1024:.1f} MB / {total/1024/1024:.1f} MB")

    def on_finished(self, success, path, error_message):
        self.download_finished.emit(success, path)
        # ダイアログを閉じる
        self.accept()

    def cancel_download(self):
        if hasattr(self, 'download_thread') and self.download_thread.isRunning():
            # 実行中のスレッドにキャンセルを要求
            self.download_thread.cancel()

# ダウンロード項目ウィジェットクラス
class DownloadItemWidget(QWidget):
    def __init__(self, download_item, parent=None):
        super().__init__(parent)
        self.download_item = download_item
        # UIの構築
        layout = QHBoxLayout()
        layout.setSpacing(10) # ウィジェット間のスペースを調整
        self.setLayout(layout)

        self.filename_label = QLabel(self.get_filename())
        self.filename_label.setToolTip(self.get_filename())
        # ファイル名ラベルが伸縮する比率を2に設定
        layout.addWidget(self.filename_label, 2)

        self.progress_bar = QProgressBar()
        # プログレスバーが伸縮する比率を3に設定 (ファイル名より優先的に伸びる)
        layout.addWidget(self.progress_bar, 3)

        self.progress_label = QLabel()
        # 進捗ラベルの最小幅を設定してレイアウトを安定させる
        self.progress_label.setMinimumWidth(120)
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(self.progress_label)

        self.action_button = QPushButton()
        # アクションボタンの幅を固定してレイアウトを安定させる
        self.action_button.setFixedWidth(100)
        self.action_button.clicked.connect(self.perform_action)
        layout.addWidget(self.action_button)

        # PyQt6のバージョンによるAPI互換性: downloadProgressシグナルが存在する場合のみ接続
        if hasattr(self.download_item, 'downloadProgress'):
            self.download_item.downloadProgress.connect(self.update_progress)
        else:
            # シグナルがない古いバージョンの場合、プログレスバーを不定モードに設定
            self.progress_bar.setRange(0, 0)
            self.progress_label.setText("ダウンロード中...")

        self.download_item.stateChanged.connect(self.update_state) # 状態変化を監視
        self.update_state(self.download_item.state()) # 初期状態を設定

    def get_filename(self):
        # PyQt6のバージョンによるAPI互換性
        if hasattr(self.download_item, 'path'):
            return os.path.basename(self.download_item.path())
        else:
            return self.download_item.downloadFileName()
    
    def get_full_path(self):
        # API互換性
        if hasattr(self.download_item, 'path'):
            return self.download_item.path()
        else:
            return os.path.join(self.download_item.downloadDirectory(), self.download_item.downloadFileName())

    def update_progress(self, bytes_received, bytes_total):
        """プログレスバーとラベルを更新する"""
        if bytes_total > 0:
            self.progress_bar.setMaximum(int(bytes_total))
            self.progress_bar.setValue(int(bytes_received))
            self.progress_label.setText(f"{bytes_received/1024/1024:.1f} / {bytes_total/1024/1024:.1f} MB")
        else:
            self.progress_bar.setRange(0, 0) # 不定モード
            self.progress_label.setText(f"{bytes_received/1024/1024:.1f} MB")

    def update_state(self, state):
        """ダウンロードの状態に応じてUIを更新する"""
        DownloadState = type(self.download_item.state())
        if state == DownloadState.DownloadInProgress:
            self.action_button.setText("キャンセル")
            self.action_button.setEnabled(True)
        elif state == DownloadState.DownloadCompleted:
            # ダウンロード完了時にプログレスバーを100%にする
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(100)
            self.progress_label.setText("完了")
            self.action_button.setText("フォルダを開く")
            self.action_button.setEnabled(True)
        elif state == DownloadState.DownloadCancelled:
            # キャンセル時にもプログレスバーをリセット
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.progress_label.setText("キャンセル済")
            self.action_button.setEnabled(False)
        elif state == DownloadState.DownloadInterrupted:
            self.progress_label.setText("中断")
            self.action_button.setEnabled(False)

    def perform_action(self):
        """アクションボタン（キャンセル/フォルダを開く）の処理"""
        DownloadState = type(self.download_item.state())
        if self.download_item.state() == DownloadState.DownloadInProgress:
            self.download_item.cancel()
        elif self.download_item.state() == DownloadState.DownloadCompleted:
            folder = os.path.dirname(self.get_full_path())
            QDesktopServices.openUrl(QUrl.fromLocalFile(folder))

# ダウンロードマネージャークラス
class DownloadManager(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ダウンロード")
        # UIの構築 (スクロール可能なエリアにダウンロード項目を追加していく)
        self.setGeometry(400, 400, 600, 300)
        self.main_layout = QVBoxLayout(self)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        self.main_layout.addWidget(scroll_area)
        container = QWidget()
        self.downloads_layout = QVBoxLayout(container)
        self.downloads_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        scroll_area.setWidget(container)

    def add_download_item(self, download_item):
        """新しいダウンロード項目をリストの先頭に追加する"""
        widget = DownloadItemWidget(download_item)
        self.downloads_layout.insertWidget(0, widget)

    def closeEvent(self, event):
        """ウィンドウが閉じられたとき、実際には非表示にする"""
        self.hide()
        event.ignore()

# メインブラウザウィンドウクラス
class BrowserWindow(QMainWindow):
    # シグナル: 広告ブロックリストの更新が完了したときに (成功/失敗, メッセージ) を送信
    blocklist_update_finished = pyqtSignal(bool, str) 

    def __init__(self, profile):
        super().__init__()

        self.setWindowIcon(QIcon(resource_path('equa.ico')))

        self.profile = profile
        self.is_private = self.profile.isOffTheRecord()

        # 各種ダイアログやスレッドの参照を保持
        self.settings_dialog = None
        self.update_download_dialog = None
        # 開発者ツールウィンドウを管理するための辞書
        self.dev_tools_windows = {}

        # SPA遷移用の擬似プログレスバータイマーと関連変数
        self._spa_progress_timer = QTimer(self)
        self._spa_progress_timer.setInterval(50)  # 50msごとに更新
        self._spa_progress_timer.timeout.connect(self._update_spa_progress)
        self.update_thread = None
        self.fullscreen_request = None # 全画面リクエストを保持

        self.setWindowTitle("EQUA - ウェブ閲覧と使いやすさのHybrid")
        self.setGeometry(100, 100, 1024, 768)
        
        # アプリケーションデータディレクトリのパスを取得
        self.data_path = os.path.join(PORTABLE_BASE_PATH, DATA_DIR_NAME)
        if not os.path.exists(self.data_path):
            os.makedirs(self.data_path)

        # QSettingsを初期化し、ポータブルなiniファイルから設定をロード
        settings_path = os.path.join(PORTABLE_BASE_PATH, SETTINGS_FILE_NAME)
        self.settings = QSettings(settings_path, QSettings.Format.IniFormat)
        
        # 広告ブロッカーのインスタンスを作成
        self.ad_blocker = AdBlockInterceptor(self)
        # 広告ブロック設定を適用
        self.set_ad_blocking(self.settings.value("ad_block_enabled", True, type=bool))

        # ウィンドウへのファイルのドラッグ＆ドロップを有効化
        self.setAcceptDrops(True)

        self.default_new_tab_url = self.settings.value("default_new_tab_url", "https://www.google.com")

        # 履歴DBとブックマークのファイルパスを初期化
        self.history_db_path = os.path.join(self.data_path, "history.sqlite")
        self.bookmarks = []
        self.bookmarks_file = os.path.join(self.data_path, "bookmarks.json")

        # タブグループの情報 (名前と色) を保持する辞書
        self.groups = {}

        # 通常モードの場合のみ履歴DBの初期化とブックマークのロード
        if not self.is_private:
            self.init_history_db()
            self.load_bookmarks()

        # ダウンロードマネージャーの初期化
        self.download_manager = DownloadManager(self)

        # タブウィジェットを作成
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        self.tabs.setDocumentMode(True)
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True) # タブのドラッグ移動を有効化
        # タブのテキストが長すぎる場合に省略記号(...)を表示
        self.tabs.tabBar().setElideMode(Qt.TextElideMode.ElideRight)

        # 起動時に保存されたタブの表示位置を適用
        position_name = self.settings.value("tab_position", "上")
        position_map = {
            "左": QTabWidget.TabPosition.West,
            "右": QTabWidget.TabPosition.East,
            "上": QTabWidget.TabPosition.North,
            "下": QTabWidget.TabPosition.South,
        }
        qt_position = position_map.get(position_name, QTabWidget.TabPosition.North)
        self.tabs.setTabPosition(qt_position)
        # タブが少ないときに引き伸ばされないようにする
        self.tabs.tabBar().setExpanding(False)
        
        # タブ関連のシグナルとスロットを接続
        self.tabs.tabCloseRequested.connect(self.close_current_tab)
        # タブバーの右クリックメニュー
        self.tabs.tabBar().setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tabs.tabBar().customContextMenuRequested.connect(self.show_tab_context_menu)
        self.tabs.currentChanged.connect(self.update_navigation_state)

        # ツールバーの代わりに新しいウィジェットとレイアウトを使用
        navigation_widget = QWidget() # ナビゲーションバー用のコンテナウィジェット
        navigation_widget.setObjectName("NavigationBar") # スタイルシート適用のため
        navigation_layout = QHBoxLayout()
        navigation_layout.setContentsMargins(0, 0, 0, 0) # レイアウトの余白をなくす
        navigation_layout.setSpacing(4) # ウィジェット間の間隔を調整
        navigation_widget.setLayout(navigation_layout)

        # 戻るボタン
        self.back_button = QPushButton()
        self.back_button.setIcon(qta.icon('fa5s.arrow-left', color='#D8DEE9'))
        self.back_button.setToolTip("戻る")
        self.back_button.clicked.connect(self.go_back)
        self.back_button.setEnabled(False)
        navigation_layout.addWidget(self.back_button)

        # 進むボタン
        self.forward_button = QPushButton()
        self.forward_button.setIcon(qta.icon('fa5s.arrow-right', color='#D8DEE9'))
        self.forward_button.setToolTip("進む")
        self.forward_button.clicked.connect(self.go_forward)
        self.forward_button.setEnabled(False)
        navigation_layout.addWidget(self.forward_button)

        # リロードボタン
        self.reload_button = QPushButton()
        self.reload_button.setIcon(qta.icon('fa5s.redo', color='#D8DEE9'))
        self.reload_button.setToolTip("リロード")
        self.reload_button.clicked.connect(self.reload_page)
        navigation_layout.addWidget(self.reload_button)

        # 新しいタブを開くボタンをURLバーの左に追加
        self.new_tab_button = QPushButton()
        self.new_tab_button.setIcon(qta.icon('fa5s.plus', color='#D8DEE9'))
        self.new_tab_button.setToolTip("新しいタブを開く")
        self.new_tab_button.clicked.connect(lambda: self.add_new_tab())
        navigation_layout.addWidget(self.new_tab_button)

        # アドレスバーを作成
        self.url_bar = QLineEdit()
        self.url_bar.returnPressed.connect(self.navigate_to_url)
        self.url_bar.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        navigation_layout.addWidget(self.url_bar)

        # ハンバーガーメニューボタン
        self.menu_button = QPushButton()
        self.menu_button.setIcon(qta.icon('fa5s.bars', color='#D8DEE9'))
        self.menu_button.setToolTip("メニュー")
        main_menu = QMenu(self) # メニューボタンに表示するメニューを作成

        # 新しいタブ アクション
        new_tab_action = QAction(qta.icon('fa5s.plus'), "新しいタブ", self)
        new_tab_action.triggered.connect(lambda: self.add_new_tab())
        main_menu.addAction(new_tab_action)

        # ファイルを開く アクション
        self.open_file_action = QAction(qta.icon('fa5s.folder-open'), "ファイルを開く...", self)
        self.open_file_action.triggered.connect(self.open_file)
        main_menu.addAction(self.open_file_action)

        # 通常モードの場合のみ「新しいプライベートウィンドウ」を追加
        if not self.is_private:
            private_window_action = QAction(qta.icon('fa5s.user-secret'), "新しいプライベートウィンドウを開く", self)
            private_window_action.triggered.connect(self.open_private_window)
            main_menu.addAction(private_window_action)

        main_menu.addSeparator()

        # ブックマークアクション
        bookmark_action = QAction(qta.icon('fa5s.bookmark'), "ブックマーク", self)
        bookmark_action.triggered.connect(self.show_bookmark_window)
        main_menu.addAction(bookmark_action)

        # 履歴アクション
        history_action = QAction(qta.icon('fa5s.history'), "履歴", self)
        history_action.triggered.connect(self.show_history_window)
        main_menu.addAction(history_action)

        # ダウンロードアクション
        download_action = QAction(qta.icon('fa5s.download'), "ダウンロード", self)
        download_action.triggered.connect(self.download_manager.show)
        main_menu.addAction(download_action)

        main_menu.addSeparator()

        # 設定アクション
        settings_action = QAction(qta.icon('fa5s.cogs'), "設定", self)
        settings_action.triggered.connect(self.show_settings_dialog)
        main_menu.addAction(settings_action)

        main_menu.addSeparator()

        # アップデート確認アクション
        self.update_action = QAction("アップデートを確認", self)
        self.update_action.triggered.connect(self.manual_check_for_updates)
        main_menu.addAction(self.update_action)

        self.menu_button.setMenu(main_menu)
        navigation_layout.addWidget(self.menu_button)
        
        # QToolBarを作成し、そこにナビゲーションウィジェットを追加
        self.navigation_bar = self.addToolBar("ナビゲーション")
        # saveState/restoreStateがツールバーを正しく識別できるように、一意のオブジェクト名を設定する
        # これにより、終了時に 'objectName' not set for QToolBar という警告が出るのを防ぐ
        self.navigation_bar.setObjectName("NavigationBarToolBar")
        self.navigation_bar.addWidget(navigation_widget)
        # ナビゲーションバーを右クリックした際に表示されるメニュー項目を非表示にし、
        # 誤ってツールバーを非表示にしてしまう事故を防ぐ
        self.navigation_bar.toggleViewAction().setVisible(False)

        # デフォルトプロファイルでダウンロード要求をハンドル
        # これにより、このプロファイルを使用するすべてのQWebEngineViewインスタンスでダウンロードが捕捉される
        self.profile.downloadRequested.connect(self.handle_download_request)

        # テーマに基づいてアイコンの色などを初期化
        self.update_theme_elements()

        # セッション復元、または初期タブの追加
        session_restore_enabled = self.settings.value("session_restore_enabled", True, type=bool)
        # プライベートモードでなく、セッション復元が有効な場合
        if not self.is_private and session_restore_enabled:
            urls = self.settings.value("session/urls", [])
            current_index = self.settings.value("session/current_index", 0, type=int)
            if urls:
                for url in urls:
                    self.add_new_tab(QUrl(url), "読み込み中...")
                # 最初のデフォルトタブが残っている場合は閉じる
                if len(urls) > 0 and self.tabs.count() > len(urls): # 復元したタブの他に余分なタブがあれば
                     self.close_current_tab(0)
                if 0 <= current_index < self.tabs.count():
                    self.tabs.setCurrentIndex(current_index)
            else:
                # 復元するセッションがない場合はデフォルトページを開く
                self.add_new_tab(QUrl(self.default_new_tab_url), "新しいタブ")
        else:
            # プライベートモードまたは設定が無効な場合はデフォルトページを開く
            self.add_new_tab(QUrl(self.default_new_tab_url), "新しいタブ")

        # 起動時の広告ブロックリスト自動更新
        if not self.is_private and self.settings.value("ad_block_autoupdate_enabled", True, type=bool): # プライベートモードでなく、自動更新が有効な場合
            self.start_blocklist_update(silent=True)

        # プライベートモードのUI設定
        if self.is_private:
            self.setWindowTitle(f"{self.windowTitle()} (プライベート)")

        # 起動時にアップデートをチェック (通常ウィンドウのみ)
        if not self.is_private:
            # update_check_threadをインスタンス変数として保持しないとGCされる可能性がある
            self.update_check_thread = None
            if self.settings.value("update_check_enabled", True, type=bool):
                self.check_for_updates()
        
        # ウィンドウサイズと位置を復元 (通常ウィンドウのみ)
        if not self.is_private and self.settings.value("window_geometry_restore_enabled", True, type=bool):
            geometry = self.settings.value("windowGeometry")
            if geometry:
                self.restoreGeometry(geometry)
            state = self.settings.value("windowState")
            if state:
                self.restoreState(state)

    def set_ad_blocking(self, enabled):
        """広告ブロックの有効/無効を切り替える"""
        self.ad_block_enabled = enabled
        if enabled:
            self.profile.setUrlRequestInterceptor(self.ad_blocker)
            self.ad_blocker.load_domains() # ドメインリストを再読み込み
        else:
            self.profile.setUrlRequestInterceptor(None) # インターセプターを解除
        
        if not self.is_private: # プライベートモードでは設定を保存しない
            self.settings.setValue("ad_block_enabled", enabled)

    def init_history_db(self):
        """履歴データベースを初期化し、テーブルが存在しない場合は作成する"""
        try:
            conn = sqlite3.connect(self.history_db_path)
            cursor = conn.cursor()
            # 履歴テーブルを作成
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    url TEXT NOT NULL UNIQUE,
                    title TEXT,
                    last_visit_time TEXT NOT NULL,
                    visit_count INTEGER NOT NULL DEFAULT 1
                )
            """)
            conn.commit()
        except sqlite3.Error as e:
            print(f"履歴データベースの初期化に失敗しました: {e}")
        finally:
            if conn:
                conn.close()

    def load_bookmarks(self):
        """bookmarks.jsonファイルからブックマークを読み込むメソッド"""
        if os.path.exists(self.bookmarks_file):
            try:
                with open(self.bookmarks_file, "r", encoding="utf-8") as f:
                    self.bookmarks = json.load(f)
            except json.JSONDecodeError:
                self.bookmarks = []
        else:
            self.bookmarks = []

    def save_bookmarks(self):
        """現在のブックマークをbookmarks.jsonファイルに保存するメソッド"""
        with open(self.bookmarks_file, "w", encoding="utf-8") as f:
            json.dump(self.bookmarks, f, ensure_ascii=False, indent=4)

    def update_history_entry(self, url, title):
        """URLとタイトルを履歴データベースに追加または更新する"""
        # プライベートモードやデータ削除設定が有効な場合は何もしない
        if self.is_private or self.settings.value("privacy/clear_on_exit", False, type=bool):
            return

        url_str = url.toString()
        # about: や空のURLは保存しない
        if not url_str or url_str.startswith("about:"):
            return

        # タイトルが空の場合はURLをタイトルとして使用
        title_str = title if title else url_str
        
        try:
            conn = sqlite3.connect(self.history_db_path)
            cursor = conn.cursor()
            
            # UPSERT (INSERT or UPDATE) クエリ
            # URLが競合した場合、訪問回数をインクリメントし、タイトルと最終訪問日時を更新
            query = """
                INSERT INTO history (url, title, last_visit_time, visit_count)
                VALUES (?, ?, ?, 1)
                ON CONFLICT(url) DO UPDATE SET
                    title = excluded.title,
                    last_visit_time = excluded.last_visit_time,
                    visit_count = visit_count + 1;
            """
            cursor.execute(query, (url_str, title_str, datetime.now().isoformat()))
            conn.commit()
        except sqlite3.Error as e:
            print(f"履歴の更新に失敗しました: {e}")
        finally:
            if conn:
                conn.close()

    def clear_browsing_data(self):
        """このウィンドウに関連する閲覧データを削除する"""
        # プライベートモードではデータは保存されないので何もしない
        if self.is_private:
            return

        # Cookieとキャッシュをクリア
        self.profile.cookieStore().deleteAllCookies()
        self.profile.clearHttpCache()

        # 履歴データベースをクリア
        try:
            conn = sqlite3.connect(self.history_db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM history")
            conn.commit()
        except sqlite3.Error as e:
            print(f"履歴データベースのクリアに失敗しました: {e}")
        finally:
            if conn:
                conn.close()

        # ダウンロードマネージャーのリストをクリア
        while self.download_manager.downloads_layout.count() > 0:
            item = self.download_manager.downloads_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def _create_new_browser(self, set_as_current=True, label="新しいタブ"):
        """ヘルパー: 新しいブラウザビューとページを作成し、タブに追加して返す"""
        browser = QWebEngineView()
        page = SilentWebEnginePage(self.profile, browser)
        # 各種設定とシグナル/スロット接続
        browser.setPage(page)
        browser.setProperty("group_name", None)

        settings = browser.settings()
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.PdfViewerEnabled, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.FullScreenSupportEnabled, True)
        # JavaScript関連の設定を明示的に有効化。一部サイトのレンダリング問題を解決するため。
        settings.setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.JavascriptCanOpenWindows, True)
        settings.setAttribute(QWebEngineSettings.WebAttribute.LocalStorageEnabled, True)
        # HTTPSページでHTTPコンテンツ(スクリプトなど)の実行を許可する。Mixed Contentによる機能不全を防ぐため。
        settings.setAttribute(QWebEngineSettings.WebAttribute.AllowRunningInsecureContent, True)

        # 全画面表示リクエストをハンドル
        page.fullScreenRequested.connect(self._handle_fullscreen_request)
        # 音声状態の変更をハンドル
        page.recentlyAudibleChanged.connect(lambda audible, p=page: self.handle_audio_state_changed(p))
        page.audioMutedChanged.connect(lambda muted, p=page: self.handle_audio_state_changed(p))

        browser.urlChanged.connect(lambda q, browser=browser: self.handle_url_changed(q, browser)) # URL変更をハンドル
        browser.titleChanged.connect(lambda title, browser=browser: self.tabs.setTabText(self.tabs.indexOf(browser), title)) # タブのタイトルを更新
        # ページの読み込み進捗をハンドル
        page.loadStarted.connect(lambda browser=browser: self.handle_load_started(browser))
        page.loadProgress.connect(lambda progress, browser=browser: self.handle_load_progress(progress, browser))
        page.loadFinished.connect(lambda ok, browser=browser: self.handle_load_finished(ok, browser))

        i = self.tabs.addTab(browser, label) # タブウィジェットに追加
        if set_as_current:
            self.tabs.setCurrentIndex(i)
        self.update_tab_visuals(i)

        return browser, page

    def add_new_tab(self, qurl=None, label="新しいタブ"):
        """新しいタブを追加するメソッド"""
        if qurl is None:
            qurl = QUrl(self.default_new_tab_url)

        # ラベルがデフォルトの「新しいタブ」で、URLがローカルファイルの場合、
        # タブの初期ラベルをファイル名に設定する
        if label == "新しいタブ" and qurl.isLocalFile():
            label = qurl.fileName()

        browser, page = self._create_new_browser(label=label)
        browser.setUrl(qurl)

    def close_current_tab(self, index):
        """現在のタブを閉じるメソッド"""
        widget_to_close = self.tabs.widget(index)
        if not widget_to_close: return

        # このタブに関連付けられた開発者ツールウィンドウがあれば閉じる
        if widget_to_close in self.dev_tools_windows:
            dev_window = self.dev_tools_windows.pop(widget_to_close, None)
            if dev_window:
                dev_window.close() # 開発者ツールも一緒に閉じる

        group_name = widget_to_close.property("group_name")

        if self.tabs.count() < 2: # 最後のタブを閉じるときはウィンドウごと閉じる
            self.close()
            return

        self.tabs.removeTab(index)
        # QWebEngineViewを明示的に削除し、音声再生などを停止させる
        widget_to_close.deleteLater()

        # グループが空になったかチェック
        if group_name:
            is_group_empty = True
            for i in range(self.tabs.count()):
                if self.tabs.widget(i).property("group_name") == group_name:
                    is_group_empty = False
                    break
            if is_group_empty:
                if group_name in self.groups:
                    del self.groups[group_name]

    def go_back(self):
        """現在のタブで前のページに戻る"""
        current_browser = self.tabs.currentWidget()
        if current_browser:
            current_browser.back()

    def go_forward(self):
        """現在のタブで次のページに進む"""
        current_browser = self.tabs.currentWidget()
        if current_browser:
            current_browser.forward()

    def reload_page(self):
        """現在のタブをリロードする"""
        current_browser = self.tabs.currentWidget()
        if current_browser:
            current_browser.reload()

    def _update_spa_progress(self):
        """擬似プログレスバーの値を更新するタイマーイベント"""
        # プログレスの進み方を調整
        if self._spa_progress_value < 60:
            self._spa_progress_value += 15
        elif self._spa_progress_value < 90:
            self._spa_progress_value += 5
        else:
            self._spa_progress_value += 1
        
        if self._spa_progress_value >= 100:
            self._spa_progress_value = 100
            self._spa_progress_timer.stop()

        self.update_progress_bar(self._spa_progress_value)

    def _start_spa_progress(self):
        """擬似的なプログレスバーを開始する"""
        if self._spa_progress_timer.isActive():
            self._spa_progress_timer.stop()
        self._spa_progress_value = 0
        self.update_progress_bar(self._spa_progress_value)
        self._spa_progress_timer.start()

    def handle_url_changed(self, q, browser):
        """URLが変更されたときの処理 (SPA遷移を含む)"""
        # この変更が現在のタブで起きたものか確認
        if browser == self.tabs.currentWidget():
            # QWebEnginePageのloadStarted/loadFinishedが発行されない遷移をSPA遷移とみなす
            # loadProgressが100の状態でのURL変更はSPA遷移の可能性が高い
            progress = browser.property("loadProgress")
            if progress is None or progress == 100:
                self._start_spa_progress()

            # UIを更新
            self.update_navigation_state()

    def update_navigation_state(self):
        """現在のタブの状態に合わせてUI（URLバー、タイトル、ナビゲーションボタン）を更新する"""
        # SPAタイマーが動いていたら、タブ切り替え時に停止する
        if self._spa_progress_timer.isActive():
            self._spa_progress_timer.stop()

        current_browser = self.tabs.currentWidget()
        if current_browser:
            # URLバーとウィンドウタイトルを更新
            self.url_bar.setText(current_browser.url().toString())
            self.url_bar.setCursorPosition(0)
            self.setWindowTitle(f"{current_browser.title()} - EQUA" if current_browser.title() else "EQUA - ウェブ閲覧と使いやすさのHybrid")
            self.back_button.setEnabled(current_browser.page().history().canGoBack())
            self.forward_button.setEnabled(current_browser.page().history().canGoForward())
            # プログレスバーの状態を更新
            progress = current_browser.property("loadProgress")
            self.update_progress_bar(progress if progress is not None and progress < 100 else 100)
        else:
            self.url_bar.setText("")
            self.setWindowTitle("EQUA - ウェブ閲覧と使いやすさのHybrid")
            self.back_button.setEnabled(False)
            self.forward_button.setEnabled(False)
            self.update_progress_bar(100) # プログレスバーをリセット

    def navigate_to_url(self):
        """アドレスバーのURLに移動するメソッド"""
        current_browser = self.tabs.currentWidget()
        if not current_browser:
            return

        url_text = self.url_bar.text().strip()
        if not url_text:
            return

        # 入力されたテキストがURLか検索クエリかを判定するヒューリスティック
        # 1. ' 'を含まず'.'を含む (example.com)
        # 2. 'localhost'である
        # 3. 有効なスキームを持つ (http://..., file://...)
        is_url = (' ' not in url_text and '.' in url_text) or url_text.lower() == 'localhost'
        qurl = QUrl(url_text)
        if qurl.scheme():
            is_url = True

        if is_url:
            if not qurl.scheme(): # スキームがなければhttpsを付与
                qurl.setScheme("https")
            current_browser.setUrl(qurl)
        else:
            # 検索クエリとして処理
            search_engine_name = self.settings.value("search_engine", "Google")
            search_url_template = SEARCH_ENGINES.get(search_engine_name, "https://www.google.com/search?q={}")
            encoded_query = urllib.parse.quote_plus(url_text)
            search_url = search_url_template.format(encoded_query)
            current_browser.setUrl(QUrl(search_url))

    def show_settings_dialog(self):
        """新しいタブのデフォルトURLを設定するためのダイアログを表示するメソッド"""
        if self.settings_dialog is None: # ダイアログがまだ開かれていない場合
            self.settings_dialog = SettingsDialog(self)
            self.blocklist_update_finished.connect(self.settings_dialog.on_update_finished)
            self.settings_dialog.finished.connect(self.on_settings_dialog_closed)
            self.settings_dialog.show()
        else:
            self.settings_dialog.raise_()
            self.settings_dialog.activateWindow()

    def show_history_window(self):
        """履歴ウィンドウを開くメソッド"""
        history_dialog = HistoryWindow(self)
        history_dialog.exec()

    def show_bookmark_window(self):
        """ブックマークウィンドウを開くメソッド"""
        bookmark_dialog = BookmarkWindow(self)
        bookmark_dialog.exec()

    def on_settings_dialog_closed(self):
        """設定ダイアログが閉じられたときに参照をクリアする"""
        if self.settings_dialog: # シグナルを切断し、参照をNoneにする
            try:
                self.blocklist_update_finished.disconnect(self.settings_dialog.on_update_finished)
            except TypeError:
                pass
            self.settings_dialog = None

    def show_tab_context_menu(self, pos):
        """タブの右クリックメニューを表示する"""
        index = self.tabs.tabBar().tabAt(pos)
        if index == -1:
            return

        current_widget = self.tabs.widget(index)
        page = current_widget.page()

        menu = QMenu(self)
        
        # ミュート/ミュート解除アクション
        if page.recentlyAudible() or page.isAudioMuted():
            if page.isAudioMuted():
                mute_action = QAction(qta.icon('fa5s.volume-up', color=self.theme_colors['icon_color']), "タブのミュートを解除", self)
                mute_action.triggered.connect(lambda: page.setAudioMuted(False))
            else:
                mute_action = QAction(qta.icon('fa5s.volume-mute', color=self.theme_colors['icon_color']), "タブをミュート", self)
                mute_action.triggered.connect(lambda: page.setAudioMuted(True))
            menu.addAction(mute_action)
            menu.addSeparator()

        # 「タブをグループに追加」サブメニュー
        add_to_group_menu = menu.addMenu("タブをグループに追加")
        
        new_group_action = QAction("新しいグループ", self)
        new_group_action.triggered.connect(lambda: self.add_tab_to_new_group(index))
        add_to_group_menu.addAction(new_group_action)
        
        if self.groups:
            add_to_group_menu.addSeparator()
            for group_name in self.groups.keys():
                action = QAction(group_name, self)
                action.triggered.connect(lambda checked, name=group_name: self.add_tab_to_group(index, name))
                add_to_group_menu.addAction(action)

        # タブがグループに属している場合の操作メニュー
        current_group = current_widget.property("group_name")
        if current_group:
            menu.addSeparator()
            
            remove_action = QAction("グループから削除", self)
            remove_action.triggered.connect(lambda: self.remove_tab_from_group(index))
            menu.addAction(remove_action)
            
            rename_action = QAction("グループの名前を変更", self)
            rename_action.triggered.connect(lambda: self.handle_rename_group_from_menu(current_group))
            menu.addAction(rename_action)
            
            close_group_action = QAction(f"「{current_group}」グループのタブを閉じる", self)
            close_group_action.triggered.connect(lambda: self.close_group(current_group))
            menu.addAction(close_group_action)

        menu.addSeparator()

        # 開発者ツール
        dev_tools_action = QAction(qta.icon('fa5s.code', color=self.theme_colors['icon_color']), "開発者ツール", self)
        dev_tools_action.triggered.connect(lambda: self.open_dev_tools(index))
        menu.addAction(dev_tools_action)

        # ページキャプチャ
        capture_action = QAction(qta.icon('fa5s.camera', color=self.theme_colors['icon_color']), "ページをキャプチャ", self)
        capture_action.triggered.connect(lambda: self.capture_page(index))
        menu.addAction(capture_action)

        menu.exec(self.tabs.tabBar().mapToGlobal(pos))

    def open_file(self):
        """ファイルダイアログを開き、選択されたファイルを新しいタブで開く"""
        # ユーザーのホームディレクトリを初期位置に設定
        home_dir = os.path.expanduser("~")
        # サポートするファイルタイプをフィルタに設定
        file_filter = "ウェブページとPDF (*.html *.htm *.pdf);;すべてのファイル (*.*)"
        path, _ = QFileDialog.getOpenFileName(self, "ファイルを開く", home_dir, file_filter)

        if path:
            self.add_new_tab(QUrl.fromLocalFile(path))

    def open_dev_tools(self, index):
        """指定されたタブの開発者ツールを開く"""
        browser = self.tabs.widget(index)
        if not browser:
            return

        # 既に開いている場合はそれをアクティブにする
        if browser in self.dev_tools_windows:
            dev_window = self.dev_tools_windows.get(browser)
            if dev_window:
                dev_window.raise_()
                dev_window.activateWindow()
                return

        page = browser.page()

        # 新しいウィンドウを作成して開発者ツールを表示
        dev_window = QMainWindow(self)
        dev_window.setWindowTitle(f"開発者ツール - {browser.title()}")
        dev_tools_view = QWebEngineView()
        dev_window.setCentralWidget(dev_tools_view)

        page.setDevToolsPage(dev_tools_view.page()) # ページの開発者ツールを新しいウィンドウに表示
        
        # 開発者ツールウィンドウが閉じられたときに辞書から削除する
        dev_window.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        dev_window.destroyed.connect(lambda: self.dev_tools_windows.pop(browser, None))

        self.dev_tools_windows[browser] = dev_window
        dev_window.show()

    def capture_page(self, index):
        """指定されたタブの表示内容を画像として保存する"""
        browser = self.tabs.widget(index)
        if not browser:
            return

        # ファイル保存ダイアログを表示
        default_filename = f"capture_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        path, _ = QFileDialog.getSaveFileName(
            self, "ページをキャプチャ", default_filename, "PNG画像 (*.png);;JPEG画像 (*.jpg *.jpeg)"
        )

        if path:
            pixmap = browser.grab()
            if pixmap.save(path):
                QMessageBox.information(self, "成功", f"ページをキャプチャしました:\n{path}")
            else:
                QMessageBox.critical(self, "エラー", "キャプチャの保存に失敗しました。")

    def handle_rename_group_from_menu(self, old_name):
        """右クリックメニューからグループ名を変更するためのハンドラ"""
        new_name, ok = QInputDialog.getText(self, "グループ名の変更", f"「{old_name}」の新しい名前:", text=old_name)
        if ok and new_name and new_name != old_name and new_name not in self.groups:
            self.rename_group(old_name, new_name)

    def add_tab_to_new_group(self, index):
        """タブを新しいグループに追加する"""
        group_name, ok = QInputDialog.getText(self, "新しいグループ", "グループ名を入力してください:")
        if ok and group_name:
            if group_name in self.groups:
                self.add_tab_to_group(index, group_name) # 既存のグループに追加
            else:
                color = QColorDialog.getColor(title="グループの色を選択")
                if color.isValid():
                    self.groups[group_name] = color
                    self.add_tab_to_group(index, group_name)

    def add_tab_to_group(self, index, group_name):
        """タブを既存のグループに追加する"""
        widget = self.tabs.widget(index)
        widget.setProperty("group_name", group_name) # プロパティにグループ名を設定
        self.update_tab_visuals(index) # タブの外観を更新

    def remove_tab_from_group(self, index):
        """タブをグループから削除する"""
        widget = self.tabs.widget(index)
        widget.setProperty("group_name", None) # プロパティをリセット
        self.update_tab_visuals(index)

    def rename_group(self, old_name, new_name):
        """グループの名前を変更する（ロジックのみ）"""
        if old_name in self.groups and new_name and new_name not in self.groups:
            self.groups[new_name] = self.groups.pop(old_name)
            for i in range(self.tabs.count()):
                widget = self.tabs.widget(i)
                if widget.property("group_name") == old_name:
                    widget.setProperty("group_name", new_name)
                    self.update_tab_visuals(i)

    def change_group_color(self, group_name, new_color):
        """グループの色を変更する"""
        if group_name in self.groups:
            self.groups[group_name] = new_color
            for i in range(self.tabs.count()):
                widget = self.tabs.widget(i)
                if widget.property("group_name") == group_name:
                    self.update_tab_visuals(i)

    def ungroup_tabs(self, group_name):
        """指定されたグループを解散する（タブは閉じない）"""
        if group_name in self.groups:
            for i in range(self.tabs.count()):
                widget = self.tabs.widget(i)
                if widget.property("group_name") == group_name:
                    widget.setProperty("group_name", None)
                    self.update_tab_visuals(i)
            del self.groups[group_name]

    def close_group(self, group_name):
        """指定されたグループのタブをすべて閉じる"""
        indices_to_close = []
        for i in range(self.tabs.count()):
            if self.tabs.widget(i).property("group_name") == group_name:
                indices_to_close.append(i)
        
        # インデックスが大きい方から削除する (削除によるインデックスのズレを防ぐため)
        for i in sorted(indices_to_close, reverse=True):
            self.close_current_tab(i)

    def handle_audio_state_changed(self, page):
        """音声の状態が変化したタブのアイコンを更新する"""
        for i in range(self.tabs.count()):
            widget = self.tabs.widget(i)
            if widget and widget.page() == page:
                self.update_tab_visuals(i)
                break

    def update_tab_visuals(self, index):
        """タブの外観（グループの色のアイコン、またはデフォルトアイコン）を更新する"""
        widget = self.tabs.widget(index)
        if not widget: return
        page = widget.page()

        icon = None
        # 優先度1: 音声の状態 (ミュート/再生中)
        if page.isAudioMuted():
            icon = qta.icon('fa5s.volume-mute', color=self.theme_colors['icon_color'])
        elif page.recentlyAudible():
            icon = qta.icon('fa5s.volume-up', color=self.theme_colors['icon_color'])
        
        # 優先度2: グループの状態 (グループに属しているか)
        if not icon:
            group_name = widget.property("group_name")
            if group_name and group_name in self.groups:
                color = self.groups[group_name] # グループの色でアイコンを作成
                pixmap = QPixmap(16, 16)
                pixmap.fill(color)
                icon = QIcon(pixmap)

        # 優先度3: デフォルト
        if not icon:
            icon = qta.icon(
                'fa5s.globe-americas',
                color=self.theme_colors['icon_color'],
                color_active=self.theme_colors['icon_active_color']
            )
        
        self.tabs.setTabIcon(index, icon)

    def export_bookmarks(self):
        """ブックマークをHTMLファイルにエクスポートする"""
        path, _ = QFileDialog.getSaveFileName(self, "ブックマークをエクスポート", "", "HTMLファイル (*.html)")
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write('<!DOCTYPE NETSCAPE-Bookmark-file-1>\n')
                    f.write('<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">\n')
                    f.write('<TITLE>Bookmarks</TITLE>\n')
                    f.write('<H1>Bookmarks</H1>\n')
                    f.write('<DL><p>\n')
                    for bookmark in self.bookmarks:
                        f.write(f'    <DT><A HREF="{bookmark["url"]}">{bookmark["title"]}</A>\n')
                    f.write('</DL><p>\n')
                QMessageBox.information(self, "成功", "ブックマークのエクスポートが完了しました。")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"エクスポート中にエラーが発生しました:\n{e}")

    def import_bookmarks(self):
        """HTMLファイルからブックマークをインポートする"""
        path, _ = QFileDialog.getOpenFileName(self, "ブックマークをインポート", "", "HTMLファイル (*.html *.htm)")
        if path:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                parser = BookmarkHTMLParser()
                parser.feed(content)
                
                imported_bookmarks = parser.bookmarks # 解析結果を取得
                if not imported_bookmarks:
                    QMessageBox.warning(self, "警告", "ファイルからブックマークが見つかりませんでした。")
                    return

                existing_urls = {b['url'] for b in self.bookmarks}
                new_bookmarks_count = 0
                for bookmark in imported_bookmarks:
                    if bookmark['url'] not in existing_urls: # 重複を避ける
                        self.bookmarks.insert(0, bookmark)
                        new_bookmarks_count += 1
                
                self.save_bookmarks()
                QMessageBox.information(self, "完了", f"{new_bookmarks_count}件の新しいブックマークをインポートしました。")
            except Exception as e:
                QMessageBox.critical(self, "エラー", f"インポート中にエラーが発生しました:\n{e}")

    def handle_download_request(self, download):
        """
        ウェブエンジンからのダウンロード要求を処理し、保存ダイアログを開きます。
        """
        suggested_filename = download.suggestedFileName()
        
        default_download_path = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DownloadLocation)
        download_dir = self.settings.value("download_path", default_download_path)
        
        try:
            os.makedirs(download_dir, exist_ok=True)
        except OSError:
            QMessageBox.warning(self, "ダウンロードエラー", 
                                f"指定されたダウンロード先フォルダにアクセスできませんでした:\n{download_dir}\n\n"
                                "デフォルトのダウンロードフォルダを使用します。")
            download_dir = default_download_path
            os.makedirs(download_dir, exist_ok=True)
            
        suggested_path = os.path.join(download_dir, suggested_filename)
        file_path, _ = QFileDialog.getSaveFileName(self, "ファイルを保存", suggested_path)

        if file_path:
            if hasattr(download, 'setPath'):
                download.setPath(file_path)
            else:
                directory = os.path.dirname(file_path)
                filename = os.path.basename(file_path)
                download.setDownloadDirectory(directory)
                download.setDownloadFileName(filename)

            download.accept()
            self.download_manager.add_download_item(download)
            self.download_manager.show()
            self.download_manager.raise_()
        else:
            download.cancel()

    def change_theme(self, theme_name):
        """テーマを変更する"""
        self.settings.setValue("theme", theme_name)
        # グローバルな関数を呼び出して全ウィンドウに適用
        apply_application_theme(theme_name)

    def change_tab_position(self, position_name):
        """タブの表示位置を変更する"""
        self.settings.setValue("tab_position", position_name)
        apply_tab_position() # グローバル関数を呼び出して全ウィンドウに適用

    def change_search_engine(self, engine_name):
        """検索エンジンを変更する"""
        self.settings.setValue("search_engine", engine_name)

    def closeEvent(self, a0: QCloseEvent):
        """ウィンドウを閉じる際に履歴とブックマークを保存するイベントハンドラ"""
        # グローバルリストからこのウィンドウの参照を削除
        if self in windows:
            windows.remove(self)

        # ウィンドウを閉じる際に全画面を解除
        if self.isFullScreen():
            current_browser = self.tabs.currentWidget()
            if current_browser:
                current_browser.page().triggerAction(QWebEnginePage.WebAction.ExitFullScreen)

        # 通常モードの場合のみ各種情報を保存
        if not self.is_private:
            # 終了時にデータを削除する設定が有効な場合
            if self.settings.value("privacy/clear_on_exit", False, type=bool):
                self.clear_browsing_data()
                # セッションやウィンドウサイズは保存しないので、保存済みの情報をクリア
                self.settings.remove("session/urls")
                self.settings.remove("session/current_index")
                self.settings.remove("windowGeometry")
                self.settings.remove("windowState")
            else:
                # 通常の保存処理
                # ウィンドウのサイズと位置を保存
                if self.settings.value("window_geometry_restore_enabled", True, type=bool):
                    self.settings.setValue("windowGeometry", self.saveGeometry())
                    self.settings.setValue("windowState", self.saveState())
                # セッション情報を保存
                if self.settings.value("session_restore_enabled", True, type=bool):
                    urls = [self.tabs.widget(i).url().toString() for i in range(self.tabs.count())]
                    self.settings.setValue("session/urls", urls)
                    self.settings.setValue("session/current_index", self.tabs.currentIndex())

            self.save_bookmarks()
        super().closeEvent(a0)

    def open_private_window(self):
        """新しいプライベートウィンドウを開く"""
        # 新しいオフザレコードプロファイルでウィンドウを作成
        private_profile = QWebEngineProfile()
        # User-Agentを設定して互換性を向上
        private_profile.setHttpUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        # プライベートウィンドウも現在の広告ブロック設定に従う
        # current_ad_block_setting = self.settings.value("ad_block_enabled", True, type=bool) # この行は不要
        private_window = BrowserWindow(profile=private_profile)
        windows.append(private_window) # ウィンドウへの参照を保持
        private_window.show()

    # --- ドラッグ＆ドロップ関連のイベントハンドラ ---
    def dragEnterEvent(self, event):
        """ドラッグされたデータがウィンドウに入ったときに呼び出される"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        """データがウィンドウにドロップされたときに呼び出される"""
        for url in event.mimeData().urls():
            self.add_new_tab(url)

    def _handle_fullscreen_request(self, request):
        """ウェブページからの全画面表示リクエストを処理する"""
        request.accept() # リクエストを受け入れる
        # toggleOn() は、全画面に入るべきか(True)、出るべきか(False)を返します。
        if request.toggleOn():
            self.fullscreen_request = request # 終了時に使うのでリクエストを保持
            self._toggle_fullscreen_ui(True) # UIを全画面モードに切り替え
        else:
            self.fullscreen_request = None # リクエストをクリア
            self._toggle_fullscreen_ui(False)

    def _toggle_fullscreen_ui(self, on):
        """全画面表示に合わせてUIの表示/非表示を切り替える"""
        if on:
            self.navigation_bar.hide() # ナビゲーションバーとタブバーを非表示
            self.tabs.tabBar().hide()
            self.showFullScreen()
        else:
            self.navigation_bar.show() # 再表示
            self.tabs.tabBar().show()
            self.showNormal()

    def keyPressEvent(self, event):
        """キーボードイベントを処理する (Escキーでの全画面終了)"""
        if event.key() == Qt.Key.Key_Escape and self.isFullScreen():
            # 現在のページのWebActionをトリガーして全画面を終了させます。
            # これにより、ページから fullScreenRequested(toggleOn=False) が発行されます。
            current_browser = self.tabs.currentWidget()
            if current_browser:
                current_browser.page().triggerAction(QWebEnginePage.WebAction.ExitFullScreen)
        else:
            super().keyPressEvent(event)

    def update_theme_colors(self):
        """現在のテーマ設定に基づいて色の辞書を更新する"""
        theme_setting = self.settings.value("theme", "自動")
        actual_theme = theme_setting
        if theme_setting == "自動":
            actual_theme = get_windows_theme()
        
        self.theme_colors = THEME_COLORS.get(actual_theme, THEME_COLORS["ダーク"])
        
        # スタイルシートでカバーできない動的な色も設定
        self.theme_colors["link_color"] = "#5E81AC" if actual_theme == "ライト" else "#88C0D0"
        self.theme_colors["url_bar_bg_color"] = "#FFFFFF" if actual_theme == "ライト" else "#2E3440"
        self.theme_colors["text_color"] = "#2E3440" if actual_theme == "ライト" else "#D8DEE9"
        self.theme_colors["url_bar_border_color"] = "#D8DEE9" if actual_theme == "ライト" else "#4C566A"
        self.theme_colors["url_bar_border_focus_color"] = "#5E81AC" if actual_theme == "ライト" else "#88C0D0"

    def update_theme_elements(self):
        """テーマ変更時に、スタイルシートだけでは変わらない要素（アイコンなど）を更新する"""
        self.update_theme_colors()

        icon_color = self.theme_colors['icon_color']
        # ナビゲーションバーのアイコン
        self.back_button.setIcon(qta.icon('fa5s.arrow-left', color=icon_color))
        self.forward_button.setIcon(qta.icon('fa5s.arrow-right', color=icon_color))
        self.reload_button.setIcon(qta.icon('fa5s.redo', color=icon_color))
        self.new_tab_button.setIcon(qta.icon('fa5s.plus', color=icon_color))
        self.menu_button.setIcon(qta.icon('fa5s.user-secret' if self.is_private else 'fa5s.bars', color=icon_color))

        # メニュー内のアイコン
        if hasattr(self, 'open_file_action'):
            self.open_file_action.setIcon(qta.icon('fa5s.folder-open', color=icon_color))
        if hasattr(self, 'update_action'):
            self.update_action.setIcon(qta.icon('fa5s.sync-alt', color=icon_color))

        # 全てのタブのアイコンを更新
        for i in range(self.tabs.count()):
            self.update_tab_visuals(i)

        # 設定ダイアログが開いていれば、そちらも更新
        if self.settings_dialog:
            self.settings_dialog.update_theme_elements()

    def start_blocklist_update(self, silent=False):
        """ブロックリストの更新処理を開始する"""
        if self.update_thread and self.update_thread.isRunning():
            return

        url = self.settings.value("ad_block_update_url", DEFAULT_ADBLOCK_LIST_URL) # 設定から更新URLを取得

        self.update_thread = UpdateBlocklistThread(url)
        self.update_thread.finished.connect(
            lambda success, content, error_msg: self.finish_blocklist_update(success, content, error_msg, silent)
        )
        self.update_thread.start()

    def finish_blocklist_update(self, success, content, error_message, silent=False):
        """ブロックリストの更新処理が完了した後に呼び出される"""
        message = ""
        if success:
            try:
                user_defined_rules = []
                user_rules_marker = "[User Defined Rules]"
                # 既存のファイルからユーザー定義ルールを読み込む
                if os.path.exists(self.ad_blocker.block_list_path):
                    with open(self.ad_blocker.block_list_path, 'r', encoding='utf-8') as f:
                        in_user_section = False
                        for line in f:
                            if line.strip() == user_rules_marker:
                                in_user_section = True
                                continue # マーカー自体は含めない
                            if in_user_section:
                                user_defined_rules.append(line.strip())

                # 新しいリストとユーザー定義ルールを結合して書き込む
                with open(self.ad_blocker.block_list_path, 'w', encoding='utf-8') as f:
                    f.write(content) # ダウンロードしたコンテンツを書き込む
                    # 末尾に改行がない場合に備える
                    if content and not content.endswith('\n'):
                        f.write('\n')
                    # ユーザー定義ルールセクションを追加
                    f.write(f'\n{user_rules_marker}\n')
                    for rule in user_defined_rules:
                        if rule: # 空行は無視
                            f.write(rule + '\n')
                for window in windows:
                    if window.ad_block_enabled:
                        window.ad_blocker.load_domains() # 全ウィンドウのブロッカーをリロード
                self.settings.setValue("ad_block_last_updated", datetime.now().isoformat())
                
                if not silent: # 手動更新の場合のみメッセージを作成
                    new_domains_count = 0
                    for line in content.splitlines():
                        if line.strip() and not line.startswith(('!', '[', '#', '@')) and line.startswith('||'):
                            new_domains_count += 1
                    message = f"{new_domains_count}個のドメインでリストを更新しました。"
            except Exception as e:
                success = False
                if not silent:
                    message = f"リストの保存中にエラーが発生しました。\n\nエラー: {e}"
        else:
            if not silent:
                message = f"広告ブロックリストの更新に失敗しました。\n\nエラー: {error_message}"
        
        # 設定ダイアログに結果を通知 (silentがTrueの場合は空メッセージ)
        self.blocklist_update_finished.emit(success, message)

    def check_for_updates(self):
        """アプリケーションのアップデートを非同期でチェックする"""
        self.update_check_thread = UpdateCheckThread(GITHUB_REPO_OWNER, GITHUB_REPO_NAME)
        self.update_check_thread.finished.connect(self.on_update_check_finished)
        self.update_check_thread.start()

    def on_update_check_finished(self, success, latest_version, release_url, asset_url, error_message):
        """アップデートチェック完了時の処理"""
        if not success:
            print(f"アップデートチェックに失敗しました: {error_message}")
            return

        if not latest_version or not release_url:
            print("アップデート情報の取得に失敗しました: バージョンまたはURLが空です。")
            return

        try:
            # packaging.versionを使って、セマンティックバージョニングに沿った堅牢な比較を行う
            if parse_version(latest_version) > parse_version(__version__):
                self.show_update_notification(latest_version, release_url, asset_url)
        except InvalidVersion as e:
            print(f"バージョン番号の比較に失敗しました: current='{__version__}', latest='{latest_version}', error: {e}")

    def show_update_notification(self, new_version, release_url, asset_url):
        """アップデート通知のメッセージボックスを表示する"""
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Information)
        msg_box.setWindowTitle("アップデートのお知らせ")
        msg_box.setText(f"新しいバージョン {new_version} が利用可能です。")
        msg_box.setInformativeText("今すぐダウンロードしてアップデートしますか？")
        msg_box.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg_box.setDefaultButton(QMessageBox.StandardButton.Yes)
        if msg_box.exec() == QMessageBox.StandardButton.Yes:
            if asset_url: # OSに対応するアセットが見つかっていた場合
                self.start_update_download(asset_url)
            else:
                QMessageBox.warning(self, "エラー", "アップデート用のファイルが見つかりませんでした。ダウンロードページを開きます。")
                QDesktopServices.openUrl(QUrl(release_url))

    def start_update_download(self, asset_url):
        """アップデートファイルのダウンロードを開始する"""
        if self.update_download_dialog and self.update_download_dialog.isVisible():
            return

        # ダウンロードするファイル名を取得
        filename = os.path.basename(urllib.parse.urlparse(asset_url).path)

        # アップデートファイルは "downloads" フォルダに保存する
        save_dir = self.settings.value("download_path", os.path.join(PORTABLE_BASE_PATH, "downloads"))
        save_path = os.path.join(save_dir, filename)

        self.update_download_dialog = UpdateDownloadDialog(self)
        self.update_download_dialog.download_finished.connect(self.on_update_download_finished) # ダウンロード完了シグナルを接続
        self.update_download_dialog.start_download(asset_url, save_path)
        self.update_download_dialog.exec()
    
    def on_update_download_finished(self, success, downloaded_path):
        """アップデートファイルのダウンロード完了後の処理"""
        if success:
            QMessageBox.information(self, "ダウンロード完了",
                                    "アップデートファイルのダウンロードが完了しました。\n\n"
                                    "現在のアプリケーションを終了し、ダウンロードした新しい 'EQUA-Portable.exe' を実行してください。\n\n"
                                    "ダウンロード先フォルダを開きます。")
            download_folder = os.path.dirname(downloaded_path)
            QDesktopServices.openUrl(QUrl.fromLocalFile(download_folder))
        else:
            QMessageBox.critical(self, "アップデート失敗", "ダウンロードに失敗しました。")

    def manual_check_for_updates(self):
        """手動でアップデートを確認する"""
        # update_check_threadをインスタンス変数として保持しないとGCされる可能性がある
        self.update_check_thread = UpdateCheckThread(GITHUB_REPO_OWNER, GITHUB_REPO_NAME)
        self.update_check_thread.finished.connect(self.on_manual_update_check_finished)
        self.update_check_thread.start()

    def on_manual_update_check_finished(self, success, latest_version, release_url, asset_url, error_message):
        """手動アップデートチェック完了時の処理"""
        if not success:
            QMessageBox.warning(self, "アップデート確認", f"アップデートの確認に失敗しました。\n\nエラー: {error_message}")
            return

        if not latest_version or not release_url:
            QMessageBox.warning(self, "アップデート確認", "アップデート情報の取得に失敗しました。")
            return

        try:
            # packaging.versionを使って堅牢なバージョン比較を行う
            if parse_version(latest_version) > parse_version(__version__):
                self.show_update_notification(latest_version, release_url, asset_url)
            else:
                QMessageBox.information(self, "アップデート確認", "お使いのバージョンは最新です。")
        except InvalidVersion as e:
            QMessageBox.critical(self, "エラー", f"バージョン番号の比較に失敗しました。\n\nエラー: {e}")

    # --- ページ読み込みプログレスバー関連のハンドラ ---
    def handle_load_started(self, browser):
        """ページの読み込みが開始されたときの処理"""
        # SPA遷移用のタイマーが動いていれば停止する
        if self._spa_progress_timer.isActive():
            self._spa_progress_timer.stop()
        browser.setProperty("loadProgress", 0)
        if browser == self.tabs.currentWidget():
            self.update_progress_bar(0)

    def handle_load_progress(self, progress, browser):
        """ページの読み込みが進んだときの処理"""
        browser.setProperty("loadProgress", progress)
        if browser == self.tabs.currentWidget():
            self.update_progress_bar(progress)

    def handle_load_finished(self, ok, browser):
        """ページの読み込みが完了したときの処理"""
        # SPA遷移用のタイマーが動いていれば停止する
        if self._spa_progress_timer.isActive():
            self._spa_progress_timer.stop()
        browser.setProperty("loadProgress", 100)
        if browser == self.tabs.currentWidget():
            self.update_progress_bar(100)
        
        # 読み込みが成功した場合、履歴を更新
        if ok:
            self.update_history_entry(browser.url(), browser.title())

    def update_progress_bar(self, progress):
        """アドレスバーの背景を更新してプログレスバーとして表示する"""
        # テーマから色を取得
        bg_color = self.theme_colors.get('url_bar_bg_color')
        text_color = self.theme_colors.get('text_color')
        border_color = self.theme_colors.get('url_bar_border_color')
        border_focus_color = self.theme_colors.get('url_bar_border_focus_color')

        # 色情報がなければ何もしない
        if not all([bg_color, text_color, border_color, border_focus_color]):
            return

        # ベースとなるスタイル（パディングや角丸など）
        base_style_parts = f"""
            color: {text_color};
            border: 1px solid {border_color};
            border-radius: 13px;
            padding: 3px 12px;
            font-size: 10pt;
        """

        if progress >= 100:
            # 読み込み完了時は通常の背景色に戻す
            style = f"""
                QLineEdit {{ background-color: {bg_color}; {base_style_parts} }}
                QLineEdit:focus {{ border: 1px solid {border_focus_color}; }}
            """
        else:
            # 読み込み中はグラデーションでプログレスを表示
            progress_color = self.theme_colors.get('progress_bar_color')
            ratio = progress / 100.0
            
            # グラデーションの境界をくっきりさせるための微小なオフセット
            stop_point = ratio + 0.0001 if ratio > 0 else 0

            style = f"""
                QLineEdit {{
                    background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                        stop:0 {progress_color}, stop:{ratio} {progress_color}, 
                        stop:{stop_point} {bg_color}, stop:1 {bg_color});
                    {base_style_parts}
                }}
                QLineEdit:focus {{ border: 1px solid {border_focus_color}; }}
            """
        self.url_bar.setStyleSheet(style)

def cleanup_before_quit():
    """アプリケーション終了前にすべてのウィンドウを閉じる"""
    # イテレート中にリストが変更される可能性があるため、コピーを作成
    for window in list(windows):
        window.close()

def apply_cookie_policy():
    """アプリケーション全体のCookieポリシーを設定から読み込んで適用する"""
    settings = QSettings(os.path.join(PORTABLE_BASE_PATH, SETTINGS_FILE_NAME), QSettings.Format.IniFormat)
    # 保存されたenumの値を取得。デフォルトはAllowPersistentCookies
    default_policy_value = QWebEngineProfile.PersistentCookiesPolicy.AllowPersistentCookies.value
    policy_value = settings.value("privacy/cookie_policy_value", default_policy_value, type=int)

    # 値に対応するenumメンバーを直接生成する
    qt_policy = QWebEngineProfile.PersistentCookiesPolicy(policy_value)

    # 永続プロファイルに適用
    if persistent_profile:
        persistent_profile.setPersistentCookiesPolicy(qt_policy)
    
    # 現在開いている通常ウィンドウにも適用 (念のため)
    for window in windows:
        if not window.is_private:
            window.profile.setPersistentCookiesPolicy(qt_policy)

def apply_application_theme(theme_name):
    """アプリケーション全体にテーマを適用する"""
    actual_theme_name = theme_name if theme_name != "自動" else "ダーク"
    if theme_name == "自動": # 「自動」の場合はOSのテーマを取得
        actual_theme_name = get_windows_theme()

    # アプリケーションインスタンスにスタイルシートを適用
    QApplication.instance().setStyleSheet(THEMES.get(actual_theme_name, DARK_STYLESHEET))
    
    # タブ位置も適用
    apply_tab_position()

    # スタイルシートの変更後に、動的に色を設定しているウィジェットを更新
    for window in windows:
        window.update_theme_elements()

def apply_tab_position():
    """全ウィンドウのタブ表示位置を更新する"""
    settings = QSettings(os.path.join(PORTABLE_BASE_PATH, SETTINGS_FILE_NAME), QSettings.Format.IniFormat)
    position_name = settings.value("tab_position", "上")

    position_map = {
        "左": QTabWidget.TabPosition.West,
        "右": QTabWidget.TabPosition.East,
        "上": QTabWidget.TabPosition.North,
        "下": QTabWidget.TabPosition.South,
    }
    qt_position = position_map.get(position_name, QTabWidget.TabPosition.North)

    for window in windows:
        window.tabs.setTabPosition(qt_position)
        # タブの伸縮ポリシーを常にFalseに設定する。
        # これにより、タブの数が少ないときにタブが不必要に引き伸ばされるのを防ぎ、
        # タブの幅をスタイルシートで定義されたサイズに保つ。
        # タブが多すぎて表示しきれない場合は、自動的にスクロールボタンが表示される。
        window.tabs.tabBar().setExpanding(False)

# --- アプリケーションのエントリーポイント ---
if __name__ == '__main__':
    # QApplicationインスタンスを作成する前に、設定に基づいて環境変数を設定
    # ポータブル版では、実行ファイルと同じ場所にあるiniファイルを使用
    settings_path = os.path.join(PORTABLE_BASE_PATH, SETTINGS_FILE_NAME)
    settings = QSettings(settings_path, QSettings.Format.IniFormat)
    # デフォルトは有効(True)。設定が無効(False)の場合に --disable-gpu を追加
    if not settings.value("hw_accel_enabled", True, type=bool):
        # 既存のフラグに追記する形にする
        existing_flags = os.environ.get("QTWEBENGINE_CHROMIUM_FLAGS", "")
        os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = f"{existing_flags} --disable-gpu".strip()

    # QApplicationインスタンスを作成
    app = QApplication(sys.argv) 
    app.setApplicationName("EQUA")
    app.setOrganizationName("StudioNosa")
    app.setWindowIcon(QIcon(resource_path('equa.ico')))

    # 垂直タブのテキストを水平に描画するカスタムスタイルを適用
    # OSネイティブのスタイルに依存しないように、Fusionスタイルをベースにする
    # これにより、OSがライトモードでもアプリがダークモードの際にアイコンが黒くなる問題を回避する
    fusion_style = QStyleFactory.create("Fusion")
    app.setStyle(HorizontalTextTabStyle(fusion_style))

    # アプリケーションが終了する直前にクリーンアップ処理を接続
    app.aboutToQuit.connect(cleanup_before_quit)

    
    # 通常モード用の永続プロファイルを作成・設定します。
    # デフォルトプロファイルに依存するのではなく、専用のプロファイルを作成することで、
    # 設定が他のコンポーネントに影響されることを防ぎ、より確実にデータを保存します。
    persistent_profile = QWebEngineProfile("persistent")
    # PyInstallerでexe化した場合、__file__のパス解決が不安定になるため、
    # User-Agentを一般的なChromeのものに設定し、サイト互換性を向上させる
    # 一部のサイトでJavaScriptの動作がUser-Agentに依存する場合があるため
    persistent_profile.setHttpUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    # ポータブル化のため、プロファイルデータも実行ファイルからの相対パスに保存します。
    data_path = os.path.join(PORTABLE_BASE_PATH, DATA_DIR_NAME)
    profile_path = os.path.join(data_path, "profile_data")
    if not os.path.exists(profile_path):
        os.makedirs(profile_path)
    persistent_profile.setPersistentStoragePath(profile_path)
    apply_cookie_policy() # 設定に基づいてCookieポリシーを適用

    # テーマ設定を読み込んで適用
    current_theme = settings.value("theme", "自動")
    apply_application_theme(current_theme)

    # 作成した永続プロファイルをメインウィンドウに渡して起動します。
    main_window = BrowserWindow(profile=persistent_profile) 
    windows.append(main_window) # 最初のウィンドウの参照を保持
    main_window.show()
    
    # アプリケーションのイベントループを開始
    sys.exit(app.exec())
