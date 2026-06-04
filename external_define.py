
from pathlib import Path
import configparser

# ──────────────
# オプション定義
# ──────────────
ARG_DESCRIPTION = 'PPTXファイルをMP4動画に変換する'
OPTION_DEFS = [
    #    name                type       default                required        store_true        frame(GUI用)            help
    dict(name='file',        type=str,  default='input.pptx',  required=True,  store_true=False, frame='ファイル設定',   help='入力PPTXファイルパス'),
    dict(name='output',      type=str,  default='output.mp4',  required=True,  store_true=False, frame='ファイル設定',   help='出力MP4ファイルパス'),
    dict(name='dpi',         type=int,  default=150,           required=False, store_true=False, frame='基本設定',       help='PNG解像度'),
    dict(name='quality',     type=int,  default=5,             required=False, store_true=False, frame='基本設定',       help='動画品質（1-31、値が小さいほど高品質）'),
    dict(name='lang',        type=str,  default='ja',          required=False, store_true=False, frame='基本設定',       help='音声言語コード'),
    dict(name='voicevox',    type=bool, default=False,         required=False, store_true=True,  frame='VOICEVOX設定',   help='VOICEVOX音声モード'),
    dict(name='voicevoxid',  type=int,  default=3,             required=False, store_true=False, frame='VOICEVOX設定',   help='VOICEVOX話者ID'),
    dict(name='creditimg',   type=str,  default=None,          required=False, store_true=False, frame='クレジット設定', help='クレジット画像パス'),
    dict(name='creditbg',    type=str,  default=None,          required=False, store_true=False, frame='クレジット設定', help='クレジット背景色（#FF6600のカラーコード）'),
    dict(name='creditcolor', type=str,  default=None,          required=False, store_true=False, frame='クレジット設定', help='クレジットテキスト色（#FF6600のカラーコード）'),
    dict(name='debug',       type=bool, default=False,         required=False, store_true=True,  frame='その他',         help='デバッグモード'),
]
CONFIG_DEFAULT = "config.ini"

# ──────────────
# iniファイルの読み書き関数
# ──────────────
# OPTION_DEFSの型定義に従ってiniファイルを読み込む
def load_ini(config_path: str) -> dict:
    config = configparser.ConfigParser()
    if not Path(config_path).exists():
        return {}

    config.read(config_path, encoding='utf-8')
    section = 'settings'
    if section not in config:
        return {}

    ini = config[section]
    result = {}

    for opt in OPTION_DEFS:
        key = opt['name']
        if key not in ini:
            continue
        if opt['type'] == int:
            result[key] = ini.getint(key)
        elif opt['type'] == bool or opt['store_true']:
            result[key] = ini.getboolean(key)
        else:
            result[key] = ini[key]

    return result

# 現在の設定値をiniとして書き出す
def save_ini(config_path: str, merged: dict) -> None:
    config = configparser.ConfigParser()
    config['settings'] = {}

    for opt in OPTION_DEFS:
        key = opt['name']
        value = merged.get(key)
        if value is None:
            config['settings'][key] = ''
        else:
            config['settings'][key] = str(value).lower() if isinstance(value, bool) else str(value)

    with open(config_path, 'w', encoding='utf-8') as f:
        config.write(f)
