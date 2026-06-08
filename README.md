slide2movie（PPTX → 音声付きMP4生成）
======================================

PPTXスライドを **画像（PNG）化** し、各スライドの **ノート（優先）or スライド文字** から **音声生成**（VOICEVOX または gTTS）して、最後に **FFmpegでMP4** に結合します。  
**FFmpegをインストール**、もしくは、**`slide2movie.exe`/`slide2movie.py` と同階層にffmpeg/ffprobeの実行ファイル**が必要です。

---

## 処理の流れ（概要）

- **引数優先順位**  
引数 → INI → default の優先順位で実行します。  
config.iniが存在しない場合、新たに生成します。
- **スライドをPNG化**  
Windows+PowerPoint(COM) → LibreOffice(headless) → python-pptx(Pillow) の順でPNG化を試します。
  - WindowsはPowerPoint COMでPNG書き出し。
  - LibreOffice headlessでPNG変換。
  - python-pptx + Pillowで簡易描画。
- **音声生成（スライドごと）**
  - ノート優先（なければスライド上テキスト結合）。
  - `voicevox=True`：VOICEVOX（ローカルAPI）でWAV出力。
  - `voicevox=False`：gTTSでMP3出力。
  - 空スライドは `None` 扱いです。（後段で無音を補います）
- **クレジットスライド対応**
  - VOICEVOXでは話者IDからキャラ/スタイル名を取得し、無音クレジット用音声も追加。
  - クレジットPNGを生成して末尾に追加。
  - `creditimg`：クレジット画像を指定。
  - `creditbg`,`creditcolor`：クレジットの背景色と文字色を指定。
- **動画合成**
  - concat demuxer用に各PNGの `duration` を作成し、FFmpegでMP4化。
  - 映像：`mpeg4`、音声：`aac`、fps=30、`yuv420p`へ変換。
  - `debug=False` の場合は中間ファイルを削除。

---

## 出力

処理中は一時生成として、基本的にスライド順で次を作ります。

- `slides_png/slide_###.png`
- `slides_audio/audio_###.mp3` または `slides_audio/audio_###.wav`

さらにクレジット用として：

- `slides_png/slide_credit.png`
- `slides_audio/audio_credit.wav`（クレジット無音 1.0秒を追加）

最終成果物：

- 指定した `output.mp4`

---

## 設定（OPTION_DEFS / config.ini）

設定項目は `external_define.py` の `OPTION_DEFS` に定義されています。  
また、`config.ini`（`CONFIG_DEFAULT` は `config.ini`）の `settings` 節から読み込む形式です。

### OPTION_DEFS（主な項目）

| オプション    | 説明                                   | default      | 必須  |
|:--------------|:---------------------------------------|:-------------|:------|
| `file`        | 入力ファイルパス                       | `input.pptx` | 必須  |
| `output`      | 出力ファイルパス                       | `output.mp4` | 必須  |
| `config`      | 設定ファイルパス                       | `config.ini` | オプション |
| `dpi`         | Pillow使用時の解像度                   | `150`        | オプション |
| `quality`     | 動画品質（1-31、値が小さいほど高品質） | `5`          | オプション |
| `lang`        | gTTS使用時の言語指定                   | `ja`         | オプション |
| `voicevox`    | VOICEVOXの使用有無                     | `False`      | オプション |
| `voicevoxid`  | VOICEVOX話者IDの指定                   | `3`          | オプション |
| `creditimg`   | クレジット画像のファイルパス           | `None`       | オプション |
| `creditbg`    | VOICEVOXのクレジットの背景色           | `None`       | オプション |
| `creditcolor` | VOICEVOXのクレジットの文字色           | `None`       | オプション |
| `debug`       | デバッグ用ログ出力と中間ファイル保存   | `False`      | オプション |

### config.ini の例

```ini
[settings]
file = input.pptx
output = output.mp4
dpi = 150
quality = 5
lang = ja
voicevox = false
voicevoxid = 3
creditimg =
creditbg =
creditcolor =
debug = false
```

- `bool` は `lowercase` で保存されます。
- `None` は空文字として扱われます。

---

## 使い方（CLI）

`slide2movie.py` を引数で実行します（引数 → INI → default の優先順位）。

例：

```bash
python slide2movie.py --file input.pptx --output output.mp4
```

configファイルを指定する場合：

```bash
python slide2movie.py --config new_config.ini
```

VOICEVOXで生成する場合：

```bash
# VOICEVOXを起動した状態で
python slide2movie.py --file input.pptx --output output.mp4 --voicevox --voicevoxid 3
```

デバッグ有効（中間ファイルを残す）：

```bash
python slide2movie.py --file input.pptx --output output.mp4 --debug
```

---

## 使い方（GUI）

`slide2movie_gui.py` は `slide2movie.py` をラップするGUIです。単体では動画生成できません。

- `OPTION_DEFS` を元に入力欄・チェックボックス・数値スピン等が自動生成されます。
- `config.ini` を読み込み、初期値が反映されます。
- 「▶ 実行」で実行ファイルを起動します。
- 実行ファイルは同梱の `slide2movie.exe` があればそれを優先し、なければ `slide2movie.py` を起動します。

---

## 注意点 / よくあるポイント

- **VOICEVOXを使うと強制的に**、話者IDに対応するキャラ/スタイル名をクレジット文言に反映します。クレジット非表示はできません。
- **空スライド**は音声生成が `None` になり得るため、結合側で **無音（2秒）** を補います
- 動画の長さは、基本的に **音声長** に合わせてPNGの表示時間が決まります
- `debug=True` を使うと中間ファイル（例：concatリストや結合音声）を残して確認しやすくなります
- 実行時にFFmpegを使用します。  
   **事前にFFmpegのインストール**、もしくは、**`slide2movie.exe`/`slide2movie.py` と同階層にffmpeg/ffprobeの実行ファイル**が必要です。  
   ffmpeg/ffprobeを同梱する場合はLGPLビルド版を使用し、以下のファイルを同梱してください。
   - COPYING_GPLv3.txt
   - COPYING-LGPLv3.txt

---

## pythonコード

- `slide2movie.py`：本体（PPTX→PNG→音声→MP4）
- `slide2movie_gui.py`：Tkinter GUI
- `external_define.py`：共通設定（OPTION_DEFS / ini入出力）