import os
from pptx import Presentation
from gtts import gTTS
from pydub import AudioSegment
from pathlib import Path
import shutil

# スクリプトと同階層のffmpegを使用
_BASE_DIR = Path(__file__).parent
AudioSegment.converter = str(_BASE_DIR / "ffmpeg.exe")
AudioSegment.ffmpeg    = str(_BASE_DIR / "ffmpeg.exe")
AudioSegment.ffprobe   = str(_BASE_DIR / "ffprobe.exe")

# ──────────────────────────────────────────
# 1. PPTXの各スライドをPNG画像に変換
# ──────────────────────────────────────────
# PPTXの各スライドをPNG画像に変換します。
# LibreOfficeを使わず、python-pptx + Pillowで実装。
# Args:
#     pptx_path (str): 入力PPTXファイルパス
#     output_dir (str): 出力ディレクトリ
#     dpi (int): 画像解像度（デフォルト150）
# Returns:
#     list[str]: 生成されたPNGファイルパスのリスト（スライド順）
def pptx_to_pngs(pptx_path, output_dir="slides_png", dpi=150):
    from PIL import Image, ImageDraw, ImageFont
    os.makedirs(output_dir, exist_ok=True)
    prs = Presentation(pptx_path)

    # スライドサイズ取得（EMU → ピクセル変換）
    emu_per_inch = 914400
    slide_width_px  = int(prs.slide_width  / emu_per_inch * dpi)
    slide_height_px = int(prs.slide_height / emu_per_inch * dpi)

    png_paths = []

    for i, slide in enumerate(prs.slides):
        # スライドの背景色を取得（取得できない場合は白）
        try:
            bg_color = slide.background.fill.fore_color.rgb
            bg = tuple(int(bg_color[j:j+2], 16) for j in (0, 2, 4))
        except Exception:
            bg = (255, 255, 255)

        img = Image.new("RGB", (slide_width_px, slide_height_px), color=bg)
        draw = ImageDraw.Draw(img)

        # テキストシェイプを描画（簡易レンダリング）
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if not text:
                    continue
                # 位置をEMU → ピクセルに変換
                x = int(shape.left  / emu_per_inch * dpi) if shape.left  else 10
                y = int(shape.top   / emu_per_inch * dpi) if shape.top   else 10
                draw.text((x, y), text, fill=(0, 0, 0))

        output_path = os.path.join(output_dir, f"slide_{i+1:03d}.png")
        img.save(output_path, "PNG", dpi=(dpi, dpi))
        print(f"PNG保存: {output_path}")
        png_paths.append(output_path)

    return png_paths

# PowerPoint COMオブジェクトを使ってスライドをPNGに変換する方法（Windows限定）
# Args:
#     pptx_path (str): 入力PPTXファイルパス
#     output_dir (str): 出力ディレクトリ
# Returns:
#     list[str]: 生成されたPNGファイルパスのリスト（スライド順）
def pptx_to_pngs_com(pptx_path, output_dir="slides_png"):
    import win32com.client

    # 出力ディレクトリを削除して再作成（初期化）
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)

    abs_pptx = str(Path(pptx_path).resolve())
    abs_out  = str(Path(output_dir).resolve())

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = 1

    try:
        pres = ppt.Presentations.Open(abs_pptx)
        pres.Export(abs_out, "PNG")
        slide_count = pres.Slides.Count
        pres.Close()
    finally:
        ppt.Quit()

    # スライド順にソートしてパスリストを構築
    png_files = sorted(
        Path(abs_out).glob("*.PNG"),
        key=lambda p: p.stem  # "スライド1", "スライド2" ... の順
    )

    # slide_001.png 形式にリネーム
    png_paths = []
    for i, src in enumerate(png_files, start=1):
        dst = Path(abs_out) / f"slide_{i:03d}.png"
        src.rename(dst)
        png_paths.append(str(dst))

    return png_paths


# ──────────────────────────────────────────
# 2. スライドテキストから音声ファイルを生成
# ──────────────────────────────────────────
# 各スライドのテキストからMP3音声を生成します。
# Args:
#     pptx_path (str): 入力PPTXファイルパス
#     audio_dir (str): 音声ファイルの出力ディレクトリ
#     lang (str): 音声言語コード（デフォルト 'ja'）
#     voicevox (bool): VOICEVOX音声モード（デフォルト False）
#     voicevox=True の場合はローカルVOICEVOX APIを使用（WAV出力）。
#     voicevox=False の場合はgTTSを使用（MP3出力）。
# Returns:
#     list[str|None]: 各スライドの音声ファイルパス（空スライドはNone）
def generate_audio_files(pptx_path, audio_dir="slides_audio", lang="ja", voicevox=False, voicevoxid=3):

    # 出力ディレクトリを削除して再作成（初期化）
    if os.path.exists(audio_dir):
        shutil.rmtree(audio_dir)
    os.makedirs(audio_dir)

    prs = Presentation(pptx_path)
    audio_paths = []

    for i, slide in enumerate(prs.slides):
        # スライドノートがある場合はノートテキストを優先して取得
        notes_text = ""
        if slide.has_notes_slide:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()

        if notes_text:
            text = notes_text
            print(f"スライド {i+1}: ノートからテキスト取得")
        else:
            # スライド上の全テキストを結合して取得
            text = " ".join(
                shape.text for shape in slide.shapes
                if hasattr(shape, "text")
            ).strip()
            print(f"スライド {i+1}: スライド本文からテキスト取得")



        if text:
            # voicevoxがローカルで起動しているか確認
            voicevox = is_voicevox_running()
            if voicevox:
                import requests
                VOICEVOX_URL = "http://localhost:50021"
                VOICEVOX_SPEAKER = voicevoxid  # 話者ID

                # VOICEVOX API で音声生成（WAV）
                try:
                    query = requests.post(
                        f"{VOICEVOX_URL}/audio_query",
                        params={"text": text, "speaker": VOICEVOX_SPEAKER}
                    ).json()

                    audio_bytes = requests.post(
                        f"{VOICEVOX_URL}/synthesis",
                        params={"speaker": VOICEVOX_SPEAKER},
                        json=query
                    ).content

                    audio_path = os.path.join(audio_dir, f"audio_{i+1:03d}.wav")
                    with open(audio_path, "wb") as f:
                        f.write(audio_bytes)
                    print(f"音声保存（VOICEVOX）: {audio_path}")
                    audio_paths.append(audio_path)

                except Exception as e:
                    print(f"スライド {i+1}: VOICEVOX音声生成失敗 ({e}) → Noneとして処理")
                    audio_paths.append(None)
            else:
                tts = gTTS(text=text, lang=lang, slow=False)
                audio_path = os.path.join(audio_dir, f"audio_{i+1:03d}.mp3")
                tts.save(audio_path)
                print(f"音声保存: {audio_path}")
                audio_paths.append(audio_path)
        else:
            audio_paths.append(None)

    return audio_paths

# VOICEVOXがローカルで起動しているか確認する関数
def is_voicevox_running(url="http://localhost:50021"):
    import requests
    try:
        res = requests.get(f"{url}/version", timeout=2)
        return res.status_code == 200
    except requests.exceptions.ConnectionError:
        return False


# ──────────────────────────────────────────
# 3. 音声ファイルを結合してWAVに出力
# ──────────────────────────────────────────
# 複数のMP3音声ファイルを結合してWAVファイルを生成します。
# Args:
#     audio_paths (list[str|None]): 音声ファイルパスのリスト
#     output_path (str): 結合後のWAVファイルパス
# Returns:
#     str: 出力WAVファイルパス
def combine_audio(audio_paths, output_path="combined_audio.wav"):
    combined = AudioSegment.empty()
    for path in audio_paths:
        if path and os.path.exists(path):
            combined += AudioSegment.from_mp3(path)
        else:
            # 空スライドは2秒の無音を挿入
            combined += AudioSegment.silent(duration=2000)

    combined.export(output_path, format="wav")
    print(f"音声結合完了: {output_path}")
    return output_path


# ──────────────────────────────────────────
# 4. FFmpegでPNG + 音声を動画に合成
# ──────────────────────────────────────────
# PNG画像リストと音声ファイルをFFmpegで動画に合成します。
# 各スライドの表示時間は対応する音声の長さに合わせます。
# Args:
#     png_paths (list[str]): PNGファイルパスのリスト（スライド順）
#     audio_paths (list[str|None]): 音声ファイルパスのリスト
#     output_mp4 (str): 出力MP4ファイルパス
#     debug (bool): デバッグモード
def create_video_ffmpeg(png_paths, audio_paths, output_mp4="output.mp4", debug=False):
    import subprocess
    # FFmpegのパスを取得（同ディレクトリのffmpeg.exeを想定）
    _BASE_DIR = Path(__file__).parent
    ffmpeg_path = str(_BASE_DIR / "ffmpeg.exe")

    # 各スライドの音声長を取得
    durations = []
    for path in audio_paths:
        if path and os.path.exists(path):
            seg = AudioSegment.from_mp3(path)
            durations.append(len(seg) / 1000.0)
        else:
            durations.append(2.0)  # 空スライドは2秒
    total_duration = sum(durations)  # durationリストの合計
    print(f"合計再生時間: {total_duration:.3f} 秒")

    # concat demuxerファイルを生成（スライドごとに個別duration）
    concat_file = "concat_list.txt"
    with open(concat_file, "w", encoding="utf-8") as f:
        for png, duration in zip(png_paths, durations):
            f.write(f"file '{os.path.abspath(png)}'\n")
            f.write(f"duration {duration:.3f}\n")
        # 最後のフレームを明示（FFmpegのconcat demuxer仕様）
        f.write(f"file '{os.path.abspath(png_paths[-1])}'\n")

    # 音声を結合してWAVに出力
    combined_audio_path = "combined_audio.wav"
    combined_audio = combine_audio(audio_paths, output_path=combined_audio_path)

    print(f"結合された音声の長さ: {len(AudioSegment.from_wav(combined_audio_path)) / 1000.0} 秒")

    cmd = [
        ffmpeg_path,                # 実行するffmpegのパス（スクリプトと同階層のffmpeg.exe）
        "-y",                       # 出力ファイルが既に存在する場合、確認なしで上書き
        "-f", "concat",             # 入力フォーマットとしてconcat demuxerを使用
        "-safe", "0",               # concat_list.txt内の絶対パスを許可（デフォルトは相対パスのみ）
        "-i", concat_file,          # 入力①：concat_list.txt（スライドPNGと各表示時間を定義したファイル）
        "-i", combined_audio,       # 入力②：結合済み音声ファイル（WAV）
        "-c:v", "libx264",          # 映像コーデックにH.264を使用
        "-pix_fmt", "yuv420p",      # ピクセルフォーマットをYUV 4:2:0に変換（広い互換性のため）
        "-r", "30",                 # フレームレートを30fpsに設定
        "-t", str(total_duration),  # 動画の総再生時間を指定（concat_list.txtのduration合計）
        "-c:a", "aac",              # 音声コーデックにAACを使用
        "-map", "0:v:0",            # 入力①（concat_list.txt）の映像ストリーム0番を出力に使用
        "-map", "1:a:0",            # 入力②（combined_audio）の音声ストリーム0番を出力に使用
        output_mp4                  # 出力ファイルパス
    ]
    subprocess.run(cmd, check=True)
    print(f"動画生成完了: {output_mp4}")

    # デバッグモードのみ中間ファイルを保持
    if not debug:
        os.remove(concat_file)
        os.remove(combined_audio)
        print("中間ファイルを削除しました。")
    else:
        print(f"[DEBUG] 中間ファイルを保持しています: {concat_file}, {combined_audio}")


# ──────────────────────────────────────────
# 5. メイン処理（統合実行）
# ──────────────────────────────────────────
# PPTXファイルを動画（MP4）に変換します。
# LibreOfficeを使わず、python-pptx + Pillow + gTTS + FFmpegで実装。

# Args:
#     pptx_path (str): 入力PPTXファイルパス
#     output_mp4 (str): 出力MP4ファイルパス
#     dpi (int): PNG解像度
#     lang (str): 音声言語コード
#     png_dir (str): PNG一時保存ディレクトリ
#     audio_dir (str): 音声一時保存ディレクトリ
#     voicevox (bool): VOICEVOX音声モード
#     voicevoxid (int): VOICEVOX話者ID(デフォルト3: ずんだもんノーマル)
#     debug (bool): デバッグモード（中間ファイルを保持）
def pptx_to_video(
    pptx_path,
    output_mp4="output.mp4",
    dpi=150,
    lang="ja",
    png_dir="slides_png",
    audio_dir="slides_audio",
    voicevox=False,
    voicevoxid=3,
    debug=False,
):
    print("=== STEP 1: PNG変換 ===")
    png_paths = pptx_to_pngs_com(pptx_path, output_dir=png_dir)

    print("\n=== STEP 2: 音声生成 ===")
    audio_paths = generate_audio_files(pptx_path, audio_dir=audio_dir, lang=lang, voicevox=voicevox, voicevoxid=voicevoxid)

    print("\n=== STEP 3: 動画合成 ===")
    create_video_ffmpeg(png_paths, audio_paths, output_mp4=output_mp4, debug=debug)

    # 各スライドの音声ファイルを削除（デバッグ時は保持）
    if not debug:
        for p in png_paths:
            if p and os.path.exists(p):
                os.remove(p)
        print("画像ファイルを削除しました。")
        for p in audio_paths:
            if p and os.path.exists(p):
                os.remove(p)
        print("音声ファイルを削除しました。")
    else:
        print("[DEBUG] 画像ファイルを保持しています。")
        print("[DEBUG] 音声ファイルを保持しています。")

    print(f"\n✅ 完了: {output_mp4}")


# ──────────────────────────────────────────
# 実行
# ──────────────────────────────────────────
def main():
    args = doArgParse()
    print(f'指定された引数: {args}')
    
    pptx_to_video(
        pptx_path=args['file'],
        output_mp4=args['output'],
        dpi=args['dpi'],
        lang=args['lang'],
        voicevox=args['voicevox'],
        voicevoxid=args['voicevoxid'],
        debug=args['debug'],
    )

# 相対パス取得
def getRelativePath(filePath):
    from pathlib import Path
    if(Path(filePath).is_absolute()):
        return Path(filePath).relative_to(Path.cwd())
    else:
        return filePath

# 絶対パス取得
def getAbsolutePath(filePath):
    from pathlib import Path
    if(Path(filePath).is_absolute()):
        return filePath
    else:
        return Path(filePath).resolve()

# 引数解析
def doArgParse():
    import argparse
    parser = argparse.ArgumentParser(description='PPTXファイルをMP4動画に変換する')
    parser.add_argument('--file',   required=True,  help='ファイルパス（例: /file/to/path.pptx）')
    parser.add_argument('--output', required=True, help='出力ファイルパス（例: /file/to/path.mp4）')
    parser.add_argument('--voicevox', action='store_true', help='VOICEVOX音声モード（ローカルVOICEVOX APIを使用してWAV出力）')
    parser.add_argument('--voicevoxid', type=int, default=3, help='VOICEVOX音声モード（ローカルVOICEVOX APIを使用してWAV出力）')
    parser.add_argument('--dpi', type=int, default=150, help='PNG解像度（例: 150）')
    parser.add_argument('--lang', type=str, default='ja', help='音声言語コード（例: ja）')
    parser.add_argument('--debug', action='store_true', help='デバッグモード（中間ファイルを保持）')
    args = parser.parse_args()
    return vars(args)

if __name__ == '__main__':
    main()