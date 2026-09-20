import csv
import io
import json
import os
import re
import subprocess
import urllib.request
import time
from pathlib import Path

# --- 設定 ---
# 【方法1】Googleスプレッドシートの「Webに公開(CSV)」URLを設定する場合（ここにURLを貼るだけでCSV出力不要に！）
GOOGLE_SHEETS_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRY-pZ4TwzwctKR0TRSO8xlwdRn1xc7xESgCsu2BeW-J8r7U3ri_GeO5WVmbGue8NO_1m0xYY1t1WUM/pub?gid=689538152&single=true&output=csv"

# 【方法2】ローカルのExcel/CSVファイルを使う場合（URLが空欄のときに自動使用されます）
EXCEL_FILE = "songs.xlsx"
CSV_FILE = "songs.csv"

OUTPUT_JSON = "tracks.json"
ASSETS_DIR = Path("assets")

ASSETS_DIR.mkdir(exist_ok=True)


def parse_time_to_seconds(time_str) -> int | None:
    """時間文字列（0:30 や 90 など）を秒数に変換"""
    if time_str is None or not str(time_str).strip():
        return None

    time_str = str(time_str).strip()

    # 数値（または float 相当の文字列）の場合
    try:
        val = float(time_str)
        return int(val)
    except ValueError:
        pass

    parts = time_str.split(":")
    if len(parts) == 2:
        minutes, seconds = map(int, parts)
        return minutes * 60 + seconds
    elif len(parts) == 3:
        hours, minutes, seconds = map(int, parts)
        return hours * 3600 + minutes * 60 + seconds

    return None


def parse_volume(vol_str, default: float = 1.0) -> float:
    """音量設定の変換"""
    if vol_str is None or not str(vol_str).strip():
        return default

    vol_clean = str(vol_str).strip().rstrip("%")
    try:
        val = float(vol_clean)
        if val > 1.0:
            val = val / 100.0
        return max(0.0, min(1.0, val))
    except ValueError:
        return default


def is_checked(val) -> bool:
    """チェックが入っているか判定"""
    if not val:
        return False
    val_clean = str(val).strip().lower()
    return val_clean in [
        "1",
        "1.0",
        "true",
        "x",
        "o",
        "y",
        "yes",
        "v",
        "✓",
        "check",
    ]


def extract_youtube_id(url: str) -> str:
    """YouTube ID の抽出"""
    match = re.search(r"(?:v=|\/)([a-zA-Z0-9_-]{11})", url)
    if match:
        return match.group(1)
    return ""


def download_thumbnail(yt_id: str, save_path: Path):
    """サムネイル画像取得"""
    urls = [
        f"https://img.youtube.com/vi/{yt_id}/maxresdefault.jpg",
        f"https://img.youtube.com/vi/{yt_id}/hqdefault.jpg",
    ]
    for url in urls:
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": "Mozilla/5.0"}
            )
            with (
                urllib.request.urlopen(req) as resp,
                open(save_path, "wb") as out_file,
            ):
                out_file.write(resp.read())
            if save_path.stat().st_size > 3000:
                return
        except Exception:
            continue


def load_rows() -> tuple[list[dict], str]:
    """データソース（Googleスプレッドシート / Excel / CSV）から行データを取得"""
    
    def get_check_col(fieldnames: list[str]) -> str:
        if not fieldnames:
            return ""
        # 優先度の高いチェック列名を探す
        for col in fieldnames:
            if col.strip().lower() in ["update", "check", "更新", "dl", "取得", "フラグ"]:
                return col
        return fieldnames[-1] # 見つからなければ従来の通り末尾の列を使用

    # 1. Googleスプレッドシートから読み込み
    if GOOGLE_SHEETS_URL.strip():
        print("🌐 Googleスプレッドシートからデータを取得中...")
        req = urllib.request.Request(
            GOOGLE_SHEETS_URL, headers={"User-Agent": "Mozilla/5.0"}
        )
        with urllib.request.urlopen(req) as resp:
            content = resp.read().decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(content))
        fieldnames = reader.fieldnames or []
        return list(reader), get_check_col(fieldnames)

    # 3. CSV (.csv) から読み込み
    if os.path.exists(CSV_FILE):
        print(f"📄 {CSV_FILE} を読み込み中...")
        with open(CSV_FILE, mode="r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            fieldnames = reader.fieldnames or []
            return list(reader), get_check_col(fieldnames)

    print(
        "❌ エラー: GoogleスプレッドシートのURL設定、songs.xlsx、songs.csv のいずれも見つかりません。"
    )
    return [], ""


def process_song(index: int, row: dict, check_col_name: str) -> dict | None:
    # CSVの通し番号列（id, no, 通し番号等）があればそれを読み込み、無ければループ番号を使用
    raw_id = row.get("id") or row.get("no") or row.get("通し番号") or index
    try:
        track_id = f"{int(raw_id):03d}"
    except (ValueError, TypeError):
        track_id = f"{index:03d}"

    url = str(row.get("url", "") or "").strip()
    title = str(row.get("title", "") or "").strip() or f"Track {track_id}"
    artist = str(row.get("artist", "") or "").strip() or "Unknown Artist"

    if not url:
        return None

    start_time_val = parse_time_to_seconds(row.get("start_time"))
    end_time_val = parse_time_to_seconds(row.get("end_time"))

    # 秒数が未入力の曲はスキップ
    if start_time_val is None and end_time_val is None:
        return None

    start_time = start_time_val if start_time_val is not None else 0
    end_time = end_time_val if end_time_val is not None else start_time + 30

    volume = parse_volume(row.get("volume"))

    yt_id = extract_youtube_id(url)
    if not yt_id:
        print(f"[{track_id}] スキップ: 有効なYouTube URLではありません ({url})")
        return None

    audio_filename = f"{track_id}.mp3"
    cover_filename = f"{track_id}.jpg"

    audio_path = ASSETS_DIR / audio_filename
    cover_path = ASSETS_DIR / cover_filename

    raw_check_val = row.get(check_col_name, "")
    should_update = is_checked(raw_check_val)
    files_exist = audio_path.exists() and cover_path.exists()

    # チェックがなく、かつファイル未作成の場合は完全スキップ
    if not should_update and not files_exist:
        print(f"[{track_id}] スキップ (未取得・チェックなし): {title} - {artist}")
        return None

    # チェックが入っている場合はダウンロード・更新実行
    if should_update:
        reason = (
            "新規作成 (チェックあり)"
            if not files_exist
            else "更新・再作成 (チェックあり)"
        )
        print(
            f"\n[{track_id}] {reason}: {title} - {artist} ({start_time}秒〜{end_time}秒)"
        )

        section_arg = f"*{start_time}-{end_time}"
        cmd = [
            "yt-dlp",
            "-x",
            "--audio-format",
            "mp3",
            "--postprocessor-args",
            f"ffmpeg:-ss {start_time} -to {end_time}",  # ← ffmpeg側で確実に切り出し
            "--force-overwrites",
            "-o",
            str(audio_path),
            url,
        ]

        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print(f"   ✓ MP3更新完了")
                break
            except subprocess.CalledProcessError as e:
                if attempt < max_retries:
                    print(f"   ⚠️ MP3取得一時エラー ({attempt}/{max_retries}回目)。2秒後に再試行します...")
                    time.sleep(2)
                else:
                    print(f"   ✗ MP3取得失敗 (最大リトライ回数到達): {e}")

        download_thumbnail(yt_id, cover_path)
        print(f"  ✓ ジャケット画像更新完了")
    else:
        print(f"[{track_id}] 既存ファイル使用: {title} - {artist}")

    track_data = {
        "id": track_id,
        "title": title,
        "artist": artist,
        "youtubeId": yt_id,
        "cover": f"assets/{cover_filename}",
        "audio": f"assets/{audio_filename}",
        "startTime": 0,
        "volume": volume,
    }

    return track_data


def main():
    rows, check_col_name = load_rows()
    if not rows:
        return

    tracks = []
    for i, row in enumerate(rows, start=1):
        track_data = process_song(i, row, check_col_name)
        if track_data:
            tracks.append(track_data)

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(tracks, f, ensure_ascii=False, indent=2)

    print(f"\n🎉 完了！計 {len(tracks)} 曲を `{OUTPUT_JSON}` に出力しました。")


if __name__ == "__main__":
    main()