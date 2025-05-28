import re
import subprocess
from pathlib import Path
import statistics

from PIL import Image
import imagehash
from tqdm import tqdm


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def make_timestamp(idx: int, interval_seconds: int) -> str:
    timestamp = idx * interval_seconds
    hours = timestamp // 3600
    minutes = (timestamp % 3600) // 60
    seconds = timestamp % 60
    return f"{hours:d}-{minutes:02d}-{seconds:02d}"


def sort_timestamp(k: Path) -> str:
    timestamp = k.name.split("_").pop()
    if not timestamp:
        return k.name
    h, x = timestamp.split("-", 1)
    return f"{int(h):04d}-{x}"


def extract_video_id(input_url_or_id: str) -> str:
    if re.fullmatch(r"[A-Za-z0-9_-]{11}", input_url_or_id):
        return input_url_or_id
    patterns = [
        r"(?:v=|\/)([A-Za-z0-9_-]{11})(?:[&?\/]|$)",
    ]
    for pat in patterns:
        m = re.search(pat, input_url_or_id)
        if m:
            return m.group(1)
    return ""


def extract_frames_ffmpeg(
    video_path: Path, frame_dir: Path, interval_seconds=3
) -> list[Path]:
    temp_pattern = frame_dir / "frame_%04d.jpg"
    ffmpeg_cmd = [
        "ffmpeg",
        "-i",
        video_path,
        "-vf",
        f"fps=1/{interval_seconds}",
        "-q:v",
        "2",
        str(temp_pattern),
    ]
    subprocess.run(ffmpeg_cmd, check=True)
    frames = sorted(frame_dir.glob("frame_[0-9][0-9][0-9][0-9].jpg"))
    renamed_frames = []
    for idx, frame_path in enumerate(frames):
        timestamp = make_timestamp(idx, interval_seconds)
        new_name = f"frame_{timestamp}.jpg"
        new_path = frame_path.with_name(new_name)
        if new_path.exists():
            new_path.unlink()
        frame_path.rename(new_path)
        renamed_frames.append(new_path)
    return renamed_frames


def filter_unique_images(
    image_paths: list[Path], hash_diff_threshold: int | None = None, fps_interval=3
) -> list[tuple[Path, int]]:
    hashes = []
    for idx, path in enumerate(tqdm(image_paths, desc="Hashing images")):
        with Image.open(path) as img:
            hash_obj = imagehash.average_hash(img)
            hashes.append((path, hash_obj, idx))
    if hash_diff_threshold is None and len(hashes) > 1:
        diffs = [abs(hashes[i][1] - hashes[i - 1][1]) for i in range(1, len(hashes))]
        mean = statistics.mean(diffs)
        stdev = statistics.stdev(diffs) if len(diffs) > 1 else 0
        hash_diff_threshold = max(1, int(mean / 2))
        print(
            f"ℹ️ Auto-calculated hash_diff_threshold: {hash_diff_threshold} using mean/2 with mean={mean:.2f} and stdev={stdev:.2f}"
        )
    elif hash_diff_threshold is None:
        hash_diff_threshold = 5
        print(f"ℹ️ Using default hash_diff_threshold: {hash_diff_threshold}")

    unique_images = []
    last_hash = None
    duplicate_start = None
    duplicate_end = None

    for i, (path, curr_hash, idx) in enumerate(hashes):
        if last_hash is None or abs(curr_hash - last_hash) > hash_diff_threshold:
            if duplicate_start is not None and duplicate_end is not None:
                start_sec = (duplicate_start - 1) * fps_interval
                end_sec = duplicate_end * fps_interval
                start_ts = f"{start_sec // 60:02.0f}:{start_sec % 60:02.0f}"
                end_ts = f"{end_sec // 60:02.0f}:{end_sec % 60:02.0f}"
                print(f"🗑️ Removed duplicate frames from {start_ts} to {end_ts}")
                duplicate_start = None
                duplicate_end = None
            seconds = idx * fps_interval
            unique_images.append((path, seconds))
            last_hash = curr_hash
        else:
            if duplicate_start is None:
                duplicate_start = idx
            duplicate_end = idx
    if duplicate_start is not None and duplicate_end is not None:
        start_sec = duplicate_start * fps_interval
        end_sec = duplicate_end * fps_interval
        start_ts = f"{start_sec // 60:02.0f}:{start_sec % 60:02.0f}"
        end_ts = f"{end_sec // 60:02.0f}:{end_sec % 60:02.0f}"
        print(f"🗑️ Removed duplicate frames from {start_ts} to {end_ts}")
    return unique_images
