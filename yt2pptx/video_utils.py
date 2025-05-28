import re
import subprocess
from pathlib import Path
import statistics

from PIL import Image
import imagehash
from tqdm import tqdm


def sanitize_filename(name: str) -> str:
    """Sanitize a filename by replacing invalid characters with underscores.

    This function replaces characters that are not allowed in filenames
    on most file systems, such as Windows and Unix-like systems.

    Args:
        name (str): The name to sanitize.

    Returns:
        str: Sanitized filename with invalid characters replaced by underscores.
    """
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def make_timestamp(timestamp: int, is_filename=False) -> str:
    """Convert a timestamp in seconds to a formatted string.

    The format is "h:mm:ss" or "m:ss" if the hour is zero.
    If `is_filename` is True, it uses '-' instead of ':' for compatibility with filenames.

    Args:
        timestamp (int): The timestamp in seconds to convert.
        is_filename (bool, optional): If True, formats the timestamp for use in filenames.
        Defaults to False.

    Returns:
        str: Formatted timestamp string.
    """
    hour = timestamp // 3600
    format_dict = {
        "h": hour,
        "m": (timestamp % 3600) // 60,
        "s": timestamp % 60,
        "x": "-" if is_filename else ":",
    }

    format_string = "{h:d}{x}{m:02d}{x}{s:02d}"
    if not is_filename and hour == 0:
        format_string = "{m:d}{x}{s:02d}"

    return format_string.format(**format_dict)


def sort_timestamp(k: Path) -> str:
    """Extract and format the timestamp from a filename for sorting purposes.

    This function assumes the filename contains a timestamp in the format
    "h-mm-ss" or "m-ss", where 'h' is hours, 'm' is minutes, and 's' is seconds.
    If the timestamp is not present, it returns the original filename.

    Args:
        k (Path): The Path object representing the file.

    Returns:
        str: Formatted timestamp string that can be used for sorting,
        or the original filename if no timestamp is found.
    """
    timestamp = k.name.split("_").pop()
    if not timestamp:
        return k.name
    h, x = timestamp.split("-", 1)

    return f"{int(h):04d}-{x}"


def extract_video_id(input_url_or_id: str) -> str:
    """Extract the YouTube video ID from a URL or ID string.

    This function checks if the input is already a valid video ID (11 characters long).
    If it is not, it uses regular expressions to search for the video ID in the input string.
    If no valid ID is found, it returns an empty string.

    Args:
        input_url_or_id (str): The input string that may contain a YouTube video ID or URL.

    Returns:
        str: The extracted YouTube video ID if found, otherwise an empty string.
    """
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
    """Extract frames from a video file using ffmpeg.

    This function uses ffmpeg to extract frames from a video file at specified intervals.
    It saves the frames in a specified directory with a naming pattern that includes the timestamp
    in the format "frame_h-mm-ss.jpg".
    The frames are extracted at a rate of one frame every `interval_seconds`.

    ffmpeg is required to be installed and available in the system PATH.

    Args:
        video_path (Path): Path to the video file from which frames will be extracted.
        frame_dir (Path): Directory where the extracted frames will be saved.
        interval_seconds (int, optional): Interval in seconds at which frames will be extracted.
        Defaults to 3 seconds.

    Returns:
        list[Path]: A list of Path objects representing the extracted frames.
    Raises:
        subprocess.CalledProcessError: If the ffmpeg command fails to execute.
    """
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
        timestamp = make_timestamp(idx * interval_seconds, True)
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
    """Filter unique images based on perceptual hashing.

    This function computes the average perceptual hash for each image in the provided list.
    It then compares the hashes of consecutive images to determine if they are unique
    based on a specified threshold. If the difference between hashes is greater than the
    threshold, the image is considered unique and added to the result list.

    Args:
        image_paths (list[Path]): List of image file paths to be processed.
        hash_diff_threshold (int | None, optional): Threshold for hash difference to consider images unique.
        If None, it will be auto-calculated based on the mean and standard deviation of hash differences.
        Defaults to None.
        fps_interval (int, optional): Interval in seconds for the frame rate, used to calculate timestamps.
        Defaults to 3 seconds.

    Returns:
        list[tuple[Path, int]]: A list of tuples where each tuple contains the path to a unique image
        and the corresponding timestamp in seconds.
    Raises:
        ValueError: If no images are provided in the image_paths list.
        statistics.StatisticsError: If there is not enough data to calculate mean or standard deviation.
    """
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

    def remove_duplicates():
        nonlocal duplicate_start, duplicate_end
        if duplicate_start is not None and duplicate_end is not None:
            start_sec = (duplicate_start - 1) * fps_interval
            end_sec = duplicate_end * fps_interval
            start_ts = make_timestamp(start_sec)
            end_ts = make_timestamp(end_sec)
            print(f"🗑️ Removed duplicate frames from {start_ts} to {end_ts}")
            duplicate_start = None
            duplicate_end = None

    for path, curr_hash, idx in hashes:
        if last_hash is None or abs(curr_hash - last_hash) > hash_diff_threshold:
            remove_duplicates()
            seconds = idx * fps_interval
            unique_images.append((path, seconds))
            last_hash = curr_hash
        else:
            if duplicate_start is None:
                duplicate_start = idx
            duplicate_end = idx
    remove_duplicates()

    return unique_images
