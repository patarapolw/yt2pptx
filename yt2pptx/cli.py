import sys
from pathlib import Path
import platform
from os import startfile
import subprocess

import yt_dlp

from .video_utils import (
    sanitize_filename,
    extract_video_id,
    extract_frames_ffmpeg,
    filter_unique_images,
)
from .pptx_utils import create_pptx_from_images_with_timestamps


def download_youtube_video(
    input_url_or_id: str, out_dir: Path
) -> tuple[Path, str, str]:
    """Download a YouTube video and return the path, title, and video ID.

    If the video is already downloaded, it will return the existing path and title.

    Args:
        input_url_or_id (str): YouTube video URL or ID.
        out_dir (Path): Output directory where the video will be saved.

    Returns:
        tuple[Path, str, str]: A tuple containing the path to the downloaded video,
        the title of the video, and the video ID.
    """
    video_id = extract_video_id(input_url_or_id)
    video_folder = out_dir / video_id
    video_folder.mkdir(exist_ok=True)
    title_file = video_folder / ".title.txt"
    final_path = out_dir / f"{video_id}.mp4"
    if final_path.exists() and title_file.exists():
        title = title_file.read_text("utf-8").strip()
        print(f"✅ Video already downloaded: {final_path} with title '{title}'")
        return final_path, title, video_id

    video_info = {}

    def get_info_hook(d: dict) -> None:
        """Hook to get video information during download."""
        if d.get("status") == "finished":
            video_info["title"] = d["info_dict"].get("title", "video")
            video_info["id"] = d["info_dict"].get("id", "")

    ydl_opts = {
        "outtmpl": str(final_path),
        "format": "/".join(
            [
                "best[width<=?1920][height<=?1080]",
                "bestvideo[width<=?1920][height<=?1080]+bestaudio",
            ]  # ][::-1]
        ),
        "merge_output_format": "mp4",
        "progress_hooks": [get_info_hook],
        "quiet": True,
        # "sub_langs": "zh-Hans",
    }
    print("🔽 Downloading video...")

    with yt_dlp.YoutubeDL(ydl_opts) as ydl:  # pyright: ignore[reportArgumentType]
        info = ydl.extract_info(input_url_or_id, download=True)
        if info is not None:
            video_info["title"] = info.get("title", "video")
            video_info["id"] = info.get("id", "")
        else:
            video_info["title"] = "video"
            video_info["id"] = ""
    title = sanitize_filename(str(video_info["title"]))
    title_file.write_text(title, encoding="utf-8")

    return final_path, title, video_id


def youtube_to_pptx_cache_frames(
    out_dir: Path,
    video_file: Path,
    pptx_output: Path,
    *,  # Use * to enforce keyword-only arguments after this point
    fps_interval: int,
) -> None:
    video_id = video_file.stem
    frames_folder = out_dir / video_id

    print(f"🎞 Extracting frames using interval: {fps_interval} seconds")
    extracted_images = extract_frames_ffmpeg(
        video_file, frames_folder, interval_seconds=fps_interval
    )
    print("🧹 Filtering duplicate frames...")
    unique_images, _ = filter_unique_images(extracted_images, fps_interval=fps_interval)

    print("🧾 Creating PowerPoint...")
    create_pptx_from_images_with_timestamps(
        unique_images, pptx_output, video_id, video_path=video_file
    )


def parse_args(argv: list[str]) -> tuple[str, str | None, int | None]:
    """Parse command line arguments.

    Args:
        argv (list[str]): List of command line arguments.
    If the first argument is a YouTube URL or ID, it will be used as the input.
    If the second argument is provided, it will be used as the custom base name for the output file.
    Optionally, an interval can be specified with -i=SECONDS or --interval=SECONDS.
    If no interval is specified, it defaults to 3 seconds.
    If no input URL or ID is provided, it will print usage instructions and exit.

    Returns:
        tuple[str, str | None, int]: A tuple containing the input URL or ID,
        the custom base name (or None if not provided), and the FPS interval in seconds.
    """
    input_url_or_id = None
    custom_base = None
    fps_interval = None
    for arg in argv[1:]:
        if arg.startswith("-i=") or arg.startswith("--interval="):
            try:
                fps_interval = int(arg.split("=", 1)[1])
            except Exception:
                print("❌ Invalid interval value. Use -i=SECONDS or --interval=SECONDS")
                sys.exit(1)
        elif input_url_or_id is None:
            input_url_or_id = arg
        elif custom_base is None:
            custom_base = sanitize_filename(arg)

    return input_url_or_id or "", custom_base, fps_interval


def main() -> None:
    """Main function to run the command line interface for yt2pptx."""
    input_url_or_id, custom_base, fps_interval = parse_args(sys.argv)
    if not input_url_or_id:
        print(
            "❌ Usage: python -m yt2pptx.cli <YouTube_URL_or_ID> [output_base_name] [-i=SECONDS|--interval=SECONDS]"
        )
        sys.exit(1)

    if fps_interval is None:
        fps_interval = 2  # Default second interval

    out_dir = Path() / "out"
    out_dir.mkdir(exist_ok=True)
    video_file, video_title, _ = download_youtube_video(input_url_or_id, out_dir)

    base_name = custom_base or video_title
    pptx_output = out_dir / f"{base_name}.pptx"

    youtube_to_pptx_cache_frames(
        out_dir, video_file, pptx_output, fps_interval=fps_interval
    )

    def open_pptx(filepath):
        if platform.system() == "Windows":
            startfile(filepath)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", filepath])
        else:  # Linux and others
            subprocess.run(["xdg-open", filepath])

    open_pptx(pptx_output)


if __name__ == "__main__":
    main()
