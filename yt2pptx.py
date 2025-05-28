import os
import sys
import subprocess
from pathlib import Path
from PIL import Image
import imagehash
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import yt_dlp
from tqdm import tqdm
import re


def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)


def make_timestamp(idx, interval_seconds):
    timestamp = idx * interval_seconds
    hours = timestamp // 3600
    minutes = (timestamp % 3600) // 60
    seconds = timestamp % 60
    return f"{hours:d}-{minutes:02d}-{seconds:02d}"


def sort_timestamp(k: Path):
    timestamp = k.name.split("_").pop()
    if not timestamp:
        return k
    h, x = timestamp.split("-", 1)
    return f"{int(h):04d}-{x}"


def download_youtube_video(input_url_or_id, output_path=None):
    video_info = {}

    def get_info_hook(d):
        if d.get("status") == "finished":
            video_info["title"] = d["info_dict"].get("title", "video")
            video_info["id"] = d["info_dict"].get("id", "")

    # Get video info first to determine video_id
    with yt_dlp.YoutubeDL({"quiet": True}) as ydl:
        info = ydl.extract_info(input_url_or_id, download=False)
        if info is not None:
            video_info["title"] = info.get("title", "video")
            video_info["id"] = info.get("id", "")
        else:
            video_info["title"] = "video"
            video_info["id"] = ""

    video_id = video_info["id"]
    title = sanitize_filename(video_info["title"])
    out_dir = "out"
    os.makedirs(out_dir, exist_ok=True)
    final_path = os.path.join(out_dir, f"{video_id}.mp4")

    if os.path.exists(final_path):
        print(f"✅ Video already downloaded: {final_path}")
        return final_path, title, video_id

    temp_path = final_path  # Download directly to final_path

    ydl_opts = {
        "outtmpl": temp_path,
        "format": "bestvideo+bestaudio/best",
        "merge_output_format": "mp4",
        "progress_hooks": [get_info_hook],
        "quiet": True,
    }

    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        ydl.download([input_url_or_id])

    return final_path, title, video_id


def extract_frames_ffmpeg(video_path, output_folder, interval_seconds=3):
    os.makedirs(output_folder, exist_ok=True)
    temp_pattern = os.path.join(output_folder, "frame_%04d.jpg")
    ffmpeg_cmd = [
        "ffmpeg",
        "-i",
        video_path,
        "-vf",
        f"fps=1/{interval_seconds}",
        "-q:v",
        "2",
        temp_pattern,
    ]
    subprocess.run(ffmpeg_cmd, check=True)

    # Rename frames to include timestamp
    frames = sorted(Path(output_folder).glob("frame_*.jpg"))
    renamed_frames = []
    for idx, frame_path in enumerate(frames):
        timestamp = make_timestamp(idx, interval_seconds)
        new_name = f"frame_{timestamp}.jpg"
        new_path = frame_path.with_name(new_name)
        frame_path.rename(new_path)
        renamed_frames.append(new_path)
    return renamed_frames


def filter_unique_images(image_paths, hash_diff_threshold=5, fps_interval=3):
    hashes = []
    for idx, path in enumerate(tqdm(image_paths, desc="Hashing images")):
        with Image.open(path) as img:
            hashes.append((path, imagehash.average_hash(img), idx))

    unique_images = []
    last_hash = None
    duplicate_start = None
    duplicate_end = None

    for i, (path, curr_hash, idx) in enumerate(hashes):
        if last_hash is None or curr_hash - last_hash > hash_diff_threshold:
            # Print duplicate range if any
            if duplicate_start is not None and duplicate_end is not None:
                start_sec = duplicate_start * fps_interval
                end_sec = duplicate_end * fps_interval
                start_ts = f"{start_sec // 60:02.0f}:{start_sec % 60:02.0f}"
                end_ts = f"{end_sec // 60:02.0f}:{end_sec % 60:02.0f}"

                frame_range = (
                    f"{start_ts}-{end_ts}" if start_sec != end_sec else start_ts
                )

                print(f"🗑️ Removed duplicate frames {frame_range}")
                duplicate_start = None
                duplicate_end = None

            seconds = idx * fps_interval
            unique_images.append((path, seconds))
            last_hash = curr_hash
        else:
            if duplicate_start is None:
                duplicate_start = idx
            duplicate_end = idx

    # Print any remaining duplicate range at the end
    if duplicate_start is not None and duplicate_end is not None:
        start_sec = duplicate_start * fps_interval
        end_sec = duplicate_end * fps_interval
        start_ts = f"{start_sec // 60:02.0f}:{start_sec % 60:02.0f}"
        end_ts = f"{end_sec // 60:02.0f}:{end_sec % 60:02.0f}"
        print(f"🗑️ Removed duplicate frames from {start_ts} to {end_ts}")

    return unique_images


def create_pptx_from_images_with_timestamps(
    image_tuples, output_pptx, video_id, fps_interval
):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]

    for img_path, seconds in tqdm(image_tuples, desc="Creating slides"):
        slide = prs.slides.add_slide(blank_slide_layout)
        slide.shapes.add_picture(
            str(img_path), Inches(0), Inches(0), width=prs.slide_width
        )

        # Timestamp and hyperlink
        timestamp = f"{seconds // 60:02.0f}:{seconds % 60:02.0f}"
        youtube_link = f"https://www.youtube.com/watch?v={video_id}&t={seconds}s"

        left = Inches(0.3)
        top = prs.slide_height - Inches(0.7)  # type: ignore
        width = Inches(3)
        height = Inches(0.5)

        textbox = slide.shapes.add_textbox(left, top, width, height)  # type: ignore
        text_frame = textbox.text_frame
        p = text_frame.paragraphs[0]
        run = p.add_run()
        run.text = f"Jump to {timestamp} on YouTube"

        font = run.font
        font.size = Pt(12)
        font.bold = True
        font.underline = True
        font.color.rgb = RGBColor(0, 102, 204)

        run.hyperlink.address = youtube_link

    prs.save(output_pptx)
    print(f"✅ PowerPoint saved: {output_pptx}")


def parse_args(argv):
    input_url_or_id = None
    custom_base = None
    fps_interval = 3  # default

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
    return input_url_or_id, custom_base, fps_interval


# === MAIN ===
if __name__ == "__main__":
    input_url_or_id, custom_base, fps_interval = parse_args(sys.argv)
    if not input_url_or_id:
        print(
            "❌ Usage: python yt2pptx.py <YouTube_URL_or_ID> [output_base_name] [-i=SECONDS|--interval=SECONDS]"
        )
        sys.exit(1)

    out_dir = "out"
    os.makedirs(out_dir, exist_ok=True)

    print("🔽 Downloading video...")
    video_file, video_title, video_id = download_youtube_video(input_url_or_id)
    base_name = custom_base or video_title

    frames_folder = os.path.join(out_dir, f"{video_id}_frames")
    interval_marker = os.path.join(frames_folder, f".interval_{fps_interval}s")

    pptx_output = os.path.join(out_dir, f"{base_name}.pptx")

    # Only extract frames if not already done with this interval
    if not (os.path.isdir(frames_folder) and os.path.isfile(interval_marker)):
        print("🎞 Extracting frames...")
        extracted_images = extract_frames_ffmpeg(
            video_file, frames_folder, interval_seconds=fps_interval
        )
        # Mark extraction interval
        with open(interval_marker, "w") as f:
            f.write(str(fps_interval))
    else:
        print(
            f"🎞 Using previously extracted frames in '{frames_folder}' (interval={fps_interval}s)"
        )
        extracted_images = sorted(Path(frames_folder).glob("frame_*.jpg"))

    print("🧹 Filtering duplicate frames...")
    unique_images = filter_unique_images(extracted_images, fps_interval=fps_interval)

    print("🧾 Creating PowerPoint...")
    create_pptx_from_images_with_timestamps(
        unique_images, pptx_output, video_id, fps_interval
    )
