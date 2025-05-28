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


def download_youtube_video(input_url_or_id, output_path=None):
    video_info = {}

    def get_info_hook(d):
        if d.get("status") == "finished":
            video_info["title"] = d["info_dict"].get("title", "video")
            video_info["id"] = d["info_dict"].get("id", "")

    temp_path = "temp_video.mp4" if output_path is None else output_path

    ydl_opts = {
        "outtmpl": temp_path,
        "format": "bestvideo+bestaudio/best",
        "merge_output_format": "mp4",
        "progress_hooks": [get_info_hook],
        "quiet": True,
    }

    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
        info = ydl.extract_info(input_url_or_id, download=True)
        video_info["title"] = info.get("title", "video")
        video_info["id"] = info.get("id", "")

    title = sanitize_filename(video_info["title"])
    final_path = output_path or f"{title}.mp4"

    if temp_path != final_path:
        os.rename(temp_path, final_path)

    return final_path, title, video_info["id"]


def extract_frames_ffmpeg(video_path, output_folder, interval_seconds=3):
    os.makedirs(output_folder, exist_ok=True)
    output_pattern = os.path.join(output_folder, "frame_%04d.jpg")
    ffmpeg_cmd = [
        "ffmpeg",
        "-i",
        video_path,
        "-vf",
        f"fps=1/{interval_seconds}",
        "-q:v",
        "2",
        output_pattern,
    ]
    subprocess.run(ffmpeg_cmd, check=True)
    return sorted(Path(output_folder).glob("frame_*.jpg"))


def filter_unique_images(image_paths, hash_diff_threshold=5):
    unique_images = []
    prev_hash = None
    for path in tqdm(image_paths, desc="Filtering duplicates"):
        with Image.open(path) as img:
            curr_hash = imagehash.average_hash(img)
            if prev_hash is None or abs(curr_hash - prev_hash) > hash_diff_threshold:
                unique_images.append(path)
                prev_hash = curr_hash
            else:
                os.remove(path)
    return unique_images


def create_pptx_from_images_with_timestamps(
    image_paths, output_pptx, video_id, fps_interval
):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]

    for idx, img_path in enumerate(tqdm(image_paths, desc="Creating slides")):
        slide = prs.slides.add_slide(blank_slide_layout)
        slide.shapes.add_picture(
            str(img_path), Inches(0), Inches(0), width=prs.slide_width
        )

        # Timestamp and hyperlink
        seconds = idx * fps_interval
        timestamp = f"{seconds // 60:02.0f}:{seconds % 60:02.0f}"
        youtube_link = f"https://www.youtube.com/watch?v={video_id}&t={seconds}s"

        left = Inches(0.3)
        top = prs.slide_height - Inches(0.7)
        width = Inches(3)
        height = Inches(0.5)

        textbox = slide.shapes.add_textbox(left, top, width, height)
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


# === MAIN ===
if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("❌ Usage: python yt2pptx.py <YouTube_URL_or_ID> [output_base_name]")
        sys.exit(1)

    input_url_or_id = sys.argv[1]
    custom_base = sanitize_filename(sys.argv[2]) if len(sys.argv) > 2 else None
    fps_interval = 3  # seconds between extracted frames

    out_dir = "out"
    os.makedirs(out_dir, exist_ok=True)

    print("🔽 Downloading video...")
    video_file, video_title, video_id = download_youtube_video(input_url_or_id)
    base_name = custom_base or video_title

    pptx_output = os.path.join(out_dir, f"{base_name}.pptx")
    frames_folder = os.path.join(out_dir, f"{base_name}_frames")

    print("🎞 Extracting frames...")
    extracted_images = extract_frames_ffmpeg(
        video_file, frames_folder, interval_seconds=fps_interval
    )

    print("🧹 Filtering duplicate frames...")
    unique_images = filter_unique_images(extracted_images)

    print("🧾 Creating PowerPoint...")
    create_pptx_from_images_with_timestamps(
        unique_images, pptx_output, video_id, fps_interval
    )
