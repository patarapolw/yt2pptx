from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Length, Pt
from pptx.dml.color import RGBColor
from tqdm import tqdm

from .video_utils import timestamp_to_seconds


def create_pptx_from_images_with_timestamps(
    image_paths: list[Path], output_pptx: Path, video_id: str
) -> None:
    """Create a PowerPoint presentation from a list of images with timestamps.

    This function takes a list of tuples containing image paths and their corresponding
    timestamps in seconds, and creates a PowerPoint presentation where each slide contains
    an image and a hyperlink to the YouTube video at the specified timestamp.

    Args:
        image_paths (list[Path]): A list of paths to the image file.
        output_pptx (Path): The path where the PowerPoint file will be saved.
        video_id (str): The YouTube video ID to create hyperlinks for the timestamps.
    """
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    for img_path in tqdm(image_paths, desc="Creating slides"):
        slide = prs.slides.add_slide(blank_slide_layout)
        slide.shapes.add_picture(
            str(img_path), Inches(0), Inches(0), width=prs.slide_width
        )
        timestamp = img_path.stem.split("_")[-1].replace("-", ":")
        zero_hours = "0:"
        if timestamp.startswith(zero_hours):
            timestamp = timestamp[len(zero_hours) :]

        seconds = timestamp_to_seconds(timestamp)
        youtube_link = f"https://www.youtube.com/watch?v={video_id}&t={seconds}s"
        left = Inches(0.3)
        top = (
            Length(prs.slide_height - Inches(0.7)) if prs.slide_height else Inches(0.5)
        )
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

    prs.save(str(output_pptx))
    print(f"✅ PowerPoint saved: {output_pptx}")
