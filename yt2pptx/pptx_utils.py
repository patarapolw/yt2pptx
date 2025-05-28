from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Length, Pt
from pptx.dml.color import RGBColor
from tqdm import tqdm

from .video_utils import make_timestamp


def create_pptx_from_images_with_timestamps(
    image_tuples: list[tuple[Path, int]], output_pptx: Path, video_id: str
) -> None:
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]
    for img_path, seconds in tqdm(image_tuples, desc="Creating slides"):
        slide = prs.slides.add_slide(blank_slide_layout)
        slide.shapes.add_picture(
            str(img_path), Inches(0), Inches(0), width=prs.slide_width
        )
        timestamp = make_timestamp(seconds)
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
