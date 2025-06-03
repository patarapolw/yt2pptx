# yt2pptx

A Python script to convert a YouTube video (or any video supported by [yt-dlp](https://github.com/yt-dlp/yt-dlp)) into a PowerPoint presentation. The script downloads the video, extracts frames at regular intervals, removes near-duplicate frames, and creates a .pptx file with each unique frame as a slide. Each slide includes a timestamp and a clickable link to jump to that moment on YouTube.

The PowerPoint presentation can be edited later, such as removing duplicated frames or adding notes.

## Features

- Accepts any input supported by yt-dlp (YouTube URL, video ID, playlist, etc.)
- Extracts frames every N seconds
- Removes near-duplicate frames using perceptual hashing
- Generates a PowerPoint presentation with each unique frame as a slide
- Adds a timestamp and clickable YouTube link to each slide

## Requirements

Install dependencies with:

```sh
pip install yt-dlp pillow imagehash python-pptx tqdm
```

Make sure [ffmpeg](https://ffmpeg.org/) is installed and available in your system PATH.

## Usage

```sh
python -m yt2pptx.cli <YouTube_URL_or_ID> [output_base_name] [-i=SECONDS|--interval=SECONDS]
```

- `<YouTube_URL_or_ID>`: Any input accepted by yt-dlp (YouTube URL, video ID, playlist, etc.)
- `[output_base_name]`: (Optional) Custom base name for output files
- `-i=SECONDS` or `--interval=SECONDS`: (Optional, can be anywhere in the arguments) Set seconds between frames (default: 2)

### Example

```sh
python -m yt2pptx.cli https://www.youtube.com/watch?v=dQw4w9WgXcQ
```

This will create in the `out/` directory:
- `out/video_title.pptx` — the generated PowerPoint
- `out/{video_id}/` — folder with extracted frames, generated scripts, and video title (where `{video_id}` is the YouTube video ID)

You can also specify a custom base name and/or change the interval:

```sh
python -m yt2pptx.cli dQw4w9WgXcQ MySlides --interval=5
```

### Post-processing

PowerPoint can be edited to delete more duplicated frames. If some frames are missing, links to YouTube at the timestamp, and buttons to `ffplay` (bundled with `ffmpeg`) can be clicked to manually find the missing frames, or more accurate ones.

## License

MIT License

## Acknowledgements

- [yt-dlp](https://github.com/yt-dlp/yt-dlp)
- [ffmpeg](https://ffmpeg.org/)
- [python-pptx](https://python-pptx.readthedocs.io/)
- [imagehash](https://github.com/JohannesBuchner/imagehash)
