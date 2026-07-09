# process_photos.py

A self-contained Python 3 script that turns a folder of JPEG photos into an annotated 4K H.264 slideshow and a Word manifest document.

Copy the script into any trip photo folder and run it — the trip name, output filenames, and trip date range are all auto-derived from the folder itself and the photos' own EXIF timestamps, so no per-trip editing is required.

## What it does

For each photo, the script:

- Extracts GPS coordinates and camera EXIF data
- Looks up elevation via the USGS Elevation Point Query Service
- Reverse-geocodes a human-readable area name (Nominatim + OpenStreetMap Overpass)
- Downloads a USGS topo map tile inset
- Burns caption overlays into two captioned image variants (full info + nearest-area only)
- Sets aside photos with missing or clearly wrong EXIF dates on a review list, instead of silently misfiling them into the wrong day

Processing is incremental (only new/changed photos are reprocessed), and it can optionally build a black-background PowerPoint deck alongside the manifest and slideshow.

## Requirements

- Python 3.8+ with `pillow`, `python-docx`, `python-pptx`, `requests`
- ImageMagick (`convert`, `identify` on PATH)
- FFmpeg with `libx264`
- `xfce4-terminal` (optional, for watching the slideshow encode)

## Quick start

```
cp process_photos.py /path/to/your/photos/
cd /path/to/your/photos
python3 process_photos.py
```

The script processes every JPEG in the folder, rebuilds the manifest, then prompts whether to encode the 4K slideshow.

## Common options

| Flag | Effect |
|---|---|
| `--force` | Reprocess all photos, ignoring the cache |
| `--manifest` | Rebuild the manifest only, no image processing |
| `--slideshow` / `--no-slideshow` | Always / never encode the slideshow |
| `--powerpoint` | Also build a black-background `.pptx` deck |
| `--trip-start YYYY-MM-DD --trip-end YYYY-MM-DD` | Override auto-detected trip date range |
| `--limit N` | Only process the first N photos chronologically |

Run `python3 process_photos.py --help` for the complete list of options, including camera-attribution, date-correction, and Nominatim-throttling flags.
