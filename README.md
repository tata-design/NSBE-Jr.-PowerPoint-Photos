# NSBE Jr. End-of-Year PowerPoint

Yes — this workspace can help you turn a folder of photos into a presentable PowerPoint slideshow.

## What this project does

The included script creates a widescreen `.pptx` slideshow with:

- a title slide
- one centered photo per slide with a clean frame
- a closing thank-you slide

## How to use it

1. Put your photos in a folder named `photos/`
2. Install the Python dependencies:

   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

3. Generate the slideshow:

   ```bash
   python generate_slideshow.py
   ```

4. Open the generated file at `output/end-of-year-celebration.pptx`

## Optional custom title

```bash
python generate_slideshow.py \
  --title "NSBE Jr. End of Year Celebration" \
  --subtitle "Celebrating a great year together"
```

## Best results for a polished slideshow

- Use clear, high-resolution photos
- Rename photos in the order you want them shown
- Mix group shots, events, and candid moments
- Review the final PowerPoint and add any personal notes or transitions you want in PowerPoint itself