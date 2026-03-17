# morse_now

This repository now includes a utility to convert images into editable `.docx` files using OCR.

## Convert image to editable Word file

1. Install Tesseract OCR (system dependency):
   - Ubuntu/Debian: `sudo apt-get install tesseract-ocr`
2. Install Python dependencies:
   - `pip install -r requirements.txt`
3. Run conversion:
   - Local image: `python ocr_to_word.py /absolute/path/to/image.jpg -o /absolute/path/to/output.docx`
   - Image URL: `python ocr_to_word.py "https://github.com/user-attachments/assets/3979b1a4-5680-4d91-84e1-ba36ba7f7728" -o /absolute/path/to/output.docx`

The converter keeps OCR text editable and preserves layout approximately by line positioning and spacing, while also applying detected text colors.
