# compress_pdf_advanced.py
import io
import sys

import pikepdf
from PIL import Image

UNSUPPORTED_IMAGE_KEYS = (
    pikepdf.Name.SMask,
    pikepdf.Name.Mask,
    pikepdf.Name.ImageMask,
    pikepdf.Name.Decode,
)


def replace_with_jpeg(img_obj, image_quality):
    pil_image = pikepdf.PdfImage(img_obj).as_pil_image()

    if pil_image.width > 1920 or pil_image.height > 1920:
        pil_image.thumbnail((1920, 1920), Image.Resampling.LANCZOS)

    if pil_image.mode == 'L':
        colorspace = pikepdf.Name.DeviceGray
    else:
        pil_image = pil_image.convert('RGB')
        colorspace = pikepdf.Name.DeviceRGB

    buffer = io.BytesIO()
    pil_image.save(buffer, format='JPEG', quality=image_quality, optimize=True)
    img_obj.write(buffer.getvalue(), filter=pikepdf.Name.DCTDecode)
    img_obj.Width, img_obj.Height = pil_image.size
    img_obj.ColorSpace = colorspace
    img_obj.BitsPerComponent = 8


def compress_pdf_advanced(input_path, output_path, image_quality=85):
    """
    Compress PDF with image optimization.

    Args:
        input_path: Path to input PDF
        output_path: Path to output PDF
        image_quality: JPEG quality (0-100, default 85)
    """
    compressed_count = 0
    skipped_count = 0
    seen = set()

    with pikepdf.Pdf.open(input_path) as pdf:
        for page in pdf.pages:
            for img_key, img_obj in page.images.items():
                if img_obj.objgen in seen:
                    continue
                seen.add(img_obj.objgen)

                if any(key in img_obj for key in UNSUPPORTED_IMAGE_KEYS):
                    skipped_count += 1
                    continue

                try:
                    replace_with_jpeg(img_obj, image_quality)
                    compressed_count += 1
                except Exception as error:
                    skipped_count += 1
                    print(f"Warning: could not compress image {img_key}: {error}", file=sys.stderr)

        pdf.save(output_path, compress_streams=True)

    print(f"Advanced compressed PDF saved to: {output_path}")
    print(f"Images compressed: {compressed_count}, Skipped: {skipped_count}")

# Usage
if __name__ == "__main__":
    compress_pdf_advanced("original-source-file.pdf", "compressed_destinaltion-file.pdf", image_quality=25)
