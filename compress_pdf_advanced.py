# compress_pdf_advanced.py
import pikepdf
from PIL import Image
import io

def compress_pdf_advanced(input_path, output_path, image_quality=85):
    """
    Compress PDF with image optimization.
    
    Args:
        input_path: Path to input PDF
        output_path: Path to output PDF
        image_quality: JPEG quality (0-100, default 85)
    """
    pdf = pikepdf.Pdf.open(input_path)
    
    compressed_count = 0
    skipped_count = 0
    
    for page in pdf.pages:
        for img_key, img_obj in page.images.items():
            try:
                # Extract image - some streams may not be filterable
                raw_image = img_obj
                
                # Try to read the image bytes
                try:
                    image_bytes = raw_image.read_bytes()
                except Exception:
                    # Skip unfilterable streams (already compressed or encrypted)
                    skipped_count += 1
                    continue
                
                pil_image = Image.open(io.BytesIO(image_bytes))
                
                # Compress image if it's large
                if pil_image.width > 1920 or pil_image.height > 1920:
                    pil_image.thumbnail((1920, 1920), Image.Resampling.LANCZOS)
                
                # Convert to RGB if needed
                if pil_image.mode == 'RGBA':
                    pil_image = pil_image.convert('RGB')
                
                # Save compressed image
                img_buffer = io.BytesIO()
                pil_image.save(img_buffer, format='JPEG', quality=image_quality, optimize=True)
                img_buffer.seek(0)
                
                # Replace image in PDF
                raw_image.write(img_buffer.read(), filter=pikepdf.Name("/DCTDecode"))
                compressed_count += 1
                
            except Exception as e:
                skipped_count += 1
                continue
    
    # Save with compression
    pdf.save(output_path, compress_streams=True)
    pdf.close()
    
    print(f"Advanced compressed PDF saved to: {output_path}")
    print(f"Images compressed: {compressed_count}, Skipped: {skipped_count}")

# Usage
if __name__ == "__main__":
    compress_pdf_advanced("scan-Merged.pdf", "compressed_scan-Merged_advanced.pdf", image_quality=25)