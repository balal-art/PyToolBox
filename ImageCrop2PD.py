from PIL import Image
import os

def optimize_images_for_pdf(input_folder, output_pdf_path, left_px, right_px, top_px, bottom_px, target_dpi=150):
    """
    Crop, optimize and convert images to PDF with specific DPI
    
    :param input_folder: Path to folder containing images
    :param output_pdf_path: Output PDF file path
    :param target_dpi: Target DPI resolution (default: 150)
    """
    cropped_images = []
    
    for filename in sorted(os.listdir(input_folder)):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
            try:
                with Image.open(os.path.join(input_folder, filename)) as img:
                    # Crop image
                    width, height = img.size
                    left = left_px
                    right = width - right_px
                    top = top_px
                    bottom = height - bottom_px
                    
                    if left >= right or top >= bottom:
                        continue
                    
                    cropped_img = img.crop((left, top, right, bottom)).convert('RGB')
                    
                    # Calculate new size for target DPI
                    if 'dpi' in img.info:
                        original_dpi = img.info['dpi'][0]
                        scale_factor = target_dpi / original_dpi
                        new_size = (int(cropped_img.width * scale_factor), 
                                   int(cropped_img.height * scale_factor))
                        cropped_img = cropped_img.resize(new_size, Image.LANCZOS)
                    
                    cropped_images.append(cropped_img)
                    print(f"Processed: {filename}")
                    
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
    
    if cropped_images:
        # Save with optimized parameters
        cropped_images[0].save(
            output_pdf_path,
            save_all=True,
            append_images=cropped_images[1:],
            quality=85,  # Reduced quality for smaller size
            dpi=(target_dpi, target_dpi),  # Set target DPI
            optimize=True
        )
        print(f"Optimized PDF created at: {output_pdf_path}")
        print(f"Final size: {os.path.getsize(output_pdf_path)/1024:.2f} KB")
    else:
        print("No valid images processed")

# مثال استفاده:
optimize_images_for_pdf(
    input_folder=r"C:\Users\Ali\Desktop\2025-06-03",
    output_pdf_path=r"C:\Users\Ali\Desktop\output_optimized.pdf",
    left_px=20,
    right_px=20,
    top_px=0,
    bottom_px=350,
    target_dpi=150  # تنظیم DPI به 150
)