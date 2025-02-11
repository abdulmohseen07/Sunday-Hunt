from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os
from fpdf import FPDF

# Function to read names from Excel
def read_names_from_excel(file_path):
    try:
        df = pd.read_excel(file_path, usecols=[0], header=None)
        names = df.dropna().iloc[:, 0].tolist()
        return names
    except Exception as e:
        print(f"\u274c Error reading Excel file: {e}")
        return []

# Function to generate certificates with correct aspect ratio
def generate_certificates(template_path, output_folder, names, font_path="AlexBrush.ttf", font_size=100, text_position=(750, 510)):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    if not os.path.exists(template_path):
        print(f"\u274c Error: Template file '{template_path}' not found.")
        return

    for name in names:
        try:
            # Open template
            img = Image.open(template_path).convert("RGB")
            width, height = img.size  # Get actual width & height
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype(font_path, font_size)

            # Calculate text width for centering
            bbox = draw.textbbox((0, 0), name, font=font)
            text_width, text_height = bbox[2] - bbox[0], bbox[3] - bbox[1]
            adjusted_position = (text_position[0] - text_width // 2, text_position[1] - text_height // 4)

            draw.text(adjusted_position, name, fill="white", font=font)  # Changed font color to white

            # Save as temporary image
            temp_img_path = os.path.join(output_folder, f"certificate_{name.replace(' ', '_')}.png")
            img.save(temp_img_path, format="PNG")

            # Convert to PDF with proper size
            pdf_output_path = os.path.join(output_folder, f"certificate_{name.replace(' ', '_')}.pdf")
            pdf = FPDF(unit="pt", format=[width, height])  # Use actual image size
            pdf.add_page()
            pdf.image(temp_img_path, 0, 0, width, height)  # Maintain aspect ratio
            pdf.output(pdf_output_path)

            # Remove temporary PNG file
            os.remove(temp_img_path)

            print(f"\u2705 Generated: {pdf_output_path}")
        except Exception as e:
            print(f"\u274c Error generating certificate for {name}: {e}")

# File Paths
template_path = r"C:\Users\khan mohseen\Desktop\Certificate Generator\Certificate.png"
excel_file = r"C:\Users\khan mohseen\Desktop\Certificate Generator\TypeSprint register.xlsx"
output_folder = "certificates"
font_path = r"C:\Users\khan mohseen\Desktop\Certificate Generator\AlexBrush-Regular.ttf"

# Read names from Excel and generate certificates
names = read_names_from_excel(excel_file)
if names:
    generate_certificates(template_path, output_folder, names, font_path)
