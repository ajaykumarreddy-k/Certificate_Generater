import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os

def generate_certificates():
    """
    Generate certificates by reading names from Excel and placing them on template image
    """
    
    # File paths (all in current directory)
    excel_file = 'all_names_list.xlsx'
    template_file = 'White Gold Black Simple Minimalist Certificate Of Participation (1).png'  # or .jpg
    output_folder = 'certificates'
    
    try:
        # Check if required files exist
        if not os.path.exists(excel_file):
            print(f"❌ Error: {excel_file} not found in current directory")
            print("Please create an Excel file with participant names in column 'Name'")
            return
        
        if not os.path.exists(template_file):
            print(f"❌ Error: {template_file} not found in current directory")
            print("Please add your certificate template image to current directory")
            return
        
        # Read participant names from Excel
        print("📖 Reading participant names from Excel...")
        df = pd.read_excel(excel_file)
        
        # Check if 'Name' column exists
        if 'Name' not in df.columns:
            print("❌ Error: Excel file must have a 'Name' column")
            print("Current columns:", df.columns.tolist())
            return
        
        names = df['Name'].dropna().tolist()  # Remove any empty names
        
        if not names:
            print("❌ Error: No names found in Excel file")
            return
        
        print(f"✅ Found {len(names)} participants")
        
        # Load certificate template
        print("🖼️  Loading certificate template...")
        template = Image.open(template_file)
        print(f"Template size: {template.size[0]} x {template.size[1]} pixels")
        
        # Set up font - Using size 50 (same as before)
        font_size = 80
        font_loaded = False
        
        # Try multiple font options in order of preference
        font_options = [
            "OleoScript-Regular.ttf",
            "fonts/OleoScript-Regular.ttf", 
            "OleoScript.ttf",
            "fonts/OleoScript.ttf",
            "OleoScript-Bold.ttf",
            "fonts/OleoScript-Bold.ttf",
            # Windows script fonts
            "C:/Windows/Fonts/SCRIPTBL.TTF",  # Script MT Bold
            "C:/Windows/Fonts/BRUSHSCI.TTF",  # Brush Script MT
            "C:/Windows/Fonts/VLADIMIR.TTF",  # Vladimir Script
            "C:/Windows/Fonts/VIVALDII.TTF",  # Vivaldi
            # macOS script fonts
            "/System/Library/Fonts/Brush Script MT.ttf",
            "/Library/Fonts/Brush Script MT.ttf",
            # Linux script fonts
            "/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf",
            # Fallback to common fonts
            "times.ttf",
            "arial.ttf",
        ]
        
        for font_path in font_options:
            try:
                font = ImageFont.truetype(font_path, font_size)
                print(f"✅ Successfully loaded font: {font_path}")
                font_loaded = True
                break
            except:
                continue
        
        if not font_loaded:
            try:
                font = ImageFont.truetype("times.ttf", font_size)
                print("⚠️  Using Times font as fallback")
            except:
                font = ImageFont.load_default()
                print("⚠️  Using default system font")
                print("💡 For best results, add 'OleoScript-Regular.ttf' to your folder")
        
        # Create output directory
        os.makedirs(output_folder, exist_ok=True)
        print(f"📁 Created/using output folder: {output_folder}")
        
        # Generate certificates
        print("\n🎓 Generating certificates...")
        successful = 0
        
        for i, name in enumerate(names, 1):
            try:
                # Create a copy of template
                certificate = template.copy()
                draw = ImageDraw.Draw(certificate)
                
                # Calculate text position (centered horizontally)
                bbox = draw.textbbox((0, 0), name, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                
                # Position text on the dotted line in the new certificate template
                # Adjusted coordinates for the new white certificate template
                x = (template.width - text_width) // 2  # Center horizontally
                y = 600  # Positioned on the dotted line (adjusted for new template)
                
                # Draw the name on certificate with BLACK color (changed from white)
                draw.text((x, y), name, fill='black', font=font)
                
                # Create filename as participant_name.png
                clean_name = str(name).replace(" ", "_").replace("/", "_").replace("\\", "_")
                clean_name = "".join(c for c in clean_name if c.isalnum() or c in "._-")
                filename = f'{output_folder}/{clean_name}.png'
                
                # Handle duplicate names by adding number
                original_filename = filename
                counter = 1
                while os.path.exists(filename):
                    name_part = clean_name
                    filename = f'{output_folder}/{name_part}_{counter}.png'
                    counter += 1
                
                # Save certificate
                certificate.save(filename, 'PNG')
                print(f"  ✅ {i:2d}. {name} -> {os.path.basename(filename)}")
                successful += 1
                
            except Exception as e:
                print(f"  ❌ {i:2d}. Error generating certificate for {name}: {e}")
        
        print(f"\n🎉 Successfully generated {successful} out of {len(names)} certificates!")
        print(f"📂 All certificates saved in '{output_folder}' folder")
        
    except FileNotFoundError as e:
        print(f"❌ File not found: {e}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")

def main():
    """
    Main function - shows instructions and runs certificate generator
    """
    print("🎓 Certificate Generator - New White Template")
    print("=" * 50)
    print("\nRequired files in current directory:")
    print("1. all_names_list.xlsx - Excel file with 'Name' column")
    print("2. A6 - 2.png - Your NEW certificate template image")
    print("3. OleoScript-Regular.ttf - Oleo Script font file")
    print("\nSettings: Black font, size 50, positioned on dotted line")
    print("\nStarting generation...\n")
    
    generate_certificates()

if __name__ == "__main__":
    main()