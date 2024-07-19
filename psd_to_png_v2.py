from PIL import Image
import psd_tools
import os

def convert_psd_to_png(directory, output_base_directory):
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.lower().endswith('.psd'):
                file_path = os.path.join(root, filename)
                psd = psd_tools.PSDImage.open(file_path)
                composite = psd.compose()

                relative_path = os.path.relpath(root, directory)
                output_directory = os.path.join(output_base_directory, relative_path)
                os.makedirs(output_directory, exist_ok=True)

                png_path = os.path.join(output_directory, os.path.splitext(filename)[0] + '.png')
                composite.save(png_path)

                print(f"Converted {filename} to {png_path}")

def main():
    while True:
        directory = input("Enter the base directory path containing PSD files: ")
        output_base_directory = os.path.join(directory, 'png')
        os.makedirs(output_base_directory, exist_ok=True)

        convert_psd_to_png(directory, output_base_directory)

        continue_choice = input("Do you want to convert files in another directory? (yes/no): ").strip().lower()
        if continue_choice != 'yes':
            break

if __name__ == "__main__":
    main()