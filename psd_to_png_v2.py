from PIL import Image
import psd_tools
import os
import logging

# logging
def setup_logging(output_base_directory):
    log_path = os.path.join(output_base_directory, 'log.txt')
    logging.basicConfig(filename=log_path, level=logging.DEBUG, 
                        format='%(asctime)s - %(levelname)s - %(message)s')

def convert_psd_to_png(directory, output_base_directory):
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.lower().endswith('.psd'):
                file_path = os.path.join(root, filename)
                try:
                    psd = psd_tools.PSDImage.open(file_path)
                    composite = psd.compose()

                    relative_path = os.path.relpath(root, directory)
                    output_directory = os.path.join(output_base_directory, relative_path)
                    os.makedirs(output_directory, exist_ok=True)

                    png_path = os.path.join(output_directory, os.path.splitext(filename)[0] + '.png')
                    composite.save(png_path)

                    print(f"Converted {filename} to {png_path}")
                    logging.info(f"Successfully converted {file_path} to {png_path}")
                except Exception as e:
                    print(f"Error converting {filename}: {e}")
                    logging.error(f"Failed to convert {file_path}: {e}")

def main():
    while True:
        directory = input("Enter the base directory path containing PSD files: ")
        output_base_directory = os.path.join(directory, 'png')
        os.makedirs(output_base_directory, exist_ok=True)
        
        setup_logging(output_base_directory)

        convert_psd_to_png(directory, output_base_directory)

        continue_choice = input("Do you want to convert files in another directory? (yes/no): ").strip().lower()
        if continue_choice != 'yes':
            break

if __name__ == "__main__":
    main()