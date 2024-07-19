from PIL import Image
import psd_tools
import os

def convert_psd_to_png(directory):
    output_directory = os.path.join(directory, 'png')
    os.makedirs(output_directory, exist_ok=True)

    for filename in os.listdir(directory):
        if filename.lower().endswith('.psd'):
            file_path = os.path.join(directory, filename)

            psd = psd_tools.PSDImage.open(file_path)

            composite = psd.compose()

            png_path = os.path.join(output_directory, os.path.splitext(filename)[0] + '.png')
            composite.save(png_path)

            print(f"Converted {filename} to {png_path}")

def main():
    while True:
        directory = input("Enter the directory path containing PSD files: ")

        convert_psd_to_png(directory)

        continue_choice = input("Do you want to convert files in another directory? (yes/no): ").strip().lower()
        if continue_choice != 'yes':
            break

if __name__ == "__main__":
    main()