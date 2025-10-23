import pandas as pd
import os
import argparse
import tempfile
import shutil
import zipfile

def convert_xls_to_xlsx(input_path, rename=None, output_dir=None):
    try:
        xls = pd.read_excel(input_path, sheet_name=None, dtype=str)
        basename_with_ext = os.path.basename(input_path)
        basename = rename if rename else os.path.splitext(basename_with_ext)[0]
        output_path = os.path.join(os.path.dirname(input_path), basename + ".xlsx")
        
        if output_dir:
          os.makedirs(output_dir, exist_ok=True)
          output_path = os.path.join(output_dir, basename + ".xlsx")
        else:
          output_path = os.path.join(os.path.dirname(input_path), basename + ".xlsx")

        print(f"Converting {input_path} to {output_path}")

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in xls.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Converted: {input_path} → {output_path} | Sheets: {len(xls)}")

    except Exception as e:
        print(f"Failed to convert {input_path}: {e}")

def process_directory(dir_path, output_dir, recursive=False, include_zip=False):
    for root, _, files in os.walk(dir_path):
      for file in files:
        file_ext = os.path.splitext(file)[1].lower()
        full_path = os.path.join(root, file)
      
      if file_ext == ".zip" and include_zip:
        process_zip(full_path, output_dir, recursive, include_zip)
      elif file_ext in {".xls", ".xlsx"} or file_ext == "":
        convert_xls_to_xlsx(full_path)
    
      if not recursive:
          break
      

def process_zip(zip_path, output_dir, recursive=False, include_zip=False):
    temp_dir = tempfile.mkdtemp()
    try:
      with zipfile.ZipFile(zip_path, 'r') as zip_ref:
          zip_ref.extractall(temp_dir)
      print(f"Extracted ZIP {zip_path} to {temp_dir}")
      process_directory(temp_dir, output_dir, recursive, include_zip)
    except Exception as e:
      print(f"Error extracting ZIP {zip_path}: {e}")
    finally:
      shutil.rmtree(temp_dir)

def main():
    parser = argparse.ArgumentParser(description='Convert xls files to xlsx')
    parser.add_argument('input_file', help='Path to the input file')
    parser.add_argument('--rename', help='Rename the output file')
    parser.add_argument('--recursive', help='Recursively process all files in the directory', action='store_true')
    parser.add_argument('--include-zip', help='Extract and read file in zip', action='store_true')
    parser.add_argument('--output', help='Output directory')

    args = parser.parse_args()

    input_file = args.input_file
    output_dir = os.path.dirname(input_file)

    if args.output:
        output_dir = args.output

    if os.path.isfile(input_file):
       file_ext = os.path.splitext(input_file)[1].lower()
       if file_ext == ".zip" and args.include_zip:
        process_zip(input_file, output_dir, args.recursive, args.include_zip)
       elif file_ext in {".xls", ".xlsx"} or file_ext == "":
        convert_xls_to_xlsx(input_file, args.rename, output_dir)
       elif file_ext == ".zip" and not args.include_zip:
        print(f"Error: File {input_file} is a zip file, but --include-zip is not specified")
    elif os.path.isdir(input_file):
        process_directory(input_file, output_dir, args.recursive, args.include_zip)
    else:
        print(f"Error: Invalid input file or directory: {input_file}")
          


if __name__ == "__main__":
    main()