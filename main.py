import pandas as pd
import os
import argparse

def convert_xls_to_xlsx(input_path, rename=None):
  if not os.path.exists(input_path):
    print(f"File not found: {input_path}")
    return

  basename_with_ext = os.path.basename(input_path)
  basename = None
  
  if rename:
    basename = rename
  else:
    basename =  os.path.splitext(basename_with_ext)[0]

  output_path = os.path.join(os.path.dirname(input_path), basename + ".xlsx")

  try:
    xls = pd.read_excel(input_path, engine=None, sheet_name=None, dtype=str)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
          for sheet_name, df in xls.items():
              df.to_excel(writer, sheet_name=sheet_name, index=False)
      
    print(f"Successfully converted: {input_path} to {output_path}")
    print(f"Total Sheets: {len(xls)}")
  except Exception as e:
    print(f"Error converting {input_path} to {output_path}: {e}")

def process_directory(dir_path):
    for root, _, files in os.walk(dir_path):
        for file in files:
            if file.lower().endswith(('.xls', '.xlsx')) or not os.path.splitext(file)[1]:
                full_path = os.path.join(root, file)
                print(f"Processing: {full_path}")
                convert_xls_to_xlsx(full_path)

def main():
    parser = argparse.ArgumentParser(description='Convert xls to xlsx')
    parser.add_argument('input_file', help='Path to the input file')
    parser.add_argument('--rename', help='Rename the output file')
    args = parser.parse_args()

    convert_xls_to_xlsx(args.input_file, args.rename)

if __name__ == "__main__":
    main()