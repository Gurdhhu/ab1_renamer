from re import findall, sub
import pandas as pd
from os import listdir


def change_internal_name(path_to_infile, new_name):
    """
    Decodes the .ab1 file, finds the default internal sequence name,
    substitutes it with the provided new name and returns a list of
    byte strings.
    """
    new_internal_name = new_name.encode("utf-8")
    substitute = bytes([len(new_internal_name)]) + new_internal_name
    default_name_found = False

    with open(path_to_infile, "rb") as f:
        lines = f.readlines()

    byte_line_list = []

    for line in lines:
        if b'\x06KB.bcp' not in line:
            byte_line_list.append(line)
        else:
            default_name = findall(b'(\\w+)\\x06KB.bcp', line)[0]
            default_name_found = True
            pattern = bytes([len(default_name)]) + default_name
            line_edited = sub(pattern, substitute, line)
            byte_line_list.append(line_edited)
    if not default_name_found:
        print(f"Warning: {path_to_infile} - internal name was not found!")

    return byte_line_list


def save_renamed_ab1(path_to_outfile, byte_line_list):
    """
    Saves the byte_line_list into path_to_outfile in .ab1 format.
    """
    with open(path_to_outfile, "wb") as out:
        out.writelines(byte_line_list)
        # print(f"Renamed ab1 file saved to {path_to_outfile}\n")


def find_header_line(file_path, sheet_name=0):
    """
    Finds the header line in an Excel file, defined as the first line where
    at least the first two columns are non-empty. Returns the header line
    position (0-indexed) and the list of column headers.
    """
    try:
        df_temp = pd.read_excel(file_path, header=None, sheet_name=sheet_name)
    except Exception as e:
        raise ValueError(f"Could not read sheet '{sheet_name}': {str(e)}")

    for idx in range(len(df_temp)):
        row = df_temp.iloc[idx]
        col0 = str(row[0]).strip() if pd.notnull(row[0]) else ''
        col1 = str(row[1]).strip() if pd.notnull(row[1]) else ''

        if col0 != '' and col1 != '':
            headers = [str(cell).strip() if pd.notnull(cell) else '' for cell in row.values]
            return idx, headers

    raise ValueError(f"No valid header line found in sheet '{sheet_name}'. ")


def create_mapping(
        file_path,
        header_line_pos,
        input_col="Macrogen",
        output_col="Real name",
        sheet_name=0
):
    """
    Creates a dictionary mapping values from the input column to the output column,
    returns it as dictionary and DataFrame.
    - file_path: Path to the Excel file.
    - header_line_pos: Position of the header line (from find_header_line).
    - input_col: Name of the column with keys (default: "Macrogen").
    - output_col: Name of the column with values (default: "Real name").
    """
    # Read the Excel file, skipping rows before the header line
    df = pd.read_excel(
        file_path,
        skiprows=range(header_line_pos),  # Skip all rows before the header
        header=0,  # Use the first remaining row as header
        sheet_name=sheet_name
    )

    # Validate required columns
    missing_cols = []
    if input_col not in df.columns:
        missing_cols.append(input_col)
    if output_col not in df.columns:
        missing_cols.append(output_col)
    if missing_cols:
        raise ValueError(f"Columns not found: {', '.join(missing_cols)}")

    # Drop rows with missing values in either key column
    df_clean = df.dropna(subset=[input_col, output_col])

    # Clean the input filenames by removing extensions
    df_clean[input_col] = df_clean[input_col].str.rsplit('.', n=1).str[0]

    # Create the dictionary
    mapping = df_clean.set_index(input_col)[output_col].to_dict()

    return mapping, df_clean


def get_ab1_file_list(path_to_ab1_folder):
    """
    returns a list of .ab1 files in the provided directory
    """
    ab1_list = []
    for file in listdir(path_to_ab1_folder):
        if file.endswith(".ab1"):
            # print(file)
            ab1_list.append(file)
    return ab1_list


def sanitize_filename(name):
    """Replace invalid filename characters with underscores"""
    invalid_chars = r'\/:*?"<>|'
    replace_char = '_'

    # Replace invalid characters
    for char in invalid_chars:
        name = name.replace(char, replace_char)

    # Remove leading/trailing whitespace and dots
    name = name.strip().rstrip('.')

    # Replace spaces with underscores if needed
    name = name.replace(' ', replace_char)

    # Truncate to reasonable length
    max_length = 200
    return name[:max_length]
