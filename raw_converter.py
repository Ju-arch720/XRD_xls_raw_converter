import os
import pandas as pd
import numpy as np

def convert_xls_to_txt(input_file, output_file=None):
    """
    Convert XLS file to TXT file
    
    Parameters:
    -----------
    input_file : str
        Path to the input XLS file
    output_file : str, optional
        Path for the output TXT file
    
    Returns:
    --------
    str
        Path to the created TXT file
    """
    # Read the XLS file
    try:
        df = pd.read_excel(input_file, engine='openpyxl')
    except Exception as e:
        raise ValueError(f"Could not read XLS file: {e}")
    
    # Identify potential angle and intensity columns
    possible_angle_cols = [col for col in df.columns if any(x in col.lower() for x in ['angle', '2theta', 'two-theta'])]
    possible_intensity_cols = [col for col in df.columns if any(x in col.lower() for x in ['intensity', 'counts', 'int'])]
    
    # Default to first two columns if no clear match
    angle_col = possible_angle_cols[0] if possible_angle_cols else df.columns[0]
    intensity_col = possible_intensity_cols[0] if possible_intensity_cols else df.columns[1]
    
    # Prepare output filename
    if output_file is None:
        output_file = os.path.splitext(input_file)[0] + '.txt'
    
    # Write to TXT file
    df[[angle_col, intensity_col]].to_csv(output_file, sep='\t', index=False, header=False)
    
    print(f"Converted {input_file} to TXT: {output_file}")
    return output_file

def convert_txt_to_xy(input_file, output_file=None):
    """
    Convert TXT file to XY file
    
    Parameters:
    -----------
    input_file : str
        Path to the input TXT file
    output_file : str, optional
        Path for the output XY file
    
    Returns:
    --------
    str
        Path to the created XY file
    """
    # Read the TXT file
    try:
        df = pd.read_csv(input_file, sep='\t', header=None, names=['Angle', 'Intensity'])
    except Exception as e:
        raise ValueError(f"Could not read TXT file: {e}")
    
    # Prepare output filename
    if output_file is None:
        output_file = os.path.splitext(input_file)[0] + '.xy'
    
    # Write to XY file
    df.to_csv(output_file, sep='\t', index=False, header=False)
    
    print(f"Converted TXT to XY: {output_file}")
    return output_file

def full_conversion_workflow(input_file):
    """
    Complete conversion workflow:
    1. XLS to TXT
    2. TXT to XY
    
    Parameters:
    -----------
    input_file : str
        Path to the input XLS file
    
    Returns:
    --------
    dict
        Paths to intermediate and final files
    """
    # Step 1: Convert XLS to TXT
    txt_file = convert_xls_to_txt(input_file)
    
    # Step 2: Convert TXT to XY
    xy_file = convert_txt_to_xy(txt_file)
    
    return {
        'input_file': input_file,
        'txt_file': txt_file,
        'xy_file': xy_file
    }

# Example usage
if __name__ == "__main__":
    # Replace with your input XLS file path
    input_file = "Arumetsa\Aru_XRD.xlsx"
    
    # Run full conversion workflow
    conversion_results = full_conversion_workflow(input_file)
    
    # Print out the files created
    print("\nConversion Complete!")
    for key, value in conversion_results.items():
        print(f"{key}: {value}")