import pandas as pd
import openpyxl
import re
import sys

# Prompt user for file name
input_file = input("Enter the input Excel file: ")

# Load selected sheet into DataFrame
df = pd.read_excel(input_file)

required_cols = {"Year", "Make", "Model", "Trim", "Engine", "Notes"}
if not required_cols.issubset(df.columns):
  raise ValueError("Missing required columns in the file.")

try:
  # Load the Excel file and show available sheets
  xls = pd.ExcelFile(input_file)

  # Functions to clean and extract info
  def clean_trim(trim):
    """Extract submodel, body type, and body number from Trim."""
    
    if isinstance(trim, str):
      return None, None, None  # case where trim isn't a string

    # Split the trim string
    parts = trim.split()
    
    if len(parts) > 2:
      submodel = parts[0]  # The first part is usually the submodel
      body_type = parts[1] if len(parts) > 1 else None  
      body_number = parts[2] if len(parts) > 2 else None  

      # Search for a number before the hyphen
      match = re.match(r'(\d)-', parts[2])
      if match:
        body_number = match.group(1)  # Extract the number before the hyphen
    else:
      print("Insufficient Trim Input")
      return None, None, None 
    
    return submodel, body_type, body_number
    
  def clean_engine(engine):
    """Extracts engine specifications into separate attributes without units."""
    if not isinstance(engine, str):
        return None, None, None, None, None, None, None
    
    # Extract liters (e.g., '2.0L' → '2.0')
    liter_match = re.search(r'(\d+\.\d+)L', engine)
    liters = liter_match.group(1) if liter_match else None

    # Extract CC (e.g., '1998CC' → '1998')
    cc_match = re.search(r'(\d{3,5})CC', engine)
    cc = cc_match.group(1) if cc_match else None

    # Extract CID (e.g., '122Cu. In.' → '122')
    cid_match = re.search(r'(\d+)Cu\. In\.', engine)
    cid = cid_match.group(1) if cid_match else None

    # Extract Cylinders (e.g., 'l4', 'V6' → '4', '6')
    cyl_match = re.search(r'(?:l|V)(\d+)', engine, re.IGNORECASE)
    cylinders = cyl_match.group(1) if cyl_match else None  # Only the number

    # Extract Fuel Type (Assuming it’s always 'GAS' for now)
    fuel_type = "GAS" if "GAS" in engine else None

    # Extract Cylinder Head Type (e.g., 'DOHC', 'SOHC', 'OHV')
    head_match = re.search(r'(DOHC|SOHC|OHV|OHC)', engine, re.IGNORECASE)
    cylinder_head_type = head_match.group(0) if head_match else None

    # Extract Aspiration (e.g., 'Turbocharged', 'Naturally Aspirated')
    aspiration_match = re.search(r'(Turbocharged|Naturally Aspirated)', engine, re.IGNORECASE)
    aspiration = aspiration_match.group(0) if aspiration_match else None

    return liters, cc, cid, cylinders, fuel_type, cylinder_head_type, aspiration

  # Apply functions to the DataFrame
  df[['Submodel', 'Body Type', 'Body Number']] = df['Trim'].apply(lambda x: pd.Series(clean_trim(x)))

  df[['Liters', 'CC', 'CID', 'Cylinders', 'Fuel Type', 'Cylinder Head Type', 'Aspiration']] = df['Engine'].apply(lambda x: pd.Series(clean_engine(x)))

  # Saving to a new sheet in the same Excel file
  with pd.ExcelWriter(input_file, engine='openpyxl', mode='a') as writer: # mode a: ModifiedData sheet should not exist
    df.to_excel(writer, sheet_name='ModifiedData', index=False)
    print(f"Data processed and saved to a new sheet 'ModifiedData' in {input_file}")
  
except FileNotFoundError:
  print("Error: File not found. Please check the file name and try again.")
  sys.exit() 

except Exception as e: 
  print(f"An error occured: {e}")
  sys.exit() 

