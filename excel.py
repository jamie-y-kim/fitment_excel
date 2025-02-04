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
    else:
      return None, None, None  # Handle the case where trim isn't a string
    
  # def extract_engine_info(engine):
  #     """Extracts displacement, cylinder type, and fuel type from Engine string."""
  #     if isinstance(engine, str):
  #         parts = engine.split()
  #         displacement = parts[0] if "L" in parts[0] else None
  #         cylinder = parts[1] if "V" in parts[1] or "l4" in parts[1] else None
  #         fuel_type = parts[3] if len(parts) > 3 else None
  #         return displacement, cylinder, fuel_type
  #     return None, None, None

  # Apply functions to the DataFrame
  df[['Submodel', 'Body Type', 'Body Number']] = df['Trim'].apply(lambda x: pd.Series(clean_trim(x)))

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

