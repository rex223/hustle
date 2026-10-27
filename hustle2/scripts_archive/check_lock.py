import os

file_path = os.path.abspath('RENTAL DATASHEET 8 OCT.xlsm')
try:
    with open(file_path, 'r+b') as f:
        print("File is NOT locked! We have read/write access.")
except PermissionError as e:
    print(f"File is LOCKED by another process: {e}")
except Exception as e:
    print(f"Error opening file: {e}")
