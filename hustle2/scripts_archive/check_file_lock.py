import os

file_path = os.path.abspath('RENTAL DATASHEET 8 OCT.xlsm')
try:
    with open(file_path, 'r+b') as f:
        print("File is NOT locked! Can be written directly.")
except Exception as e:
    print(f"File IS locked: {e}")
