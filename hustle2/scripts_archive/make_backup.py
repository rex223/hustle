import shutil
import os
from datetime import datetime

timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
src = os.path.abspath('RENTAL DATASHEET 19 OCT.xlsm')
backup_dir = os.path.abspath('backups')
os.makedirs(backup_dir, exist_ok=True)
dst = os.path.join(backup_dir, f'RENTAL_DATASHEET_19_OCT_backup_{timestamp}.xlsm')

shutil.copyfile(src, dst)
print("Backup created successfully at:", dst)
