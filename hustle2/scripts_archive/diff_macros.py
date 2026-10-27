import os
import difflib

dir1 = 'vba_26_aug'
dir2 = 'vba_3_sep'

files1 = set(os.listdir(dir1))
files2 = set(os.listdir(dir2))

all_files = sorted(files1.union(files2))
found_diffs = False

for f in all_files:
    path1 = os.path.join(dir1, f)
    path2 = os.path.join(dir2, f)
    
    if f not in files1:
        print(f"File {f} is ONLY in {dir2}")
        found_diffs = True
        continue
    if f not in files2:
        print(f"File {f} is ONLY in {dir1}")
        found_diffs = True
        continue
        
    with open(path1, 'r', encoding='utf-8') as file1:
        lines1 = file1.readlines()
    with open(path2, 'r', encoding='utf-8') as file2:
        lines2 = file2.readlines()
        
    diff = list(difflib.unified_diff(lines1, lines2, fromfile=f'26_AUG/{f}', tofile=f'3_SEP/{f}'))
    if diff:
        found_diffs = True
        print(f"--- Differences in {f} ---")
        for line in diff:
            print(line.strip('\n'))
        print()

if not found_diffs:
    print("NO DIFFERENCES FOUND between the macros.")
