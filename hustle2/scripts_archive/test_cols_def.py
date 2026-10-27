import win32com.client
import os
import datetime

# Map the exact data for July and August
# Standardized 44 columns:
# 1: NAME, 2: SALARIES, 3: TYPE, 4: DATE, 5: PAID STATUS
# 6: DATE, 7: ADV, 8: AMOUNT, 9: PAID
# 10: DATE, 11: CESC / URP, 12: AMOUNT, 13: PAID
# 14: DATE, 15: CESC / RESD, 16: AMOUNT, 17: PAID
# 18: DATE, 19: LABOUR/URP, 20: AMOUNT, 21: PAID
# 22: DATE, 23: DIGRUM / URP, 24: AMOUNT, 25: PAID
# 26: DATE, 27: ANIMAL FEED, 28: AMOUNT, 29: PAID
# 30: DATE, 31: FUEL/GAS, 32: AMOUNT, 33: PAID
# 34: DATE, 35: MEDICAL, 36: AMOUNT, 37: PAID
# 38: DATE, 39: LOCATION, 40: REPAIR, 41: AMOUNT, 42: PAID
# 43: RELIGION, 44: AMOUNT

HEADERS = [
    (1, "NAME"), (2, "SALARIES"), (3, "TYPE"), (4, "DATE"), (5, "PAID STATUS"),
    (6, "DATE"), (7, "ADV"), (8, "AMOUNT"), (9, "PAID"),
    (10, "DATE"), (11, "CESC / URP"), (12, "AMOUNT"), (13, "PAID"),
    (14, "DATE"), (15, "CESC / RESD"), (16, "AMOUNT"), (17, "PAID"),
    (18, "DATE"), (19, "LABOUR/URP"), (20, "AMOUNT"), (21, "PAID"),
    (22, "DATE"), (23, "DIGRUM / URP"), (24, "AMOUNT"), (25, "PAID"),
    (26, "DATE"), (27, "ANIMAL FEED"), (28, "AMOUNT"), (29, "PAID"),
    (30, "DATE"), (31, "FUEL/GAS"), (32, "AMOUNT"), (33, "PAID"),
    (34, "DATE"), (35, "MEDICAL"), (36, "AMOUNT"), (37, "PAID"),
    (38, "DATE"), (39, "LOCATION"), (40, "REPAIR"), (41, "AMOUNT"), (42, "PAID"),
    (43, "RELIGION"), (44, "AMOUNT")
]

DATE_COLS = [4, 6, 10, 14, 18, 22, 26, 30, 34, 38]
AMOUNT_COLS = [2, 8, 12, 16, 20, 24, 28, 32, 36, 41, 44]
TEXT_COLS = [1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39, 40, 42, 43]

print("Defined columns structure. Standardizing script...")
