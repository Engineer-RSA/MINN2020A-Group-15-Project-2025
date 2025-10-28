# MINN2020A-Group-15-Project-2025
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import json
from datetime import datetime
import re
import numpy as np
import tkintermapview
import warnings
import hashlib
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ----------------------------
# File paths
# ----------------------------
FILE_6_4 = "6.4. Production_of_Mineral_Raw_Materials_of_individual_Countries_by_Minerals.xlsx"
USERS_FILE = "users.json"
CMO_MONTHLY_FILE = "CMO-Historical-Data-Monthly (1).xlsx"
CMO_ANNUAL_FILE = "CMO-Historical-Data-Annual (1).xlsx"

# ----------------------------
# Configuration
# ----------------------------
COUNTRIES = {
    "South Africa": "🇿🇦", "Namibia": "🇳🇦", "Zambia": "🇿🇲", "DR Congo": "🇨🇩",
    "Zimbabwe": "🇿🇼", "Botswana": "🇧🇼", "Ghana": "🇬🇭", "Mali": "🇲🇱",
    "Tanzania": "🇹🇿", "Sudan": "🇸🇩", "Ethiopia": "🇪🇹", "Mozambique": "🇲🇿",
    "Nigeria": "🇳🇬", "Morocco": "🇲🇦", "Mauritania": "🇲🇷", "Senegal": "🇸🇳",
    "Uganda": "🇺🇬", "Angola": "🇦🇴", "Burkina Faso": "🇧🇫", "Kenya": "🇰🇪",
    "Egypt": "🇪🇬", "Algeria": "🇩🇿", "Libya": "🇱🇾", "Tunisia": "🇹🇳",
    "Cameroon": "🇨🇲", "Ivory Coast": "🇨🇮"
}
