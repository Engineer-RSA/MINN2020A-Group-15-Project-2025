# --- Automatic Dependency Installation Helper ---
import subprocess
import sys
import importlib
import os
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

def install_and_import(package_name, import_name=None):
    """
    Attempts to import a package, installs it via pip if not found.
    """
    if import_name is None:
        import_name = package_name
    try:
        # Attempt to import the module
        module = importlib.import_module(import_name)
        print(f"✓ Found required package: {package_name}")
        return module
    except ImportError:
        print(f"⚠ Missing required package: {package_name}. Installing...")
        try:
            # Use subprocess to call pip install
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print(f"✓ Successfully installed {package_name}")
            # Now try to import again after installation
            module = importlib.import_module(import_name)
            print(f"✓ Successfully imported {package_name} after installation.")
            return module
        except subprocess.CalledProcessError as e:
            print(f"✗ Failed to install {package_name}: {e}")
            sys.exit(1)
        except ImportError as e:
            print(f"✗ Failed to import {package_name} even after installation: {e}")
            sys.exit(1)




# --- End of Automatic Dependency Installation Helper ---

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

AFRICA_COUNTRIES = [
    "Algeria", "Angola", "Benin", "Botswana", "Burkina Faso", "Burundi", "Cameroon",
    "Cape Verde", "Central African Republic", "Chad", "Comoros", "Congo", "DR Congo",
    "Djibouti", "Egypt", "Equatorial Guinea", "Eritrea", "Eswatini", "Ethiopia",
    "Gabon", "Gambia", "Ghana", "Guinea", "Guinea-Bissau", "Ivory Coast", "Kenya",
    "Lesotho", "Liberia", "Libya", "Madagascar", "Malawi", "Mali", "Mauritania",
    "Mauritius", "Morocco", "Mozambique", "Namibia", "Niger", "Nigeria", "Rwanda",
    "Sao Tome and Principe", "Senegal", "Seychelles", "Sierra Leone", "Somalia",
    "South Africa", "South Sudan", "Sudan", "Tanzania", "Togo", "Tunisia", "Uganda",
    "Zambia", "Zimbabwe"
]

CRITICAL_MINERALS = [
    "Cobalt", "Lithium", "Nickel", "Copper", "Graphite", "Manganese",
    "Rare Earths", "Platinum", "Gold", "Diamonds", "Phosphates", "Uranium",
    "Iron", "Aluminum", "Tin", "Zinc", "Lead"
]

ALLOWED_INSTITUTIONS = ["Wits", "UJ", "UNISA", "UP", "UCT", "Stellenbosch", "Rhodes"]

# Primary Admin Configuration
PRIMARY_ADMINS = ["admin1", "admin2", "admin3", "admin4"]
PRIMARY_ADMIN_PIN = "2024"  # Unique PIN for primary admins

# ----------------------------
# Color Scheme
# ----------------------------
COLORS = {
    "primary": "#1a3e2d",
    "secondary": "#2d5c47",
    "accent": "#d4af37",
    "highlight": "#e94560",
    "text_light": "#ffffff",
    "text_muted": "#b0b0b0",
    "success": "#4CAF50",
    "warning": "#FF9800",
    "error": "#f44336",
    "card_bg": "#1f3a32",
    "input_bg": "#162a24",
    "border": "#3a5b4d",
    "background": "#0d1f18",
    "gold_accent": "#d4af37",
    "crystal_red": "#c92a2a",
    "crystal_blue": "#2563eb",
    "crystal_green": "#16a34a"
}

# ----------------------------
# Security Functions
# ----------------------------
def hash_password(password):
    """Hash password using SHA-256"""
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(password, hashed):
    """Verify password against hash"""
    return hash_password(password) == hashed

def is_primary_admin(username):
    """Check if user is one of the primary admins"""
    return username in PRIMARY_ADMINS

def get_admin_count():
    """Count how many admin users exist"""
    users = load_users()
    admin_count = 0
    for user_data in users.values():
        if user_data.get("role", "").lower() == "admin":
            admin_count += 1
    return admin_count

def can_create_admin(current_username, requested_role):
    """Check if admin creation is allowed"""
    if requested_role.lower() != "admin":
        return True
    
    users = load_users()
    current_user = users.get(current_username, {})
    
    # Only existing admins can create new admin accounts
    if current_user.get("role", "").lower() != "admin":
        return False
    
    # Primary admins can create unlimited admins
    if is_primary_admin(current_username):
        return True
    
    # Non-primary admins can only create admins if total is less than 10
    return get_admin_count() < 10

# ----------------------------
# CMO Data Loading
# ----------------------------
def load_cmo_data():
    try:
        annual_df = pd.read_excel(CMO_ANNUAL_FILE, sheet_name=0, skiprows=4)
        monthly_df = pd.read_excel(CMO_MONTHLY_FILE, sheet_name=0, skiprows=4)

        annual_df.columns = [str(c).strip() for c in annual_df.columns]
        monthly_df.columns = [str(c).strip() for c in monthly_df.columns]

        annual_df = annual_df.rename(columns={annual_df.columns[0]: 'Year'})
        monthly_df = monthly_df.rename(columns={monthly_df.columns[0]: 'DateStr'})

        annual_df['Year'] = pd.to_numeric(annual_df['Year'], errors='coerce')
        annual_df = annual_df.dropna(subset=['Year'])
        annual_df['Year'] = annual_df['Year'].astype(int)

        monthly_df['Date'] = pd.to_datetime(monthly_df['DateStr'].str.replace('M', '-'), format='%Y-%m')

        cmo_minerals = {
            "Copper": ["COPPER", "KCOPPER"],
            "Nickel": ["NICKEL", "KNICKEL"],
            "Aluminum": ["ALUMINUM", "KALUMINUM"],
            "Lead": ["LEAD", "KLEAD"],
            "Tin": ["Tin", "KTin"],
            "Zinc": ["Zinc", "KZinc"],
            "Gold": ["GOLD", "KGOLD"],
            "Platinum": ["PLATINUM", "KPLATINUM"],
            "Iron Ore": ["IRON_ORE", "KIRON_ORE"],
            "Phosphate Rock": ["PHOSROCK", "KPHOSROCK"],
        }

        annual_cols = ['Year'] + [col for cols in cmo_minerals.values() for col in cols if col in annual_df.columns]
        monthly_cols = ['Date', 'DateStr'] + [col for cols in cmo_minerals.values() for col in cols if col in monthly_df.columns]

        annual_clean = annual_df[annual_cols].copy()
        monthly_clean = monthly_df[monthly_cols].copy()

        return annual_clean, monthly_clean, cmo_minerals

    except Exception as e:
        print(f"⚠️ Error loading CMO data: {e}")
        empty_df = pd.DataFrame(columns=["Year"])
        return empty_df, pd.DataFrame(columns=["Date", "DateStr"]), {}

cmo_annual_df, cmo_monthly_df, CMO_MINERALS = load_cmo_data()
CMO_MINERAL_NAMES = list(CMO_MINERALS.keys())

# ----------------------------
# User Management
# ----------------------------
def initialize_users_file():
    """Initialize users file with primary admins if empty"""
    if not os.path.exists(USERS_FILE):
        initial_users = {}
        
        # Create primary admin accounts with default passwords
        for admin in PRIMARY_ADMINS:
            initial_users[admin] = {
                "password": hash_password("Admin123!"),  # Default password for primary admins
                "role": "Admin",
                "date_created": str(datetime.now()),
                "name": f"Primary {admin.capitalize()}",
                "surname": "Administrator",
                "age": "30",
                "institution": "Wits",
                "id_number": "8001015000089",
                "nationality": "South Africa",
                "gender": "MALE",
                "email": f"{admin}@mineralxafrica.com",
                "is_primary": True,
                "status": "active"
            }
        
        with open(USERS_FILE, 'w') as f:
            json.dump(initial_users, f, indent=4)

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_users(users_dict):
    with open(USERS_FILE, 'w') as f:
        json.dump(users_dict, f, indent=4)

def validate_email(email):
    if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
        return False, "Email must be a valid email address"
    return True, "Valid email"

def is_email_unique(email, current_username=None):
    users = load_users()
    for username, user_data in users.items():
        if user_data.get("email") == email:
            if current_username is None or username != current_username:
                return False
    return True

def validate_sa_id_number(id_number):
    if not re.match(r'^\d{13}$', id_number):
        return False, "ID must be exactly 13 digits"
    first_two = int(id_number[:2])
    if not ((0 <= first_two <= 5) or (26 <= first_two <= 99)):
        return False, "ID must start with numbers between 00-05 or 26-99"
    month = int(id_number[2:4])
    day = int(id_number[4:6])
    if not (1 <= month <= 12):
        return False, "Invalid month in ID number"
    if not (1 <= day <= 31):
        return False, "Invalid day in ID number"
    if month in [4, 6, 9, 11] and day > 30:
        return False, "Invalid day for this month"
    if month == 2 and day > 29:
        return False, "Invalid day for February"
    return True, "Valid SA ID"

def calculate_age_from_id(id_number):
    try:
        year_part = int(id_number[:2])
        month = int(id_number[2:4])
        day = int(id_number[4:6])
        birth_year = 2000 + year_part if 0 <= year_part <= 7 else 1900 + year_part
        today = datetime.now()
        age = today.year - birth_year
        if today.month < month or (today.month == month and today.day < day):
            age -= 1
        return age
    except:
        return None

def register_user(username, password, role, name, surname, age, institution, id_number, nationality, gender, email, current_admin=None):
    users = load_users()
    
    if username in users:
        return False, "Username already exists!"
    
    # Check admin creation permissions
    if role.lower() == "admin":
        if not can_create_admin(current_admin, role):
            admin_count = get_admin_count()
            if current_admin and is_primary_admin(current_admin):
                return False, f"Admin creation failed. Please contact system administrator."
            else:
                return False, f"Admin creation restricted. Maximum admin limit reached or insufficient privileges."
    
    email_valid, email_msg = validate_email(email)
    if not email_valid:
        return False, email_msg
    
    if not is_email_unique(email):
        return False, "Email is already registered by another user!"
    
    if institution not in ALLOWED_INSTITUTIONS:
        return False, f"Institution must be one of: {', '.join(ALLOWED_INSTITUTIONS)}"
    
    try:
        age_int = int(age)
        if age_int < 18 or age_int > 120:
            return False, "Age must be between 18 and 120!"
    except ValueError:
        return False, "Age must be a valid number!"
    
    id_valid, id_msg = validate_sa_id_number(id_number)
    if not id_valid:
        return False, id_msg
    
    calculated_age = calculate_age_from_id(id_number)
    if calculated_age is not None and abs(calculated_age - age_int) > 1:
        return False, f"Age ({age_int}) doesn't match ID number (calculated age: {calculated_age})"
    
    users[username] = {
        "password": hash_password(password),
        "role": role,
        "date_created": str(datetime.now()),
        "name": name,
        "surname": surname,
        "age": age,
        "institution": institution,
        "id_number": id_number,
        "nationality": nationality,
        "gender": gender,
        "email": email,
        "is_primary": username in PRIMARY_ADMINS,
        "status": "active",
        "created_by": current_admin
    }
    save_users(users)
    return True, f"User {username} registered successfully as {role}!"

def authenticate_user(login_input, password, admin_pin=None):
    users = load_users()
    email_valid, _ = validate_email(login_input)
    
    if email_valid:
        for username, user_data in users.items():
            if user_data.get("email") == login_input and user_data.get("status") == "active":
                if verify_password(password, user_data["password"]):
                    # Check if primary admin requires PIN
                    if user_data.get("is_primary") and admin_pin != PRIMARY_ADMIN_PIN:
                        return False, "pin_required", username
                    return True, user_data["role"], username
    else:
        if login_input in users and users[login_input].get("status") == "active":
            if verify_password(password, users[login_input]["password"]):
                # Check if primary admin requires PIN
                if users[login_input].get("is_primary") and admin_pin != PRIMARY_ADMIN_PIN:
                    return False, "pin_required", login_input
                return True, users[login_input]["role"], login_input
    return False, None, None

def update_user_status(username, status, admin_username):
    """Update user status (active/suspended)"""
    users = load_users()
    if username not in users:
        return False, "User not found!"
    
    if users[username].get("is_primary"):
        return False, "Cannot modify primary admin accounts!"
    
    users[username]["status"] = status
    users[username]["modified_by"] = admin_username
    users[username]["modified_date"] = str(datetime.now())
    save_users(users)
    return True, f"User {username} {status} successfully!"

def change_user_role(username, new_role, admin_username):
    """Change user role"""
    users = load_users()
    if username not in users:
        return False, "User not found!"
    
    if users[username].get("is_primary"):
        return False, "Cannot modify primary admin roles!"
    
    if new_role.lower() == "admin" and not can_create_admin(admin_username, new_role):
        return False, "Admin creation restricted. Maximum admin limit reached or insufficient privileges."
    
    users[username]["role"] = new_role
    users[username]["modified_by"] = admin_username
    users[username]["modified_date"] = str(datetime.now())
    save_users(users)
    return True, f"User {username} role changed to {new_role} successfully!"

def reset_user_password(username, new_password, admin_username):
    """Reset user password (admin function)"""
    users = load_users()
    if username not in users:
        return False, "User not found!"
    
    users[username]["password"] = hash_password(new_password)
    users[username]["password_reset_by"] = admin_username
    users[username]["password_reset_date"] = str(datetime.now())
    save_users(users)
    return True, f"Password for {username} reset successfully!"

# ----------------------------
# Sample Data
# ----------------------------
def create_sample_data():
    print("Creating comprehensive sample data for demonstration...")
    sample_countries = ["South Africa", "DR Congo", "Zambia", "Zimbabwe", "Ghana",
                        "Namibia", "Botswana", "Morocco", "Mali", "Tanzania",
                        "Egypt", "Algeria", "Nigeria", "Kenya", "Angola"]
    sample_minerals = ["Cobalt", "Lithium", "Copper", "Gold", "Platinum",
                       "Diamonds", "Manganese", "Uranium", "Iron", "Aluminum",
                       "Phosphates", "Graphite", "Nickel", "Tin", "Zinc"]
    base_values = {
        "South Africa": {"Platinum": 120000, "Gold": 130000, "Diamonds": 8500000, "Manganese": 6800000, "Iron": 78000000},
        "DR Congo": {"Cobalt": 130000, "Copper": 2200000, "Gold": 45000, "Tin": 85000},
        "Zambia": {"Copper": 850000, "Cobalt": 8000, "Lead": 45000},
        "Zimbabwe": {"Lithium": 1200, "Platinum": 15000, "Gold": 25000, "Diamonds": 4200000},
        "Ghana": {"Gold": 130000, "Bauxite": 850000, "Manganese": 2800000},
        "Namibia": {"Uranium": 5500, "Diamonds": 2200000, "Zinc": 85000},
        "Botswana": {"Diamonds": 22000000, "Copper": 45000, "Nickel": 12000},
        "Morocco": {"Phosphates": 38000000, "Lead": 65000},
        "Mali": {"Gold": 65000},
        "Tanzania": {"Gold": 55000, "Tanzanite": 1200, "Diamonds": 1800000},
        "Egypt": {"Phosphates": 6500000, "Iron": 3500000},
        "Algeria": {"Iron": 2800000, "Zinc": 45000},
        "Nigeria": {"Tin": 85000, "Lead": 65000},
        "Kenya": {"Soda Ash": 450000, "Fluorspar": 120000},
        "Angola": {"Diamonds": 8500000, "Iron": 6500000}
    }
    data = []
    for year in [2019, 2020, 2021, 2022, 2023]:
        for country in sample_countries:
            for mineral in sample_minerals:
                if country in base_values and mineral in base_values[country]:
                    production = base_values[country][mineral] * (0.9 + 0.2 * (year - 2019) / 4)
                else:
                    production = hash(f"{country}{mineral}{year}") % 50000 + 1000
                data.append({
                    'Country': country,
                    'Mineral': mineral,
                    'Year': year,
                    'Production': int(production)
                })
    return pd.DataFrame(data)

# ----------------------------
# Mineral Data Loading
# ----------------------------
def load_mineral_data():
    if not os.path.exists(FILE_6_4):
        print(f"⚠️ File not found: {FILE_6_4}")
        return create_sample_data()
    try:
        xls = pd.ExcelFile(FILE_6_4, engine="openpyxl")
        sheet_names = xls.sheet_names

        def process_sheet(sheet_name):
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=1, engine="openpyxl")
                df = df.dropna(how="all")
                df.columns = [str(c).strip() for c in df.columns]
                country_col = df.columns[0]
                year_cols = [c for c in df.columns[1:] if re.search(r"(20\d{2})", str(c))]
                if not year_cols:
                    year_cols = [c for c in df.columns[1:] if pd.api.types.is_numeric_dtype(df[c]) or str(c).strip().isdigit()]
                if not year_cols:
                    return pd.DataFrame(columns=["Country", "Mineral", "Year", "Production"])
                df_long = df.melt(id_vars=[country_col], value_vars=year_cols, var_name="Year", value_name="Production")
                df_long = df_long.rename(columns={country_col: "Country"})
                df_long = df_long[~df_long["Country"].astype(str).str.strip().str.lower().isin(["total", "world"])]
                df_long["Year"] = df_long["Year"].astype(str).str.extract(r"(20\d{2})", expand=False)
                df_long = df_long[df_long["Year"].notna()].copy()
                df_long["Year"] = df_long["Year"].astype(int)
                df_long["Mineral"] = sheet_name
                df_long["Production"] = pd.to_numeric(df_long["Production"], errors="coerce")
                df_long = df_long.dropna(subset=["Production"]).copy()
                df_long["Country"] = df_long["Country"].astype(str).str.strip()
                return df_long[["Country", "Mineral", "Year", "Production"]]
            except Exception as e:
                print(f"Error processing sheet {sheet_name}: {e}")
                return pd.DataFrame(columns=["Country", "Mineral", "Year", "Production"])

        rows = []
        for s in sheet_names:
            df_s = process_sheet(s)
            if not df_s.empty:
                rows.append(df_s)
        if rows:
            combined = pd.concat(rows, ignore_index=True)
            combined = combined[combined["Country"].isin(AFRICA_COUNTRIES)]
            combined = combined[combined["Year"] >= 2019]
            combined = combined[combined["Mineral"].isin(CRITICAL_MINERALS)]
            return combined
        else:
            print("No valid data found in any sheet")
            return create_sample_data()
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return create_sample_data()

flat_critical_africa = load_mineral_data()

# ----------------------------
# Splash Screen
# ----------------------------
class SplashScreen:
    def __init__(self, root):
        self.root = root
        self.root.title("MineralXAfrica - Loading...")
        self.root.geometry("1000x800")
        self.root.resizable(False, False)
        self.root.configure(bg=COLORS["background"])

        try:
            from PIL import Image, ImageTk
            logo_path = "mineralxafrica_logo.png"
            if os.path.exists(logo_path):
                img = Image.open(logo_path)
                img = img.resize((800, 800), Image.Resampling.LANCZOS)
                img = img.convert("RGBA")
                alpha = Image.new("L", img.size, 255)
                img.putalpha(alpha)
                self.logo_img = ImageTk.PhotoImage(img)
                bg_canvas = tk.Canvas(self.root, width=1000, height=800, bg=COLORS["background"], highlightthickness=0)
                bg_canvas.pack(fill="both", expand=True)
                self.bg_image = bg_canvas.create_image(500, 400, image=self.logo_img, anchor="center")
                self.canvas = bg_canvas
            else:
                self.canvas = tk.Canvas(self.root, width=1000, height=800, bg=COLORS["background"], highlightthickness=0)
                self.canvas.pack(fill="both", expand=True)
        except ImportError:
            self.canvas = tk.Canvas(self.root, width=1000, height=800, bg=COLORS["background"], highlightthickness=0)
            self.canvas.pack(fill="both", expand=True)

        self.root.after(3000, self.fade_out_logo)

    def fade_out_logo(self):
        self.alpha = 255
        self.fade_step()

    def fade_step(self):
        if self.alpha > 0:
            self.alpha -= 10
            try:
                from PIL import Image, ImageTk
                logo_path = "mineralxafrica_logo.png"
                if os.path.exists(logo_path):
                    img = Image.open(logo_path)
                    img = img.resize((800, 800), Image.Resampling.LANCZOS)
                    img = img.convert("RGBA")
                    alpha = Image.new("L", img.size, self.alpha)
                    img.putalpha(alpha)
                    self.logo_img = ImageTk.PhotoImage(img)
                    self.canvas.itemconfig(self.bg_image, image=self.logo_img)
            except ImportError:
                pass
            self.root.after(50, self.fade_step)
        else:
            self.canvas.destroy()
            login_system = LoginSystem()
            self.root.destroy()

# ----------------------------
# Enhanced Login System with Better User Guidance
# ----------------------------
class LoginSystem:
    def __init__(self):
        self.root = None
        self.registration_window = None
        self.admin_pin_window = None
        self.login_window()

    def login_window(self):
        self.root = tk.Tk()
        self.root.title("MineralXAfrica - Login System")
        self.root.geometry("1000x800")
        self.root.resizable(False, False)
        self.root.configure(bg=COLORS["background"])

        main_container = tk.Frame(self.root, bg=COLORS["primary"], bd=0, relief="flat")
        main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # Header
        header_frame = tk.Frame(main_container, bg=COLORS["primary"])
        header_frame.pack(fill="x", pady=(0, 20))
        title_label = tk.Label(header_frame, text="💎 MineralXAfrica Portal",
                               font=("Arial", 24, "bold"),
                               fg=COLORS["text_light"], bg=COLORS["primary"])
        title_label.pack(pady=10)
        subtitle_label = tk.Label(header_frame, text="African Critical Minerals Intelligence Platform",
                                  font=("Arial", 12),
                                  fg=COLORS["text_muted"], bg=COLORS["primary"])
        subtitle_label.pack()

        # Content Frame
        content_frame = tk.Frame(main_container, bg=COLORS["primary"])
        content_frame.pack(fill="both", expand=True)

        # Login Panel (Left) - 60% width
        login_panel = tk.Frame(content_frame, bg=COLORS["secondary"], bd=2, relief="ridge",
                              highlightbackground=COLORS["border"], highlightthickness=1)
        login_panel.pack(side="left", fill="both", expand=True, padx=(0, 10))

        # Registration Panel (Right) - 40% width
        reg_panel = tk.Frame(content_frame, bg=COLORS["secondary"], bd=2, relief="ridge",
                            highlightbackground=COLORS["border"], highlightthickness=1)
        reg_panel.pack(side="right", fill="both", expand=True, padx=(10, 0))

        self.setup_login_panel(login_panel)
        self.setup_registration_panel(reg_panel)

        # Footer
        self.setup_footer(main_container)

        self.root.bind('<Return>', lambda event: self.login())
        self.root.mainloop()

    def setup_login_panel(self, parent):
        login_header = tk.Frame(parent, bg=COLORS["accent"], height=50)
        login_header.pack(fill="x", pady=(0, 20))
        login_header.pack_propagate(False)
        tk.Label(login_header, text="🔐 User Login", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["accent"]).pack(expand=True)

        content_frame = tk.Frame(parent, bg=COLORS["secondary"], padx=30, pady=20)
        content_frame.pack(fill="both", expand=True)

        # Login Instructions
        instructions_text = """Please enter your login credentials:

• Username OR Email address
• Your password
• Click the eye icon to show/hide password

For Primary Admins (admin1, admin2, admin3, admin4):
You will be prompted for an additional security PIN after password verification."""
        
        instructions_label = tk.Label(content_frame, text=instructions_text, font=("Arial", 9),
                                     fg=COLORS["text_light"], bg=COLORS["secondary"],
                                     justify="left", wraplength=400)
        instructions_label.pack(pady=(0, 20))

        # Username/Email Field
        username_frame = tk.Frame(content_frame, bg=COLORS["secondary"])
        username_frame.pack(fill="x", pady=(0, 15))
        
        tk.Label(username_frame, text="Username or Email:", font=("Arial", 11, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["secondary"]).pack(anchor="w")
        
        tk.Label(username_frame, text="Enter your registered username OR email address", font=("Arial", 8),
                 fg=COLORS["text_muted"], bg=COLORS["secondary"]).pack(anchor="w")
        
        self.login_entry = self.create_styled_entry(username_frame)
        self.login_entry.pack(fill="x", pady=(5, 0))
        self.login_entry.focus()

        # Password Field with Show/Hide
        password_frame = tk.Frame(content_frame, bg=COLORS["secondary"])
        password_frame.pack(fill="x", pady=(0, 10))
        
        tk.Label(password_frame, text="Password:", font=("Arial", 11, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["secondary"]).pack(anchor="w")
        
        tk.Label(password_frame, text="Enter your account password", font=("Arial", 8),
                 fg=COLORS["text_muted"], bg=COLORS["secondary"]).pack(anchor="w")
        
        # Password entry with show/hide button
        password_input_frame = tk.Frame(password_frame, bg=COLORS["secondary"])
        password_input_frame.pack(fill="x", pady=(5, 0))
        
        self.password_entry = self.create_styled_entry(password_input_frame, show="•")
        self.password_entry.pack(side="left", fill="x", expand=True)
        
        # Show/Hide password button
        self.show_password_btn = tk.Button(password_input_frame, text="👁️", font=("Arial", 10),
                                          bg=COLORS["primary"], fg=COLORS["text_light"],
                                          relief="flat", width=3, command=self.toggle_password_visibility)
        self.show_password_btn.pack(side="right", padx=(5, 0))
        self.password_visible = False

        # Login button
        login_btn = tk.Button(content_frame, text="Login to Dashboard", font=("Arial", 12, "bold"),
                              bg=COLORS["highlight"], fg=COLORS["text_light"],
                              activebackground=COLORS["accent"], activeforeground=COLORS["text_light"],
                              relief="flat", padx=20, pady=12, cursor="hand2", command=self.login)
        login_btn.pack(fill="x", pady=20)

        # Test credentials button
        test_btn = tk.Button(content_frame, text="Fill Test Credentials", font=("Arial", 10),
                             bg=COLORS["warning"], fg=COLORS["text_light"],
                             activebackground=COLORS["accent"], activeforeground=COLORS["text_light"],
                             relief="flat", padx=10, pady=6, cursor="hand2", command=self.fill_test_credentials)
        test_btn.pack(fill="x", pady=5)

        # Admin PIN info
        admin_info = tk.Label(content_frame, text=f"Primary Admin PIN: {PRIMARY_ADMIN_PIN} (for testing)", 
                             font=("Arial", 8), fg=COLORS["text_muted"], bg=COLORS["secondary"])
        admin_info.pack(pady=(10, 0))

    def setup_registration_panel(self, parent):
        reg_header = tk.Frame(parent, bg=COLORS["accent"], height=50)
        reg_header.pack(fill="x", pady=(0, 20))
        reg_header.pack_propagate(False)
        tk.Label(reg_header, text="📝 New User Registration", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["accent"]).pack(expand=True)

        content_frame = tk.Frame(parent, bg=COLORS["secondary"], padx=30, pady=40)
        content_frame.pack(fill="both", expand=True)

        # Registration description
        desc_text = """New to MineralXAfrica?

Create your account to access:
• Mineral production analytics
• Interactive Africa maps
• Market intelligence data
• Custom dashboards and reports

Click below to start the registration process where you'll provide:
• Personal information
• Institutional details
• South African ID verification
• Contact information"""
        
        desc_label = tk.Label(content_frame, text=desc_text, font=("Arial", 10),
                             fg=COLORS["text_light"], bg=COLORS["secondary"],
                             justify="left", wraplength=300)
        desc_label.pack(pady=(0, 30))

        # Admin creation warning
        users = load_users()
        admin_exists = any(user_data.get("role", "").lower() == "admin" for user_data in users.values())
        
        if admin_exists:
            warning_text = "🔒 Admin accounts can only be created by existing administrators. Self-registration is for Investor/Researcher roles only."
            warning_label = tk.Label(content_frame, text=warning_text, font=("Arial", 9),
                                   fg=COLORS["warning"], bg=COLORS["secondary"],
                                   justify="center", wraplength=300)
            warning_label.pack(pady=(0, 20))

        # Register button
        register_btn = tk.Button(content_frame, text="Start Registration", font=("Arial", 12, "bold"),
                                bg=COLORS["success"], fg=COLORS["text_light"],
                                activebackground=COLORS["accent"], activeforeground=COLORS["text_light"],
                                relief="flat", padx=20, pady=15, cursor="hand2", 
                                command=self.open_registration_form)
        register_btn.pack(fill="x", pady=10)

        # Quick info
        info_text = """Already have an account?
Use the login form on the left to access your dashboard."""
        info_label = tk.Label(content_frame, text=info_text, font=("Arial", 9),
                             fg=COLORS["text_muted"], bg=COLORS["secondary"],
                             justify="center")
        info_label.pack(pady=(20, 0))

    def setup_footer(self, parent):
        footer_frame = tk.Frame(parent, bg=COLORS["primary"])
        footer_frame.pack(fill="x", pady=(20, 0))

        stats_frame = tk.Frame(footer_frame, bg=COLORS["card_bg"], padx=10, pady=10)
        stats_frame.pack(fill="x", pady=10)
        stats_data = [("500+", "Mineral Analysts"), ("50+", "African Countries"), ("1M+", "Data Points")]
        for i, (number, label) in enumerate(stats_data):
            stat = tk.Frame(stats_frame, bg=COLORS["card_bg"])
            stat.grid(row=0, column=i, padx=20, sticky="nsew")
            tk.Label(stat, text=number, font=("Arial", 14, "bold"),
                     fg=COLORS["highlight"], bg=COLORS["card_bg"]).pack()
            tk.Label(stat, text=label, font=("Arial", 9),
                     fg=COLORS["text_muted"], bg=COLORS["card_bg"]).pack()

        footer_label = tk.Label(footer_frame, text="© 2024 Mineral Analytics Platform | Secure Access Required",
                                font=("Arial", 9),
                                fg=COLORS["text_muted"], bg=COLORS["primary"])
        footer_label.pack(pady=(10, 0))

    def toggle_password_visibility(self):
        """Toggle password visibility with visual feedback"""
        if self.password_visible:
            self.password_entry.config(show="•")
            self.show_password_btn.config(text="👁️", bg=COLORS["primary"])
            self.password_visible = False
        else:
            self.password_entry.config(show="")
            self.show_password_btn.config(text="🔒", bg=COLORS["accent"])
            self.password_visible = True

    def show_admin_pin_dialog(self, username):
        """Show PIN dialog for primary admin authentication"""
        if self.admin_pin_window and self.admin_pin_window.winfo_exists():
            self.admin_pin_window.lift()
            return
        
        self.admin_pin_window = tk.Toplevel(self.root)
        self.admin_pin_window.title("Primary Admin Security Verification")
        self.admin_pin_window.geometry("400x300")
        self.admin_pin_window.configure(bg=COLORS["background"])
        self.admin_pin_window.resizable(False, False)
        self.admin_pin_window.transient(self.root)
        self.admin_pin_window.grab_set()
        
        # Center the PIN window
        self.admin_pin_window.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.admin_pin_window.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.admin_pin_window.winfo_height() // 2)
        self.admin_pin_window.geometry(f"+{x}+{y}")
        
        main_frame = tk.Frame(self.admin_pin_window, bg=COLORS["primary"], padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)
        
        # Header
        tk.Label(main_frame, text="🔒 Primary Admin Verification", font=("Arial", 16, "bold"),
                fg=COLORS["text_light"], bg=COLORS["primary"]).pack(pady=(0, 10))
        
        tk.Label(main_frame, text=f"Welcome, {username}", font=("Arial", 12),
                fg=COLORS["text_muted"], bg=COLORS["primary"]).pack(pady=(0, 20))
        
        # Instructions
        instructions = """Primary Administrator Access requires additional security verification.

Please enter the Primary Admin PIN to continue:"""
        
        tk.Label(main_frame, text=instructions, font=("Arial", 10),
                fg=COLORS["text_light"], bg=COLORS["primary"], justify="center", wraplength=350).pack(pady=(0, 20))
        
        # PIN Entry
        pin_frame = tk.Frame(main_frame, bg=COLORS["primary"])
        pin_frame.pack(pady=10)
        
        tk.Label(pin_frame, text="Admin PIN:", font=("Arial", 11, "bold"),
                fg=COLORS["text_light"], bg=COLORS["primary"]).pack()
        
        self.pin_entry = self.create_styled_entry(pin_frame, show="•", width=15, justify="center")
        self.pin_entry.pack(pady=5)
        self.pin_entry.focus()
        
        # Show/Hide PIN button
        pin_btn_frame = tk.Frame(main_frame, bg=COLORS["primary"])
        pin_btn_frame.pack(pady=5)
        
        self.show_pin_btn = tk.Button(pin_btn_frame, text="👁️ Show PIN", font=("Arial", 9),
                                     bg=COLORS["primary"], fg=COLORS["text_light"],
                                     command=self.toggle_pin_visibility)
        self.show_pin_btn.pack()
        self.pin_visible = False
        
        # Buttons
        btn_frame = tk.Frame(main_frame, bg=COLORS["primary"])
        btn_frame.pack(pady=20)
        
        tk.Button(btn_frame, text="Verify PIN", font=("Arial", 11, "bold"),
                 bg=COLORS["success"], fg=COLORS["text_light"],
                 command=lambda: self.verify_admin_pin(username), width=12).pack(side="left", padx=5)
        
        tk.Button(btn_frame, text="Cancel", font=("Arial", 11),
                 bg=COLORS["error"], fg=COLORS["text_light"],
                 command=self.cancel_admin_login, width=12).pack(side="left", padx=5)
        
        # Test info
        tk.Label(main_frame, text=f"Test PIN: {PRIMARY_ADMIN_PIN}", font=("Arial", 8),
                fg=COLORS["text_muted"], bg=COLORS["primary"]).pack(pady=(10, 0))

    def toggle_pin_visibility(self):
        """Toggle PIN visibility"""
        if self.pin_visible:
            self.pin_entry.config(show="•")
            self.show_pin_btn.config(text="👁️ Show PIN")
            self.pin_visible = False
        else:
            self.pin_entry.config(show="")
            self.show_pin_btn.config(text="🔒 Hide PIN")
            self.pin_visible = True

    def verify_admin_pin(self, username):
        """Verify the admin PIN and complete login"""
        pin = self.pin_entry.get().strip()
        
        if not pin:
            messagebox.showerror("Error", "Please enter the Admin PIN!")
            return
        
        if pin != PRIMARY_ADMIN_PIN:
            messagebox.showerror("Error", "Invalid Admin PIN! Access denied.")
            self.pin_entry.delete(0, tk.END)
            self.pin_entry.focus()
            return
        
        # PIN is correct, complete the login
        if self.admin_pin_window:
            self.admin_pin_window.destroy()
        
        messagebox.showinfo("Login Successful", f"Welcome Primary Admin {username}!")
        self.root.destroy()
        dash_root = tk.Tk()
        DashboardApp(dash_root, username, "Admin")
        dash_root.mainloop()

    def cancel_admin_login(self):
        """Cancel the admin login process"""
        if self.admin_pin_window:
            self.admin_pin_window.destroy()
        messagebox.showinfo("Login Cancelled", "Primary admin login was cancelled.")

    def create_styled_entry(self, parent, **kwargs):
        entry = tk.Entry(parent, font=("Arial", 10), bg=COLORS["input_bg"], fg=COLORS["text_light"],
                         insertbackground=COLORS["text_light"], relief="flat", bd=2,
                         highlightbackground=COLORS["border"], highlightcolor=COLORS["highlight"],
                         highlightthickness=1, **kwargs)
        return entry

    def show_loading(self, message="Loading..."):
        loading_window = tk.Toplevel(self.root)
        loading_window.title("Please Wait")
        loading_window.geometry("300x150")
        loading_window.configure(bg=COLORS["primary"])
        loading_window.transient(self.root)
        loading_window.grab_set()
        loading_window.resizable(False, False)
        
        x = (self.root.winfo_screenwidth() // 2) - (loading_window.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (loading_window.winfo_height() // 2)
        loading_window.geometry(f"+{x}+{y}")
        
        tk.Label(loading_window, text=message, font=("Arial", 12),
                 fg=COLORS["text_light"], bg=COLORS["primary"]).pack(expand=True, pady=20)
        progress = ttk.Progressbar(loading_window, mode='indeterminate')
        progress.pack(fill="x", padx=50, pady=10)
        progress.start()
        return loading_window

    def fill_test_credentials(self):
        """Fill test credentials for demonstration"""
        self.login_entry.delete(0, tk.END)
        self.login_entry.insert(0, "admin1")
        self.password_entry.delete(0, tk.END)
        self.password_entry.insert(0, "Admin123!")
        
        messagebox.showinfo("Test Credentials", 
                           "Primary Admin credentials filled!\n\n"
                           "Username: admin1\n"
                           "Password: Admin123!\n\n"
                           f"After clicking Login, you'll need to enter the Primary Admin PIN: {PRIMARY_ADMIN_PIN}")

    def login(self):
        """Handle user login with enhanced guidance"""
        login_input = self.login_entry.get().strip()
        password = self.password_entry.get().strip()
        
        # Clear any previous error highlighting
        self.login_entry.config(highlightcolor=COLORS["highlight"])
        self.password_entry.config(highlightcolor=COLORS["highlight"])
        
        # Validation with clear error messages
        if not login_input and not password:
            messagebox.showerror("Input Required", 
                                "Please enter both:\n"
                                "• Your username OR email address\n"
                                "• Your password")
            self.login_entry.focus()
            return
        elif not login_input:
            messagebox.showerror("Username/Email Required", 
                                "Please enter your username OR email address")
            self.login_entry.focus()
            return
        elif not password:
            messagebox.showerror("Password Required", 
                                "Please enter your password")
            self.password_entry.focus()
            return
        
        # Show loading
        loading_window = self.show_loading("Verifying credentials...")
        self.root.update()
        
        # Authenticate user (without PIN first)
        success, status, username = authenticate_user(login_input, password)
        loading_window.destroy()
        
        if success:
            # Regular user login successful
            messagebox.showinfo("Login Successful", f"Welcome {username}! Role: {status}")
            self.root.destroy()
            dash_root = tk.Tk()
            DashboardApp(dash_root, username, status)
            dash_root.mainloop()
        elif status == "pin_required":
            # Primary admin requires PIN verification
            self.show_admin_pin_dialog(username)
        else:
            # Login failed
            messagebox.showerror("Login Failed", 
                                "Invalid username/email or password!\n\n"
                                "Please check:\n"
                                "• Username/email spelling\n"
                                "• Password correctness\n"
                                "• Caps Lock status\n"
                                "• Account status (contact admin if suspended)")

    def open_registration_form(self):
        """Open a separate window for user registration"""
        if self.registration_window and self.registration_window.winfo_exists():
            self.registration_window.lift()
            return

        self.registration_window = tk.Toplevel(self.root)
        self.registration_window.title("MineralXAfrica - User Registration")
        self.registration_window.geometry("700x800")
        self.registration_window.configure(bg=COLORS["background"])
        self.registration_window.resizable(False, False)
        self.registration_window.transient(self.root)
        self.registration_window.grab_set()

        # Center the registration window
        self.registration_window.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.registration_window.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.registration_window.winfo_height() // 2)
        self.registration_window.geometry(f"+{x}+{y}")

        main_frame = tk.Frame(self.registration_window, bg=COLORS["primary"], padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)

        # Header
        header = tk.Frame(main_frame, bg=COLORS["primary"])
        header.pack(fill="x", pady=(0, 20))
        tk.Label(header, text="📝 Complete User Registration", font=("Arial", 18, "bold"),
                fg=COLORS["text_light"], bg=COLORS["primary"]).pack()
        tk.Label(header, text="Fill in all required fields to create your account", font=("Arial", 10),
                fg=COLORS["text_muted"], bg=COLORS["primary"]).pack()

        # Create scrollable form
        canvas = tk.Canvas(main_frame, bg=COLORS["primary"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=COLORS["primary"])

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Registration form fields
        self.setup_registration_form(scrollable_frame)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Make mouse wheel scroll work
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def setup_registration_form(self, parent):
        """Setup the complete registration form in the separate window"""
        form_frame = tk.Frame(parent, bg=COLORS["primary"], padx=10, pady=10)
        form_frame.pack(fill="both", expand=True)

        fields = [
            ("Username:", "reg_username_entry", "Unique identifier for login", {}),
            ("Password:", "reg_password_entry", "At least 8 chars, 1 upper, 1 lower, 1 number", {"show": "•"}),
            ("Confirm Password:", "reg_confirm_password_entry", "Re-enter your password", {"show": "•"}),
            ("Email:", "reg_email_entry", "Your email address", {}),
            ("Role:", "reg_role_combo", "Investor or Researcher", {"values": ["Investor", "Researcher"], "state": "readonly"}),
            ("First Name:", "reg_name_entry", "Your legal first name", {}),
            ("Surname:", "reg_surname_entry", "Your legal surname", {}),
            ("Age:", "reg_age_entry", "Must be 18–120", {}),
            ("Gender:", "reg_gender_combo", "MALE, FEMALE, OTHER", {"values": ["MALE", "FEMALE", "OTHER"], "state": "readonly"}),
            ("ID Number:", "reg_id_entry", "South African ID (13 digits, valid format)", {}),
            ("Nationality:", "reg_nationality_combo", "Select your country", {"values": list(COUNTRIES.keys()), "state": "readonly"}),
            ("Institution:", "reg_institution_combo", "Select your institution", {"values": ALLOWED_INSTITUTIONS, "state": "readonly"})
        ]

        # Check admin restrictions - regular users can't create admin accounts
        users = load_users()
        admin_exists = any(user_data.get("role", "").lower() == "admin" for user_data in users.values())
        
        for i, (label_text, attr_name, help_text, combo_args) in enumerate(fields):            
            row_frame = tk.Frame(form_frame, bg=COLORS["primary"])
            row_frame.pack(fill="x", pady=8)
            
            tk.Label(row_frame, text=label_text, font=("Arial", 10, "bold"),
                    fg=COLORS["text_light"], bg=COLORS["primary"], width=15, anchor="w").pack(side="left", padx=(0, 10))
            
            if "combo" in attr_name:
                combo = ttk.Combobox(row_frame, font=("Arial", 10), **combo_args)
                if attr_name == "reg_role_combo":
                    combo.set("Investor")  # Default to Investor
                elif attr_name == "reg_gender_combo":
                    combo.set("MALE")
                elif attr_name == "reg_nationality_combo":
                    combo.set("South Africa")
                elif attr_name == "reg_institution_combo":
                    combo.set("Wits")
                combo.configure(width=25)
                setattr(self, attr_name, combo)
                combo.pack(side="left", fill="x", expand=True)
            else:
                entry = self.create_styled_entry(row_frame, width=30, **combo_args)
                setattr(self, attr_name, entry)
                entry.pack(side="left", fill="x", expand=True)
            
            tk.Label(row_frame, text=help_text, font=("Arial", 8),
                    fg=COLORS["text_muted"], bg=COLORS["primary"]).pack(side="left", padx=(10, 0))

        # Password strength label
        self.reg_strength_label = tk.Label(form_frame, text="Password Strength: None", font=("Arial", 8),
                                          fg=COLORS["text_muted"], bg=COLORS["primary"])
        self.reg_strength_label.pack(pady=(10, 0))

        # Bind password strength check
        if hasattr(self, 'reg_password_entry'):
            self.reg_password_entry.bind('<KeyRelease>', self.check_reg_password_strength)

        # Register button
        register_btn = tk.Button(form_frame, text="Create Account", font=("Arial", 12, "bold"),
                                bg=COLORS["success"], fg=COLORS["text_light"],
                                activebackground=COLORS["accent"], activeforeground=COLORS["text_light"],
                                relief="flat", padx=20, pady=10, cursor="hand2", 
                                command=self.process_registration)
        register_btn.pack(pady=20)

        # Close button
        close_btn = tk.Button(form_frame, text="Close", font=("Arial", 10),
                             bg=COLORS["error"], fg=COLORS["text_light"],
                             activebackground=COLORS["accent"], activeforeground=COLORS["text_light"],
                             relief="flat", padx=15, pady=5, cursor="hand2",
                             command=lambda: self.registration_window.destroy())
        close_btn.pack(pady=(0, 10))

    def check_reg_password_strength(self, event=None):
        """Check password strength for registration form"""
        password = self.reg_password_entry.get()
        if not password:
            self.reg_strength_label.config(text="Password Strength: None", fg=COLORS["text_muted"])
            return
        
        strength = 0
        if len(password) >= 8: strength += 1
        if re.search(r"[A-Z]", password): strength += 1
        if re.search(r"[a-z]", password): strength += 1
        if re.search(r"[0-9]", password): strength += 1
        if re.search(r"[!@#$%^&*(),.?\":{}|<>]", password): strength += 1
        
        strength_levels = ["Very Weak", "Weak", "Fair", "Good", "Strong", "Very Strong"]
        colors = [COLORS["error"], COLORS["error"], COLORS["warning"], COLORS["warning"], COLORS["success"], COLORS["success"]]
        
        strength_text = f"Password Strength: {strength_levels[min(strength, 5)]}"
        self.reg_strength_label.config(text=strength_text, fg=colors[min(strength, 5)])

    def process_registration(self):
        """Process the registration form submission"""
        # Get all field values
        username = self.reg_username_entry.get().strip()
        password = self.reg_password_entry.get().strip()
        confirm_password = self.reg_confirm_password_entry.get().strip()
        email = self.reg_email_entry.get().strip()
        role = self.reg_role_combo.get().strip()
        name = self.reg_name_entry.get().strip()
        surname = self.reg_surname_entry.get().strip()
        age = self.reg_age_entry.get().strip()
        institution = self.reg_institution_combo.get().strip()
        id_number = self.reg_id_entry.get().strip()
        nationality = self.reg_nationality_combo.get().strip()
        gender = self.reg_gender_combo.get().strip()

        # Validation
        if not all([username, password, confirm_password, email, name, surname, age, id_number]):
            messagebox.showerror("Error", "All fields are required!")
            return

        if password != confirm_password:
            messagebox.showerror("Error", "Passwords do not match!")
            return

        if len(password) < 8:
            messagebox.showerror("Error", "Password must be at least 8 characters long!")
            return

        # Show loading
        loading_window = self.show_loading("Creating your account...")
        self.registration_window.update()

        # Register user (no current_admin for self-registration)
        success, message = register_user(username, password, role, name, surname, age, institution, id_number, nationality, gender, email, None)
        
        loading_window.destroy()

        if success:
            messagebox.showinfo("Success", message)
            self.registration_window.destroy()
            # Clear registration fields for next time
            self.clear_registration_fields()
        else:
            messagebox.showerror("Error", message)

    def clear_registration_fields(self):
        """Clear all registration fields"""
        fields = ['reg_username_entry', 'reg_password_entry', 'reg_confirm_password_entry', 
                 'reg_email_entry', 'reg_name_entry', 'reg_surname_entry', 'reg_age_entry', 'reg_id_entry']
        
        for field in fields:
            if hasattr(self, field):
                getattr(self, field).delete(0, tk.END)

# ----------------------------
# Dashboard Application
# ----------------------------
class DashboardApp:
    def __init__(self, root, username, role):
        self.root = root
        self.username = username
        self.role = role
        self.root.title(f"MineralXAfrica Dashboard - {username} ({role})")
        self.root.geometry("1400x900")
        self.root.configure(bg=COLORS["background"])
        users = load_users()
        user_info = users.get(username, {})
        self.full_name = f"{user_info.get('name', '')} {user_info.get('surname', '')}".strip()
        self.create_menu_bar()
        tab_control = ttk.Notebook(root)
        self.tab_dashboard = ttk.Frame(tab_control)
        self.tab_analytics = ttk.Frame(tab_control)
        self.tab_interactive = ttk.Frame(tab_control)
        self.tab_africa_map = ttk.Frame(tab_control)
        self.tab_data = ttk.Frame(tab_control)
        self.tab_profile = ttk.Frame(tab_control)
        self.tab_market = ttk.Frame(tab_control)
        self.tab_admin = ttk.Frame(tab_control) if role.lower() == "admin" else None
        
        tab_control.add(self.tab_dashboard, text="📊 Dashboard")
        tab_control.add(self.tab_analytics, text="📈 Analytics")
        tab_control.add(self.tab_interactive, text="🔍 Explorer")
        tab_control.add(self.tab_africa_map, text="🗺️ Africa Map")
        tab_control.add(self.tab_data, text="💎 Data")
        tab_control.add(self.tab_profile, text="👤 Profile")
        tab_control.add(self.tab_market, text="🌍 Market Intelligence")
        if self.tab_admin:
            tab_control.add(self.tab_admin, text="⚙️ Admin Panel")
        
        tab_control.pack(expand=1, fill="both")
        self.create_dashboard_tab()
        self.create_analytics_tab()
        self.create_interactive_tab()
        self.create_africa_map_tab()
        self.create_data_tab()
        self.create_profile_tab()
        self.create_market_tab()
        if self.tab_admin:
            self.create_admin_tab()

    def create_menu_bar(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Export Data", command=self.export_data)
        file_menu.add_separator()
        file_menu.add_command(label="Logout", command=self.logout)
        menubar.add_cascade(label="File", menu=file_menu)
        
        if self.role.lower() == "admin":
            admin_menu = tk.Menu(menubar, tearoff=0)
            admin_menu.add_command(label="User Management", command=self.open_user_management)
            admin_menu.add_command(label="Admin Tools", command=self.open_admin_tools)
            menubar.add_cascade(label="Admin", menu=admin_menu)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        self.root.config(menu=menubar)

    def create_dashboard_tab(self):
        main_frame = tk.Frame(self.tab_dashboard, bg=COLORS["background"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        welcome_name = self.full_name if self.full_name else self.username
        tk.Label(main_frame, text=f"Welcome, {welcome_name}!", font=("Arial", 18, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["background"]).pack(anchor="w", pady=(0, 5))
        tk.Label(main_frame, text="African Critical Minerals Dashboard", font=("Arial", 14),
                 fg=COLORS["text_muted"], bg=COLORS["background"]).pack(anchor="w", pady=(0, 20))
        stats_frame = tk.Frame(main_frame, bg=COLORS["background"])
        stats_frame.pack(fill="x", pady=10)
        if flat_critical_africa.empty:
            tk.Label(stats_frame, text="No mineral data available", font=("Arial", 12), fg="red").pack(pady=20)
            return
        total_production = flat_critical_africa["Production"].sum()
        country_count = flat_critical_africa["Country"].nunique()
        mineral_count = flat_critical_africa["Mineral"].nunique()
        latest_year = flat_critical_africa["Year"].max()
        stats_data = [
            ("Total Production", f"{total_production:,.0f} tons", COLORS["crystal_green"]),
            ("Countries", f"{country_count}", COLORS["crystal_blue"]),
            ("Minerals Tracked", f"{mineral_count}", COLORS["crystal_red"]),
            ("Latest Data", f"{latest_year}", COLORS["gold_accent"])
        ]
        for i, (title, value, color) in enumerate(stats_data):
            card = tk.Frame(stats_frame, relief='raised', borderwidth=2, bg=color)
            card.grid(row=0, column=i, padx=10, pady=10, sticky='nsew')
            tk.Label(card, text=title, bg=color, fg='white', font=("Arial", 10, "bold")).pack(pady=(15, 5))
            tk.Label(card, text=value, bg=color, fg='white', font=("Arial", 14, "bold")).pack(pady=(0, 15))
        for i in range(len(stats_data)):
            stats_frame.columnconfigure(i, weight=1)

    def create_analytics_tab(self):
        main_frame = tk.Frame(self.tab_analytics, bg=COLORS["background"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        tk.Label(main_frame, text="Mineral Production Analytics", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        if flat_critical_africa.empty:
            tk.Label(main_frame, text="No data available for analytics").pack()
            return
        control_frame = tk.Frame(main_frame, bg=COLORS["background"])
        control_frame.pack(fill="x", pady=10)
        tk.Label(control_frame, text="Chart Type:", bg=COLORS["background"], fg=COLORS["text_light"]).pack(side="left", padx=5)
        chart_var = tk.StringVar(value="Production by Mineral")
        chart_combo = ttk.Combobox(control_frame, textvariable=chart_var,
                                   values=["Production by Mineral", "Production by Country",
                                           "Top Minerals", "Country Comparison", "Yearly Trends"],
                                   state="readonly", width=20)
        chart_combo.pack(side="left", padx=5)
        chart_frame = tk.Frame(main_frame, bg=COLORS["background"])
        chart_frame.pack(fill="both", expand=True)

        def update_chart(*args):
            for widget in chart_frame.winfo_children():
                widget.destroy()
            fig, ax = plt.subplots(figsize=(12, 8), facecolor=COLORS["background"])
            chart_type = chart_var.get()
            if chart_type == "Production by Mineral":
                data = flat_critical_africa.groupby("Mineral")["Production"].sum().sort_values(ascending=False)
                bars = ax.bar(data.index, data.values, color='skyblue', alpha=0.7)
                ax.set_title("Total Production by Mineral", fontsize=14, fontweight='bold', color=COLORS["text_light"])
                ax.set_ylabel("Production (tons)", fontsize=12, color=COLORS["text_light"])
                plt.xticks(rotation=45, ha='right', color=COLORS["text_light"])
                for bar in bars:
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height, f'{height:,.0f}', ha='center', va='bottom', fontsize=9, color=COLORS["text_light"])
            elif chart_type == "Production by Country":
                data = flat_critical_africa.groupby("Country")["Production"].sum().sort_values(ascending=False).head(15)
                bars = ax.bar(data.index, data.values, color='lightgreen', alpha=0.7)
                ax.set_title("Top 15 Countries by Production", fontsize=14, fontweight='bold', color=COLORS["text_light"])
                ax.set_ylabel("Production (tons)", fontsize=12, color=COLORS["text_light"])
                plt.xticks(rotation=45, ha='right', color=COLORS["text_light"])
                for bar in bars:
                    height = bar.get_height()
                    ax.text(bar.get_x() + bar.get_width()/2., height, f'{height:,.0f}', ha='center', va='bottom', fontsize=9, color=COLORS["text_light"])
            elif chart_type == "Top Minerals":
                data = flat_critical_africa.groupby("Mineral")["Production"].sum().nlargest(8)
                ax.pie(data.values, labels=data.index, autopct='%1.1f%%', startangle=90, colors=plt.cm.Set3(range(len(data))))
                ax.set_title("Top 8 Minerals Production Distribution", fontsize=14, fontweight='bold', color=COLORS["text_light"])
            elif chart_type == "Country Comparison":
                top_countries = flat_critical_africa.groupby("Country")["Production"].sum().nlargest(5).index
                filtered_data = flat_critical_africa[flat_critical_africa["Country"].isin(top_countries)]
                pivot_data = filtered_data.pivot_table(index="Country", columns="Mineral", values="Production", aggfunc="sum")
                pivot_data.plot(kind="bar", ax=ax, stacked=True, color=plt.cm.Set2(range(len(pivot_data.columns))))
                ax.set_title("Mineral Production by Top Countries", fontsize=14, fontweight='bold', color=COLORS["text_light"])
                ax.set_ylabel("Production (tons)", fontsize=12, color=COLORS["text_light"])
                ax.legend(title="Minerals", bbox_to_anchor=(1.05, 1), loc='upper left', facecolor=COLORS["card_bg"], edgecolor=COLORS["border"], fontsize=10)
            elif chart_type == "Yearly Trends":
                minerals = sorted(flat_critical_africa["Mineral"].unique())
                if minerals:
                    mineral_frame = tk.Frame(chart_frame, bg=COLORS["background"])
                    mineral_frame.pack(pady=5)
                    tk.Label(mineral_frame, text="Select Mineral:", font=("Arial", 10), fg=COLORS["text_light"], bg=COLORS["background"]).pack(side="left")
                    mineral_var = tk.StringVar(value=minerals[0])
                    mineral_combo = ttk.Combobox(mineral_frame, textvariable=mineral_var, values=minerals, state="readonly", width=20)
                    mineral_combo.pack(side="left", padx=5)

                    def plot_yearly_trend():
                        for widget in chart_frame.winfo_children():
                            if widget != mineral_frame:
                                widget.destroy()
                        fig, ax = plt.subplots(figsize=(12, 6), facecolor=COLORS["background"])
                        selected_mineral = mineral_var.get()
                        mineral_data = flat_critical_africa[flat_critical_africa["Mineral"] == selected_mineral]
                        yearly_data = mineral_data.groupby("Year")["Production"].sum()
                        if not yearly_data.empty:
                            ax.plot(yearly_data.index, yearly_data.values, marker='o', linewidth=2, markersize=8, color='purple')
                            ax.set_title(f"Yearly Production Trend: {selected_mineral}", fontsize=14, fontweight='bold', color=COLORS["text_light"])
                            ax.set_ylabel("Production (tons)", fontsize=12, color=COLORS["text_light"])
                            ax.set_xlabel("Year", fontsize=12, color=COLORS["text_light"])
                            ax.grid(True, alpha=0.3, color=COLORS["text_muted"])
                            for x, y in zip(yearly_data.index, yearly_data.values):
                                ax.text(x, y, f'{int(y):,}', ha='center', va='bottom', fontsize=9, color=COLORS["text_light"])
                        else:
                            ax.text(0.5, 0.5, f"No data for {selected_mineral}", transform=ax.transAxes, ha='center', color=COLORS["text_light"])
                        plt.tight_layout()
                        canvas = FigureCanvasTkAgg(fig, master=chart_frame)
                        canvas.draw()
                        canvas.get_tk_widget().pack(fill="both", expand=True)
                    mineral_combo.bind('<<ComboboxSelected>>', lambda e: plot_yearly_trend())
                    plot_yearly_trend()
                else:
                    ax.text(0.5, 0.5, "No mineral data available", transform=ax.transAxes, ha='center', color=COLORS["text_light"])
            plt.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
        chart_combo.bind('<<ComboboxSelected>>', update_chart)
        update_chart()

    def create_interactive_tab(self):
        main_frame = tk.Frame(self.tab_interactive, bg=COLORS["background"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        tk.Label(main_frame, text="Interactive Mineral Explorer", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        if flat_critical_africa.empty:
            tk.Label(main_frame, text="No data available for interactive exploration").pack()
            return
        controls_frame = tk.Frame(main_frame, bg=COLORS["background"])
        controls_frame.pack(fill="x", pady=10)
        tk.Label(controls_frame, text="Mineral:", bg=COLORS["background"], fg=COLORS["text_light"]).grid(row=0, column=0, padx=5, sticky="w")
        mineral_var = tk.StringVar()
        minerals = sorted(flat_critical_africa["Mineral"].unique())
        mineral_combo = ttk.Combobox(controls_frame, textvariable=mineral_var, values=minerals, state="readonly", width=20)
        mineral_combo.set(minerals[0] if minerals else "")
        mineral_combo.grid(row=0, column=1, padx=5)
        tk.Label(controls_frame, text="Year:", bg=COLORS["background"], fg=COLORS["text_light"]).grid(row=0, column=2, padx=5, sticky="w")
        year_var = tk.StringVar()
        years = sorted(flat_critical_africa["Year"].unique())
        year_combo = ttk.Combobox(controls_frame, textvariable=year_var, values=years, state="readonly", width=10)
        year_combo.set(years[-1] if years else "")
        year_combo.grid(row=0, column=3, padx=5)
        chart_frame = tk.Frame(main_frame, bg=COLORS["background"])
        chart_frame.pack(fill="both", expand=True)

        def update_interactive_chart(*args):
            mineral = mineral_var.get()
            year = int(year_var.get()) if year_var.get() else None
            if not mineral or not year:
                return
            for widget in chart_frame.winfo_children():
                widget.destroy()
            df_filtered = flat_critical_africa[(flat_critical_africa["Mineral"] == mineral) &
                                              (flat_critical_africa["Year"] == year)].sort_values("Production", ascending=False)
            if df_filtered.empty:
                tk.Label(chart_frame, text=f"No data for {mineral} in {year}", fg="red", bg=COLORS["background"]).pack()
                return
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6), facecolor=COLORS["background"])
            bars = ax1.bar(df_filtered["Country"], df_filtered["Production"], color='teal', alpha=0.7)
            ax1.set_title(f"{mineral} Production by Country ({year})", fontsize=14, fontweight='bold', color=COLORS["text_light"])
            ax1.set_ylabel("Production (tons)", fontsize=12, color=COLORS["text_light"])
            ax1.tick_params(axis='x', rotation=45, colors=COLORS["text_light"])
            for bar in bars:
                height = bar.get_height()
                ax1.text(bar.get_x() + bar.get_width()/2., height, f'{height:,.0f}', ha='center', va='bottom', fontsize=9, color=COLORS["text_light"])
            top_countries = df_filtered.head(8)
            ax2.pie(top_countries["Production"], labels=top_countries["Country"], autopct='%1.1f%%', startangle=90, colors=plt.cm.Set3(range(len(top_countries))))
            ax2.set_title(f"Top 8 Countries - {mineral} ({year})", fontsize=14, fontweight='bold', color=COLORS["text_light"])
            plt.tight_layout()
            canvas = FigureCanvasTkAgg(fig, master=chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            total = df_filtered["Production"].sum()
            avg = df_filtered["Production"].mean()
            max_country = df_filtered.iloc[0]["Country"]
            max_value = df_filtered.iloc[0]["Production"]
            stats_text = f"""Summary for {mineral} ({year}):
• Total Production: {total:,.0f} tons
• Average per Country: {avg:,.0f} tons
• Top Producer: {max_country} ({max_value:,.0f} tons)
• Countries with Data: {len(df_filtered)}"""
            stats_label = tk.Label(chart_frame, text=stats_text, font=("Arial", 10), justify="left",
                                   bg="#f0f0f0", padx=10, pady=10)
            stats_label.pack(fill="x", pady=10)
        mineral_combo.bind('<<ComboboxSelected>>', update_interactive_chart)
        year_combo.bind('<<ComboboxSelected>>', update_interactive_chart)
        if minerals and years:
            update_interactive_chart()

    def create_africa_map_tab(self):
        main_frame = tk.Frame(self.tab_africa_map, bg=COLORS["background"])
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        tk.Label(main_frame, text="🗺️ African Mineral Production Map", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=5)
        if flat_critical_africa.empty:
            tk.Label(main_frame, text="No data available for map visualization").pack()
            return
        controls_frame = tk.Frame(main_frame, bg=COLORS["background"])
        controls_frame.pack(fill="x", pady=5)
        tk.Label(controls_frame, text="Mineral:", bg=COLORS["background"], fg=COLORS["text_light"]).pack(side="left", padx=5)
        mineral_var = tk.StringVar()
        minerals = sorted(flat_critical_africa["Mineral"].unique())
        mineral_combo = ttk.Combobox(controls_frame, textvariable=mineral_var, values=minerals, state="readonly", width=20)
        mineral_combo.set(minerals[0] if minerals else "")
        mineral_combo.pack(side="left", padx=5)
        tk.Label(controls_frame, text="Year:", bg=COLORS["background"], fg=COLORS["text_light"]).pack(side="left", padx=5)
        year_var = tk.StringVar()
        years = sorted(flat_critical_africa["Year"].unique())
        year_combo = ttk.Combobox(controls_frame, textvariable=year_var, values=years, state="readonly", width=10)
        year_combo.set(years[-1] if years else "")
        year_combo.pack(side="left", padx=5)

        # Map style selector
        tk.Label(controls_frame, text="Map Style:", bg=COLORS["background"], fg=COLORS["text_light"]).pack(side="left", padx=(20, 5))
        map_style_var = tk.StringVar(value="CartoDB Positron")
        map_style_combo = ttk.Combobox(controls_frame, textvariable=map_style_var,
                                      values=["CartoDB Positron", "OpenStreetMap", "Stamen Toner"],
                                      state="readonly", width=18)
        map_style_combo.pack(side="left", padx=5)

        map_frame = tk.LabelFrame(main_frame, text="Interactive Map of Africa", padx=5, pady=5, bg=COLORS["background"], fg=COLORS["text_light"])
        map_frame.pack(fill="both", expand=True, pady=10)
        map_widget = tkintermapview.TkinterMapView(map_frame, width=800, height=600, corner_radius=0)
        map_widget.pack(fill="both", expand=True)
        map_widget.set_position(-8.7832, 34.5085)
        map_widget.set_zoom(4)

        def set_map_style(style):
            if style == "CartoDB Positron":
                map_widget.set_tile_server("https://a.basemaps.cartocdn.com/light_all/{z}/{x}/{y}.png")
            elif style == "Stamen Toner":
                map_widget.set_tile_server("https://stamen-tiles.a.ssl.fastly.net/toner/{z}/{x}/{y}.png")
            else:
                map_widget.set_tile_server("https://a.tile.openstreetmap.org/{z}/{x}/{y}.png")

        set_map_style("CartoDB Positron")

        country_coords = {
            "Algeria": (28.0339, 2.8123), "Angola": (-11.2027, 17.8739), "Benin": (9.3077, 2.3158),
            "Botswana": (-22.3285, 24.6849), "Burkina Faso": (12.2383, -1.5616), "Burundi": (-3.3731, 29.9189),
            "Cameroon": (7.3697, 12.3547), "Cape Verde": (16.0021, -24.0132), "Central African Republic": (6.6111, 20.9394),
            "Chad": (15.4542, 18.7322), "Comoros": (-11.6455, 43.3333), "Congo": (-0.2280, 15.8277),
            "DR Congo": (-4.0383, 21.7587), "Djibouti": (11.8251, 42.5903), "Egypt": (26.8206, 30.8025),
            "Equatorial Guinea": (1.6508, 10.2679), "Eritrea": (15.1794, 39.7823), "Eswatini": (-26.5225, 31.4659),
            "Ethiopia": (9.1450, 40.4897), "Gabon": (-0.8037, 11.6094), "Gambia": (13.4432, -15.3101),
            "Ghana": (7.9465, -1.0232), "Guinea": (9.9456, -9.6966), "Guinea-Bissau": (11.8037, -15.1804),
            "Ivory Coast": (7.5400, -5.5471), "Kenya": (-0.0236, 37.9062), "Lesotho": (-29.6100, 28.2336),
            "Liberia": (6.4281, -9.4295), "Libya": (26.3351, 17.2283), "Madagascar": (-18.7669, 46.8691),
            "Malawi": (-13.2543, 34.3015), "Mali": (17.5707, -3.9962), "Mauritania": (21.0079, -10.9408),
            "Mauritius": (-20.3484, 57.5522), "Morocco": (31.7917, -7.0926), "Mozambique": (-18.6657, 35.5296),
            "Namibia": (-22.9576, 18.4904), "Niger": (17.6078, 8.0817), "Nigeria": (9.0820, 8.6753),
            "Rwanda": (-1.9403, 29.8739), "Sao Tome and Principe": (0.1864, 6.6131), "Senegal": (14.4974, -14.4524),
            "Seychelles": (-4.6796, 55.4920), "Sierra Leone": (8.4607, -11.7799), "Somalia": (5.1521, 46.1996),
            "South Africa": (-30.5595, 22.9375), "South Sudan": (6.8770, 31.3070), "Sudan": (12.8628, 30.2176),
            "Tanzania": (-6.3690, 34.8888), "Togo": (8.6195, 0.8248), "Tunisia": (33.8869, 9.5375),
            "Uganda": (1.3733, 32.2903), "Zambia": (-13.1339, 27.8493), "Zimbabwe": (-19.0154, 29.1549)
        }
        markers = []

        def update_map_style(*args):
            set_map_style(map_style_var.get())
            update_map()

        def update_map(*args):
            for marker in markers:
                marker.delete()
            markers.clear()
            mineral = mineral_var.get()
            year = int(year_var.get()) if year_var.get() else None
            if not mineral or not year:
                return
            df_filtered = flat_critical_africa[(flat_critical_africa["Mineral"] == mineral) &
                                              (flat_critical_africa["Year"] == year)].copy()
            if df_filtered.empty:
                return
            country_production = df_filtered.groupby("Country")["Production"].sum()
            for country, production in country_production.items():
                if country in country_coords:
                    lat, lon = country_coords[country]
                    marker = map_widget.set_marker(lat, lon,
                                                   text=f"{country}\n{production:,.0f} tons",
                                                   marker_color_circle="red",
                                                   marker_color_outside="darkred")
                    markers.append(marker)

        mineral_combo.bind('<<ComboboxSelected>>', update_map)
        year_combo.bind('<<ComboboxSelected>>', update_map)
        map_style_combo.bind('<<ComboboxSelected>>', update_map_style)
        if minerals and years:
            update_map()

    def create_data_tab(self):
        main_frame = tk.Frame(self.tab_data, bg=COLORS["background"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        tk.Label(main_frame, text="Mineral Data Explorer", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        if flat_critical_africa.empty:
            tk.Label(main_frame, text="No data available").pack()
            return
        filter_frame = tk.Frame(main_frame, bg=COLORS["background"])
        filter_frame.pack(fill="x", pady=10)
        tk.Label(filter_frame, text="Country:", bg=COLORS["background"], fg=COLORS["text_light"]).grid(row=0, column=0, padx=5)
        country_var = tk.StringVar(value="All")
        countries = ["All"] + sorted(flat_critical_africa["Country"].unique())
        country_combo = ttk.Combobox(filter_frame, textvariable=country_var, values=countries, state="readonly")
        country_combo.grid(row=0, column=1, padx=5)
        tk.Label(filter_frame, text="Mineral:", bg=COLORS["background"], fg=COLORS["text_light"]).grid(row=0, column=2, padx=5)
        mineral_var = tk.StringVar(value="All")
        minerals = ["All"] + sorted(flat_critical_africa["Mineral"].unique())
        mineral_combo = ttk.Combobox(filter_frame, textvariable=mineral_var, values=minerals, state="readonly")
        mineral_combo.grid(row=0, column=3, padx=5)
        tk.Label(filter_frame, text="Year:", bg=COLORS["background"], fg=COLORS["text_light"]).grid(row=0, column=4, padx=5)
        year_var = tk.StringVar(value="All")
        years = ["All"] + sorted(flat_critical_africa["Year"].unique())
        year_combo = ttk.Combobox(filter_frame, textvariable=year_var, values=years, state="readonly")
        year_combo.grid(row=0, column=5, padx=5)
        table_frame = tk.Frame(main_frame, bg=COLORS["background"])
        table_frame.pack(fill="both", expand=True)
        columns = ("Country", "Mineral", "Year", "Production")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150)
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def update_table(*args):
            for item in tree.get_children():
                tree.delete(item)
            filtered_data = flat_critical_africa.copy()
            country = country_var.get()
            mineral = mineral_var.get()
            year = year_var.get()
            if country != "All":
                filtered_data = filtered_data[filtered_data["Country"] == country]
            if mineral != "All":
                filtered_data = filtered_data[filtered_data["Mineral"] == mineral]
            if year != "All":
                filtered_data = filtered_data[filtered_data["Year"] == int(year)]
            for _, row in filtered_data.iterrows():
                tree.insert("", "end", values=(row["Country"], row["Mineral"], row["Year"], f"{row['Production']:,.0f}"))
        country_combo.bind('<<ComboboxSelected>>', update_table)
        mineral_combo.bind('<<ComboboxSelected>>', update_table)
        year_combo.bind('<<ComboboxSelected>>', update_table)
        update_table()

    def create_profile_tab(self):
        main_frame = tk.Frame(self.tab_profile, bg=COLORS["background"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        tk.Label(main_frame, text="User Profile", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        users = load_users()
        user_info = users.get(self.username, {})
        if not user_info:
            tk.Label(main_frame, text="No profile information available").pack()
            return
        profile_frame = tk.LabelFrame(main_frame, text="Personal Information", font=("Arial", 12, "bold"),
                                      padx=20, pady=20, bg=COLORS["card_bg"], fg=COLORS["text_light"])
        profile_frame.pack(fill="x", padx=20, pady=10)
        details = [
            ("Username:", self.username),
            ("Email:", user_info.get("email", "")),
            ("Role:", user_info.get("role", "")),
            ("Full Name:", f"{user_info.get('name', '')} {user_info.get('surname', '')}"),
            ("Age:", user_info.get("age", "")),
            ("Gender:", user_info.get("gender", "")),
            ("Nationality:", user_info.get("nationality", "")),
            ("Institution:", user_info.get("institution", "")),
            ("ID Number:", user_info.get("id_number", "")),
            ("Date Registered:", user_info.get("date_created", "").split('.')[0]),
            ("Status:", user_info.get("status", "active")),
            ("Primary Admin:", "Yes" if user_info.get("is_primary") else "No")
        ]
        for i, (label, value) in enumerate(details):
            tk.Label(profile_frame, text=label, font=("Arial", 11, "bold"),
                     fg=COLORS["text_light"], bg=COLORS["card_bg"]).grid(row=i, column=0, sticky="w", padx=10, pady=5)
            tk.Label(profile_frame, text=value, font=("Arial", 11),
                     fg=COLORS["text_light"], bg=COLORS["card_bg"]).grid(row=i, column=1, sticky="w", padx=10, pady=5)

    def create_market_tab(self):
        main_frame = tk.Frame(self.tab_market, bg=COLORS["background"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        tk.Label(main_frame, text="🌍 Market Intelligence Dashboard", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        if CMO_MINERAL_NAMES:
            control_frame = tk.Frame(main_frame, bg=COLORS["background"])
            control_frame.pack(fill="x", pady=(0, 15))
            tk.Label(control_frame, text="Commodity:", bg=COLORS["background"], fg=COLORS["text_light"]).pack(side="left", padx=(0, 5))
            self.market_mineral_var = tk.StringVar(value=CMO_MINERAL_NAMES[0])
            mineral_combo = ttk.Combobox(control_frame, textvariable=self.market_mineral_var,
                                         values=CMO_MINERAL_NAMES, state="readonly", width=20)
            mineral_combo.pack(side="left", padx=5)
            tk.Label(control_frame, text="Frequency:", bg=COLORS["background"], fg=COLORS["text_light"]).pack(side="left", padx=(15, 5))
            self.market_freq_var = tk.StringVar(value="Annual")
            freq_combo = ttk.Combobox(control_frame, textvariable=self.market_freq_var,
                                      values=["Annual", "Monthly"], state="readonly", width=10)
            freq_combo.pack(side="left", padx=5)
            self.market_chart_frame = tk.Frame(main_frame, bg=COLORS["background"])
            self.market_chart_frame.pack(fill="both", expand=True)
            self.market_mineral_var.trace("w", lambda *args: self.plot_market_data())
            self.market_freq_var.trace("w", lambda *args: self.plot_market_data())
            self.plot_market_data()
        else:
            tk.Label(main_frame, text="No commodity price data available.\nPlease ensure CMO Excel files are in the working directory.",
                     fg="red", bg=COLORS["background"], font=("Arial", 12), justify="center").pack(expand=True)

    def create_admin_tab(self):
        """Create the admin panel tab with advanced user management"""
        main_frame = tk.Frame(self.tab_admin, bg=COLORS["background"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        tk.Label(main_frame, text="⚙️ Administrator Panel", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        
        # Admin statistics
        users = load_users()
        total_users = len(users)
        admin_users = sum(1 for u in users.values() if u.get("role", "").lower() == "admin")
        active_users = sum(1 for u in users.values() if u.get("status", "active") == "active")
        suspended_users = total_users - active_users
        
        stats_frame = tk.Frame(main_frame, bg=COLORS["background"])
        stats_frame.pack(fill="x", pady=10)
        
        stats_data = [
            ("Total Users", f"{total_users}", COLORS["crystal_blue"]),
            ("Admin Users", f"{admin_users}", COLORS["gold_accent"]),
            ("Active Users", f"{active_users}", COLORS["crystal_green"]),
            ("Suspended Users", f"{suspended_users}", COLORS["crystal_red"])
        ]
        
        for i, (title, value, color) in enumerate(stats_data):
            card = tk.Frame(stats_frame, relief='raised', borderwidth=2, bg=color)
            card.grid(row=0, column=i, padx=10, pady=10, sticky='nsew')
            tk.Label(card, text=title, bg=color, fg='white', font=("Arial", 10, "bold")).pack(pady=(15, 5))
            tk.Label(card, text=value, bg=color, fg='white', font=("Arial", 14, "bold")).pack(pady=(0, 15))
        
        for i in range(len(stats_data)):
            stats_frame.columnconfigure(i, weight=1)
        
        # Admin actions frame
        actions_frame = tk.LabelFrame(main_frame, text="Administrative Actions", 
                                    font=("Arial", 12, "bold"), padx=15, pady=15,
                                    bg=COLORS["card_bg"], fg=COLORS["text_light"])
        actions_frame.pack(fill="x", pady=20)
        
        btn1 = tk.Button(actions_frame, text="👥 User Management", font=("Arial", 11),
                        bg=COLORS["primary"], fg=COLORS["text_light"],
                        command=self.open_user_management, width=20, height=2)
        btn1.grid(row=0, column=0, padx=10, pady=5)
        
        btn2 = tk.Button(actions_frame, text="🛠️ Admin Tools", font=("Arial", 11),
                        bg=COLORS["secondary"], fg=COLORS["text_light"],
                        command=self.open_admin_tools, width=20, height=2)
        btn2.grid(row=0, column=1, padx=10, pady=5)
        
        btn3 = tk.Button(actions_frame, text="📊 System Stats", font=("Arial", 11),
                        bg=COLORS["accent"], fg=COLORS["text_light"],
                        command=self.show_system_stats, width=20, height=2)
        btn3.grid(row=0, column=2, padx=10, pady=5)
        
        # Primary admin warning
        if is_primary_admin(self.username):
            warning_frame = tk.Frame(main_frame, bg=COLORS["warning"])
            warning_frame.pack(fill="x", pady=10)
            tk.Label(warning_frame, text="⚠️ PRIMARY ADMIN PRIVILEGES: You have full system access", 
                    font=("Arial", 10, "bold"), bg=COLORS["warning"], fg="black").pack(pady=5)

    def plot_market_data(self):
        for widget in self.market_chart_frame.winfo_children():
            widget.destroy()
        mineral = self.market_mineral_var.get()
        freq = self.market_freq_var.get()
        candidate_cols = CMO_MINERALS.get(mineral, [])
        if freq == "Annual":
            df = cmo_annual_df
            date_col = 'Year'
        else:
            df = cmo_monthly_df
            date_col = 'Date'
        col_found = None
        for col in candidate_cols:
            if col in df.columns:
                col_found = col
                break
        if not col_found:
            tk.Label(self.market_chart_frame, text=f"No price data for {mineral}", fg="red", bg=COLORS["background"]).pack()
            return
        plot_data = df[[date_col, col_found]].dropna().copy()
        plot_data = plot_data.sort_values(date_col)
        plot_data = plot_data[plot_data[col_found] > 0]
        if plot_data.empty:
            tk.Label(self.market_chart_frame, text=f"No valid price data for {mineral}", fg="orange", bg=COLORS["background"]).pack()
            return
        fig, ax = plt.subplots(figsize=(12, 6), facecolor=COLORS["background"])
        ax.plot(plot_data[date_col], plot_data[col_found],
                marker='o', markersize=2, linewidth=2, color='#1f77b4', alpha=0.8)
        title = f"{mineral} Price Trend ({freq})"
        ax.set_title(title, fontsize=14, fontweight='bold', color=COLORS["text_light"])
        ax.set_ylabel("Price (USD)", fontsize=12, color=COLORS["text_light"])
        ax.grid(True, alpha=0.3, color=COLORS["text_muted"])
        if freq == "Monthly":
            ax.tick_params(axis='x', rotation=45, colors=COLORS["text_light"])
        plt.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=self.market_chart_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    def export_data(self):
        if flat_critical_africa.empty:
            messagebox.showwarning("Export", "No data to export!")
            return
        filename = f"mineral_data_export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv"
        flat_critical_africa.to_csv(filename, index=False)
        messagebox.showinfo("Export Successful", f"Data exported to {filename}")

    def open_user_management(self):
        if self.role.lower() != "admin":
            messagebox.showwarning("Access Denied", "Admin privileges required!")
            return
        
        win = tk.Toplevel(self.root)
        win.title("User Management - Advanced")
        win.geometry("1200x700")
        win.configure(bg=COLORS["background"])
        
        # Header
        header_frame = tk.Frame(win, bg=COLORS["primary"])
        header_frame.pack(fill="x", pady=10)
        tk.Label(header_frame, text="👥 Advanced User Management", font=("Arial", 16, "bold"),
                 fg=COLORS["text_light"], bg=COLORS["primary"]).pack(pady=5)
        tk.Label(header_frame, text="Manage user accounts, roles, and permissions", font=("Arial", 10),
                 fg=COLORS["text_muted"], bg=COLORS["primary"]).pack(pady=2)
        
        # Control frame
        control_frame = tk.Frame(win, bg=COLORS["background"])
        control_frame.pack(fill="x", pady=10, padx=20)
        
        tk.Button(control_frame, text="Create New User", bg=COLORS["success"], fg="white",
                 command=lambda: self.create_user_dialog(win)).pack(side="left", padx=5)
        tk.Button(control_frame, text="Refresh", bg=COLORS["primary"], fg="white",
                 command=lambda: self.refresh_user_table(tree)).pack(side="left", padx=5)
        
        # User table
        table_frame = tk.Frame(win, bg=COLORS["background"])
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        columns = ("Username", "Email", "Role", "Status", "Primary", "Created", "Last Modified")
        tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=20)
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120)
        
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Populate table
        self.refresh_user_table(tree)
        
        # Action buttons frame
        action_frame = tk.Frame(win, bg=COLORS["background"])
        action_frame.pack(fill="x", pady=10, padx=20)
        
        tk.Button(action_frame, text="Edit User", bg=COLORS["warning"], fg="white",
                 command=lambda: self.edit_user_dialog(tree, win)).pack(side="left", padx=5)
        tk.Button(action_frame, text="Suspend/Activate", bg=COLORS["accent"], fg="white",
                 command=lambda: self.toggle_user_status(tree)).pack(side="left", padx=5)
        tk.Button(action_frame, text="Reset Password", bg=COLORS["secondary"], fg="white",
                 command=lambda: self.reset_password_dialog(tree)).pack(side="left", padx=5)
        tk.Button(action_frame, text="Change Role", bg=COLORS["primary"], fg="white",
                 command=lambda: self.change_role_dialog(tree)).pack(side="left", padx=5)
        tk.Button(action_frame, text="View Details", bg=COLORS["crystal_blue"], fg="white",
                 command=lambda: self.view_user_details(tree)).pack(side="left", padx=5)

    def refresh_user_table(self, tree):
        """Refresh the user table with current data"""
        for item in tree.get_children():
            tree.delete(item)
        
        users = load_users()
        for username, user_data in users.items():
            tree.insert("", "end", values=(
                username,
                user_data.get("email", ""),
                user_data.get("role", ""),
                user_data.get("status", "active"),
                "Yes" if user_data.get("is_primary") else "No",
                user_data.get("date_created", "").split('.')[0],
                user_data.get("modified_date", "Never")
            ))

    def create_user_dialog(self, parent):
        """Dialog for creating new users (admin only)"""
        dialog = tk.Toplevel(parent)
        dialog.title("Create New User")
        dialog.geometry("500x600")
        dialog.configure(bg=COLORS["background"])
        dialog.transient(parent)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (parent.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (parent.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        tk.Label(dialog, text="Create New User Account", font=("Arial", 14, "bold"),
                fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        
        # Form fields
        fields = [
            ("Username:", "entry"),
            ("Password:", "entry", {"show": "•"}),
            ("Email:", "entry"),
            ("Role:", "combo", {"values": ["Investor", "Researcher", "Admin"], "state": "readonly"}),
            ("First Name:", "entry"),
            ("Surname:", "entry"),
            ("Age:", "entry"),
            ("Gender:", "combo", {"values": ["MALE", "FEMALE", "OTHER"], "state": "readonly"}),
            ("ID Number:", "entry"),
            ("Nationality:", "combo", {"values": list(COUNTRIES.keys()), "state": "readonly"}),
            ("Institution:", "combo", {"values": ALLOWED_INSTITUTIONS, "state": "readonly"})
        ]
        
        entries = {}
        for i, (label, field_type, *extra_args) in enumerate(fields):
            frame = tk.Frame(dialog, bg=COLORS["background"])
            frame.pack(fill="x", padx=20, pady=5)
            
            tk.Label(frame, text=label, font=("Arial", 10), 
                    fg=COLORS["text_light"], bg=COLORS["background"], width=12, anchor="w").pack(side="left")
            
            if field_type == "entry":
                kwargs = extra_args[0] if extra_args else {}
                entry = tk.Entry(frame, font=("Arial", 10), width=30, **kwargs)
                entry.pack(side="left", fill="x", expand=True)
                entries[label] = entry
            elif field_type == "combo":
                kwargs = extra_args[0] if extra_args else {}
                combo = ttk.Combobox(frame, font=("Arial", 10), width=28, **kwargs)
                if "Role" in label:
                    combo.set("Investor")
                    # Check admin creation permissions
                    if not can_create_admin(self.username, "admin"):
                        combo['values'] = ["Investor", "Researcher"]
                elif "Gender" in label:
                    combo.set("MALE")
                elif "Nationality" in label:
                    combo.set("South Africa")
                elif "Institution" in label:
                    combo.set("Wits")
                combo.pack(side="left", fill="x", expand=True)
                entries[label] = combo
        
        def create_user():
            # Gather data
            username = entries["Username:"].get().strip()
            password = entries["Password:"].get().strip()
            email = entries["Email:"].get().strip()
            role = entries["Role:"].get().strip()
            name = entries["First Name:"].get().strip()
            surname = entries["Surname:"].get().strip()
            age = entries["Age:"].get().strip()
            gender = entries["Gender:"].get().strip()
            id_number = entries["ID Number:"].get().strip()
            nationality = entries["Nationality:"].get().strip()
            institution = entries["Institution:"].get().strip()
            
            # Validate
            if not all([username, password, email, name, surname, age, id_number]):
                messagebox.showerror("Error", "All fields are required!")
                return
            
            # Create user
            success, message = register_user(username, password, role, name, surname, age, 
                                           institution, id_number, nationality, gender, email, self.username)
            if success:
                messagebox.showinfo("Success", message)
                dialog.destroy()
                if hasattr(self, 'refresh_user_table'):
                    self.refresh_user_table(getattr(self, 'user_tree', None))
            else:
                messagebox.showerror("Error", message)
        
        # Buttons
        btn_frame = tk.Frame(dialog, bg=COLORS["background"])
        btn_frame.pack(fill="x", pady=20)
        
        tk.Button(btn_frame, text="Create User", bg=COLORS["success"], fg="white",
                 command=create_user).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Cancel", bg=COLORS["error"], fg="white",
                 command=dialog.destroy).pack(side="left", padx=10)

    def edit_user_dialog(self, tree, parent):
        """Edit user details"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select a user to edit.")
            return
        
        username = tree.item(selection[0])['values'][0]
        users = load_users()
        user_data = users.get(username)
        
        if not user_data:
            messagebox.showerror("Error", "User not found!")
            return
        
        if user_data.get("is_primary"):
            messagebox.showwarning("Restricted", "Cannot edit primary admin accounts!")
            return
        
        # Create edit dialog (similar to create_user_dialog but with existing data)
        messagebox.showinfo("Edit User", f"Edit functionality for {username} would be implemented here.")

    def toggle_user_status(self, tree):
        """Toggle user status between active and suspended"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select a user.")
            return
        
        username = tree.item(selection[0])['values'][0]
        users = load_users()
        user_data = users.get(username)
        
        if not user_data:
            messagebox.showerror("Error", "User not found!")
            return
        
        if user_data.get("is_primary"):
            messagebox.showwarning("Restricted", "Cannot modify primary admin accounts!")
            return
        
        current_status = user_data.get("status", "active")
        new_status = "suspended" if current_status == "active" else "active"
        
        success, message = update_user_status(username, new_status, self.username)
        if success:
            messagebox.showinfo("Success", message)
            self.refresh_user_table(tree)
        else:
            messagebox.showerror("Error", message)

    def reset_password_dialog(self, tree):
        """Reset user password"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select a user.")
            return
        
        username = tree.item(selection[0])['values'][0]
        users = load_users()
        user_data = users.get(username)
        
        if not user_data:
            messagebox.showerror("Error", "User not found!")
            return
        
        # Simple password reset dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Reset Password")
        dialog.geometry("300x200")
        dialog.configure(bg=COLORS["background"])
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text=f"Reset Password for {username}", font=("Arial", 12, "bold"),
                fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        
        tk.Label(dialog, text="New Password:", fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=5)
        password_entry = tk.Entry(dialog, show="•", font=("Arial", 10))
        password_entry.pack(pady=5)
        
        def reset_password():
            new_password = password_entry.get().strip()
            if len(new_password) < 8:
                messagebox.showerror("Error", "Password must be at least 8 characters!")
                return
            
            success, message = reset_user_password(username, new_password, self.username)
            if success:
                messagebox.showinfo("Success", message)
                dialog.destroy()
            else:
                messagebox.showerror("Error", message)
        
        tk.Button(dialog, text="Reset Password", bg=COLORS["success"], fg="white",
                 command=reset_password).pack(pady=10)
        tk.Button(dialog, text="Cancel", bg=COLORS["error"], fg="white",
                 command=dialog.destroy).pack(pady=5)

    def change_role_dialog(self, tree):
        """Change user role"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select a user.")
            return
        
        username = tree.item(selection[0])['values'][0]
        users = load_users()
        user_data = users.get(username)
        
        if not user_data:
            messagebox.showerror("Error", "User not found!")
            return
        
        if user_data.get("is_primary"):
            messagebox.showwarning("Restricted", "Cannot modify primary admin roles!")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Change User Role")
        dialog.geometry("300x200")
        dialog.configure(bg=COLORS["background"])
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text=f"Change Role for {username}", font=("Arial", 12, "bold"),
                fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        
        tk.Label(dialog, text="New Role:", fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=5)
        role_var = tk.StringVar(value=user_data.get("role", "Investor"))
        role_combo = ttk.Combobox(dialog, textvariable=role_var, 
                                 values=["Investor", "Researcher", "Admin"], state="readonly")
        role_combo.pack(pady=5)
        
        def change_role():
            new_role = role_var.get()
            success, message = change_user_role(username, new_role, self.username)
            if success:
                messagebox.showinfo("Success", message)
                self.refresh_user_table(tree)
                dialog.destroy()
            else:
                messagebox.showerror("Error", message)
        
        tk.Button(dialog, text="Change Role", bg=COLORS["success"], fg="white",
                 command=change_role).pack(pady=10)
        tk.Button(dialog, text="Cancel", bg=COLORS["error"], fg="white",
                 command=dialog.destroy).pack(pady=5)

    def view_user_details(self, tree):
        """View detailed user information"""
        selection = tree.selection()
        if not selection:
            messagebox.showwarning("Selection Required", "Please select a user.")
            return
        
        username = tree.item(selection[0])['values'][0]
        users = load_users()
        user_data = users.get(username)
        
        if not user_data:
            messagebox.showerror("Error", "User not found!")
            return
        
        # Create detailed view dialog
        dialog = tk.Toplevel(self.root)
        dialog.title(f"User Details - {username}")
        dialog.geometry("500x400")
        dialog.configure(bg=COLORS["background"])
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Create scrollable text widget
        text_frame = tk.Frame(dialog, bg=COLORS["background"])
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        text_widget = tk.Text(text_frame, wrap="word", bg=COLORS["input_bg"], fg=COLORS["text_light"],
                             font=("Arial", 10), padx=10, pady=10)
        scrollbar = tk.Scrollbar(text_frame, command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Format user data for display
        details = f"""USER DETAILS FOR: {username}
{'='*50}

BASIC INFORMATION:
• Username: {username}
• Email: {user_data.get('email', 'N/A')}
• Role: {user_data.get('role', 'N/A')}
• Status: {user_data.get('status', 'active')}
• Primary Admin: {'Yes' if user_data.get('is_primary') else 'No'}

PERSONAL INFORMATION:
• Full Name: {user_data.get('name', '')} {user_data.get('surname', '')}
• Age: {user_data.get('age', 'N/A')}
• Gender: {user_data.get('gender', 'N/A')}
• Nationality: {user_data.get('nationality', 'N/A')}
• Institution: {user_data.get('institution', 'N/A')}
• ID Number: {user_data.get('id_number', 'N/A')}

SYSTEM INFORMATION:
• Date Created: {user_data.get('date_created', 'N/A')}
• Created By: {user_data.get('created_by', 'System')}
• Last Modified: {user_data.get('modified_date', 'Never')}
• Modified By: {user_data.get('modified_by', 'N/A')}
• Password Reset: {user_data.get('password_reset_date', 'Never')}
• Reset By: {user_data.get('password_reset_by', 'N/A')}

SECURITY INFORMATION:
• Password Hash: {user_data.get('password', 'N/A')[:50]}...
• Account Type: {'Primary Admin' if user_data.get('is_primary') else 'Standard User'}
"""
        
        text_widget.insert("1.0", details)
        text_widget.config(state="disabled")  # Make read-only
        
        tk.Button(dialog, text="Close", bg=COLORS["primary"], fg="white",
                 command=dialog.destroy).pack(pady=10)

    def open_admin_tools(self):
        """Open advanced admin tools"""
        if self.role.lower() != "admin":
            messagebox.showwarning("Access Denied", "Admin privileges required!")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Admin Tools")
        dialog.geometry("600x400")
        dialog.configure(bg=COLORS["background"])
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="🛠️ Administrator Tools", font=("Arial", 16, "bold"),
                fg=COLORS["text_light"], bg=COLORS["background"]).pack(pady=10)
        
        # Admin tools frame
        tools_frame = tk.Frame(dialog, bg=COLORS["card_bg"], padx=20, pady=20)
        tools_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Tool buttons
        tools = [
            ("🔍 Audit Log", "View system audit log", lambda: self.show_audit_log()),
            ("📊 Database Stats", "Database statistics", lambda: self.show_db_stats()),
            ("🔄 System Backup", "Create system backup", lambda: self.create_backup()),
            ("🚀 Performance", "System performance", lambda: self.show_performance()),
            ("🔐 Security", "Security settings", lambda: self.show_security()),
            ("📋 Logs", "System logs", lambda: self.show_system_logs())
        ]
        
        for i, (text, desc, command) in enumerate(tools):
            btn = tk.Button(tools_frame, text=text, font=("Arial", 11),
                          bg=COLORS["primary"], fg=COLORS["text_light"],
                          command=command, width=15, height=2)
            btn.grid(row=i//2, column=i%2, padx=10, pady=10, sticky="nsew")
            tk.Label(tools_frame, text=desc, font=("Arial", 8),
                    fg=COLORS["text_muted"], bg=COLORS["card_bg"]).grid(row=i//2, column=i%2, pady=(40, 0))
        
        # Configure grid weights
        for i in range(3):
            tools_frame.rowconfigure(i, weight=1)
        for i in range(2):
            tools_frame.columnconfigure(i, weight=1)

    def show_system_stats(self):
        """Show system statistics"""
        users = load_users()
        total_users = len(users)
        admin_count = get_admin_count()
        primary_admin_count = sum(1 for u in users.values() if u.get("is_primary"))
        active_users = sum(1 for u in users.values() if u.get("status", "active") == "active")
        
        stats_text = f"""SYSTEM STATISTICS
{'='*20}

USER STATISTICS:
• Total Users: {total_users}
• Admin Users: {admin_count}
• Primary Admins: {primary_admin_count}
• Active Users: {active_users}
• Suspended Users: {total_users - active_users}

ADMIN LIMITS:
• Primary Admins: {len(PRIMARY_ADMINS)} (Fixed)
• Maximum Regular Admins: 10
• Current Regular Admins: {admin_count - primary_admin_count}

DATA STATISTICS:
• Mineral Records: {len(flat_critical_africa)}
• Countries Covered: {flat_critical_africa['Country'].nunique()}
• Minerals Tracked: {flat_critical_africa['Mineral'].nunique()}
• Data Years: {sorted(flat_critical_africa['Year'].unique())}

YOUR PRIVILEGES:
• Username: {self.username}
• Role: {self.role}
• Primary Admin: {'Yes' if is_primary_admin(self.username) else 'No'}
• Can Create Admins: {'Yes' if can_create_admin(self.username, 'admin') else 'No'}
"""
        
        messagebox.showinfo("System Statistics", stats_text)

    def show_audit_log(self):
        messagebox.showinfo("Audit Log", "Audit log functionality would be implemented here.")
    
    def show_db_stats(self):
        messagebox.showinfo("Database Stats", "Database statistics functionality would be implemented here.")
    
    def create_backup(self):
        messagebox.showinfo("Backup", "System backup functionality would be implemented here.")
    
    def show_performance(self):
        messagebox.showinfo("Performance", "Performance monitoring would be implemented here.")
    
    def show_security(self):
        messagebox.showinfo("Security", "Security settings would be implemented here.")
    
    def show_system_logs(self):
        messagebox.showinfo("System Logs", "System logs would be implemented here.")

    def show_about(self):
        about_text = f"""African Critical Minerals Dashboard
Version: 4.0 (Enhanced Security)

User: {self.username} ({self.role})
Primary Admin: {'Yes' if is_primary_admin(self.username) else 'No'}

SECURITY FEATURES:
• 4 Primary Admin System
• Role-based Access Control
• Password Hashing
• Admin Creation Restrictions
• User Status Management

Data Sources: BGR Mineral Raw Materials Database
© 2024 Mineral Analytics Platform"""
        messagebox.showinfo("About", about_text)

    def logout(self):
        if messagebox.askyesno("Logout", "Are you sure you want to logout?"):
            self.root.destroy()
            start_login_system()

# --- Main Execution ---
def start_login_system():
    """Initialize the application and start the login system."""
    print("MineralXAfrica Dashboard - Enhanced Security Edition")
    print(f"Primary Admin PIN: {PRIMARY_ADMIN_PIN}")
    
    # Ensure user file exists
    initialize_users_file() 
    
    # Instead of a splash screen, show a quick info message
    # Note: This requires a temporary hidden root window
    temp_root = tk.Tk()
    temp_root.withdraw() # Hide the temporary window
    
    # Show initialization message
    messagebox.showinfo(
        "MineralXAfrica Initializing",
        "Loading application data...\n\n"
        "Please wait, the login screen will appear shortly.\n"
        "© 2024 Mineral Analytics Platform"
    )
    
    # Destroy the temporary root
    temp_root.destroy()
    
    # Start the main login system
    LoginSystem()

if __name__ == "__main__":
    start_login_system()



