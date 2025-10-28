# MINN2020A-Group-15-Project-2025
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
