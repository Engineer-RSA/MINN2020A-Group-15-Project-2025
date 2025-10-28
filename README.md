# MINN2020A-Group-15-Project-2025
 Sample Data
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
