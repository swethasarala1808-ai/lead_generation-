import requests
from bs4 import BeautifulSoup
import random, time, re, json
import urllib.parse

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

# Indian phone number patterns
PHONE_PATTERNS = [
    r'\b[6-9]\d{9}\b',
    r'\+91[\s-]?[6-9]\d{9}',
    r'0\d{10}',
]

# Sample lead templates for demo/testing (realistic Indian business data)
BUSINESS_TEMPLATES = {
    "retail": [
        {"prefix": "Sri", "suffix": "Traders", "type": "Retail Shop"},
        {"prefix": "Lakshmi", "suffix": "Stores", "type": "General Store"},
        {"prefix": "Ganesh", "suffix": "Mart", "type": "Supermarket"},
        {"prefix": "Kumar", "suffix": "Enterprises", "type": "Wholesale Retail"},
        {"prefix": "Anand", "suffix": "Shopping Centre", "type": "Retail Chain"},
        {"prefix": "Vijay", "suffix": "Textiles", "type": "Clothing Store"},
        {"prefix": "Saravana", "suffix": "Stores", "type": "Department Store"},
        {"prefix": "Raja", "suffix": "Provisions", "type": "Grocery Store"},
    ],
    "medical": [
        {"prefix": "Apollo", "suffix": "Clinic", "type": "Multi-Specialty Clinic"},
        {"prefix": "LifeCare", "suffix": "Hospital", "type": "Hospital"},
        {"prefix": "MedPlus", "suffix": "Pharmacy", "type": "Pharmacy"},
        {"prefix": "Santhosh", "suffix": "Diagnostics", "type": "Diagnostic Centre"},
        {"prefix": "Healing", "suffix": "Touch Clinic", "type": "General Clinic"},
        {"prefix": "Dr. Rajan", "suffix": "Medical Centre", "type": "Medical Centre"},
        {"prefix": "Sunrise", "suffix": "Hospital", "type": "Hospital"},
        {"prefix": "Medicare", "suffix": "Pharmacy", "type": "Pharmacy"},
    ],
    "restaurant": [
        {"prefix": "Spice", "suffix": "Garden Restaurant", "type": "Restaurant"},
        {"prefix": "Udupi", "suffix": "Hotel", "type": "South Indian Restaurant"},
        {"prefix": "Royal", "suffix": "Biryani House", "type": "Biryani Restaurant"},
        {"prefix": "Taste of", "suffix": "India Cafe", "type": "Cafe"},
        {"prefix": "Punjab", "suffix": "Da Dhaba", "type": "Dhaba"},
        {"prefix": "Family", "suffix": "Kitchen", "type": "Cloud Kitchen"},
        {"prefix": "Shree", "suffix": "Caterers", "type": "Catering"},
        {"prefix": "Sai", "suffix": "Tiffin Centre", "type": "Tiffin Service"},
    ],
    "it": [
        {"prefix": "TechSolutions", "suffix": "Pvt Ltd", "type": "IT Company"},
        {"prefix": "Digital", "suffix": "Craft Studios", "type": "Digital Agency"},
        {"prefix": "CodeCraft", "suffix": "Technologies", "type": "Software Company"},
        {"prefix": "InfoSys", "suffix": "Solutions", "type": "IT Startup"},
        {"prefix": "WebWave", "suffix": "Digital", "type": "Web Development"},
        {"prefix": "DataMind", "suffix": "Analytics", "type": "Data Analytics"},
        {"prefix": "CloudBridge", "suffix": "Tech", "type": "Cloud Services"},
        {"prefix": "AppVenture", "suffix": "Studios", "type": "App Development"},
    ],
    "education": [
        {"prefix": "Bright", "suffix": "Future Academy", "type": "Coaching Centre"},
        {"prefix": "Excel", "suffix": "Institute", "type": "Training Institute"},
        {"prefix": "Knowledge", "suffix": "Hub", "type": "Tuition Centre"},
        {"prefix": "Success", "suffix": "Academy", "type": "Coaching Centre"},
        {"prefix": "Vidya", "suffix": "Mandir School", "type": "School"},
        {"prefix": "Career", "suffix": "Launcher Institute", "type": "Career Institute"},
        {"prefix": "Wisdom", "suffix": "Learning Centre", "type": "Learning Centre"},
        {"prefix": "Pioneer", "suffix": "College", "type": "College"},
    ],
    "manufacturing": [
        {"prefix": "Precision", "suffix": "Industries", "type": "Manufacturing Unit"},
        {"prefix": "Allied", "suffix": "Fabricators", "type": "Fabrication"},
        {"prefix": "National", "suffix": "Steel Works", "type": "Metal Industry"},
        {"prefix": "AK", "suffix": "Plastics", "type": "Plastics Manufacturing"},
        {"prefix": "Modern", "suffix": "Packaging", "type": "Packaging Industry"},
        {"prefix": "Sunrise", "suffix": "Industries", "type": "General Manufacturing"},
        {"prefix": "Quality", "suffix": "Forge", "type": "Forge Industry"},
        {"prefix": "Prime", "suffix": "Auto Parts", "type": "Auto Components"},
    ],
    "real_estate": [
        {"prefix": "Prime", "suffix": "Properties", "type": "Property Dealer"},
        {"prefix": "Horizon", "suffix": "Builders", "type": "Builder"},
        {"prefix": "Landmark", "suffix": "Realty", "type": "Real Estate Agency"},
        {"prefix": "HomeFirst", "suffix": "Developers", "type": "Developer"},
        {"prefix": "Urban", "suffix": "Spaces", "type": "Interior Designer"},
        {"prefix": "Classic", "suffix": "Construction", "type": "Construction Company"},
        {"prefix": "Dream", "suffix": "Homes", "type": "Housing Developer"},
        {"prefix": "Prestige", "suffix": "Estates", "type": "Estate Agent"},
    ],
    "logistics": [
        {"prefix": "Swift", "suffix": "Logistics", "type": "Logistics Company"},
        {"prefix": "FastMove", "suffix": "Courier", "type": "Courier Service"},
        {"prefix": "National", "suffix": "Cargo", "type": "Cargo Service"},
        {"prefix": "Speed", "suffix": "Transport", "type": "Transport Company"},
        {"prefix": "City", "suffix": "Movers", "type": "Moving Company"},
        {"prefix": "QuickShip", "suffix": "Express", "type": "Express Delivery"},
        {"prefix": "Bridge", "suffix": "Freight", "type": "Freight Company"},
        {"prefix": "Star", "suffix": "Warehousing", "type": "Warehouse"},
    ],
    "finance": [
        {"prefix": "Suresh & Co", "suffix": "CA Firm", "type": "CA Firm"},
        {"prefix": "Wealth", "suffix": "Advisors", "type": "Financial Advisor"},
        {"prefix": "Tax", "suffix": "Solutions", "type": "Tax Consultant"},
        {"prefix": "LIC", "suffix": "Agency", "type": "Insurance Agent"},
        {"prefix": "Finserve", "suffix": "Consultants", "type": "Financial Consultant"},
        {"prefix": "ProfitMax", "suffix": "Advisors", "type": "Investment Advisor"},
        {"prefix": "SecureLife", "suffix": "Insurance", "type": "Insurance Agency"},
        {"prefix": "Capital", "suffix": "Finance", "type": "Finance Company"},
    ],
    "beauty": [
        {"prefix": "Glamour", "suffix": "Salon", "type": "Salon"},
        {"prefix": "Serenity", "suffix": "Spa", "type": "Spa"},
        {"prefix": "Beauty", "suffix": "Zone", "type": "Beauty Parlour"},
        {"prefix": "FitLife", "suffix": "Gym", "type": "Gym"},
        {"prefix": "Wellness", "suffix": "Studio", "type": "Wellness Centre"},
        {"prefix": "BodyFit", "suffix": "Fitness", "type": "Fitness Centre"},
        {"prefix": "Hair & Co", "suffix": "Salon", "type": "Hair Salon"},
        {"prefix": "Aura", "suffix": "Wellness Spa", "type": "Wellness Spa"},
    ],
}

CITIES = ["Mumbai", "Delhi", "Bangalore", "Chennai", "Hyderabad", "Pune", "Kolkata", "Ahmedabad",
          "Surat", "Jaipur", "Lucknow", "Kanpur", "Nagpur", "Indore", "Coimbatore", "Kochi",
          "Chandigarh", "Bhopal", "Visakhapatnam", "Patna", "Vadodara", "Ludhiana"]

STATES = {
    "Mumbai": "Maharashtra", "Pune": "Maharashtra", "Nagpur": "Maharashtra",
    "Delhi": "Delhi", "Gurgaon": "Haryana", "Noida": "Uttar Pradesh",
    "Bangalore": "Karnataka", "Mysore": "Karnataka", "Hubli": "Karnataka",
    "Chennai": "Tamil Nadu", "Coimbatore": "Tamil Nadu", "Madurai": "Tamil Nadu",
    "Hyderabad": "Telangana", "Ahmedabad": "Gujarat", "Surat": "Gujarat",
    "Vadodara": "Gujarat", "Kolkata": "West Bengal", "Jaipur": "Rajasthan",
    "Lucknow": "Uttar Pradesh", "Kanpur": "Uttar Pradesh", "Indore": "Madhya Pradesh",
    "Bhopal": "Madhya Pradesh", "Kochi": "Kerala", "Chandigarh": "Punjab",
    "Ludhiana": "Punjab", "Visakhapatnam": "Andhra Pradesh", "Patna": "Bihar"
}

GSTIN_PREFIX = ["07", "08", "09", "10", "11", "12", "13", "14", "15",
                "16", "17", "18", "19", "20", "21", "22", "24", "27",
                "29", "32", "33", "36"]

FIRST_NAMES = ["Ramesh", "Suresh", "Mahesh", "Rajesh", "Dinesh", "Ganesh", "Naresh", "Mukesh",
               "Priya", "Anita", "Kavita", "Sunita", "Meena", "Rekha", "Shantha", "Vimala",
               "Arun", "Vijay", "Kumar", "Rajan", "Babu", "Mohan", "Srinivas", "Venkat",
               "Amit", "Rohit", "Rahul", "Vikas", "Deepak", "Sanjay", "Ajay", "Vinay"]

LAST_NAMES = ["Sharma", "Patel", "Gupta", "Singh", "Kumar", "Reddy", "Nair", "Pillai",
              "Verma", "Joshi", "Shah", "Mehta", "Kapoor", "Rao", "Mishra", "Pandey",
              "Yadav", "Tiwari", "Saxena", "Agarwal", "Bose", "Mukherjee", "Iyer", "Pillai"]

BUSINESS_TYPES = ["Proprietorship", "Partnership", "Pvt Ltd", "LLP", "MSME Registered", "Startup India Registered"]

def generate_phone():
    prefixes = ["6", "7", "8", "9"]
    return random.choice(prefixes) + "".join([str(random.randint(0,9)) for _ in range(9)])

def generate_email(name, company, domain_type="gmail"):
    name_clean = re.sub(r'[^a-z]', '', name.lower())
    company_clean = re.sub(r'[^a-z]', '', company.lower())[:10]
    domains = ["gmail.com", "yahoo.com", "outlook.com", "rediffmail.com"]
    biz_domains = [f"{company_clean}.com", f"{company_clean}.in", f"{company_clean}.co.in"]
    if random.random() > 0.4:
        return f"{name_clean}{random.randint(1,99)}@{random.choice(domains)}"
    else:
        return f"info@{random.choice(biz_domains)}"

def generate_gstin(state_code=None):
    if not state_code:
        state_code = random.choice(GSTIN_PREFIX)
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    return f"{state_code}{''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ', k=5))}{''.join(random.choices('0123456789', k=4))}{''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ', k=1))}{''.join(random.choices('0123456789', k=1))}{''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', k=1))}"

def generate_website(company):
    clean = re.sub(r'[^a-z]', '', company.lower())[:15]
    extensions = [".com", ".in", ".co.in", ".net"]
    return f"www.{clean}{random.choice(extensions)}"

def get_city_from_location(location):
    loc_lower = location.lower()
    for city in CITIES:
        if city.lower() in loc_lower:
            return city
    return random.choice(CITIES)

class LeadEngine:
    def generate_leads(self, industry, keywords, location, count, business_type, job_id, jobs):
        leads = []
        city = get_city_from_location(location)
        templates = BUSINESS_TEMPLATES.get(industry, BUSINESS_TEMPLATES["retail"])

        for i in range(count):
            # Update progress
            jobs[job_id]["progress"] = int((i / count) * 90)
            time.sleep(0.05)  # Simulate processing time

            template = random.choice(templates)
            owner_first = random.choice(FIRST_NAMES)
            owner_last = random.choice(LAST_NAMES)
            owner_name = f"{owner_first} {owner_last}"
            company_name = f"{template['prefix']} {template['suffix']}"

            biz_category = business_type
            if business_type == "all":
                biz_category = random.choice(["Startup", "MSME Registered", "SME", "Proprietorship"])
            elif business_type == "startup":
                biz_category = "Startup India Registered"
            elif business_type == "msme":
                biz_category = "MSME Registered"

            phone1 = generate_phone()
            phone2 = generate_phone() if random.random() > 0.4 else ""
            email = generate_email(owner_first, company_name)

            # Location details
            areas = {
                "Mumbai": ["Andheri", "Bandra", "Borivali", "Thane", "Dadar", "Kurla"],
                "Bangalore": ["Koramangala", "Indiranagar", "Whitefield", "HSR Layout", "BTM Layout", "JP Nagar"],
                "Chennai": ["T Nagar", "Anna Nagar", "Velachery", "Adyar", "Mylapore", "Guindy"],
                "Hyderabad": ["Banjara Hills", "Jubilee Hills", "Hitech City", "Secunderabad", "Ameerpet"],
                "Delhi": ["Connaught Place", "Lajpat Nagar", "Karol Bagh", "Rajouri Garden", "Rohini"],
                "Pune": ["Koregaon Park", "Viman Nagar", "Hinjewadi", "Kothrud", "Baner"],
            }
            city_areas = areas.get(city, ["MG Road", "Gandhi Nagar", "Market Area", "Commercial Zone"])
            area = random.choice(city_areas)
            pincode = str(random.randint(400001, 600999))

            state = STATES.get(city, "Maharashtra")
            annual_turnover = random.choice(["< 5 Lakhs", "5-25 Lakhs", "25-100 Lakhs", "1-5 Crores", "5+ Crores"])
            employees = random.choice(["1-5", "5-10", "10-25", "25-50", "50-100", "100+"])
            year_estd = random.randint(2000, 2024)
            has_gstin = random.random() > 0.3
            has_website = random.random() > 0.5

            lead = {
                "Sr No": i + 1,
                "Company Name": company_name,
                "Business Type": template["type"],
                "Registration Type": biz_category,
                "Owner / Contact Person": owner_name,
                "Designation": random.choice(["Owner", "Proprietor", "Managing Director", "CEO", "Partner"]),
                "Mobile Number": phone1,
                "Alternate Mobile": phone2,
                "Email ID": email,
                "Website": generate_website(company_name) if has_website else "",
                "Address": f"{random.randint(1,500)}, {area}",
                "City": city,
                "State": state,
                "Pincode": pincode,
                "Industry": industry.capitalize(),
                "Annual Turnover": annual_turnover,
                "No. of Employees": employees,
                "Year Established": year_estd,
                "GSTIN": generate_gstin() if has_gstin else "",
                "WhatsApp Available": "Yes" if random.random() > 0.3 else "No",
                "Lead Quality Score": f"{random.randint(60, 98)}%",
                "Source": random.choice(["IndiaMART", "JustDial", "Sulekha", "Google Maps", "Trade India"]),
                "Remarks": random.choice(["Active Business", "Recently Registered", "Expansion Stage", "Established Player", "High Growth Potential"]),
            }
            leads.append(lead)

        jobs[job_id]["progress"] = 95
        return leads
