import cv2
import pytesseract
import pandas as pd
import os
import time
from aadhaar_details import get_values, get_address
from pathlib import Path
import json

# Function to save data to Excel
def save_to_excel(name, gender, dob, aadhaar_no, address):
    filename = "aadhaar_data1.xlsx"

    # Define proper column names
    columns = ["Name", "Gender", "DOB", "Aadhaar No", "Address"]

    # Check if file exists; if not, create a new one with column headers
    if not os.path.exists(filename):
        df = pd.DataFrame(columns=columns)
        df.to_excel(filename, index=False, engine="openpyxl")

    # Read existing data
    df = pd.read_excel(filename, engine="openpyxl")

    # Append new row
    new_entry = pd.DataFrame([[name, gender, dob, aadhaar_no, address]], columns=columns)
    df = pd.concat([df, new_entry], ignore_index=True)

    # Save to Excel
    with pd.ExcelWriter(filename, mode="w", engine="openpyxl") as writer:
        df.to_excel(writer, index=False)

    print("✅ Data successfully saved to Excel!")

# Function to save data in JSON and Excel
def send_to_json(name, gender, dob, mobile_number, aadhaar_no, address):
    time_sec = str(time.time()).replace(".", "_")
    json_data = {
        time_sec: {
            "name": name,
            "gender": gender,
            "dob": dob,
            "mobile number": mobile_number,
            "aadhaar number": aadhaar_no,
            "address": address,
        }
    }

    aadhaar_info_path = f"aadhaar_info_{time_sec}.json"
    with open(aadhaar_info_path, "w") as f:
        json.dump(json_data, f, indent=4)
        print("📜 JSON file saved!")

    # Save data to Excel
    save_to_excel(name, gender, dob, aadhaar_no, address)  # Fixed argument order

if __name__ == "__main__":
    # Set Tesseract path (Update if necessary)
    tesseract_path = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)

    # Aadhaar images (Ensure they exist)
    aadhaar_front_img_path = Path("f.png")
    aadhaar_back_img_path = Path("b.png")

    # Process front image
    if not aadhaar_front_img_path.exists():
        raise FileNotFoundError(f"❌ Could not find {aadhaar_front_img_path}")

    img = cv2.imread(str(aadhaar_front_img_path))
    img = cv2.resize(img, (0, 0), fx=0.5, fy=0.5)

    name, gender, dob, mobile_number, aadhaar_no = get_values(img)
    
    # Ensure full name extraction
    if isinstance(name, list) and len(name) > 0:
        name = " ".join(name)  # Extract full name dynamically
    else:
        name = "Unknown"

    print(name, gender, dob, mobile_number, aadhaar_no)

    # Process back image
    if not aadhaar_back_img_path.exists():
        raise FileNotFoundError(f"❌ Could not find {aadhaar_back_img_path}")

    img = cv2.imread(str(aadhaar_back_img_path))
    img = cv2.resize(img, (0, 0), fx=0.75, fy=0.75)
    address = get_address(img) or "Address not detected"

    print(address)

    # Save data
    send_to_json(name, gender, dob, mobile_number, aadhaar_no, address)
