from openpyxl.styles import numbers
from flask import Flask, render_template, request, jsonify
import json
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

subjects = {
    "DMGT": 4, "Programming in Java": 3, "DBMS": 3, "OS": 3, "AP": 3,
    "GP-II": 1, "Programming in Java Laboratory": 1, "DBMS LAB": 1, "OS LAB": 1, "Programming in C++":3
}

grade_points = {"S": 10, "A": 9, "B": 8, "C": 7, "D": 6, "E": 5, "F": 0}
total_credits = sum(subjects.values())

# Fixed student data
students = [
    {"S.No": 1, "RegNo": "23UCS003", "Name of the Student": "ABHINAV RAVEENDRAN"},
    {"S.No": 2, "RegNo": "23UCS004", "Name of the Student": "ABISHEK M"},
    {"S.No": 3, "RegNo": "23UCS005", "Name of the Student": "ADHIRAI S"},
    {"S.No": 4, "RegNo": "23UCS006", "Name of the Student": "AGILASHE A"},
    {"S.No": 5, "RegNo": "23UCS010", "Name of the Student": "AKSHAYA KRITHIKA A"},
    {"S.No": 6, "RegNo": "23UCS017", "Name of the Student": "ARNOLD A"},
    {"S.No": 7, "RegNo": "23UCS026", "Name of the Student": "BHAVANA T"},
    {"S.No": 8, "RegNo": "23UCS027", "Name of the Student": "BHUVAN SHANKAR P"},
    {"S.No": 9, "RegNo": "23UCS028", "Name of the Student": "BUELA T"},
    {"S.No": 10, "RegNo": "23UCS033", "Name of the Student": "DEVAGURU V"},
    {"S.No": 11, "RegNo": "23UCS040", "Name of the Student": "DHIVYA DHARSINI I"},
    {"S.No": 12, "RegNo": "23UCS044", "Name of the Student": "ELAVARASA S"},
    {"S.No": 13, "RegNo": "23UCS045", "Name of the Student": "EMMANUEL MATTHEW J"},
    {"S.No": 14, "RegNo": "23UCS048", "Name of the Student": "GANESH S"},
    {"S.No": 15, "RegNo": "23UCS049", "Name of the Student": "GAYATHRI J"},
    {"S.No": 16, "RegNo": "23UCS053", "Name of the Student": "GOKUL S"},
    {"S.No": 17, "RegNo": "23UCS054", "Name of the Student": "GOPIKA R"},
    {"S.No": 18, "RegNo": "23UCS058", "Name of the Student": "HARINI A"},
    {"S.No": 19, "RegNo": "23UCS061", "Name of the Student": "HARSHINI VASUTHAA P"},
    {"S.No": 20, "RegNo": "23UCS063", "Name of the Student": "HEMAPRIYA S"},
    {"S.No": 21, "RegNo": "23UCS064", "Name of the Student": "HEMASANKAR P"},
    {"S.No": 22, "RegNo": "23UCS067", "Name of the Student": "ISHWARYALAKSHMI D"},
    {"S.No": 23, "RegNo": "23UCS069", "Name of the Student": "JANCY DESSAINTS A"},
    {"S.No": 24, "RegNo": "23UCS071", "Name of the Student": "JAYASELAN R"},
    {"S.No": 25, "RegNo": "23UCS075", "Name of the Student": "KAAVIYA SALLUSTRE R"},
    {"S.No": 26, "RegNo": "23UCS076", "Name of the Student": "KAKINADA SRI TANUJA"},
    {"S.No": 27, "RegNo": "23UCS078", "Name of the Student": "KALAIVANI A"},
    {"S.No": 28, "RegNo": "23UCS084", "Name of the Student": "KAVYA B"},
    {"S.No": 29, "RegNo": "23UCS090", "Name of the Student": "KOUSHIK PATEL P"},
    {"S.No": 30, "RegNo": "23UCS092", "Name of the Student": "LAKSHMANAN S"},
    {"S.No": 31, "RegNo": "23UCS098", "Name of the Student": "MANIMUGILAN S"},
    {"S.No": 32, "RegNo": "23UCS105", "Name of the Student": "NANDINI K"},
    {"S.No": 33, "RegNo": "23UCS107", "Name of the Student": "NIRANJAN R"},
    {"S.No": 34, "RegNo": "23UCS108", "Name of the Student": "NISHAN S"},
    {"S.No": 35, "RegNo": "23UCS111", "Name of the Student": "NIVEDITA K"},
    {"S.No": 36, "RegNo": "23UCS113", "Name of the Student": "PAKALAPATI NAGA MAHA ADARSH VARMA"},
    {"S.No": 37, "RegNo": "23UCS114", "Name of the Student": "PAVANKUMAR MU"},
    {"S.No": 38, "RegNo": "23UCS116", "Name of the Student": "PERARASU P"},
    {"S.No": 39, "RegNo": "23UCS117", "Name of the Student": "POOJA R"},
    {"S.No": 40, "RegNo": "23UCS126", "Name of the Student": "RAMKUMAR S"},
    {"S.No": 41, "RegNo": "23UCS131", "Name of the Student": "ROHAN R"},
    {"S.No": 42, "RegNo": "23UCS132", "Name of the Student": "ROOPASHREE A"},
    {"S.No": 43, "RegNo": "23UCS135", "Name of the Student": "SACHIN M"},
    {"S.No": 44, "RegNo": "23UCS136", "Name of the Student": "SADHANADEVI S"},
    {"S.No": 45, "RegNo": "23UCS138", "Name of the Student": "SAMIES R"},
    {"S.No": 46, "RegNo": "23UCS142", "Name of the Student": "SHAAME U S"},
    {"S.No": 47, "RegNo": "23UCS144", "Name of the Student": "SHALINI S"},
    {"S.No": 48, "RegNo": "23UCS145", "Name of the Student": "SHARAN T"},
    {"S.No": 49, "RegNo": "23UCS149", "Name of the Student": "SILAMBARASAN T"},
    {"S.No": 50, "RegNo": "23UCS150", "Name of the Student": "SIVADARANI L"},
    {"S.No": 51, "RegNo": "23UCS154", "Name of the Student": "SRINIVASAN K"},
    {"S.No": 52, "RegNo": "23UCS155", "Name of the Student": "SUBATHRA I"},
    {"S.No": 53, "RegNo": "23UCS165", "Name of the Student": "VARSHENI"},
    {"S.No": 54, "RegNo": "23UCS167", "Name of the Student": "VINESH S"},
    {"S.No": 55, "RegNo": "23UCS172", "Name of the Student": "WILSON BERNARD W"},
    {"S.No": 56, "RegNo": "23UCS175", "Name of the Student": "YOGESHWARI"},
    {"S.No": 57, "RegNo": "23CSL005", "Name of the Student": "SRI VISHNU PRASATH"},
    {"S.No": 58, "RegNo": "23CSL004", "Name of the Student": "SHARATH RAJ"},
]


DATA_FILE = 'students_data.json'
EXCEL_FILE = 'students_grades.xlsx'

# Initialize JSON file
if not os.path.exists(DATA_FILE):
    with open(DATA_FILE, 'w') as f:
        json.dump([], f)

@app.route('/')
def home():
    return "Welcome to the CGPA Generator API"


@app.route('/submit_grades', methods=['POST'])
def submit_grades():
    data = request.get_json()
    reg_no = data.get("RegNo", "").strip().upper()

    # Search for student using RegNo
    student = next((s for s in students if s["RegNo"].upper() == reg_no), None)
    if not student:
        return jsonify({"error": "Student with RegNo not found"}), 404

    # Check grades
    for subject in subjects.keys():
        if subject not in data:
            return jsonify({"error": f"Missing grade for {subject}"}), 400
        if data[subject].strip().upper() not in grade_points:
            return jsonify({"error": f"Invalid grade {data[subject]} for {subject}"}), 400

    # GPA Calculation
    total_weighted_score = sum(grade_points[data[subj].strip().upper()] * credit for subj, credit in subjects.items())
    gpa = round(total_weighted_score / total_credits, 2)

    try:
        arrear = int(data.get("University Arrear", 0))
    except ValueError:
        arrear = 0

    # Combine and Save
    record = {
        **student,
        **{subj: data[subj].strip().upper() for subj in subjects},
        "University Arrear": arrear,
        "GPA": gpa
    }

    # Save to JSON file
    with open(DATA_FILE, 'r+') as f:
        records = json.load(f)
        records.append(record)
        f.seek(0)
        json.dump(records, f, indent=4)

    return jsonify({"message": "Grades submitted successfully", "GPA": gpa}), 200

@app.route('/generate_excel', methods=['GET'])
def generate_excel():
    with open(DATA_FILE, 'r') as f:
        records = json.load(f)

    if not records:
        return jsonify({'error': 'No data to write to Excel'}), 400

    # Define grade to points mapping
    grade_points = {"S": 10, "A": 9, "B": 8, "C": 7, "D": 6, "E": 5, "F": 0}

    total_credits = sum(subjects.values())

    # Calculate GPA for each record and add it to the record dict
    for record in records:
        total_weighted_points = 0
        for subj, credit in subjects.items():
            grade = record.get(subj, 'F').strip().upper()  # Get grade, default F
            points = grade_points.get(grade, 0)
            total_weighted_points += points * credit
        # Calculate GPA rounded to 2 decimals
        record['GPA'] = round(total_weighted_points / total_credits, 2)

    df = pd.DataFrame(records)
    ordered_columns = ['S.No', 'RegNo', 'Name of the Student'] + list(subjects.keys()) + ['GPA', 'University Arrear']
    df = df[ordered_columns]

    # Save DataFrame with grades and GPA as numbers
    df.to_excel(EXCEL_FILE, index=False)

    return jsonify({'message': 'Excel file generated with grades and calculated GPA values'}), 200

@app.route('/update_grades', methods=['PUT'])
def update_grades():
    data = request.json
    reg_no = data.get("RegNo", "").strip().upper()

    # Search for student in the original student list (loaded from your static students)
    student = next((s for s in students if s["RegNo"].upper() == reg_no), None)
    if not student:
        return jsonify({"error": "Student with RegNo not found"}), 404

    # Validate grades
    for subject in subjects.keys():
        if subject not in data:
            return jsonify({"error": f"Missing grade for {subject}"}), 400
        if data[subject].strip().upper() not in grade_points:
            return jsonify({"error": f"Invalid grade {data[subject]} for {subject}"}), 400

    # GPA Calculation
    total_weighted_score = sum(
        grade_points[data[subj].strip().upper()] * credit for subj, credit in subjects.items()
    )
    gpa = round(total_weighted_score / total_credits, 2)

    try:
        arrear = int(data.get("University Arrear", 0))
    except ValueError:
        arrear = 0

    updated_record = {
        **student,
        **{subj: data[subj].strip().upper() for subj in subjects},
        "University Arrear": arrear,
        "GPA": gpa
    }

    # Update existing record in JSON
    with open(DATA_FILE, 'r+') as f:
        records = json.load(f)
        updated = False
        for i, record in enumerate(records):
            if record['RegNo'].strip().upper() == reg_no:
                records[i] = updated_record
                updated = True
                break

        if not updated:
            records.append(updated_record)  # Add if not found (fallback)

        f.seek(0)
        f.truncate()
        json.dump(records, f, indent=4)

    return jsonify({"message": "Grades updated successfully", "GPA": gpa}), 200



if __name__ == '__main__':
    app.run(debug=True)

