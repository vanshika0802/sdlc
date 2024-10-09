import os
from flask import Flask, request, redirect, url_for, render_template, send_file
import pandas as pd
import fitz  # PyMuPDF
import re

app = Flask(__name__)

# Directory to temporarily store uploaded files (create this directory if it doesn't exist)
UPLOAD_FOLDER = 'uploaded_files'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Function to extract details from a PDF certificate
def extract_details_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = doc[0].get_text("text")
    
    name = "Unknown"
    roll_no = "Unknown"
    credits = "Unknown"
    assignment_marks = "Unknown"
    proctored_exam_marks = "Unknown"
    percentage = "Unknown"
    
    # Extract the name
    if "Data Analytics with Python" in text:
        try:
            name = text.split("Data Analytics with Python")[1].split("\n")[1].strip()
        except IndexError:
            pass

    # Extract the roll number
    roll_no_match = re.search(r'NPTEL[0-9A-Za-z]+', text)
    if roll_no_match:
        roll_no = roll_no_match.group(0).strip()

    # Extract the credits
    if "No. of credits recommended:" in text:
        try:
            credits = text.split("No. of credits recommended:")[1].split()[0].strip()
        except IndexError:
            pass

    # Extract assignment marks
    assignment_match = re.search(r'([0-9]{1,2}\.[0-9]{1,2}|[0-9]{1,2})/25', text)
    if assignment_match:
        assignment_marks = assignment_match.group(1).strip()

    # Extract proctored exam marks
    proctored_exam_match = re.search(r'([0-9]{1,2}\.[0-9]{1,2}|[0-9]{1,2})/75', text)
    if proctored_exam_match:
        proctored_exam_marks = proctored_exam_match.group(1).strip()

    # Extract the percentage
    percentage_match = re.search(r'(\d{1,3})\s*\n\s*11220', text)
    if percentage_match:
        percentage = percentage_match.group(1).strip() + "%"

    doc.close()
    return name, roll_no, credits, assignment_marks, proctored_exam_marks, percentage

# Homepage route
@app.route('/')
def index():
    return render_template('index1.html')  # Using the modified HTML file

# Route for handling file uploads and processing
@app.route('/upload', methods=['POST'])
def upload_folder():
    # Ensure files are uploaded
    if 'file' not in request.files:
        return "No file part", 400
    
    files = request.files.getlist('file')
    data = []
    
    for file in files:
        filename = file.filename
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)  # Save the uploaded file
        
        # Process PDF if it's a PDF file
        if filename.endswith(".pdf"):
            name, roll_no, credits, assignment_marks, proctored_exam_marks, percentage = extract_details_from_pdf(file_path)
            data.append([name, roll_no, credits, assignment_marks, proctored_exam_marks, percentage])
    
    # Create a DataFrame and save it to an Excel file
    df = pd.DataFrame(data, columns=["Name", "Roll Number", "Credits", "Assignment Marks", "Proctored Exam Marks", "Percentage"])
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], "Certificates_Details.xlsx")
    df.to_excel(excel_path, index=False)
    
    # Send the Excel file back to the user for download
    return send_file(excel_path, as_attachment=True)

# Success route (optional)
@app.route('/success')
def success():
    return "Files uploaded and processed successfully!"

if __name__ == '__main__':
    app.run(debug=True)
