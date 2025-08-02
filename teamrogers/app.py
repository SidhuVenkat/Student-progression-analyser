from flask import Flask, request, render_template, redirect, url_for, session, flash, send_file
import pandas as pd
import os
import io
import base64
from io import BytesIO
import matplotlib.pyplot as plt
import numpy as np
import smtplib
from email.mime.text import MIMEText
from flask import Response
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas



app = Flask(__name__)
app.secret_key = "your_secret_key_here"  # Use any random string, keep it safe


# Email configuration (replace with your actual details)
SENDER_EMAIL = "vignaninstituteofinftechnology@gmail.com"
SENDER_PASSWORD = "qjmmjfzzukagsgzl"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465

# Paths to Excel database
TEACHERS_DB = "teachers.xlsx"
STUDENTS_DB = "students.xlsx"

# Ensure database files exist
def initialize_db():
    if not os.path.exists(TEACHERS_DB):
        df = pd.DataFrame(columns=["Email", "Password", "Subject"])
        df.to_excel(TEACHERS_DB, index=False)

    if not os.path.exists(STUDENTS_DB):
        df = pd.DataFrame(columns=["Roll Number", "Class", "Name", "Password"])
        df.to_excel(STUDENTS_DB, index=False)
    else:
        # Ensure 'Homework' column exists
        df = pd.read_excel(STUDENTS_DB)
        # Remove general Homework column; homework will now be per subject
        homework_cols = [col for col in df.columns if col.endswith("_Homework")]
    if not homework_cols:
        df.to_excel(STUDENTS_DB, index=False)

initialize_db()

# Helper functions
def load_teachers():
    return pd.read_excel(TEACHERS_DB)

def save_teachers(df):
    df.to_excel(TEACHERS_DB, index=False)

def load_students():
    return pd.read_excel(STUDENTS_DB)

def save_students(df):
    df.to_excel(STUDENTS_DB, index=False)

@app.route("/")
def index():
    return render_template("index.html")


def assign_homework_auto(subject):
    df = load_students()
    condition_marks = 40
    homework_text = f"Complete revision for {subject} and submit by next class."
    subject_avg_col = f"{subject}_Average"
    subject_homework_col = f"{subject}_Homework"

    if subject_avg_col not in df.columns:
        df[subject_avg_col] = 0  # If no average, default to 0

    if subject_homework_col not in df.columns:
        df[subject_homework_col] = ""

    df.loc[df[subject_avg_col] < condition_marks, subject_homework_col] = homework_text
    save_students(df)


# ===================== TEACHER LOGIN =====================
@app.route("/teacher_login", methods=["GET", "POST"])
def teacher_login():
    if request.method == "POST":
        email = request.form["email"]
        password = request.form["password"]
        subject = request.form["subject"]

        df = load_teachers()
        teacher = df[
            (df["Email"] == email) & (df["Password"] == password) & (df["Subject"] == subject)
        ]

        if not teacher.empty:
            session["user"] = email
            session["role"] = "Teacher"
            session["subject"] = subject

            # ✅ Automatically Assign Homework
            assign_homework_auto(subject)

            return redirect(url_for("teacher_dashboard"))
        else:
            flash("Invalid credentials or subject mismatch", "danger")

    return render_template("teacher_login.html")

# ===================== STDENT LOGIN =====================
@app.route("/student_login", methods=["GET", "POST"])
def student_login():
    if request.method == "POST":
        roll_number = request.form["roll_number"].strip()
        student_class = request.form["class"].strip()
        password = request.form["password"].strip()

        df = load_students()

        # Make sure these columns are properly stripped and string typed
        df["Roll Number"] = df["Roll Number"].astype(str).str.strip()
        df["Class"] = df["Class"].astype(str).str.strip()
        df["Password"] = df["Password"].astype(str).str.strip()

        # Match student with all three fields
        student = df[
            (df["Roll Number"] == roll_number) &
            (df["Class"] == student_class) &
            (df["Password"] == password)
        ]

        if not student.empty:
            session["user"] = roll_number
            session["class"] = student_class
            session["role"] = "Student"
            return redirect(url_for("student_dashboard"))
        else:
            flash("❌ Invalid student credentials. Please check Roll Number, Class, and Password.", "error")

    return render_template("student_login.html")

@app.route("/download_students")
def download_students():
    if "user" not in session or session["role"] != "Teacher":
        return redirect(url_for("teacher_login"))
    
    file_path = "students.xlsx"  # Ensure this file exists in your project directory

    try:
        return send_file(file_path, as_attachment=True)
    except Exception as e:
        return f"Error: {e}"


# ===================== TEACHER DASHBOARD =====================
@app.route("/teacher_dashboard")
def teacher_dashboard():
    """Teacher dashboard with links to various functionalities and report status."""
    if "user" not in session or session["role"] != "Teacher":
        return redirect(url_for("teacher_login"))
    report_message = session.pop('report_message', None)
    return render_template("teacher_dashboard.html", report_message=report_message)
# ===================== STUDENT DASHBOARD =====================
@app.route("/student_dashboard", methods=["GET", "POST"])
def student_dashboard():
    if "user" not in session or session.get("role") != "Student":
        return redirect(url_for("student_login"))

    roll_number = session["user"]
    student_class = session["class"]

    df = load_students()
    if df is None:
        flash("Error loading student data.", "error")
        return redirect(url_for("student_login"))

    df["Roll Number"] = df["Roll Number"].astype(str).str.strip()
    df["Class"] = df["Class"].astype(str).str.strip()

    student_data_row = df[(df["Roll Number"] == roll_number) & (df["Class"] == student_class)]
    if student_data_row.empty:
        flash("Student record not found.", "error")
        return redirect(url_for("student_login"))

    student_data = student_data_row.iloc[0]

    selected_graph_option = request.form.get("selected_graph_option", "Test Performance")
    graph_data_base64 = None

    # Collect test marks for the table
    student_marks_table = {}
    for col in df.columns:
        if "_test" in col:
            subject = col.split("_test")[0]
            test_num = int(col.split("_test")[-1])
            mark = student_data.get(col, 'N/A')
            student_marks_table.setdefault(subject, []).append(mark)

    # Ensure all subjects have 3 test entries (fill with 'N/A' if missing)
    for subject in list(student_marks_table.keys()):
        while len(student_marks_table[subject]) < 3:
            student_marks_table[subject].append('N/A')

    # Collect averages for the table
    student_averages_table = {}
    for col in df.columns:
        if "_Average" in col:
            subject = col.replace("_Average", "")
            average = student_data.get(col, 'N/A')
            student_averages_table[subject] = average

    # Generate graph based on selection
    if selected_graph_option == "Test Performance":
        subject_tests = {}
        for col in df.columns:
            if "_test" in col:
                subject = col.split("_test")[0]
                test_num = int(col.split("_test")[-1])
                mark = student_data.get(col, None)
                if mark is not None:
                    subject_tests.setdefault(subject, {}).setdefault(test_num, mark)

        if subject_tests:
            subjects = list(subject_tests.keys())
            test_numbers = sorted(list(set(test for tests in subject_tests.values() for test in tests)))
            num_subjects = len(subjects)
            bar_width = 0.8 / num_subjects
            index = np.arange(len(test_numbers))

            plt.figure(figsize=(12, 7))
            for i, subject in enumerate(subjects):
                marks = [subject_tests[subject].get(test_num, None) for test_num in test_numbers]
                plt.bar(index + i * bar_width, [mark for mark in marks if mark is not None], bar_width, label=subject)

            plt.xlabel("Test Number")
            plt.ylabel("Marks")
            plt.title(f"Test Performance of {student_data['Name']} by Subject")
            plt.xticks(index + bar_width * (num_subjects - 1) / 2, [f"Test {num}" for num in test_numbers])
            plt.legend(title="Subject")
            plt.tight_layout()
            img_buffer = BytesIO()
            plt.savefig(img_buffer, format='png')
            img_buffer.seek(0)
            graph_data_base64 = base64.b64encode(img_buffer.read()).decode('utf-8')
            plt.close()

    elif selected_graph_option == "Average Performance":
        subject_averages_graph = {}
        for col in df.columns:
            if "_Average" in col:
                subject = col.replace("_Average", "")
                average = student_data.get(col, None)
                if average is not None:
                    subject_averages_graph[subject] = average

        if subject_averages_graph:
            subjects = list(subject_averages_graph.keys())
            averages = list(subject_averages_graph.values())

            plt.figure(figsize=(10, 6))
            plt.bar(subjects, averages, color='skyblue')
            plt.xlabel("Subject")
            plt.ylabel("Average Marks")
            plt.title(f"Average Performance of {student_data['Name']} by Subject")
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            img_buffer = BytesIO()
            plt.savefig(img_buffer, format='png')
            img_buffer.seek(0)
            graph_data_base64 = base64.b64encode(img_buffer.read()).decode('utf-8')
            plt.close()

    # Collect homework
    student_homework = []
    for col in df.columns:
        if col.endswith("_Homework"):
            subject = col.replace("_Homework", "")
            hw = student_data.get(col, "").strip()
            if hw:
                student_homework.append(f"{subject}: {hw}")

    return render_template(
        "student_dashboard.html",
        student_name=student_data["Name"],
        student_class=student_data["Class"],
        student_roll_number=student_data["Roll Number"],
        graph_data_base64=graph_data_base64,
        student_homework=student_homework,
        student_marks=student_marks_table,
        student_averages=student_averages_table,
        available_graph_options=["Test Performance", "Average Performance"],
        selected_graph_option=selected_graph_option
    )
# ===================== UPLOAD STUDENT DATA (TEACHER) =====================
@app.route("/upload_student_data", methods=["POST"])
def upload_student_data():
    if "user" not in session or session["role"] != "Teacher":
        return redirect(url_for("teacher_login"))

    if "file" not in request.files:
        flash("No file uploaded", "danger")
        return redirect(url_for("teacher_dashboard"))

    file = request.files["file"]
    if file.filename == "":
        flash("No selected file", "danger")
        return redirect(url_for("teacher_dashboard"))

    new_students = pd.read_excel(file)
    required_columns = ["Roll Number", "Class", "Name", "Password"]
    
    for col in required_columns:
        if col not in new_students.columns:
            flash(f"Invalid file format. Required columns: {', '.join(required_columns)}", "danger")
            return redirect(url_for("teacher_dashboard"))

    df = load_students()
    df["Roll Number"] = df["Roll Number"].astype(str).str.strip()

    for _, row in new_students.iterrows():
        roll_number = str(row["Roll Number"]).strip()
        row_data = row[required_columns].copy()
        row_data["Roll Number"] = roll_number
        row_data["Class"] = str(row_data["Class"]).strip()
        row_data["Password"] = str(row_data["Password"]).strip()

        if roll_number in df["Roll Number"].values:
            df.loc[df["Roll Number"] == roll_number, required_columns] = row_data
        else:
            df = pd.concat([df, pd.DataFrame([row_data])], ignore_index=True)

    save_students(df)
    flash("Student data uploaded successfully!", "success")
    return redirect(url_for("teacher_dashboard"))


# ===================== ASSIGN HOMEWORK =====================

@app.route("/assign_homework", methods=["POST"])
def assign_homework():
    if "user" not in session or session["role"] != "Teacher":
        return redirect(url_for("teacher_login"))

    class_name = request.form.get("class_name", "").strip()
    condition_type = request.form.get("condition_type")  # Get the selected condition
    condition_marks = int(request.form.get("condition_marks", 0))
    homework_text = request.form.get("homework_text", "").strip()
    subject = request.form.get("subject") or session.get("subject")

    if not class_name or not homework_text or not subject or not condition_type:
        flash("All fields are required, including subject and condition.", "error")
        return redirect(url_for("teacher_dashboard"))

    df = load_students()
    subject_avg_col = f"{subject}_Average"
    subject_homework_col = f"{subject}_Homework"

    if subject_avg_col not in df.columns:
        df[subject_avg_col] = 0
    if subject_homework_col not in df.columns:
        df[subject_homework_col] = ""

    df["Class"] = df["Class"].astype(str).str.strip()

    updated_df = assign_homework_to_students(df, subject, class_name, condition_type, condition_marks, homework_text)
    save_students(updated_df)

    flash("Homework assigned successfully!", "success")
    return redirect(url_for("teacher_dashboard"))

def assign_homework_to_students(df, subject, class_name, condition_type, threshold, homework_text):
    subject_avg_col = f"{subject}_Average"
    subject_homework_col = f"{subject}_Homework"

    for index, row in df[df["Class"] == class_name].iterrows():
        avg = row.get(subject_avg_col, 0)
        if avg != "":
            if condition_type == "less_than" and avg < threshold:
                df.at[index, subject_homework_col] = homework_text
            elif condition_type == "greater_than" and avg > threshold:
                df.at[index, subject_homework_col] = homework_text
            elif condition_type == "equal_to" and avg == threshold:
                df.at[index, subject_homework_col] = homework_text

    return df

# ===================== view stdent data =====================
@app.route("/view_student_data", methods=["POST"])
def view_student_data():
    if "user" not in session or session["role"] != "Teacher":
        return redirect(url_for("teacher_login"))

    roll_number = request.form.get("roll_number").strip()
    class_name = request.form.get("class_name_view").strip()

    # Load the students' data
    df = load_students()

    # Ensure Roll Number and Class are compared as strings
    df["Roll Number"] = df["Roll Number"].astype(str).str.strip()
    df["Class"] = df["Class"].astype(str).str.strip()

    # Find the student with the specified Roll Number and Class
    student = df[(df["Roll Number"] == roll_number) & (df["Class"] == class_name)]

    if not student.empty:
        student_data = student.iloc[0]
        excluded_columns = ["Roll Number", "Name", "Class", "DOB", "Password"]
        student_marks = {col: student_data[col] for col in df.columns if col not in excluded_columns and not col.endswith("_Homework")}

        # ✅ Handle subject-wise homework
        homework_columns = [col for col in df.columns if col.endswith("_Homework")]
        homework_list = []

        for col in homework_columns:
            val = student_data.get(col)
            if pd.notna(val) and str(val).strip():
                subject = col.replace("_Homework", "")
                homework_list.append(f"{subject}: {val}")

        if not homework_list:
            homework_list = ["No homework assigned."]

        # ✅ Return the template with student data
        return render_template(
            "teacher_dashboard.html",
            student_data={
                "Name": student_data["Name"],
                "Roll Number": student_data["Roll Number"],
                "Class": student_data["Class"],
                "Marks": student_marks,
                "Homework": homework_list
            }
        )
    else:
        flash("No student found with the provided Roll Number and Class.", "danger")
        return redirect(url_for("teacher_dashboard"))
        

# ===================== DOWNLOAD STUDENT REPORT =====================

@app.route("/download_report")
def download_report():
    if "user" not in session or session["role"] != "Student":
        return redirect(url_for("student_login"))

    df = load_students() 
    roll_number = str(session["user"]).strip()  # Convert roll_number to string and strip whitespace
    student = df[df["Roll Number"].astype(str).str.strip() == roll_number]

    if student.empty:
        flash("Student data not found.", "error")
        return redirect(url_for("student_dashboard"))

    student = student.iloc[0]

    buffer = BytesIO()
    p = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    # Title and Header
    p.setFont("Helvetica-Bold", 16)
    p.drawCentredString(width / 2.0, height - 50, "Student Progress Report")

    p.setFont("Helvetica", 12)
    p.drawString(50, height - 100, f"Name: {student['Name']}")
    p.drawString(50, height - 120, f"Class: {student['Class']}")
    p.drawString(50, height - 140, f"Roll Number: {student['Roll Number']}")

    # Marks and averages
    y = height - 180
    p.setFont("Helvetica-Bold", 12)
    p.drawString(50, y, "Subject-wise Performance:")
    y -= 20

    p.setFont("Helvetica", 11)
    for col in df.columns:
        if "_test" in col or "_Average" in col:
            val = student.get(col, "N/A")
            p.drawString(60, y, f"{col}: {val}")
            y -= 15
            if y < 50:  # Add new page if space runs out
                p.showPage()
                y = height - 50

    # Homework
    homework_cols = [col for col in df.columns if col.endswith("_Homework")]
    if homework_cols:
        y -= 20
        p.setFont("Helvetica-Bold", 12)
        p.drawString(50, y, "Assigned Homework:")
        y -= 20
        p.setFont("Helvetica", 11)
        for col in homework_cols:
            val = student.get(col, "").strip()
            if val:
                subject = col.replace("_Homework", "")
                p.drawString(60, y, f"{subject}: {val}")
                y -= 15
                if y < 50:
                    p.showPage()
                    y = height - 50

    p.save()
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{student['Roll Number']}_report.pdf",
        mimetype='application/pdf'
    )



def calculate_subject_average(row, subject):
    test_cols = [f"{subject}_test1", f"{subject}_test2", f"{subject}_test3"]
    scores = [pd.to_numeric(row[col], errors='coerce') for col in test_cols]
    scores = [s for s in scores if pd.notna(s)]
    if scores:
        return sum(scores) / len(scores)
    return ""


def update_subject_averages(df):
    subjects = set()
    for col in df.columns:
        if "_test" in col:
            subject = col.split("_test")[0]
            subjects.add(subject)

    for subject in subjects:
        test_cols = [col for col in df.columns if col.startswith(subject + "_test")]
        if test_cols:
            df[f"{subject}_Average"] = df[test_cols].mean(axis=1, skipna=True)

    return df



@app.route("/send_reports_form")
def send_reports_form():
    """Renders the form to select a class for sending reports."""
    if "user" not in session or session["role"] != "Teacher":
        return redirect(url_for("teacher_login"))

    df = load_students()
    classes = []
    if df is not None and 'Class' in df.columns:
        classes = df['Class'].unique().tolist()
    return render_template("send_reports_form.html", classes=classes)

@app.route("/send_reports", methods=["POST"])
def send_reports():
    """Handles the sending of class reports."""
    if "user" not in session or session["role"] != "Teacher":
        return redirect(url_for("teacher_login"))

    class_name = request.form.get("class_name")
    if class_name:
        report = send_class_reports(class_name)
        session['report_message'] = report
        return redirect(url_for("teacher_dashboard"))
    else:
        flash("Please select a class to send reports.", "danger")
        return redirect(url_for("teacher_dashboard"))
    
def send_class_reports(class_name):
    """Sends academic reports to parents of students in a specified class."""
    df = load_students()
    if df is None:
        return "Error: students.xlsx not found."

    class_students = df[df['Class'].str.strip() == class_name.strip()]
    if class_students.empty:
        return f"No students found for class: {class_name}"

    successful_sends = 0
    failed_sends = []

    for index, row in class_students.iterrows():
        child_name = row.get('Name', 'N/A')
        parent_email = row.get('Email', None)
        if not parent_email or pd.isna(parent_email):
            failed_sends.append((child_name, "No Email Found", "Email address is missing."))
            continue

        # Create dictionary of subjects and associated test columns
        subjects = {}
        for col in row.index:
            if '_test' in col:
                subject = col.split('_test')[0]
                subjects.setdefault(subject, []).append(col)
            elif '_Average' in col:
                subject = col.replace('_Average', '')
                subjects.setdefault(subject, []).append(col)
            elif '_Homework' in col:
                subject = col.replace('_Homework', '')
                subjects.setdefault(subject, []).append(col)

        # Compose the email content
        body = f"""Dear Parent,

Here is the academic report for {child_name} from class {class_name}:\n"""

        for subject, cols in subjects.items():
            body += f"\nSubject: {subject}\n"
            for col in sorted(cols):
                value = row.get(col, 'N/A')
                body += f"  {col}: {value}\n"

        body += "\nSincerely,\nThe School Administration"

        # Send email
        msg = MIMEText(body)
        msg['Subject'] = f"{class_name} Student Report - {child_name}"
        msg['From'] = SENDER_EMAIL
        msg['To'] = parent_email

        try:
            with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.sendmail(SENDER_EMAIL, parent_email, msg.as_string())
            print(f"Email sent to {parent_email} for {child_name} ({class_name})")
            successful_sends += 1
        except Exception as e:
            print(f"Error sending email to {parent_email} for {child_name} ({class_name}): {e}")
            failed_sends.append((child_name, parent_email, str(e)))

    report_message = f"Successfully sent {successful_sends} emails for class {class_name}."
    if failed_sends:
        report_message += "\nFailed to send emails to the following:\n"
        for name, email, error in failed_sends:
            report_message += f"- {name} ({email}): {error}\n"

    return report_message


def load_students():
    if not os.path.exists(STUDENTS_DB):
        print(f"Error: {STUDENTS_DB} not found!")
        return None
    try:
        df = pd.read_excel(STUDENTS_DB)
        # Explicitly convert the 'Class' column to string
        df['Class'] = df['Class'].astype(str)
        return df
    except Exception as e:
        print(f"Error reading {STUDENTS_DB}: {e}")
        return None
# ===================== LOGOUT =====================
@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)
