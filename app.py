
# ---------------- PDF Imports ----------------
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import pagesizes

# ---------------- Flask & Other Imports ----------------
from flask import Flask, render_template, request, redirect, url_for, send_file, session, flash
import os
import pandas as pd
import sqlite3
import math
import re
from werkzeug.security import generate_password_hash, check_password_hash
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ---------------- Email Config ----------------
EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# ---------------- App Setup ----------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev_secret_key")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "app.db")

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads", "excel_files")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")


os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ---------------- Database Init ----------------
def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            college_name TEXT,
            email TEXT UNIQUE,
            password TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS halls (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hall_id TEXT NOT NULL,
            hall_name TEXT NOT NULL,
            benches INTEGER NOT NULL,
            user_id INTEGER,
                UNIQUE(hall_id, user_id)
        )
    """)

    conn.commit()
    conn.close()

init_db()

def is_valid_college_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.(ac\.in|edu\.in|edu)$'
    return re.match(pattern, email)

# ---------------- Routes ----------------
@app.route("/")
def home():
    return redirect(url_for("login"))

# ---------------- Register ----------------
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        college = request.form.get("college")
        email = request.form.get("email")
        password = request.form.get("password")

        if not college or not email or not password:
            flash("All fields are required", "danger")
            return render_template("register.html")
        
        if not is_valid_college_email(email):
            flash("Please register using a valid college email (ac.in / edu.in / edu)", "danger")
            return render_template("register.html")

        hashed = generate_password_hash(password)

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("SELECT id FROM users WHERE email=?", (email,))
        if cur.fetchone():
            flash("Email already registered", "danger")
            conn.close()
            return render_template("register.html")

        cur.execute(
            "INSERT INTO users (college_name, email, password) VALUES (?, ?, ?)",
            (college, email, hashed)
        )
        conn.commit()
        conn.close()

        flash("Registration successful", "success")
        return redirect(url_for("login"))

    return render_template("register.html")

# ---------------- Login ----------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if "user_id" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()
        cur.execute("SELECT id, password FROM users WHERE email=?", (email,))
        user = cur.fetchone()
        conn.close()

        if user and check_password_hash(user[1], password):
            session["user_id"] = user[0]
            return redirect(url_for("dashboard"))

        flash("Invalid credentials", "danger")

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

#------Dashboard ----------------
@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    if "user_id" not in session:
        return redirect(url_for("login"))

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    if request.method == "POST":
        hall_id = request.form.get("hall_id")
        hall_name = request.form.get("hall_name")
        benches = request.form.get("benches")

        if hall_id and hall_name and benches:
            try:
                cur.execute(
                "INSERT INTO halls (hall_id, hall_name, benches, user_id) VALUES (?, ?, ?, ?)",
                (hall_id, hall_name, int(benches), session["user_id"])
                )
                conn.commit()
                flash("Hall added successfully", "success")
            except sqlite3.IntegrityError:
                flash("Hall ID already exists", "danger")
            # finally:
            #     conn.close()

    cur.execute("SELECT * FROM halls WHERE user_id=?", (session["user_id"],))
    halls = cur.fetchall()
    conn.close()

    return render_template("dashboard.html", halls=halls)
# -------------------------delete hall-------------------------
@app.route("/delete_hall/<int:hall_id>")
def delete_hall(hall_id):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute(
        "DELETE FROM halls WHERE id=? AND user_id=?",
        (hall_id, session["user_id"])
    )
    conn.commit()
    conn.close()
    return redirect(url_for("dashboard"))

# -------------------------
# Upload
# -------------------------
@app.route("/upload", methods=["GET", "POST"])
def upload_excel():
    if "user_id" not in session:
        return redirect(url_for("login"))

    if request.method == "POST":
        students = request.files.get("students_file")
        invigilators = request.files.get("invigilators_file")

        if not students or not invigilators:
            flash("Upload both files", "danger")
            return redirect(url_for("upload_excel"))

        students.save(os.path.join(UPLOAD_FOLDER, f"{session['user_id']}_students.xlsx"))
        invigilators.save(os.path.join(UPLOAD_FOLDER, f"{session['user_id']}_invigilators.xlsx"))
        # print(df.columns)

        return redirect(url_for("allocate_page"))

    return render_template("upload.html")

# -------------------------
# Allocate Page
# -------------------------
@app.route("/allocate_page")
def allocate_page():
    if "user_id" not in session:
        return redirect(url_for("login"))

    students_path = os.path.join(UPLOAD_FOLDER, f"{session['user_id']}_students.xlsx")
    if not os.path.exists(students_path):
        return redirect(url_for("upload_excel"))

    df = pd.read_excel(students_path)
    total_students = len(df)

    branch_counts = {}
    if "Branch" in df.columns:
        branch_counts = df["Branch"].value_counts().to_dict()

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("SELECT * FROM halls WHERE user_id=?", (session["user_id"],))
    halls = cur.fetchall()
    conn.close()

    return render_template(
        "allocate.html",
        halls=halls,
        total_students=total_students,
        branch_counts=branch_counts
    )

# -------------------------
# Allocation Logic
# -------------------------
@app.route("/allocate", methods=["POST"])
def allocate():
    if "user_id" not in session:
        return redirect(url_for("login"))

    selected_halls = request.form.getlist("selected_halls")

    students_per_bench = request.form.get("students_per_bench")
    
    if not students_per_bench:
        flash("Select students per bench", "danger")
        return redirect(url_for("allocate_page"))
    students_per_bench=int(students_per_bench)
    
    if not selected_halls:
        flash("Select at least one hall", "danger")
        return redirect(url_for("allocate_page"))

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    halls = []
    for hid in selected_halls:
        cur.execute(
            "SELECT hall_name, benches FROM halls WHERE id=? AND user_id=?",
            (hid, session["user_id"])
        )
        hall = cur.fetchone()
        if hall:
            halls.append({
                "hall_name": hall[0],
                "benches": int(hall[1])
            })

    conn.close()

    students_df = pd.read_excel(
        os.path.join(UPLOAD_FOLDER, f"{session['user_id']}_students.xlsx")
    )
    inv_df = pd.read_excel(
        os.path.join(UPLOAD_FOLDER, f"{session['user_id']}_invigilators.xlsx")
    )

    # Clean column names
    students_df.columns = students_df.columns.str.strip().str.title()

    # -------------------------
    # Branch Interleaving Logic
    # -------------------------

    branch_groups = {}

    for branch in students_df["Branch"].unique():
        branch_groups[branch] = students_df[
            students_df["Branch"] == branch
        ].sort_values("Roll_No").reset_index(drop=True)

    interleaved_students = []

    while any(len(group) > 0 for group in branch_groups.values()):
        for branch in list(branch_groups.keys()):
            if len(branch_groups[branch]) > 0:
                interleaved_students.append(branch_groups[branch].iloc[0])
                branch_groups[branch] = branch_groups[branch].iloc[1:]

    students_df = pd.DataFrame(interleaved_students).reset_index(drop=True)

    total_students = len(students_df)
    total_capacity = sum(
        hall["benches"] * students_per_bench 
        for hall in halls
    )

    if total_capacity < total_students:
        flash("Not enough seating capacity", "danger")
        return redirect(url_for("allocate_page"))

    # -------------------------
    # Seating Allocation
    # -------------------------

    seating = []
    idx = 0

    for hall in halls:
        hall_name = hall["hall_name"]
        benches = hall["benches"]
        for bench_no in range(1, benches + 1):
            if idx >= total_students:
                break

            if students_per_bench == 1:
                s = students_df.iloc[idx]
                seating.append({
                    "Roll_No": s["Roll_No"],
                    "Name": s["Name"],
                    "Branch": str(s.get("Branch", "")).strip(),
                    "Section": s["Section"],
                    "Email": s["Email"],
                    "Hall": hall_name,
                    "Bench_No": bench_no,
                    "Position": "Center"
                })
                idx += 1

            else:
                for pos in ["Left", "Right"]:
                    if idx >= total_students:
                        break
                    s = students_df.iloc[idx]
                    seating.append({
                        "Roll_No": s["Roll_No"],
                        "Name": s["Name"],
                        "Branch": str(s.get("Branch", "")).strip(),
                        "Section": s["Section"],
                        "Email": s["Email"],
                        "Hall": hall_name,
                        "Bench_No": bench_no,
                        "Position": pos
                    })
                    idx += 1

    seating_df = pd.DataFrame(seating)

    
    
    #------------------PDF Generation Logic------------------
    # -------- PDF Generation --------
    for hall in seating_df["Hall"].unique():
        hall_students = seating_df[seating_df["Hall"] == hall]

        pdf_path = os.path.join(OUTPUT_FOLDER, f"{hall}_Seating.pdf")
        doc = SimpleDocTemplate(pdf_path, pagesize=pagesizes.A4)
        elements = []
        styles = getSampleStyleSheet()

        elements.append(Paragraph(f"Seating Arrangement - {hall}", styles["Heading1"]))
        elements.append(Spacer(1, 12))

        data = [["Roll No", "Name", "Branch", "Section", "Bench No", "Position"]]
        for _, row in hall_students.iterrows():
            data.append([row["Roll_No"], row["Name"], row["Branch"],
                         row["Section"], row["Bench_No"], row["Position"]])

        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ]))

        elements.append(table)
        doc.build(elements)

    # -------- Invigilator Logic --------
    students_per_hall = seating_df["Hall"].value_counts().to_dict()
    invigilator_allocation = []
    inv_df = inv_df.sample(frac=1).reset_index(drop=True)
    inv_index = 0

    for hall_name, student_count in students_per_hall.items():
        inv_needed = math.ceil(student_count / 30)

        for _ in range(inv_needed):
            if inv_index >= len(inv_df):
                break
            inv_row = inv_df.iloc[inv_index].to_dict()
            inv_row["Hall"] = hall_name
            invigilator_allocation.append(inv_row)
            inv_index += 1

    assigned_inv = pd.DataFrame(invigilator_allocation)

    # -------- Save Excel --------
    out_file = os.path.join(OUTPUT_FOLDER, f"{session['user_id']}_allocation.xlsx")
    with pd.ExcelWriter(out_file) as writer:
        seating_df.to_excel(writer, index=False, sheet_name="Seating")
        assigned_inv.to_excel(writer, index=False, sheet_name="Invigilators")

    return redirect(url_for("result_page"))
    
    
# -------------------------
# Download
# -------------------------

@app.route("/result")
def result_page():
    if "user_id" not in session:
        return redirect(url_for("login"))

    allocation_path = os.path.join(OUTPUT_FOLDER, f"{session['user_id']}_allocation.xlsx")

    if not os.path.exists(allocation_path):
        return redirect(url_for("allocate_page"))

    df = pd.read_excel(allocation_path, sheet_name="Invigilators")
    total_invigilators_needed = len(df)

    return render_template("result.html", total_invigilators_needed=total_invigilators_needed)


@app.route("/download")
def download():
    file_path = os.path.join(OUTPUT_FOLDER, f"{session['user_id']}_allocation.xlsx")
    return send_file(file_path, as_attachment=True)

@app.route("/send_student_emails")
def send_student_emails():
    if "user_id" not in session:
        return redirect(url_for("login"))

    try:
        # Load allocation file
        allocation_path = os.path.join(OUTPUT_FOLDER, f"{session['user_id']}_allocation.xlsx")
        df = pd.read_excel(allocation_path, sheet_name="Seating")

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        for _, row in df.iterrows():
            student_email = row.get("Email")
            if pd.isna(student_email):
                continue

            subject = "Exam Seating Allocation"
            body = f"""
Hello {row.get('Name')},

Your exam seating details:

Hall: {row.get('Hall')}
Bench Number: {row.get('Bench_No')}
Position: {row.get('Position')}

Please report 30 minutes early.

Best Regards,
Exam Cell
"""

            msg = MIMEMultipart()
            msg["From"] = EMAIL_ADDRESS
            msg["To"] = student_email
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))

            server.send_message(msg)

        server.quit()

        flash("Student emails sent successfully!", "success")

    except Exception as e:
        flash(f"Error sending student emails: {str(e)}", "danger")

    return redirect(url_for("result_page"))



@app.route("/send_invigilator_emails")
def send_invigilator_emails():
    if "user_id" not in session:
        return redirect(url_for("login"))

    try:
        allocation_path = os.path.join(OUTPUT_FOLDER, f"{session['user_id']}_allocation.xlsx")
        df = pd.read_excel(allocation_path, sheet_name="Invigilators")

        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        for _, row in df.iterrows():
            print("Sending email to:", row["Email"])
            print("Assigned Hall:", row["Hall"])
            print("--------------------------------")

            inv_email = row.get("Email")
            if pd.isna(inv_email):
                continue

            subject = "Invigilation Duty Assigned"
            body = f"""
Hello {row.get('Invigilator_Name')},

You have been assigned invigilation duty.

Hall: {row.get('Hall')}

Please report 20 minutes before exam time.

Exam Cell
"""

            msg = MIMEMultipart()
            msg["From"] = EMAIL_ADDRESS
            msg["To"] = inv_email
            msg["Subject"] = subject
            msg.attach(MIMEText(body, "plain"))

            server.send_message(msg)

        server.quit()

        flash("Invigilator emails sent successfully!", "success")

    except Exception as e:
        flash(f"Error sending invigilator emails: {str(e)}", "danger")

    return redirect(url_for("result_page"))
# -------------------------
# Run
# -------------------------
if __name__ == "__main__":
    app.run(debug=True)