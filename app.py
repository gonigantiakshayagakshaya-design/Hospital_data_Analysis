from flask import Flask, render_template, request, redirect, url_for, send_file
import mysql.connector
import os
from decimal import Decimal, InvalidOperation
from io import BytesIO

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4, landscape

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


app = Flask(__name__)
app.secret_key = "hospital_project_secret_key"


# --------------------------------
# DATABASE CONNECTION
# --------------------------------

def get_db_connection():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password=os.getenv("DB_PASSWORD", "aimlb5"),
        database="hospital_db"
    )


# --------------------------------
# GENERATE CHARTS
# --------------------------------

def generate_charts():

    os.makedirs("static", exist_ok=True)

    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        # Disease chart
        cursor.execute("""
            SELECT disease, COUNT(*)
            FROM patients
            WHERE disease IS NOT NULL AND disease != ''
            GROUP BY disease
        """)

        disease_data = cursor.fetchall()

        plt.figure(figsize=(8, 5))

        if disease_data:
            diseases = [row[0] for row in disease_data]
            counts = [row[1] for row in disease_data]

            plt.bar(diseases, counts)
            plt.xticks(rotation=30, ha="right")
        else:
            plt.text(0.5, 0.5, "No Patient Data",
                     ha="center", va="center")

        plt.title("Patients by Disease")
        plt.xlabel("Disease")
        plt.ylabel("Number of Patients")
        plt.tight_layout()
        plt.savefig("static/disease_chart.png")
        plt.close()

        # Department chart
        cursor.execute("""
            SELECT department, COUNT(*)
            FROM patients
            WHERE department IS NOT NULL AND department != ''
            GROUP BY department
        """)

        department_data = cursor.fetchall()

        plt.figure(figsize=(8, 5))

        if department_data:
            departments = [row[0] for row in department_data]
            counts = [row[1] for row in department_data]

            plt.bar(departments, counts)
            plt.xticks(rotation=30, ha="right")
        else:
            plt.text(0.5, 0.5, "No Patient Data",
                     ha="center", va="center")

        plt.title("Patients by Department")
        plt.xlabel("Department")
        plt.ylabel("Number of Patients")
        plt.tight_layout()
        plt.savefig("static/department_chart.png")
        plt.close()

    finally:
        cursor.close()
        conn.close()


# --------------------------------
# DASHBOARD
# --------------------------------

@app.route("/")
def home():

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    try:
        cursor.execute("""
            SELECT COUNT(*) AS total
            FROM patients
        """)
        total_patients = cursor.fetchone()["total"]

        cursor.execute("""
            SELECT COUNT(*) AS total
            FROM patients
            WHERE LOWER(gender) = 'male'
        """)
        male_patients = cursor.fetchone()["total"]

        cursor.execute("""
            SELECT COUNT(*) AS total
            FROM patients
            WHERE LOWER(gender) = 'female'
        """)
        female_patients = cursor.fetchone()["total"]

        cursor.execute("""
            SELECT COALESCE(SUM(total_amount), 0) AS revenue
            FROM bills
        """)
        total_revenue = cursor.fetchone()["revenue"]

        cursor.execute("""
            SELECT * FROM patients
            ORDER BY patient_id DESC
        """)
        patients = cursor.fetchall()

    finally:
        cursor.close()
        conn.close()

    generate_charts()

    return render_template(
        "index.html",
        patients=patients,
        total_patients=total_patients,
        male_patients=male_patients,
        female_patients=female_patients,
        total_revenue=total_revenue,
        searched=False,
        searched_patient=None
    )


# --------------------------------
# ADD PATIENT
# --------------------------------

@app.route("/add_patient", methods=["GET", "POST"])
@app.route("/add", methods=["GET", "POST"])
def add_patient():

    if request.method == "POST":

        name = request.form.get("name")
        age = request.form.get("age")
        gender = request.form.get("gender")
        disease = request.form.get("disease")
        department = request.form.get("department")
        bill = request.form.get("bill", "0")

        try:
            age = int(age)
            bill = Decimal(bill)

            if age <= 0 or bill < 0:
                raise ValueError

        except (ValueError, InvalidOperation):
            return "Please enter valid age and bill values."

        conn = get_db_connection()
        cursor = conn.cursor()

        try:
            cursor.execute("""
                INSERT INTO patients
                (name, age, gender, disease, department, bill)
                VALUES (%s, %s, %s, %s, %s, %s)
            """, (
                name, age, gender, disease,
                department, bill
            ))

            conn.commit()

        finally:
            cursor.close()
            conn.close()

        return redirect(url_for("home"))

    return render_template("add_patient.html")


# --------------------------------
# UPDATE PATIENT
# --------------------------------

@app.route("/update_patient", methods=["GET", "POST"])
@app.route("/update", methods=["GET", "POST"])
def update_patient():

    if request.method == "POST":

        patient_id = request.form.get("patient_id")
        name = request.form.get("name")
        age = request.form.get("age")
        gender = request.form.get("gender")
        disease = request.form.get("disease")
        department = request.form.get("department")
        bill = request.form.get("bill", "0")

        try:
            age = int(age)
            bill = Decimal(bill)

            if age <= 0 or bill < 0:
                raise ValueError

        except (ValueError, InvalidOperation):
            return "Please enter valid age and bill values."

        conn = get_db_connection()
        cursor = conn.cursor()

        try:
            cursor.execute("""
                UPDATE patients
                SET name=%s,
                    age=%s,
                    gender=%s,
                    disease=%s,
                    department=%s,
                    bill=%s
                WHERE patient_id=%s
            """, (
                name, age, gender, disease,
                department, bill, patient_id
            ))

            conn.commit()

        finally:
            cursor.close()
            conn.close()

        return redirect(url_for("home"))

    patient_id = request.args.get("patient_id")
    patient = None

    if patient_id:

        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)

        try:
            cursor.execute("""
                SELECT *
                FROM patients
                WHERE patient_id=%s
            """, (patient_id,))

            patient = cursor.fetchone()

        finally:
            cursor.close()
            conn.close()

    return render_template(
        "update_patient.html",
        patient=patient
    )


# --------------------------------
# DELETE PATIENT
# --------------------------------

@app.route("/delete_patient", methods=["GET", "POST"])
@app.route("/delete", methods=["GET", "POST"])
def delete_patient():

    if request.method == "POST":

        patient_id = request.form.get("patient_id")

        conn = get_db_connection()
        cursor = conn.cursor()

        try:
            # Delete bills first
            cursor.execute("""
                DELETE FROM bills
                WHERE patient_id=%s
            """, (patient_id,))

            # Delete patient
            cursor.execute("""
                DELETE FROM patients
                WHERE patient_id=%s
            """, (patient_id,))

            conn.commit()

        finally:
            cursor.close()
            conn.close()

        return redirect(url_for("home"))

    return render_template("delete_patient.html")


# --------------------------------
# SEARCH PATIENT
# --------------------------------

@app.route("/search", methods=["GET", "POST"])
def search():

    search_value = request.form.get("search", "").strip()

    patient_id_value = request.form.get("patient_id", "").strip()

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    try:
        # Search by exact Patient ID if that field is submitted
        if patient_id_value:

            cursor.execute("""
                SELECT *
                FROM patients
                WHERE patient_id=%s
            """, (patient_id_value,))

            searched_patient = cursor.fetchone()

            if searched_patient:
                patients = [searched_patient]
            else:
                patients = []

            searched = True

        else:
            # Search by ID, name, disease or department
            cursor.execute("""
                SELECT *
                FROM patients
                WHERE CAST(patient_id AS CHAR) LIKE %s
                   OR name LIKE %s
                   OR disease LIKE %s
                   OR department LIKE %s
                ORDER BY patient_id DESC
            """, tuple(["%" + search_value + "%"] * 4))

            patients = cursor.fetchall()
            searched_patient = None
            searched = True

        # Dashboard statistics
        cursor.execute("""
            SELECT COUNT(*) AS total
            FROM patients
        """)
        total_patients = cursor.fetchone()["total"]

        cursor.execute("""
            SELECT COUNT(*) AS total
            FROM patients
            WHERE LOWER(gender)='male'
        """)
        male_patients = cursor.fetchone()["total"]

        cursor.execute("""
            SELECT COUNT(*) AS total
            FROM patients
            WHERE LOWER(gender)='female'
        """)
        female_patients = cursor.fetchone()["total"]

        cursor.execute("""
            SELECT COALESCE(SUM(total_amount), 0) AS revenue
            FROM bills
        """)
        total_revenue = cursor.fetchone()["revenue"]

    finally:
        cursor.close()
        conn.close()

    generate_charts()

    return render_template(
        "index.html",
        patients=patients,
        total_patients=total_patients,
        male_patients=male_patients,
        female_patients=female_patients,
        total_revenue=total_revenue,
        searched=searched,
        searched_patient=searched_patient
    )








# --------------------------------
# DOWNLOAD PDF REPORT
# --------------------------------

@app.route("/download_pdf")
def download_pdf():

    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    try:
        cursor.execute("""
            SELECT *
            FROM patients
            ORDER BY patient_id
        """)

        patients = cursor.fetchall()

    finally:
        cursor.close()
        conn.close()

    output = BytesIO()

    document = SimpleDocTemplate(
        output,
        pagesize=landscape(A4)
    )

    styles = getSampleStyleSheet()
    elements = []

    elements.append(
        Paragraph(
            "Hospital Patient Report",
            styles["Title"]
        )
    )

    elements.append(Spacer(1, 20))

    data = [[
        "ID",
        "Name",
        "Age",
        "Gender",
        "Disease",
        "Department",
        "Bill"
    ]]

    for patient in patients:
        data.append([
            str(patient["patient_id"]),
            str(patient["name"]),
            str(patient["age"]),
            str(patient["gender"] or ""),
            str(patient["disease"] or ""),
            str(patient["department"] or ""),
            str(patient["bill"] or 0)
        ])

    table = Table(data, repeatRows=1)

    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.blue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("PADDING", (0, 0), (-1, -1), 6)
    ]))

    elements.append(table)

    document.build(elements)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Hospital_Patient_Report.pdf",
        mimetype="application/pdf"
    )


# --------------------------------
# RUN APPLICATION
# --------------------------------

if __name__ == "__main__":
    app.run(debug=True)