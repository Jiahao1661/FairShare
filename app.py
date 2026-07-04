from datetime import date, datetime
from functools import wraps
from io import BytesIO
from pathlib import Path
import os
import re
import uuid

from flask import (
    Flask,
    abort,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    session,
    url_for,
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
import xlsxwriter

from database import get_db

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")
RECEIPT_UPLOAD_DIR = BASE_DIR / "static" / "uploads" / "receipts"
HOUSE_IMAGE_UPLOAD_DIR = BASE_DIR / "static" / "uploads" / "houses"
ALLOWED_RECEIPT_EXTENSIONS = {"png", "jpg", "jpeg", "pdf"}
ALLOWED_HOUSE_IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "webp"}
MAX_UPLOAD_SIZE = 5 * 1024 * 1024

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get(
    "FAIRSHARE_SECRET_KEY", "fairshare-development-secret-key"
)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE




EMAIL_PATTERN = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$")
PHONE_PATTERN = re.compile(r"^\+?[0-9][0-9 ()-]{6,19}$")


def validate_optional_contact(email, phone):
    if email and (len(email) > 254 or not EMAIL_PATTERN.fullmatch(email)):
        return "Please enter a valid email address."
    if phone and not PHONE_PATTERN.fullmatch(phone):
        return "Phone number must contain 7 to 20 characters and only use numbers, spaces, +, - or parentheses."
    return None


def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if "user_id" not in session:
            flash("Please log in first.", "error")
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped_view


def get_owned_house(house_id):
    conn = get_db()
    house = conn.execute(
        "SELECT * FROM houses WHERE id = ? AND owner_id = ?",
        (house_id, session["user_id"]),
    ).fetchone()
    conn.close()
    return house


def get_owned_bill(bill_id):
    conn = get_db()
    bill = conn.execute(
        """
        SELECT b.*, h.house_name, h.house_emoji, h.house_color
        FROM bills b
        JOIN houses h ON h.id = b.house_id
        WHERE b.id = ? AND h.owner_id = ?
        """,
        (bill_id, session["user_id"]),
    ).fetchone()
    conn.close()
    return bill


def file_signature_matches(upload, extension):
    """Check the file's actual header instead of trusting its extension alone."""
    header = upload.stream.read(16)
    upload.stream.seek(0)
    signatures = {
        "png": lambda data: data.startswith(b"\x89PNG\r\n\x1a\n"),
        "jpg": lambda data: data.startswith(b"\xff\xd8\xff"),
        "jpeg": lambda data: data.startswith(b"\xff\xd8\xff"),
        "pdf": lambda data: data.startswith(b"%PDF-"),
        "webp": lambda data: len(data) >= 12 and data[:4] == b"RIFF" and data[8:12] == b"WEBP",
    }
    checker = signatures.get(extension)
    return bool(checker and checker(header))


def allowed_receipt(upload):
    if not upload or not upload.filename or "." not in upload.filename:
        return False
    extension = upload.filename.rsplit(".", 1)[1].lower()
    return (
        extension in ALLOWED_RECEIPT_EXTENSIONS
        and file_signature_matches(upload, extension)
    )


def save_receipt(upload):
    if not upload or not upload.filename:
        return None
    if not allowed_receipt(upload):
        raise ValueError("Receipt must be a valid PNG, JPG, JPEG or PDF file.")
    RECEIPT_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    ext = upload.filename.rsplit(".", 1)[1].lower()
    original = secure_filename(upload.filename.rsplit(".", 1)[0]) or "receipt"
    filename = f"{original}-{uuid.uuid4().hex[:10]}.{ext}"
    upload.save(RECEIPT_UPLOAD_DIR / filename)
    return filename


def remove_receipt_file(filename):
    if not filename:
        return
    path = RECEIPT_UPLOAD_DIR / filename
    if path.exists() and path.is_file():
        path.unlink()


def allowed_house_image(upload):
    if not upload or not upload.filename or "." not in upload.filename:
        return False
    extension = upload.filename.rsplit(".", 1)[1].lower()
    return (
        extension in ALLOWED_HOUSE_IMAGE_EXTENSIONS
        and file_signature_matches(upload, extension)
    )


def save_house_image(upload):
    if not upload or not upload.filename:
        return None
    if not allowed_house_image(upload):
        raise ValueError("House image must be a valid PNG, JPG, JPEG or WEBP file.")
    HOUSE_IMAGE_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    ext = upload.filename.rsplit(".", 1)[1].lower()
    original = secure_filename(upload.filename.rsplit(".", 1)[0]) or "house"
    filename = f"{original}-{uuid.uuid4().hex[:10]}.{ext}"
    upload.save(HOUSE_IMAGE_UPLOAD_DIR / filename)
    return filename


def remove_house_image_file(filename):
    if not filename:
        return
    path = HOUSE_IMAGE_UPLOAD_DIR / filename
    if path.exists() and path.is_file():
        path.unlink()


def due_status(due_date, all_paid=False):
    if all_paid:
        return "Paid", "green"
    if not due_date:
        return "No due date", "gray"
    try:
        due = datetime.strptime(due_date, "%Y-%m-%d").date()
    except ValueError:
        return "Invalid date", "gray"
    today = date.today()
    if due < today:
        return "Overdue", "red"
    if due == today:
        return "Due today", "yellow"
    days = (due - today).days
    return f"Due in {days} day{'s' if days != 1 else ''}", "blue"


@app.template_filter("format_due_date")
def format_due_date(value):
    if not value:
        return "Not set"
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%m/%d/%y")
    except (TypeError, ValueError):
        return str(value)


def collect_bill_form(conn, house_id):
    housemates = conn.execute(
        "SELECT * FROM housemates WHERE house_id = ? ORDER BY id", (house_id,)
    ).fetchall()
    if not housemates:
        raise ValueError("Please add at least one housemate before creating a bill.")

    month = request.form.get("month", "").strip()
    due_date_value = request.form.get("due_date", "").strip()
    split_type = request.form.get("split_type", "")

    try:
        rent = float(request.form.get("rent", ""))
        electricity = float(request.form.get("electricity", "0") or 0)
        water = float(request.form.get("water", "0") or 0)
        indah_water = float(request.form.get("indah_water", "0") or 0)
        internet = float(request.form.get("internet", "0") or 0)
        groceries = float(request.form.get("groceries", "0") or 0)
        other_expenses = float(request.form.get("other_expenses", "0") or 0)
    except ValueError as exc:
        raise ValueError("All expense amounts must be valid numbers.") from exc

    expenses = {
        "electricity": electricity,
        "water": water,
        "indah_water": indah_water,
        "internet": internet,
        "groceries": groceries,
        "other_expenses": other_expenses,
    }
    utilities = sum(expenses.values())

    if not month or rent < 0 or any(amount < 0 for amount in expenses.values()):
        raise ValueError("Please enter a valid month and non-negative amounts.")
    if due_date_value:
        try:
            datetime.strptime(due_date_value, "%Y-%m-%d")
        except ValueError as exc:
            raise ValueError("Please enter a valid due date.") from exc

    percentages = {}
    if split_type == "custom":
        total = 0.0
        for housemate in housemates:
            try:
                percentage = float(
                    request.form.get(f"percent_{housemate['id']}", "0")
                )
            except ValueError as exc:
                raise ValueError(
                    "Custom percentages must be valid non-negative numbers."
                ) from exc
            if percentage < 0:
                raise ValueError("Custom percentages cannot be negative.")
            percentages[housemate["id"]] = percentage
            total += percentage
        if abs(total - 100.0) > 0.01:
            raise ValueError("Custom percentages must add up to 100%.")
    elif split_type != "equal":
        raise ValueError("Please choose a valid split type.")

    return (
        housemates,
        month,
        due_date_value or None,
        rent,
        utilities,
        split_type,
        percentages,
        expenses,
    )


def rebuild_bill_splits(
    conn,
    bill_id,
    housemates,
    rent,
    utilities,
    split_type,
    percentages,
    paid_by_housemate=None,
):
    paid_by_housemate = paid_by_housemate or {}
    conn.execute("DELETE FROM bill_splits WHERE bill_id = ?", (bill_id,))
    utility_per_person = utilities / len(housemates)
    for housemate in housemates:
        if split_type == "equal":
            rent_amount = rent / len(housemates)
        else:
            rent_amount = rent * percentages[housemate["id"]] / 100.0
        total_amount = rent_amount + utility_per_person
        conn.execute(
            """
            INSERT INTO bill_splits
                (bill_id, housemate_id, rent_amount, utility_amount, total_amount, paid)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                bill_id,
                housemate["id"],
                rent_amount,
                utility_per_person,
                total_amount,
                paid_by_housemate.get(housemate["id"], 0),
            ),
        )


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")

        if not username or not email or not password:
            flash("All fields are required.", "error")
            return render_template("auth/register.html")
        if password != confirm_password:
            flash("Passwords do not match.", "error")
            return render_template("auth/register.html")
        if len(password) < 6:
            flash("Password must contain at least 6 characters.", "error")
            return render_template("auth/register.html")

        conn = get_db()
        existing = conn.execute(
            "SELECT id FROM users WHERE username = ? OR email = ?",
            (username, email),
        ).fetchone()
        if existing:
            conn.close()
            flash("Username or email already exists.", "error")
            return render_template("auth/register.html")

        conn.execute(
            "INSERT INTO users (username, email, password_hash) VALUES (?, ?, ?)",
            (username, email, generate_password_hash(password)),
        )
        conn.commit()
        conn.close()
        flash("Account created successfully. Please log in.", "success")
        return redirect(url_for("login"))

    return render_template("auth/register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        user = conn.execute(
            "SELECT * FROM users WHERE username = ?", (username,)
        ).fetchone()
        conn.close()
        if user and check_password_hash(user["password_hash"], password):
            session.clear()
            session["user_id"] = user["id"]
            session["username"] = user["username"]
            return redirect(url_for("dashboard"))
        flash("Incorrect username or password.", "error")
    return render_template("auth/login.html")


@app.route("/logout")
def logout():
    session.clear()
    flash("You have logged out.", "success")
    return redirect(url_for("home"))


@app.route("/dashboard")
@login_required
def dashboard():
    conn = get_db()
    houses = conn.execute(
        """
        SELECT h.*,
               (SELECT COUNT(*) FROM housemates hm WHERE hm.house_id = h.id) AS member_count,
               (SELECT COALESCE(SUM(b.rent + b.utilities), 0) FROM bills b WHERE b.house_id = h.id) AS bill_total,
               (SELECT COUNT(*) FROM bill_splits bs JOIN bills b ON b.id = bs.bill_id WHERE b.house_id = h.id) AS payment_count,
               (SELECT COUNT(*) FROM bill_splits bs JOIN bills b ON b.id = bs.bill_id WHERE b.house_id = h.id AND bs.paid = 1) AS paid_count
        FROM houses h
        WHERE h.owner_id = ?
        ORDER BY h.id DESC
        """,
        (session["user_id"],),
    ).fetchall()

    total_houses = len(houses)
    total_housemates = conn.execute(
        """
        SELECT COUNT(*) FROM housemates hm
        JOIN houses h ON h.id = hm.house_id
        WHERE h.owner_id = ?
        """,
        (session["user_id"],),
    ).fetchone()[0]
    total_bills = conn.execute(
        """
        SELECT COALESCE(SUM(b.rent + b.utilities), 0)
        FROM bills b
        JOIN houses h ON h.id = b.house_id
        WHERE h.owner_id = ?
        """,
        (session["user_id"],),
    ).fetchone()[0]

    outstanding_payments = conn.execute(
        """
        SELECT COUNT(*)
        FROM bill_splits bs
        JOIN bills b ON b.id = bs.bill_id
        JOIN houses h ON h.id = b.house_id
        WHERE h.owner_id = ? AND bs.paid = 0
        """,
        (session["user_id"],),
    ).fetchone()[0]

    current_month = date.today().strftime("%Y-%m")
    current_month_total = conn.execute(
        """
        SELECT COALESCE(SUM(b.rent + b.utilities), 0)
        FROM bills b
        JOIN houses h ON h.id = b.house_id
        WHERE h.owner_id = ? AND b.month = ?
        """,
        (session["user_id"], current_month),
    ).fetchone()[0]

    conn.close()

    return render_template(
        "dashboard/index.html",
        houses=houses,
        total_houses=total_houses,
        total_housemates=total_housemates,
        total_bills=total_bills,
        outstanding_payments=outstanding_payments,
        current_month=current_month,
        current_month_total=current_month_total,
    )


@app.route("/quick-calculator", methods=["GET", "POST"])
def quick_calculator():
    equal_result = None
    percentage_results = []
    split_method = request.form.get("split_method", "equal")

    if request.method == "POST":
        try:
            amount = float(request.form.get("amount", ""))
            if amount < 0:
                raise ValueError

            if split_method == "equal":
                people = int(request.form.get("people", ""))
                if people <= 0:
                    raise ValueError
                equal_result = amount / people

            elif split_method == "percentage":
                names = request.form.getlist("person_name[]")
                percentages = request.form.getlist("person_percentage[]")

                if not names or len(names) != len(percentages):
                    raise ValueError

                total_percentage = 0.0
                cleaned_people = []
                for index, (name, percentage_text) in enumerate(
                    zip(names, percentages), start=1
                ):
                    name = name.strip() or f"Person {index}"
                    percentage = float(percentage_text)
                    if percentage < 0:
                        raise ValueError
                    total_percentage += percentage
                    cleaned_people.append((name, percentage))

                if abs(total_percentage - 100.0) > 0.01:
                    flash("Percentages must add up to 100%.", "error")
                else:
                    percentage_results = [
                        {
                            "name": name,
                            "percentage": percentage,
                            "amount": amount * percentage / 100.0,
                        }
                        for name, percentage in cleaned_people
                    ]
            else:
                flash("Please choose a valid split method.", "error")

        except ValueError:
            flash("Please enter valid amounts, people and percentages.", "error")

    return render_template(
        "calculator.html",
        equal_result=equal_result,
        percentage_results=percentage_results,
        split_method=split_method,
    )


@app.route("/create-house", methods=["GET", "POST"])
@login_required
def create_house():
    if request.method == "POST":
        house_name = request.form.get("house_name", "").strip()
        house_emoji = "🏠"
        house_color = request.form.get("house_color", "cyan")
        try:
            house_image = save_house_image(request.files.get("house_image"))
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template("houses/create.html")
        if not house_name:
            flash("House name is required.", "error")
            return render_template("houses/create.html")
        conn = get_db()
        conn.execute(
            "INSERT INTO houses (house_name, house_emoji, house_color, house_image, owner_id) VALUES (?, ?, ?, ?, ?)",
            (house_name, house_emoji, house_color, house_image, session["user_id"]),
        )
        conn.commit()
        conn.close()
        flash("House created successfully.", "success")
        return redirect(url_for("dashboard"))
    return render_template("houses/create.html")


@app.route("/house/<int:house_id>/edit", methods=["GET", "POST"])
@login_required
def edit_house(house_id):
    house = get_owned_house(house_id)
    if house is None:
        flash("House not found.", "error")
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        name = request.form.get("house_name", "").strip()
        color = request.form.get("house_color", "cyan")
        if not name:
            flash("House name is required.", "error")
            return render_template("houses/edit.html", house=house)
        current_image = house["house_image"]
        try:
            new_image = save_house_image(request.files.get("house_image"))
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template("houses/edit.html", house=house)
        if request.form.get("remove_house_image") == "1":
            remove_house_image_file(current_image)
            current_image = None
        if new_image:
            remove_house_image_file(current_image)
            current_image = new_image
        conn = get_db()
        conn.execute(
            "UPDATE houses SET house_name = ?, house_color = ?, house_image = ? WHERE id = ? AND owner_id = ?",
            (name, color, current_image, house_id, session["user_id"]),
        )
        conn.commit()
        conn.close()
        flash("House updated.", "success")
        return redirect(url_for("dashboard"))
    return render_template("houses/edit.html", house=house)

@app.post("/delete_house/<int:house_id>")
@login_required
def delete_house(house_id):
    house = get_owned_house(house_id)

    if house is None:
        flash(
            "House not found or you do not have permission.",
            "error",
        )
        return redirect(url_for("dashboard"))

    conn = get_db()

    conn.execute(
        """
        DELETE FROM houses
        WHERE id = ? AND owner_id = ?
        """,
        (house_id, session["user_id"]),
    )

    conn.commit()
    conn.close()
    remove_house_image_file(house["house_image"])

    flash("House deleted successfully.", "success")
    return redirect(url_for("dashboard"))

@app.route("/house/<int:house_id>/members", methods=["GET", "POST"])
@login_required
def manage_housemates(house_id):
    house = get_owned_house(house_id)
    if house is None:
        flash("House not found.", "error")
        return redirect(url_for("dashboard"))

    conn = get_db()
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip().lower()
        phone = request.form.get("phone", "").strip()
        contact_error = validate_optional_contact(email, phone)
        if contact_error:
            conn.close()
            flash(contact_error, "error")
            return redirect(url_for("manage_housemates", house_id=house_id))
        if not name:
            conn.close()
            flash("Housemate name is required.", "error")
            return redirect(url_for("manage_housemates", house_id=house_id))
        conn.execute(
            "INSERT INTO housemates (house_id, name, email, phone) VALUES (?, ?, ?, ?)",
            (house_id, name, email or None, phone or None),
        )
        conn.commit()
        flash("Housemate added.", "success")

    housemates = conn.execute(
        "SELECT * FROM housemates WHERE house_id = ? ORDER BY id", (house_id,)
    ).fetchall()
    conn.close()
    return render_template("houses/housemates.html", house=house, housemates=housemates)


@app.route(
    "/house/<int:house_id>/members/<int:housemate_id>/edit",
    methods=["GET", "POST"],
)
@login_required
def edit_housemate(house_id, housemate_id):
    house = get_owned_house(house_id)
    if house is None:
        flash("House not found.", "error")
        return redirect(url_for("dashboard"))
    conn = get_db()
    housemate = conn.execute(
        "SELECT * FROM housemates WHERE id = ? AND house_id = ?",
        (housemate_id, house_id),
    ).fetchone()
    if housemate is None:
        conn.close()
        flash("Housemate not found.", "error")
        return redirect(url_for("manage_housemates", house_id=house_id))
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip().lower()
        phone = request.form.get("phone", "").strip()
        contact_error = validate_optional_contact(email, phone)
        if contact_error:
            conn.close()
            flash(contact_error, "error")
            return render_template(
                "houses/edit_housemate.html", house=house, housemate=housemate
            )
        if not name:
            conn.close()
            flash("Housemate name is required.", "error")
            return render_template(
                "houses/edit_housemate.html", house=house, housemate=housemate
            )
        conn.execute(
            "UPDATE housemates SET name = ?, email = ?, phone = ? WHERE id = ? AND house_id = ?",
            (name, email or None, phone or None, housemate_id, house_id),
        )
        conn.commit()
        conn.close()
        flash("Housemate updated.", "success")
        return redirect(url_for("manage_housemates", house_id=house_id))
    conn.close()
    return render_template(
        "houses/edit_housemate.html", house=house, housemate=housemate
    )


@app.post("/house/<int:house_id>/members/<int:housemate_id>/remove")
@login_required
def remove_housemate(house_id, housemate_id):
    house = get_owned_house(house_id)
    if house is None:
        flash("House not found.", "error")
        return redirect(url_for("dashboard"))
    conn = get_db()
    used_in_bill = conn.execute(
        "SELECT 1 FROM bill_splits WHERE housemate_id = ? LIMIT 1", (housemate_id,)
    ).fetchone()
    if used_in_bill:
        conn.close()
        flash("This member cannot be removed because bill history exists for them.", "error")
        return redirect(url_for("manage_housemates", house_id=house_id))
    conn.execute(
        "DELETE FROM housemates WHERE id = ? AND house_id = ?",
        (housemate_id, house_id),
    )
    conn.commit()
    conn.close()
    flash("Housemate removed.", "success")
    return redirect(url_for("manage_housemates", house_id=house_id))


@app.route("/house/<int:house_id>/add-bill", methods=["GET", "POST"])
@login_required
def add_bill(house_id):
    house = get_owned_house(house_id)
    if house is None:
        flash("House not found.", "error")
        return redirect(url_for("dashboard"))

    conn = get_db()
    housemates = conn.execute(
        "SELECT * FROM housemates WHERE house_id = ? ORDER BY id", (house_id,)
    ).fetchall()
    if not housemates:
        conn.close()
        flash("Please add at least one housemate before creating a bill.", "error")
        return redirect(url_for("manage_housemates", house_id=house_id))

    if request.method == "POST":
        try:
            (
                housemates,
                month,
                due_date_value,
                rent,
                utilities,
                split_type,
                percentages,
                expenses,
            ) = collect_bill_form(conn, house_id)
            receipt_filename = save_receipt(request.files.get("receipt"))
        except ValueError as exc:
            conn.close()
            flash(str(exc), "error")
            return render_template("bills/add.html", house=house, housemates=housemates)

        cursor = conn.execute(
            """
            INSERT INTO bills
                (
                    house_id, month, rent, utilities, electricity, water,
                    indah_water, internet, groceries, other_expenses,
                    split_type, due_date, receipt_filename
                )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                house_id,
                month,
                rent,
                utilities,
                expenses["electricity"],
                expenses["water"],
                expenses["indah_water"],
                expenses["internet"],
                expenses["groceries"],
                expenses["other_expenses"],
                split_type,
                due_date_value,
                receipt_filename,
            ),
        )
        bill_id = cursor.lastrowid
        rebuild_bill_splits(
            conn,
            bill_id,
            housemates,
            rent,
            utilities,
            split_type,
            percentages,
        )
        conn.commit()
        conn.close()
        flash("Bill calculated successfully.", "success")
        return redirect(url_for("split_result", bill_id=bill_id))

    conn.close()
    return render_template("bills/add.html", house=house, housemates=housemates)


@app.route("/bill/<int:bill_id>/edit", methods=["GET", "POST"])
@login_required
def edit_bill(bill_id):
    bill = get_owned_bill(bill_id)
    if bill is None:
        flash("Bill not found.", "error")
        return redirect(url_for("dashboard"))

    conn = get_db()
    housemates = conn.execute(
        "SELECT * FROM housemates WHERE house_id = ? ORDER BY id",
        (bill["house_id"],),
    ).fetchall()
    current_percentages = {}
    if bill["rent"]:
        rows = conn.execute(
            "SELECT housemate_id, rent_amount FROM bill_splits WHERE bill_id = ?",
            (bill_id,),
        ).fetchall()
        current_percentages = {
            row["housemate_id"]: round(row["rent_amount"] / bill["rent"] * 100, 2)
            for row in rows
        }

    if request.method == "POST":
        try:
            (
                housemates,
                month,
                due_date_value,
                rent,
                utilities,
                split_type,
                percentages,
                expenses,
            ) = collect_bill_form(conn, bill["house_id"])
            new_receipt = save_receipt(request.files.get("receipt"))
        except ValueError as exc:
            conn.close()
            flash(str(exc), "error")
            return render_template(
                "bills/edit.html",
                bill=bill,
                housemates=housemates,
                current_percentages=current_percentages,
            )

        paid_rows = conn.execute(
            "SELECT housemate_id, paid FROM bill_splits WHERE bill_id = ?", (bill_id,)
        ).fetchall()
        paid_by_housemate = {row["housemate_id"]: row["paid"] for row in paid_rows}

        receipt_filename = bill["receipt_filename"]
        if request.form.get("remove_receipt") == "1":
            remove_receipt_file(receipt_filename)
            receipt_filename = None
        if new_receipt:
            remove_receipt_file(receipt_filename)
            receipt_filename = new_receipt

        conn.execute(
            """
            UPDATE bills SET
                month = ?,
                rent = ?,
                utilities = ?,
                electricity = ?,
                water = ?,
                indah_water = ?,
                internet = ?,
                groceries = ?,
                other_expenses = ?,
                split_type = ?,
                due_date = ?,
                receipt_filename = ?
            WHERE id = ?
            """,
            (
                month,
                rent,
                utilities,
                expenses["electricity"],
                expenses["water"],
                expenses["indah_water"],
                expenses["internet"],
                expenses["groceries"],
                expenses["other_expenses"],
                split_type,
                due_date_value,
                receipt_filename,
                bill_id,
            ),
        )
        rebuild_bill_splits(
            conn,
            bill_id,
            housemates,
            rent,
            utilities,
            split_type,
            percentages,
            paid_by_housemate,
        )
        conn.commit()
        conn.close()
        flash("Bill updated successfully.", "success")
        return redirect(url_for("split_result", bill_id=bill_id))

    conn.close()
    return render_template(
        "bills/edit.html",
        bill=bill,
        housemates=housemates,
        current_percentages=current_percentages,
    )


@app.route("/bill/<int:bill_id>")
@login_required
def split_result(bill_id):
    bill = get_owned_bill(bill_id)
    if bill is None:
        flash("Bill not found.", "error")
        return redirect(url_for("dashboard"))
    conn = get_db()
    splits = conn.execute(
        """
        SELECT bs.id, hm.name, hm.email, bs.rent_amount, bs.utility_amount,
               bs.total_amount, bs.paid
        FROM bill_splits bs
        JOIN housemates hm ON hm.id = bs.housemate_id
        WHERE bs.bill_id = ?
        ORDER BY hm.name
        """,
        (bill_id,),
    ).fetchall()
    conn.close()
    paid_count = sum(1 for split in splits if split["paid"])
    total_count = len(splits)
    paid_percentage = round((paid_count / total_count) * 100) if total_count else 0
    status_text, status_color = due_status(
        bill["due_date"], paid_count == total_count and total_count > 0
    )
    return render_template(
        "bills/result.html",
        bill=bill,
        splits=splits,
        paid_count=paid_count,
        total_count=total_count,
        paid_percentage=paid_percentage,
        status_text=status_text,
        status_color=status_color,
    )


@app.post("/split/<int:split_id>/toggle-paid")
@login_required
def toggle_paid(split_id):
    conn = get_db()
    split = conn.execute(
        """
        SELECT bs.id, bs.bill_id, bs.paid
        FROM bill_splits bs
        JOIN bills b ON b.id = bs.bill_id
        JOIN houses h ON h.id = b.house_id
        WHERE bs.id = ? AND h.owner_id = ?
        """,
        (split_id, session["user_id"]),
    ).fetchone()
    if split is None:
        conn.close()
        flash("Payment record not found.", "error")
        return redirect(url_for("dashboard"))
    conn.execute(
        "UPDATE bill_splits SET paid = ? WHERE id = ?",
        (0 if split["paid"] else 1, split_id),
    )
    conn.commit()
    conn.close()
    flash("Payment status updated.", "success")
    return redirect(url_for("split_result", bill_id=split["bill_id"]))


@app.route("/house/<int:house_id>/history")
@login_required
def bill_history(house_id):
    house = get_owned_house(house_id)
    if house is None:
        flash("House not found.", "error")
        return redirect(url_for("dashboard"))
    conn = get_db()
    bills = conn.execute(
        """
        SELECT b.*, (b.rent + b.utilities) AS grand_total,
               (SELECT COUNT(*) FROM bill_splits bs WHERE bs.bill_id = b.id) AS split_count,
               (SELECT COUNT(*) FROM bill_splits bs WHERE bs.bill_id = b.id AND bs.paid = 1) AS paid_count
        FROM bills b
        WHERE b.house_id = ?
        ORDER BY b.month DESC, b.id DESC
        """,
        (house_id,),
    ).fetchall()
    conn.close()
    decorated = []
    for bill in bills:
        all_paid = bill["split_count"] > 0 and bill["paid_count"] == bill["split_count"]
        status_text, status_color = due_status(bill["due_date"], all_paid)
        decorated.append((bill, status_text, status_color))
    return render_template(
        "bills/history.html", house=house, decorated_bills=decorated
    )


@app.route("/receipt/<path:filename>")
@login_required
def receipt_file(filename):
    safe = secure_filename(filename)
    if safe != filename:
        abort(404)
    conn = get_db()
    allowed = conn.execute(
        """
        SELECT 1 FROM bills b
        JOIN houses h ON h.id = b.house_id
        WHERE b.receipt_filename = ? AND h.owner_id = ?
        LIMIT 1
        """,
        (filename, session["user_id"]),
    ).fetchone()
    conn.close()
    if not allowed:
        abort(404)
    return send_from_directory(RECEIPT_UPLOAD_DIR, filename, as_attachment=False)


@app.post("/bill/<int:bill_id>/delete")
@login_required
def delete_bill(bill_id):
    bill = get_owned_bill(bill_id)
    if bill is None:
        flash("Bill not found.", "error")
        return redirect(url_for("dashboard"))
    conn = get_db()
    conn.execute("DELETE FROM bills WHERE id = ?", (bill_id,))
    conn.commit()
    conn.close()
    remove_receipt_file(bill["receipt_filename"])
    flash("Bill deleted.", "success")
    return redirect(url_for("bill_history", house_id=bill["house_id"]))


def house_report_data(house_id):
    house = get_owned_house(house_id)
    if house is None:
        return None, None, None
    conn = get_db()
    bills = conn.execute(
        """
        SELECT b.*, (b.rent + b.utilities) AS grand_total
        FROM bills b WHERE b.house_id = ? ORDER BY b.month DESC, b.id DESC
        """,
        (house_id,),
    ).fetchall()
    splits = conn.execute(
        """
        SELECT bs.bill_id, hm.name, bs.rent_amount, bs.utility_amount,
               bs.total_amount, bs.paid
        FROM bill_splits bs
        JOIN housemates hm ON hm.id = bs.housemate_id
        JOIN bills b ON b.id = bs.bill_id
        WHERE b.house_id = ?
        ORDER BY b.month DESC, hm.name
        """,
        (house_id,),
    ).fetchall()
    conn.close()
    grouped = {}
    for row in splits:
        grouped.setdefault(row["bill_id"], []).append(row)
    return house, bills, grouped


@app.route("/house/<int:house_id>/export/pdf")
@login_required
def export_house_pdf(house_id):
    house, bills, grouped = house_report_data(house_id)
    if house is None:
        flash("House not found.", "error")
        return redirect(url_for("dashboard"))

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=15 * mm,
        leftMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
    )
    styles = getSampleStyleSheet()
    story = [
        Paragraph(f"FairShare Report - {house['house_name']}", styles["Title"]),
        Paragraph(
            f"Generated on {date.today().strftime('%d %B %Y')}", styles["Normal"]
        ),
        Spacer(1, 8 * mm),
    ]

    if not bills:
        story.append(Paragraph("No bills recorded.", styles["Normal"]))
    for bill in bills:
        story.append(
            Paragraph(
                f"{bill['month']} - Total RM {bill['grand_total']:.2f}",
                styles["Heading2"],
            )
        )
        story.append(
            Paragraph(
                f"Rent RM {bill['rent']:.2f} | Monthly expenses RM {bill['utilities']:.2f} | Split: {bill['split_type'].title()} | Due: {format_due_date(bill['due_date'])}",
                styles["Normal"],
            )
        )
        story.append(
            Paragraph(
                "Expense breakdown: "
                f"Electricity RM {bill['electricity']:.2f}, "
                f"Water RM {bill['water']:.2f}, "
                f"Indah Water RM {bill['indah_water']:.2f}, "
                f"Internet RM {bill['internet']:.2f}, "
                f"Groceries RM {bill['groceries']:.2f}, "
                f"Other RM {bill['other_expenses']:.2f}",
                styles["Normal"],
            )
        )
        table_data = [["Housemate", "Rent", "Expenses", "Total", "Status"]]
        for split in grouped.get(bill["id"], []):
            table_data.append(
                [
                    split["name"],
                    f"RM {split['rent_amount']:.2f}",
                    f"RM {split['utility_amount']:.2f}",
                    f"RM {split['total_amount']:.2f}",
                    "Paid" if split["paid"] else "Unpaid",
                ]
            )
        table = Table(table_data, colWidths=[48 * mm, 28 * mm, 30 * mm, 30 * mm, 25 * mm])
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#22d3ee")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f3f4f6")]),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 7),
                    ("TOPPADDING", (0, 0), (-1, 0), 7),
                ]
            )
        )
        story.extend([Spacer(1, 3 * mm), table, Spacer(1, 7 * mm)])

    doc.build(story)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{secure_filename(house['house_name'])}-FairShare-report.pdf",
        mimetype="application/pdf",
    )


@app.route("/house/<int:house_id>/export/excel")
@login_required
def export_house_excel(house_id):
    house, bills, grouped = house_report_data(house_id)
    if house is None:
        flash("House not found.", "error")
        return redirect(url_for("dashboard"))

    buffer = BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})

    title_format = workbook.add_format(
        {
            "bold": True,
            "font_size": 18,
            "font_color": "#FFFFFF",
            "bg_color": "#7C3AED",
            "align": "center",
            "valign": "vcenter",
        }
    )
    subtitle_format = workbook.add_format(
        {
            "italic": True,
            "font_color": "#475569",
            "align": "center",
        }
    )
    header_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "#0F172A",
            "bg_color": "#F9A8D4",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        }
    )
    text_format = workbook.add_format({"border": 1})
    center_format = workbook.add_format({"border": 1, "align": "center"})
    money_format = workbook.add_format(
        {"border": 1, "num_format": '"RM" #,##0.00'}
    )
    total_format = workbook.add_format(
        {
            "bold": True,
            "border": 1,
            "bg_color": "#EDE9FE",
            "num_format": '"RM" #,##0.00',
        }
    )
    paid_format = workbook.add_format(
        {
            "border": 1,
            "align": "center",
            "font_color": "#166534",
            "bg_color": "#DCFCE7",
        }
    )
    unpaid_format = workbook.add_format(
        {
            "border": 1,
            "align": "center",
            "font_color": "#991B1B",
            "bg_color": "#FEE2E2",
        }
    )

    summary = workbook.add_worksheet("Bill Summary")
    summary.hide_gridlines(2)
    summary.set_column("A:A", 13)
    summary.set_column("B:B", 13)
    summary.set_column("C:H", 14)
    summary.set_column("I:J", 16)
    summary.merge_range("A1:J2", f"FairShare Report - {house['house_name']}", title_format)
    summary.merge_range(
        "A3:J3",
        f"Generated on {date.today().strftime('%d %B %Y')}",
        subtitle_format,
    )

    summary_headers = [
        "Month",
        "Due Date",
        "Rent",
        "Electricity",
        "Water",
        "Indah Water",
        "Internet",
        "Groceries",
        "Other",
        "Grand Total",
    ]
    summary.write_row(4, 0, summary_headers, header_format)

    row = 5
    for bill in bills:
        values = [
            bill["month"],
            format_due_date(bill["due_date"]),
            bill["rent"],
            bill["electricity"],
            bill["water"],
            bill["indah_water"],
            bill["internet"],
            bill["groceries"],
            bill["other_expenses"],
            bill["grand_total"],
        ]
        summary.write(row, 0, values[0], text_format)
        summary.write(row, 1, values[1], center_format)
        for col in range(2, 9):
            summary.write_number(row, col, float(values[col]), money_format)
        summary.write_number(row, 9, float(values[9]), total_format)
        row += 1

    if not bills:
        summary.merge_range("A6:J6", "No bills recorded.", center_format)
    else:
        summary.write(row, 8, "Overall Total", header_format)
        summary.write_formula(row, 9, f"=SUM(J6:J{row})", total_format)

    details = workbook.add_worksheet("Housemate Splits")
    details.hide_gridlines(2)
    details.set_column("A:A", 13)
    details.set_column("B:B", 22)
    details.set_column("C:E", 15)
    details.set_column("F:F", 12)
    details.merge_range("A1:F2", "Housemate Payment Details", title_format)
    details_headers = [
        "Month",
        "Housemate",
        "Rent Share",
        "Expense Share",
        "Total Share",
        "Status",
    ]
    details.write_row(4, 0, details_headers, header_format)

    detail_row = 5
    for bill in bills:
        for split in grouped.get(bill["id"], []):
            details.write(detail_row, 0, bill["month"], text_format)
            details.write(detail_row, 1, split["name"], text_format)
            details.write_number(
                detail_row, 2, float(split["rent_amount"]), money_format
            )
            details.write_number(
                detail_row, 3, float(split["utility_amount"]), money_format
            )
            details.write_number(
                detail_row, 4, float(split["total_amount"]), total_format
            )
            status = "Paid" if split["paid"] else "Unpaid"
            details.write(
                detail_row,
                5,
                status,
                paid_format if split["paid"] else unpaid_format,
            )
            detail_row += 1

    if detail_row == 5:
        details.merge_range("A6:F6", "No housemate payment records.", center_format)

    summary.freeze_panes(5, 0)
    details.freeze_panes(5, 0)
    summary.autofilter(4, 0, max(row - 1, 5), 9)
    if detail_row > 5:
        details.autofilter(4, 0, detail_row - 1, 5)

    workbook.close()
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=(
            f"{secure_filename(house['house_name'])}-FairShare-report.xlsx"
        ),
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )


@app.errorhandler(413)
def file_too_large(_error):
    flash("Uploaded file is too large. Maximum size is 5 MB.", "error")
    return redirect(request.referrer or url_for("dashboard"))

