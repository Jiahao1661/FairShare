import os
import smtplib
import ssl
import threading
import time
from datetime import date, datetime
from email.message import EmailMessage

from database import get_db


def get_gmail_settings():
    address = os.environ.get("GMAIL_ADDRESS", "").strip()
    app_password = os.environ.get("GMAIL_APP_PASSWORD", "").replace(" ", "").strip()
    return address, app_password


def send_email(sender, app_password, recipient, subject, body):
    message = EmailMessage()
    message["From"] = sender
    message["To"] = recipient
    message["Subject"] = subject
    message.set_content(body)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender, app_password)
        server.send_message(message)


def check_and_send_reminders():
    sender, app_password = get_gmail_settings()
    if not sender or not app_password:
        print(
            "[FairShare reminders] Gmail is not configured. "
            "Set GMAIL_ADDRESS and GMAIL_APP_PASSWORD."
        )
        return

    try:
        days_before = max(0, int(os.environ.get("REMINDER_DAYS_BEFORE", "3")))
    except ValueError:
        days_before = 3

    today = date.today()
    conn = get_db()
    rows = conn.execute(
        """
        SELECT
            bs.id AS split_id,
            bs.total_amount,
            hm.name AS housemate_name,
            hm.email AS housemate_email,
            b.month,
            b.due_date,
            h.house_name
        FROM bill_splits bs
        JOIN housemates hm ON hm.id = bs.housemate_id
        JOIN bills b ON b.id = bs.bill_id
        JOIN houses h ON h.id = b.house_id
        WHERE bs.paid = 0
          AND hm.email IS NOT NULL
          AND TRIM(hm.email) != ''
          AND b.due_date IS NOT NULL
          AND b.due_date != ''
        ORDER BY b.due_date
        """
    ).fetchall()

    sent_count = 0
    for row in rows:
        try:
            due = datetime.strptime(row["due_date"], "%Y-%m-%d").date()
        except ValueError:
            continue

        days_left = (due - today).days
        if days_left > days_before:
            continue

        already_sent = conn.execute(
            """
            SELECT 1
            FROM reminder_logs
            WHERE split_id = ? AND sent_on = ?
            """,
            (row["split_id"], today.isoformat()),
        ).fetchone()
        if already_sent:
            continue

        if days_left < 0:
            overdue_days = abs(days_left)
            day_word = "day" if overdue_days == 1 else "days"
            timing = f"overdue by {overdue_days} {day_word}"
        elif days_left == 0:
            timing = "due today"
        else:
            day_word = "day" if days_left == 1 else "days"
            timing = f"due in {days_left} {day_word}"

        subject = f"FairShare reminder: RM {row['total_amount']:.2f} {timing}"
        body = f"""Hi {row['housemate_name']},

This is an automatic FairShare payment reminder.

House: {row['house_name']}
Bill month: {row['month']}
Amount due: RM {row['total_amount']:.2f}
Due date: {due.strftime('%m/%d/%y')}
Status: {timing}

Please make the payment and inform the house owner once it is completed.

If you have already made the payment, please ignore this message. Thank you.

Thank you,
FairShare
"""

        try:
            send_email(
                sender,
                app_password,
                row["housemate_email"],
                subject,
                body,
            )
        except Exception as exc:
            print(
                f"[FairShare reminders] Could not email "
                f"{row['housemate_email']}: {exc}"
            )
            continue

        conn.execute(
            "INSERT INTO reminder_logs (split_id, sent_on) VALUES (?, ?)",
            (row["split_id"], today.isoformat()),
        )
        conn.commit()
        sent_count += 1

    conn.close()
    print(f"[FairShare reminders] Sent {sent_count} reminder(s).")


def reminder_loop():
    time.sleep(5)
    while True:
        try:
            check_and_send_reminders()
        except Exception as exc:
            print(f"[FairShare reminders] Unexpected error: {exc}")

        try:
            seconds = max(10, int(os.environ.get("REMINDER_CHECK_SECONDS", "3600")))
        except ValueError:
            seconds = 3600
        time.sleep(seconds)


def start_reminder_thread():
    thread = threading.Thread(
        target=reminder_loop,
        name="fairshare-reminder-worker",
        daemon=True,
    )
    thread.start()
    return thread
