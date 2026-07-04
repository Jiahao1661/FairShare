from app import app
from database import init_db
from reminder_service import start_reminder_thread


if __name__ == "__main__":
    init_db()
    start_reminder_thread()
    app.run(debug=True, use_reloader=False)
