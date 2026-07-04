# FairShare

FairShare is a small-scale web-based shared housing expense management system for students and housemates living in rented houses or apartments.

## Features

- User registration and login
- Create and manage houses
- Add and manage housemates
- Record shared bills
- Equal split and percentage split calculation
- Track paid and unpaid payments
- Upload receipt images
- Generate PDF and Excel reports
- Send payment reminder emails using Gmail settings

## Technologies Used

- Python
- Flask
- SQLite
- HTML
- Tailwind CSS
- CSS
- ReportLab
- XlsxWriter
- python-dotenv

## How to Run

1. Install Python 3.

2. Open the project folder in VS Code or terminal.

3. Create a virtual environment:

```bash
python -m venv venv
```

4. Activate the virtual environment:

```bash
venv\Scripts\activate
```

5. Install the required packages:

```bash
pip install -r requirements.txt
```

6. Complete the `.env` file.

7. Run the project:

```bash
python run.py
```

8. Open the browser and go to:

```text
http://127.0.0.1:5000
```