"""
Student Results Assessment - Flask backend

Run:   pip install -r requirements.txt
       python make_sample_questions.py   (optional - creates a sample questions.xlsx)
       python app.py
Open:  http://127.0.0.1:5000

Files:
  questions.xlsx  -> input. Col1 question, Col2-5 options, Col6 difficulty (high/medium/low),
                     Col7 correct option (A/B/C/D, 1-4, or the option text). Row 1 = header.
  results.xlsx    -> output. Created automatically; one row per student per question,
                     plus a "Summary" sheet with one row per student.
"""
import os
import random
import threading
from datetime import datetime

from flask import Flask, jsonify, render_template, request
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
QUESTIONS_FILE = os.path.join(BASE_DIR, "questions.xlsx")
RESULTS_FILE = os.path.join(BASE_DIR, "results.xlsx")
NUM_QUESTIONS = 20
LETTERS = ["A", "B", "C", "D"]

app = Flask(__name__)
file_lock = threading.Lock()  # prevents two students writing to results.xlsx at once


# ---------------------------------------------------------------- questions
def resolve_correct(raw, options):
    """Turn the 7th column value into an index 0-3. Accepts A-D, 1-4, or option text."""
    if raw is None:
        return None
    s = str(raw).strip()
    if s.upper() in LETTERS:
        return LETTERS.index(s.upper())
    if s.isdigit() and 1 <= int(s) <= 4:
        return int(s) - 1
    if s.endswith(".0") and s[:-2].isdigit() and 1 <= int(s[:-2]) <= 4:
        return int(s[:-2]) - 1
    for i, opt in enumerate(options):
        if s.lower() == opt.lower():
            return i
    return None


def load_questions():
    """Read questions.xlsx fresh each time so you can replace the file without restarting."""
    if not os.path.exists(QUESTIONS_FILE):
        raise FileNotFoundError("questions.xlsx not found next to app.py")
    wb = load_workbook(QUESTIONS_FILE, read_only=True, data_only=True)
    ws = wb.active
    questions = []
    for row in ws.iter_rows(min_row=2, values_only=True):  # skip header
        if not row or row[0] is None or str(row[0]).strip() == "":
            continue
        row = list(row) + [None] * (7 - len(row))
        options = [("" if v is None else str(v).strip()) for v in row[1:5]]
        correct = resolve_correct(row[6], options)
        if correct is None:
            continue  # skip rows whose correct option can't be understood
        difficulty = str(row[5]).strip().lower() if row[5] else "medium"
        questions.append({
            "question": str(row[0]).strip(),
            "options": options,
            "difficulty": difficulty,
            "correct": correct,
        })
    wb.close()
    return questions


# ---------------------------------------------------------------- results
def ensure_results_file():
    if os.path.exists(RESULTS_FILE):
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Responses"
    ws.append(["Submitted At", "Student Name", "Class", "Question No", "Question",
               "Selected Answer", "Correct Answer", "Result",
               "Time Taken (sec)", "Difficulty"])
    summ = wb.create_sheet("Summary")
    summ.append(["Submitted At", "Student Name", "Class", "Score", "Total Questions",
                 "Percentage", "Total Time (sec)"])
    header_fill = PatternFill("solid", start_color="1F3A5F")
    for sheet, widths in ((ws, (20, 20, 10, 12, 55, 28, 28, 10, 16, 12)),
                          (summ, (20, 20, 10, 8, 16, 12, 16))):
        for c in sheet[1]:
            c.font = Font(name="Arial", bold=True, color="FFFFFF")
            c.fill = header_fill
            c.alignment = Alignment(horizontal="center")
        for i, w in enumerate(widths):
            sheet.column_dimensions[chr(65 + i)].width = w
        sheet.freeze_panes = "A2"
    wb.save(RESULTS_FILE)


def save_result(name, cls, details):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with file_lock:
        ensure_results_file()
        wb = load_workbook(RESULTS_FILE)
        ws, summ = wb["Responses"], wb["Summary"]
        score = 0
        total_time = 0.0
        for d in details:
            ws.append([now, name, cls, d["no"], d["question"], d["selected"],
                       d["correct_answer"], d["result"], d["time"], d["difficulty"].capitalize()])
            score += d["result"] == "Correct"
            total_time += d["time"]
        pct = round(score / len(details) * 100, 1) if details else 0
        summ.append([now, name, cls, score, len(details), pct, round(total_time, 2)])
        wb.save(RESULTS_FILE)
    return score, pct


# ---------------------------------------------------------------- routes
@app.route("/")
def index():
    return render_template("index.html", total=NUM_QUESTIONS)


@app.route("/api/questions")
def api_questions():
    try:
        qs = load_questions()
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    if len(qs) > NUM_QUESTIONS:
        qs = random.sample(qs, NUM_QUESTIONS)  # pick 20 if the sheet has more
    # Send the question text (not the answer). The client sends text back; we regrade on the server.
    return jsonify({"questions": [
        {"id": q["question"], "question": q["question"], "options": q["options"]} for q in qs
    ]})


@app.route("/api/submit", methods=["POST"])
def api_submit():
    data = request.get_json(silent=True) or {}
    name = str(data.get("name", "")).strip()
    cls = str(data.get("class", "")).strip()
    answers = data.get("answers", [])
    if not name or not cls:
        return jsonify({"error": "Name and class are required."}), 400
    if not isinstance(answers, list) or not answers:
        return jsonify({"error": "No answers received."}), 400

    bank = {q["question"]: q for q in load_questions()}
    details = []
    for i, a in enumerate(answers, start=1):
        q = bank.get(a.get("id"))
        if q is None:
            continue
        sel = a.get("selected")  # index 0-3, or null if skipped
        try:
            t = max(0.0, min(float(a.get("time", 0)), 3600.0))
        except (TypeError, ValueError):
            t = 0.0
        if isinstance(sel, int) and 0 <= sel < 4:
            selected_txt = f"{LETTERS[sel]}. {q['options'][sel]}"
            result = "Correct" if sel == q["correct"] else "Wrong"
        else:
            selected_txt, result = "(not answered)", "Wrong"
        details.append({
            "no": i, "question": q["question"], "selected": selected_txt,
            "correct_answer": f"{LETTERS[q['correct']]}. {q['options'][q['correct']]}",
            "result": result, "time": round(t, 2), "difficulty": q["difficulty"],
        })
    if not details:
        return jsonify({"error": "Answers did not match any question."}), 400

    score, pct = save_result(name, cls, details)
    return jsonify({"score": score, "total": len(details), "percentage": pct})



if __name__ == "__main__":
    app.run(
        host="127.0.0.2",
        port=5000,
        debug=False,
        use_reloader=False
    )
