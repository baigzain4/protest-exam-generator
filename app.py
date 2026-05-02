import os, random, io, sqlite3
from flask import Flask, render_template, request, jsonify, send_file

try:
    from xhtml2pdf import pisa
    XHTML2PDF_AVAILABLE = True
except ImportError:
    XHTML2PDF_AVAILABLE = False
    print("xhtml2pdf not installed. Run: pip install xhtml2pdf")

app = Flask(__name__)
DB = os.path.join(os.path.dirname(__file__), "questions.db")

# ─── Helpers ─────────────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def fetch_questions(mcq_count, short_count, long_count, chapter="All"):
    conn = get_db()
    result = {}

    def query(table, col, count):
        if count <= 0:
            return []
        if chapter and chapter != "All":
            rows = conn.execute(
                f"SELECT * FROM {table} WHERE chapter=?", (chapter,)
            ).fetchall()
        else:
            rows = conn.execute(f"SELECT * FROM {table}").fetchall()
        rows = [dict(r) for r in rows]
        return random.sample(rows, min(count, len(rows)))

    mcqs  = query("mcqs",            "chapter", mcq_count)
    shorts = query("short_questions", "chapter", short_count)
    longs  = query("long_questions",  "chapter", long_count)
    conn.close()

    # Normalise keys to match existing frontend expectations
    for q in mcqs:
        q["Question"]       = q.pop("question", "")
        q["Option A"]       = q.pop("option_a", "")
        q["Option B"]       = q.pop("option_b", "")
        q["Option C"]       = q.pop("option_c", "")
        q["Option D"]       = q.pop("option_d", "")
        q["Correct Answer"] = q.pop("correct_answer", "")
        q["Chapter"]        = q.pop("chapter", "")
    for q in shorts:
        q["Question"] = q.pop("question", "")
        q["Chapter"]  = q.pop("chapter", "")
    for q in longs:
        q["Question"] = q.pop("question", "")
        q["Chapter"]  = q.pop("chapter", "")

    return {"mcqs": mcqs, "short_questions": shorts, "long_questions": longs}

def get_chapters():
    conn = get_db()
    rows = conn.execute("SELECT DISTINCT chapter FROM mcqs ORDER BY chapter").fetchall()
    conn.close()
    return [r["chapter"] for r in rows]

# ─── Routes ──────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    chapters = get_chapters()
    return render_template("index.html", chapters=chapters)

@app.route("/api/generate", methods=["POST"])
def generate():
    data    = request.json or {}
    chapter = data.get("chapter", "All")
    result  = fetch_questions(
        int(data.get("mcq_count",   0)),
        int(data.get("short_count", 0)),
        int(data.get("long_count",  0)),
        chapter=chapter,
    )
    return jsonify(result)

@app.route("/api/chapters", methods=["GET"])
def chapters_api():
    return jsonify(get_chapters())

@app.route("/api/export-pdf", methods=["POST"])
def export_pdf():
    if not XHTML2PDF_AVAILABLE:
        return jsonify({"error": "xhtml2pdf not installed. Run: pip install xhtml2pdf"}), 500

    data = request.json or {}
    html_content = data.get("html", "")
    if not html_content:
        return jsonify({"error": "No HTML content provided"}), 400

    pdf_buffer = io.BytesIO()
    pisa_status = pisa.CreatePDF(html_content, dest=pdf_buffer, encoding="utf-8")
    if pisa_status.err:
        return jsonify({"error": f"PDF generation failed: {pisa_status.err}"}), 500

    pdf_buffer.seek(0)
    return send_file(pdf_buffer, mimetype="application/pdf",
                     as_attachment=True, download_name="ProTest_Assessment.pdf")

if __name__ == "__main__":
    app.run(debug=True)
