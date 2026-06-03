
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
import base64, io, json, os, webbrowser, threading, time

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

LOGO_SMALL = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo_small_b64.txt')).read().strip()

SEC1 = ["Adopts objectives of syllabi and curricula","Selects content and prepares appropriate instructional materials & visual aids","Adopts appropriate teaching methods","Relates new lesson with students\u2019 previous knowledge and skills","Provides appropriate motivation","Presents and develops lessons","Conveys ideas clearly","Utilizes the art of questioning to develop higher level of thinking","Ensures students\u2019 participation","Addresses individual differences","Shows mastery of the subject matter","Diagnoses learner\u2019s needs","Evaluates learning outcomes","Differentiates instruction to suit students\u2019 needs","Uses a variety of effective instructional strategies and resources","Involves students in cooperative learning to enhance higher-order thinking skills","Uses a variety of formal and informal assessment strategies to guide instruction","Gives constructive and frequent feedback to students on their learning","Provides independent practice activities","Checks for understanding and adjusts instruction as needed","English oral communication skills","English written communication skills"]
SEC2 = ["Maintains a clean and orderly classroom","Maintains a classroom conducive to learning","Establishes clear expectations for classroom rules and procedures early in the school year and enforces them consistently and fairly","Maximizes instructional time.","Establishes a climate of trust and teamwork by being fair, caring, respectful, and enthusiastic","Respects students\u2019 diversity, including language, culture, race, gender and special needs","Cares about students as individuals and makes them feel valued","Promotes self-discipline","Manages disruptive behaviour, distractions and misconduct well.","Keeps criticisms fair, objective and minimal"]
SEC3 = ["Supports and participates in parent-teacher activities","Works well with the other teachers and administration","Has positive relationships with students individually and in groups","Provides a climate which opens up communication between the teacher and the parent","Uses discretion in handling confidential information and difficult situations","Informs administrators and/or appropriate personnel of school-related matters in a timely manner","Cooperates with parents in the best interest of students","Acknowledges rights of others to hold different views","Handles information about students\u2019 private lives in a proper manner","Has a positive disposition and attitude"]
SEC4 = ["Adheres to school policies and guidelines at all times","Exerts effort to improve one\u2019s knowledge and skills","Takes responsibility for his/her actions","Channels concerns to proper authorities and not in social media nor to other people who are not of concern","Observes professional ethics","Is punctual in reporting to school and in submitting requirements","Is not often absent; when absent, only for valid reasons","Willingly accepts and satisfactorily performs tasks or assignments assigned to him/her, like committee work for co-curricular activities","Does not participate in rumor mongering","Wears the prescribed uniform at all times.",None,"Honesty/Integrity","Initiative/Resourcefulness","Courtesy","Respect for Authority","Teamwork","Leadership","Stress Tolerance","Fairness/Justice","Good Grooming","Human Relations","Self Confidence"]

def no_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side in ["top","left","bottom","right"]:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "none")
        tcBorders.append(el)
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    tcPr.append(tcBorders)

def fp(para, align=WD_ALIGN_PARAGRAPH.LEFT):
    para.alignment = align
    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    # Remove any spacing added by Word
    pPr = para._p.get_or_add_pPr()
    cs = pPr.find(qn("w:contextualSpacing"))
    if cs is None:
        cs = OxmlElement("w:contextualSpacing")
        pPr.append(cs)

def ar(para, text, bold=None, italic=False, underline=False, size=12, color=None):
    run = para.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    if bold is not None: run.bold = bold
    if italic: run.italic = True
    if underline: run.underline = True
    if color: run.font.color.rgb = RGBColor.from_string(color)
    return run

def make_doc(data):
    ratings = data.get("ratings", {})
    info    = data.get("info", {})
    import os as _os, re as _re, copy as _copy
    template_path = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "template.docx")
    doc = Document(template_path)

    def fill_para(para, value_map):
        pos = 0
        for run in para.runs:
            if "_" in run.text:
                if pos in value_map:
                    v = str(value_map[pos])
                    run.text = _re.sub(r"_+", lambda m: v[:len(m.group())].ljust(len(m.group()),"_"), run.text, count=1)
                pos += 1

    name  = info.get("name","");  cls   = info.get("cls","")
    subj  = info.get("subj","");  dobs  = info.get("dateobs","")
    rater = info.get("rater",""); ptitle= info.get("ptitle","School Head")
    rdate = info.get("ratedate",""); recv= info.get("recv","")
    r1=info.get("r1",{}); r2=info.get("r2",{}); r3=info.get("r3",{}); r4=info.get("r4",{})
    total=info.get("total",""); tdesc=info.get("totalDesc","")

    fill_para(doc.paragraphs[5], {0: name, 1: cls})
    fill_para(doc.paragraphs[7], {0: subj, 1: dobs})
    for pi, r in [(11,r1),(13,r2),(15,r3),(18,r4)]:
        fill_para(doc.paragraphs[pi], {0:str(r.get("score","")),1:str(r.get("avg","")),2:str(r.get("desc","")),3:str(r.get("sub",""))})
    fill_para(doc.paragraphs[23], {0: rater, 1: rdate})
    for run in doc.paragraphs[24].runs:
        run.text = run.text.replace("School Head", ptitle)
    fill_para(doc.paragraphs[25], {0: recv})

    for sid, ti in [("s1",0),("s2",1),("s3",2),("s4",3)]:
        tbl = doc.tables[ti]; ri2 = 0
        for row_idx in range(1, len(tbl.rows)):
            row = tbl.rows[row_idx]
            if "Demonstrates" in row.cells[0].text: continue
            sel = int(ratings.get(sid,{}).get(str(ri2),0))
            for ci, v in enumerate([5,4,3,2,1]):
                cell = row.cells[ci+1]
                p_el = cell.paragraphs[0]._p
                for r in p_el.findall(qn("w:r")): p_el.remove(r)
                run = cell.paragraphs[0].add_run("\u2713" if sel==v else "")
                run.font.name = "Calibri"; run.font.size = Pt(12)
            ri2 += 1

    st = doc.tables[5]
    sum_data = [
        (r1.get("score",""),r1.get("avg",""),r1.get("sub",""),r1.get("desc","")),
        (r2.get("score",""),r2.get("avg",""),r2.get("sub",""),r2.get("desc","")),
        (r3.get("score",""),r3.get("avg",""),r3.get("sub",""),r3.get("desc","")),
        (r4.get("score",""),r4.get("avg",""),r4.get("sub",""),r4.get("desc","")),
        ("","",total,tdesc),
    ]
    for ri3, rd in enumerate(sum_data):
        row = st.rows[ri3+1]
        for ci3, val in enumerate(rd):
            cell = row.cells[ci3+1]
            p_el = cell.paragraphs[0]._p
            for r in p_el.findall(qn("w:r")): p_el.remove(r)
            run = cell.paragraphs[0].add_run(str(val))
            run.font.name = "Calibri"; run.font.size = Pt(12)
            if ri3==4: run.bold = True

    return doc


@app.route("/")
def index():
    return send_from_directory(".", "index.html")

@app.route("/<path:filename>")
def static_files(filename):
    return send_from_directory(".", filename)

@app.route("/generate-docx", methods=["POST","OPTIONS"])
def generate_docx():
    if request.method == "OPTIONS":
        return "", 204
    try:
        data = request.get_json()
        doc = make_doc(data)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document", as_attachment=True, download_name="GLA_Teacher_Performance_Rating.docx")
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

@app.route("/health")
def health():
    return jsonify({"status": "ok", "message": "GLA Word Server running"})

def open_browser():
    time.sleep(1.5)
    webbrowser.open("http://localhost:5050")

if __name__ == "__main__":
    print("=" * 50)
    print("  GLA Teacher Rating System")
    print("  Word Generation Server")
    print("=" * 50)
    print("")
    print("  Opening browser at http://localhost:5050")
    print("  Keep this window open while using the app.")
    print("")
    print("  Press Ctrl+C to stop the server.")
    print("")
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host="0.0.0.0", port=5050, debug=False)
