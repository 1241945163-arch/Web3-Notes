import os
import re
import uuid
import pythoncom
import win32com.client
from datetime import datetime
from flask import Flask, request, render_template, send_file, after_this_request
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ==================== WPS/Word 转 PDF 引擎 ====================
def wps_to_pdf(docx_path, pdf_path):
    pythoncom.CoInitialize()
    word = doc = None
    try:
        try: word = win32com.client.DispatchEx("KWPS.Application")
        except: word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        abs_docx, abs_pdf = os.path.abspath(docx_path), os.path.abspath(pdf_path)
        doc = word.Documents.Open(abs_docx)
        doc.SaveAs(abs_pdf, FileFormat=17)
    except Exception as e:
        raise Exception(f"PDF转换失败(请确保电脑已安装WPS或Office): {str(e)}")
    finally:
        if doc: doc.Close(0)
        if word: word.Quit()
        pythoncom.CoUninitialize()

# ==================== 极速正则解析引擎 ====================
def fast_regex_parse(text):
    text = re.sub(r'[（\(].*?(标题|第[一二三四五六七八九十\d]+行|居中|对齐|字体|字号|楷体|黑体|仿宋|红字|落款).*?[）\)]', '', text)
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    if not lines: return []
    
    plan = [{"align": "center", "indent": False, "runs": [{"text": lines[0], "font": "方正小标宋_GBK", "size": 22}]}]
    plan.append({"align": "left", "indent": False, "runs": [{"text": "", "font": "方正仿宋_GBK", "size": 16}]})
    
    for line in lines[1:]:
        para = {"align": "justify", "indent": True, "runs": []}
        m1 = re.match(r'^([一二三四五六七八九十]+、)(.*)$', line)
        m2 = re.match(r'^([（\(][一二三四五六七八九十]+[）\)][^。]*。?)(.*)$', line)
        m3 = re.match(r'^(\d+[.．])(.*)$', line)
        
        if m1:
            if '。' in line:
                parts = line.split('。', 1)
                para["runs"].append({"text": parts[0] + '。', "font": "方正黑体_GBK", "size": 16})
                if parts[1]: para["runs"].append({"text": parts[1], "font": "方正仿宋_GBK", "size": 16})
            else:
                para["runs"].append({"text": line, "font": "方正黑体_GBK", "size": 16})
        elif m2:
            if '。' in line:
                parts = line.split('。', 1)
                para["runs"].append({"text": parts[0] + '。', "font": "方正楷体_GBK", "size": 16})
                if parts[1]: para["runs"].append({"text": parts[1], "font": "方正仿宋_GBK", "size": 16})
            else:
                para["runs"].append({"text": line, "font": "方正楷体_GBK", "size": 16})
        elif m3:
            if '。' in line:
                parts = line.split('。', 1)
                para["runs"].append({"text": parts[0] + '。', "font": "方正仿宋_GBK", "size": 16, "bold": True})
                if parts[1]: para["runs"].append({"text": parts[1], "font": "方正仿宋_GBK", "size": 16})
            else:
                para["runs"].append({"text": line, "font": "方正仿宋_GBK", "size": 16, "bold": True})
        else:
            para["runs"].append({"text": line, "font": "方正仿宋_GBK", "size": 16})
        plan.append(para)
    return plan

# ==================== 渲染引擎 ====================
def add_border(paragraph, border_type="bottom"):
    pPr = paragraph._element.get_or_add_pPr()
    b = OxmlElement('w:pBdr')
    bd = OxmlElement(f'w:{border_type}')
    for k, v in [('w:val','single'),('w:sz','18'),('w:space','1'),('w:color','FF0000')]: bd.set(qn(k), v)
    b.append(bd)
    pPr.append(b)

def render_docx(json_plan, out_p, red_h=""):
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin, sec.bottom_margin, sec.left_margin, sec.right_margin = Mm(37), Mm(35), Mm(28), Mm(26)
    
    if red_h.strip():
        p = doc.add_paragraph(); p.alignment = 1
        r = p.add_run(red_h)
        r.font.name = 'Times New Roman'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '方正小标宋_GBK')
        r.font.size, r.font.color.rgb = Pt(36), RGBColor(255, 0, 0)
        add_border(p, "bottom")
        sec.different_first_page_header_footer = True
        add_border(sec.first_page_footer.paragraphs[0], "top")
        
    for data in json_plan:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(28.8)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        
        pPr = p._element.get_or_add_pPr()
        snap = OxmlElement('w:snapToGrid')
        snap.set(qn('w:val'), '0')
        pPr.append(snap)
        
        p.alignment = {"left":0, "center":1, "right":2, "justify":3}.get(data.get("align","justify"), 3)
        if data.get("indent"): p.paragraph_format.first_line_indent = Pt(32)
            
        for rd in data.get("runs", []):
            txt = rd.get("text", "")
            
            # 【核心修复】：直角引号强转弯引号
            txt = re.sub(r'"([^"]*)"', r'“\1”', txt)
            txt = re.sub(r"'([^']*)'", r'‘\1’', txt)
            
            if not txt: continue
            r = p.add_run(txt)
            r.font.name = 'Times New Roman'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), rd.get("font","方正仿宋_GBK"))
            
            # 【核心修复】：强制底层按中文格式渲染引号等标点
            r._element.rPr.rFonts.set(qn('w:hint'), 'eastAsia')
            
            r.font.size = Pt(rd.get("size", 16))
            r.bold = rd.get("bold", False)
            c = rd.get("color","000000").replace("#","")
            if c != "000000": r.font.color.rgb = RGBColor(int(c[:2],16), int(c[2:4],16), int(c[4:],16))
    doc.save(out_p)

# ==================== 路由逻辑 ====================
@app.route('/')
def index(): return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    rh = request.form.get('red_header', '').strip()
    fmt = request.form.get('format', 'Word')
    txt = request.form.get('text_content', '').strip()
    
    if not txt: return "正文内容不能为空", 400
        
    f_id = str(uuid.uuid4())
    docx_p = os.path.join(UPLOAD_FOLDER, f"out_{f_id}.docx")
    try:
        plan = fast_regex_parse(txt)
        render_docx(plan, docx_p, rh)
        final_p = docx_p
        
        if fmt == 'PDF':
            final_p = docx_p.replace(".docx", ".pdf")
            wps_to_pdf(docx_p, final_p)
            if os.path.exists(docx_p): os.remove(docx_p)
            
        @after_this_request
        def cl(res):
            try: os.remove(final_p)
            except: pass
            return res
            
        timestamp = datetime.now().strftime("%H时%M分%S秒")
        dl_name = f"公文_{timestamp}.{'pdf' if fmt=='PDF' else 'docx'}"
        
        return send_file(final_p, as_attachment=True, download_name=dl_name)
    except Exception as e: return str(e), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True, threaded=False)
