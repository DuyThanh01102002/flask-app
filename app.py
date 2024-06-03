from flask import Flask, request, render_template, send_file
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
from datetime import datetime
app = Flask(__name__)

def replace_text_in_paragraph(paragraph, data):
    for key, value in data.items():
        if key in paragraph.text:
            inline = paragraph.runs
            for i in range(len(inline)):
                if key in inline[i].text:
                    if key == '[dob]':  # Chuyển văn bản thành chữ nghiêng nếu là [dob]
                        inline[i].italic = True
                    inline[i].text = inline[i].text.replace(key, value)
                    if key in ['[from]', '[to]']:  # Chuyển văn bản thành chữ in đậm nếu là [from] hoặc [to]
                        inline[i].bold = True
                    inline[i].text = inline[i].text.replace(key, value)
def format_dob(dob_str):
    try:
        if not dob_str:
            return ""
        # Tách các phần của ngày tháng năm
        parts = dob_str.split('/')
        if len(parts) == 3:
            day, month, year = parts
            day = int(day) if day else 1  # Nếu không có ngày, mặc định là 1
            month = int(month) if month else 1  # Nếu không có tháng, mặc định là 1
            year = int(year)
            dob_date = datetime(year, month, day)
            return f"{dob_date.day} tháng {dob_date.month} năm {dob_date.year}"
        elif len(parts) == 2:
            month, year = parts
            month = int(month) if month else 1  # Nếu không có tháng, mặc định là 1
            year = int(year)
            dob_date = datetime(year, month, 1)  # Nếu không có ngày, mặc định là ngày 1
            return f"   tháng {dob_date.month} năm {dob_date.year}"
        elif len(parts) == 1:
            year = int(parts[0])
            return f"   tháng   năm {year}"
        else:
            return dob_str
    except ValueError:
        return dob_str
def fill_template(template_path, output_path, data):
    doc = Document(template_path)

    sections = doc.sections
    for section in sections:
        section.left_margin = Inches(1.18)  # 3 cm
        section.right_margin = Inches(0.79) # 2 cm
        section.top_margin = Inches(0.79)   # 2 cm
        section.bottom_margin = Inches(0.79) # 2 cm
    # Xử lý các đoạn văn bản thông thường và các bảng
    if '[dob]' in data:
        data['[dob]'] = format_dob(data['[dob]'])
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, data)

    
    for table in doc.tables:
        for row in table.rows: 
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_paragraph(paragraph, data)
    


    # Lưu tài liệu đã điền dữ liệu
    doc.save(output_path)  
 
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        data = {
            '[wr]': request.form.get('wr'),
            '[dob]': request.form.get('dob'),
            '[send]': request.form.get('send'),
            '[cont]': request.form.get('cont'),
            '[from]': request.form.get('from'),
            '[to]': request.form.get('to')
        }
        cont_text = request.form.get('cont').replace('\r', '\n')  # Loại bỏ ký tự \r
        cont_lines = cont_text.split('\n')
        formatted_cont = cont_lines[0]
        if len(cont_lines) >= 1:
            for line in cont_lines[1:]:
                formatted_cont += line.strip() + '\t\n\t'
              # Thêm một dòng trống vào cuối
        else:
            formatted_cont += '\t'  # Nếu chỉ có một dòng, thêm tab vào cuối dòng
        formatted_cont = formatted_cont.replace('\t\n\t\t\n\t', '\t\n\t')
        data['[cont]'] = formatted_cont
        print(f"Received data: {data}")
        template_path = r'C:\Users\Amin\Documents\Zalo Received Files\test.docx'
        output_path = r'C:\Users\Amin\Documents\Zalo Received Files\output.docx'
        
        if not os.path.exists(template_path):
            return f"Template file not found at {template_path}"

        try:
            fill_template(template_path, output_path, data)
            print(f"File saved at {output_path}")
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            print(f"An error occurred: {e}")
            return f"An error occurred: {e}"
        
    return render_template('form.html')

if __name__ == '__main__':
    app.run(debug=True)
