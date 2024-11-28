import openpyxl
from flask import Flask, request, jsonify

app = Flask(__name__)

# ฟังก์ชันเพื่อตรวจสอบข้อมูลที่ซ้ำจากไฟล์ Excel
def check_duplicates(file_path):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    values = [row[0].value for row in sheet.iter_rows(min_row=2)]  # อ่านค่าจากคอลัมน์แรก
    duplicates = [val for val in set(values) if values.count(val) > 1]
    return duplicates

# หน้าเว็บหลักสำหรับการอัปโหลดไฟล์
@app.route('/')
def index():
    return '''
        <h1>อัปโหลดไฟล์ Excel เพื่อตรวจสอบข้อมูลที่ซ้ำ</h1>
        <form method="POST" action="/check_duplicates" enctype="multipart/form-data">
            <input type="file" name="file" required>
            <button type="submit">อัปโหลดไฟล์</button>
        </form>
    '''

# เส้นทางสำหรับการตรวจสอบข้อมูลที่ซ้ำ
@app.route('/check_duplicates', methods=['POST'])
def check_duplicates_route():
    try:
        file = request.files['file']
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(f'uploaded_{filename}')
            duplicates = check_duplicates(f'uploaded_{filename}')
            
            if duplicates:
                return jsonify({'duplicates': duplicates})
            else:
                return jsonify({'message': 'ไม่มีข้อมูลที่ซ้ำ'})
        else:
            return jsonify({'error': 'ไฟล์ไม่ถูกต้อง! กรุณาอัปโหลดไฟล์ .xlsx เท่านั้น.'})
    except Exception as e:
        return jsonify({'error': 'เกิดข้อผิดพลาดในการประมวลผลไฟล์:', 'details': str(e)})

if __name__ == "__main__":
    app.run(debug=True)
    
    
