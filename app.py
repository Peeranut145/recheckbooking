from flask import Flask, request, render_template, redirect, url_for
import pandas as pd
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'  # โฟลเดอร์สำหรับเก็บไฟล์ที่อัปโหลด

# สร้างโฟลเดอร์สำหรับเก็บไฟล์ หากไม่มี
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def upload_file():
    return '''
    <!doctype html>
    <title>ตรวจสอบ Booking ซ้ำ</title>
    <h1>อัปโหลดไฟล์ Excel</h1>
    <form action="/check_duplicates" method="post" enctype="multipart/form-data">
      <input type="file" name="file">
      <input type="submit" value="อัปโหลด">
    </form>
    '''

@app.route('/check_duplicates', methods=['POST'])
def check_duplicates():
    if 'file' not in request.files:
        return 'ไม่มีไฟล์ที่อัปโหลด'

    file = request.files['file']
    if file.filename == '':
        return 'ไม่ได้เลือกไฟล์'

    # บันทึกไฟล์ที่อัปโหลด
    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    # อ่านไฟล์ Excel ด้วย Pandas
    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        return f'เกิดข้อผิดพลาดในการอ่านไฟล์: {str(e)}'

    # ตรวจสอบค่าซ้ำในคอลัมน์ Man_VoucherNo
    if 'Man_VoucherNo' not in df.columns:
        return 'ไม่พบคอลัมน์ Man_VoucherNo ในไฟล์ Excel'

    duplicates = df[df.duplicated(subset='Man_VoucherNo', keep=False)]

    # แสดงผลลัพธ์
    if duplicates.empty:
        return '<h1>ไม่พบ Booking ซ้ำในไฟล์ที่อัปโหลด</h1>'
    else:
        # ส่งผลลัพธ์ออกเป็น HTML
        return f'''
        <h1>พบ Booking ซ้ำ</h1>
        {duplicates.to_html(index=False)}
        '''

if __name__ == '__main__':
    app.run(debug=True)
