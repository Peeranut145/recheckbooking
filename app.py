from flask import Flask, request, render_template
import pandas as pd
from werkzeug.utils import secure_filename
from io import BytesIO

app = Flask(__name__)

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

    # ใช้ BytesIO แทนการบันทึกไฟล์ในระบบไฟล์
    try:
        file_stream = BytesIO(file.read())
        df = pd.read_excel(file_stream)
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
