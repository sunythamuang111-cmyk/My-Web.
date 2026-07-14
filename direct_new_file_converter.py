import pandas as pd
import os

def convert_excel_to_csv():
    # กำหนดชื่อไฟล์ต้นทางและปลายทางให้อยู่ในโฟลเดอร์เดียวกัน
    input_file = 'ระบบบันทึกข้อมูลการลาออนไลน์.xlsx'
    output_file = 'ระบบบันทึกข้อมูลการลาออนไลน์.csv'
    
    # ตรวจสอบว่ามีไฟล์ต้นทางอยู่หรือไม่
    if not os.path.exists(input_file):
        print(f"ข้อผิดพลาด: ไม่พบไฟล์ '{input_file}' ในโฟลเดอร์ปัจจุบัน")
        return

    try:
        print(f"กำลังอ่านไฟล์: {input_file} ...")
        # อ่านข้อมูลจากไฟล์ Excel
        df = pd.read_excel(input_file)
        
        # แปลงและบันทึกเป็นไฟล์ CSV 
        # ใช้ encoding='utf-8-sig' เพื่อให้รองรับภาษาไทยเมื่อเปิดด้วยโปรแกรมอื่นๆ เช่น Excel
        df.to_csv(output_file, index=False, encoding='utf-8-sig')
        
        print(f"สำเร็จ! แปลงไฟล์และบันทึกเป็น '{output_file}' เรียบร้อยแล้ว")
        
    except Exception as e:
        print(f"เกิดข้อผิดพลาดในระหว่างการแปลงไฟล์: {e}")

if __name__ == "__main__":
    convert_excel_to_csv()