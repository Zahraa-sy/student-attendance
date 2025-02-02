import tkinter as tk
from tkinter import ttk, filedialog
import webbrowser
import pandas as pd
import segno
from segno import helpers
import cv2
import pyzbar.pyzbar as pyzbar
import csv
from datetime import datetime
import os

def generate_qr_codes():
    # فتح ملف Excel
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx *.xls")])
    if not file_path:
        return

    # قراءة ملف Excel
    df = pd.read_excel(file_path)

    # إنشاء مجلد للصور إذا لم يكن موجوداً
    output_folder = os.path.dirname(file_path)
    qr_codes_folder = os.path.join(output_folder, 'qr_codes')
    if not os.path.exists(qr_codes_folder):
        os.makedirs(qr_codes_folder)

    # حلقة لإنشاء QR код لكل طالب
    for index, row in df.iterrows():
        student_name = row['Name']
        qr_content = f"Name: {student_name}"
        
        # إنشاء QR код
        qr = segno.make(qr_content, error='h')
        
        # حفظ QR код كملف svg
        filepath = os.path.join(qr_codes_folder, f"{student_name}.svg")
        qr.save(filepath, scale=5, border=0, finder_dark='#15a43a')

    status_label.config(text=" بنجاح qr تم إنشاء")

def read_qr_codes():
    camera_id = 0
    cap = cv2.VideoCapture(camera_id)

    csv_file_path = 'scanned_barcodes.csv'

    while True:
        ret, frame = cap.read()
        if not ret:
            break

    # Decode the barcodes in the frame
        barcodes = pyzbar.decode(frame)

        for barcode in barcodes:
            if barcode is not None:
                (x, y, w, h) = barcode.rect
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0,255, 0), 2)
                barcodeData = barcode.data.decode("utf-8")
                barcodeType = barcode.type
            
                text = "{} ({})".format(barcodeData, barcodeType)
                cv2.putText(frame, text, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX,
                        0.5, (255, 0, 0), 2)
                print("[INFO] Found {} barcode: {}".format(barcodeType, barcodeData))
            
                try:
                   df = pd.read_csv(csv_file_path)
                   if barcodeData not in df['barcode_data'].values:
                    # كتابة البيانات في ملف CSV إذا لم تكن موجودة
                       new_row = pd.DataFrame([[barcodeData, datetime.now().strftime("%Y-%m-%d %H:%M:%S")]], columns=['barcode_data', 'timestamp'])
                       df = pd.concat([df, new_row], ignore_index=True)
                       df.to_csv(csv_file_path, index=False)
                except FileNotFoundError:
                # كتابة البيانات في ملف CSV إذا لم يكن الملف موجودًا
                    new_row = pd.DataFrame([[barcodeData, datetime.now().strftime("%Y-%m-%d %H:%M:%S")]], columns=['barcode_data', 'timestamp'])
                    new_row.to_csv(csv_file_path, index=False)

            # Update the Tkinter interface with the scanned data
                tree.delete(*tree.get_children())
                try:
                    df = pd.read_csv(csv_file_path)
                    for index, row in df.iterrows():
                       tree.insert('', 'end', values=list(row))
                except FileNotFoundError:
                    pass

    # عرض الإطار
        cv2.imshow('frame', frame)
    
    # خروج عند الضغط على مفتاح ' ' أو إغلاق النافذة
        if cv2.waitKey(1) & 0xFF == ord(' '):
            break
# Release the camera and close all OpenCV windows after the loop ends
    cap.release()
    cv2.destroyAllWindows()

# Ensure the window is closed before exiting the loop
    if cv2.getWindowProperty('frame', cv2.WND_PROP_VISIBLE) < 1:
        cap.release()
        cv2.destroyAllWindows()


def display_csv_content():
    # تحديد مسار ملف CSV
    csv_file_path = 'students_scans.csv'
    
    try:
        with open(csv_file_path, 'r') as file:
            reader = csv.reader(file)
            data = list(reader)
            tree.delete(*tree.get_children())
            for row in data:
                tree.insert('', 'end', values=row)
    except FileNotFoundError:
        status_label.config(text="ملف CSV غير موجود.")

# إنشاء نافذة Tkinter
root = tk.Tk()
root.title("QR generater and reader _khabbazz")

# زر لتوليد QR кодات من ملف Excel
generate_button = tk.Button(root, text="من ملف اكسل qr توليد", command=generate_qr_codes)
generate_button.pack(pady=10)

# زر لقراءة QR кодات من الكاميرا
scan_button = tk.Button(root, text=" من الكاميرا qr قراءة", command=read_qr_codes)
scan_button.pack(pady=10)

# جدول لعرض محتويات ملف CSV
tree = ttk.Treeview(root, columns=('Data', 'Time'), show='headings')
tree.column('Data', width=200)
tree.column('Time', width=200)
tree.heading('Data', text='بيانات QR')
tree.heading('Time', text='الوقت')
tree.pack(pady=10)

# زر لعرض محتويات ملف CSV
display_button = tk.Button(root, text="csv عرض محتويات ملف", command=display_csv_content)
display_button.pack(pady=10)

# Add a label with Arabic text
label = tk.Label(root, text="صلى الله على محمد", font=("Arial", 28))
label.pack(pady=30)


def open_youtube_link(event):
    webbrowser.open_new("https://www.youtube.com/watch?v=taMXy4LO2Ps")

# Create a label for the YouTube link
youtube_link_label = tk.Label(root, text="youtube video about app", font=("Arial", 14), fg="blue", cursor="hand2")
youtube_link_label.pack()

# Bind the label to the mouse click event
youtube_link_label.bind("<Button-1>", open_youtube_link)

root.mainloop()
