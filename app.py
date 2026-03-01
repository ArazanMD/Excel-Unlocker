import os
import zipfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

def process_files():
    # 1. نافذة لاختيار ملف الـ ZIP الأصلي
    input_zip = filedialog.askopenfilename(
        title="اختر ملف ZIP يحتوي على ملفات الإكسل", 
        filetypes=[("Zip files", "*.zip")]
    )
    if not input_zip:
        return # المستخدم ألغى الاختيار

    # 2. نافذة لتحديد مكان واسم حفظ الملف النهائي
    output_zip = filedialog.asksaveasfilename(
        title="حفظ الملف النهائي باسم", 
        defaultextension=".zip", 
        filetypes=[("Zip files", "*.zip")]
    )
    if not output_zip:
        return # المستخدم ألغى الاختيار

    try:
        # إنشاء مجلدات مؤقتة للعمليات
        extract_dir = "temp_extracted"
        output_folder = "temp_unlocked"
        temp_dir = "temp_unzip"
        
        os.makedirs(extract_dir, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)

        # فك ضغط الملف الأصلي
        with zipfile.ZipFile(input_zip, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        # البحث عن ملفات الإكسل وفك حمايتها
        for root, _, files_list in os.walk(extract_dir):
            for file_name in files_list:
                if file_name.endswith(".xlsx"):
                    file_path = os.path.join(root, file_name)
                    os.makedirs(temp_dir, exist_ok=True)

                    # فك ضغط ملف الاكسل  
                    with zipfile.ZipFile(file_path, 'r') as zip_ref:  
                        zip_ref.extractall(temp_dir)  

                    sheet_path = os.path.join(temp_dir, "xl", "worksheets")  
                    if os.path.exists(sheet_path):  
                        for xml_file in os.listdir(sheet_path):  
                            if xml_file.endswith(".xml"):  
                                xml_full_path = os.path.join(sheet_path, xml_file)  
                                with open(xml_full_path, "r", encoding="utf-8") as f:  
                                    xml_data = f.read()  
                                
                                # إزالة تاق الحماية 
                                if "<sheetProtection" in xml_data:  
                                    start_index = xml_data.find("<sheetProtection")  
                                    end_index = xml_data.find("/>", start_index) + 2  
                                    xml_data = xml_data[:start_index] + xml_data[end_index:]  
                                    with open(xml_full_path, "w", encoding="utf-8") as f:  
                                        f.write(xml_data)  

                    # إعادة ضغط الملف بنفس البنية  
                    unlocked_file_path = os.path.join(output_folder, file_name)  
                    with zipfile.ZipFile(unlocked_file_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:  
                        for folder_name, subfolders, filenames in os.walk(temp_dir):  
                            for filename in filenames:  
                                file_path_inside = os.path.join(folder_name, filename)  
                                arcname = os.path.relpath(file_path_inside, temp_dir)  
                                new_zip.write(file_path_inside, arcname)  

                    # مسح المجلد المؤقت للإكسل الحالي
                    shutil.rmtree(temp_dir, ignore_errors=True)

        # ضغط الملفات المفتوحة في ملف ZIP نهائي
        # نقوم بإزالة صيغة .zip إذا كانت موجودة لأن مكتبة shutil تضيفها تلقائياً
        output_base = output_zip[:-4] if output_zip.endswith('.zip') else output_zip
        shutil.make_archive(output_base, 'zip', output_folder)

        # رسالة نجاح
        messagebox.showinfo("نجاح", "تم فك الحماية وتجميع الملفات بنجاح!")

    except Exception as e:
        messagebox.showerror("خطأ", f"حدث خطأ أثناء المعالجة:\n{str(e)}")

    finally:
        # تنظيف المجلدات المؤقتة من الجهاز بعد الانتهاء أو في حال حدوث خطأ
        if os.path.exists(extract_dir): shutil.rmtree(extract_dir, ignore_errors=True)
        if os.path.exists(output_folder): shutil.rmtree(output_folder, ignore_errors=True)
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir, ignore_errors=True)

# --- إعداد الواجهة الرسومية ---
root = tk.Tk()
root.title("أداة فك حماية الإكسل")
root.geometry("400x200")
root.eval('tk::PlaceWindow . center') # توسيط النافذة

label = tk.Label(root, text="برنامج لفك حماية شيتات الإكسل من ملف ZIP", font=("Arial", 12, "bold"))
label.pack(pady=30)

btn = tk.Button(root, text="اختر الملف وابدأ الفك", font=("Arial", 12), bg="#4CAF50", fg="white", command=process_files)
btn.pack(pady=10)

root.mainloop()
