import os
import zipfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

def extract_arabic_zip(zip_path, extract_to):
    """دالة مخصصة لفك ضغط الملفات بأسماء عربية بشكل صحيح"""
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        for info in zip_ref.infolist():
            try:
                # محاولة فك ترميز الويندوز العربي
                filename = info.filename.encode('cp437').decode('cp1256')
            except:
                filename = info.filename
            
            target_path = os.path.join(extract_to, filename)
            
            if info.is_dir():
                os.makedirs(target_path, exist_ok=True)
            else:
                os.makedirs(os.path.dirname(target_path), exist_ok=True)
                with zip_ref.open(info) as source, open(target_path, "wb") as target:
                    shutil.copyfileobj(source, target)

def process_files_thread(input_zip, output_zip):
    extract_dir = "temp_extracted"
    output_folder = "temp_unlocked"
    temp_dir = "temp_unzip"
    
    try:
        # 1. فك ضغط الملف الأصلي
        status_var.set("جاري استخراج الملفات من الملف المضغوط...")
        os.makedirs(extract_dir, exist_ok=True)
        os.makedirs(output_folder, exist_ok=True)
        
        extract_arabic_zip(input_zip, extract_dir)

        # 2. المعالجة
        status_var.set("جاري فك حماية شيتات الإكسل... الرجاء الانتظار")
        for root_dir, _, files_list in os.walk(extract_dir):
            for file_name in files_list:
                if file_name.endswith(".xlsx"):
                    file_path = os.path.join(root_dir, file_name)
                    os.makedirs(temp_dir, exist_ok=True)

                    with zipfile.ZipFile(file_path, 'r') as zip_ref:  
                        zip_ref.extractall(temp_dir)  

                    sheet_path = os.path.join(temp_dir, "xl", "worksheets")  
                    if os.path.exists(sheet_path):  
                        for xml_file in os.listdir(sheet_path):  
                            if xml_file.endswith(".xml"):  
                                xml_full_path = os.path.join(sheet_path, xml_file)  
                                with open(xml_full_path, "r", encoding="utf-8") as f:  
                                    xml_data = f.read()  
                                
                                if "<sheetProtection" in xml_data:  
                                    start_index = xml_data.find("<sheetProtection")  
                                    end_index = xml_data.find("/>", start_index) + 2  
                                    xml_data = xml_data[:start_index] + xml_data[end_index:]  
                                    with open(xml_full_path, "w", encoding="utf-8") as f:  
                                        f.write(xml_data)  

                    unlocked_file_path = os.path.join(output_folder, file_name)  
                    with zipfile.ZipFile(unlocked_file_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:  
                        for folder_name, subfolders, filenames in os.walk(temp_dir):  
                            for filename in filenames:  
                                file_path_inside = os.path.join(folder_name, filename)  
                                arcname = os.path.relpath(file_path_inside, temp_dir)  
                                new_zip.write(file_path_inside, arcname)  

                    shutil.rmtree(temp_dir, ignore_errors=True)

        # 3. ضغط الملفات النهائية
        status_var.set("جاري تجميع وحفظ الملفات النهائية...")
        output_base = output_zip[:-4] if output_zip.endswith('.zip') else output_zip
        shutil.make_archive(output_base, 'zip', output_folder)

        # إنهاء بنجاح
        status_var.set("✅ اكتملت العملية بنجاح!")
        root.after(0, lambda: messagebox.showinfo("نجاح", "تم فك الحماية وتجميع الملفات بنجاح يا أبو رزان!"))

    except Exception as e:
        status_var.set("❌ حدث خطأ أثناء المعالجة")
        root.after(0, lambda: messagebox.showerror("خطأ", f"حدث خطأ:\n{str(e)}"))

    finally:
        # إيقاف شريط التقدم وإعادة تفعيل الزر وتنظيف المجلدات
        root.after(0, progress.stop)
        root.after(0, lambda: btn.config(state=tk.NORMAL))
        if os.path.exists(extract_dir): shutil.rmtree(extract_dir, ignore_errors=True)
        if os.path.exists(output_folder): shutil.rmtree(output_folder, ignore_errors=True)
        if os.path.exists(temp_dir): shutil.rmtree(temp_dir, ignore_errors=True)


def start_process():
    input_zip = filedialog.askopenfilename(
        title="اختر ملف ZIP يحتوي على ملفات الإكسل", 
        filetypes=[("Zip files", "*.zip")]
    )
    if not input_zip:
        return

    output_zip = filedialog.asksaveasfilename(
        title="حفظ الملف النهائي باسم", 
        defaultextension=".zip", 
        filetypes=[("Zip files", "*.zip")]
    )
    if not output_zip:
        return

    # تشغيل الواجهة التفاعلية
    btn.config(state=tk.DISABLED)
    progress.start(15) # تشغيل شريط التقدم
    status_var.set("جاري التجهيز...")
    
    # تشغيل المعالجة في مسار منفصل حتى لا تتجمد الشاشة
    threading.Thread(target=process_files_thread, args=(input_zip, output_zip), daemon=True).start()


# --- إعداد الواجهة الرسومية (تطوير أبو رزان) ---
root = tk.Tk()
root.title("أداة فك حماية الإكسل - تطوير أبو رزان")
root.geometry("450x300")
root.configure(bg="#f4f4f4")

# توسيط النافذة في الشاشة
window_width = 450
window_height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = int((screen_width / 2) - (window_width / 2))
y = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# النصوص في الواجهة
title_label = tk.Label(root, text="برنامج لفك حماية شيتات الإكسل", font=("Arial", 14, "bold"), bg="#f4f4f4", fg="#333333")
title_label.pack(pady=(20, 5))

dev_label = tk.Label(root, text="تطوير: أبو رزان", font=("Arial", 11, "italic"), bg="#f4f4f4", fg="#0066cc")
dev_label.pack(pady=(0, 20))

# الزر الرئيسي
btn = tk.Button(root, text="📂 اختر الملف وابدأ الفك", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", 
                padx=20, pady=5, cursor="hand2", command=start_process)
btn.pack(pady=10)

# شريط التقدم
style = ttk.Style()
style.theme_use('default')
style.configure("TProgressbar", thickness=15)
progress = ttk.Progressbar(root, style="TProgressbar", orient="horizontal", length=300, mode="indeterminate")
progress.pack(pady=15)

# نص الحالة المتغير
status_var = tk.StringVar()
status_var.set("في انتظار اختيار الملف...")
status_label = tk.Label(root, textvariable=status_var, font=("Arial", 10), bg="#f4f4f4", fg="#555555")
status_label.pack(pady=5)

root.mainloop()
