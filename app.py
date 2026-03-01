import os
import zipfile
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

def process_files_thread(input_zip, output_zip):
    temp_unzip_dir = "temp_unzip_dir"
    temp_original_xlsx = "temp_orig.xlsx"
    temp_fixed_xlsx = "temp_fixed.xlsx"
    
    try:
        status_var.set("جاري التجهيز للعمل...")
        
        # نفتح الملف الأصلي للقراءة، والملف النهائي للكتابة
        with zipfile.ZipFile(input_zip, 'r') as z_in, zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
            infolist = z_in.infolist()
            
            for info in infolist:
                # 1. فك تشفير الاسم العربي وإصلاحه جذرياً
                if info.flag_bits & 0x800:
                    fixed_name = info.filename 
                else:
                    try:
                        fixed_name = info.filename.encode('cp437').decode('cp1256')
                    except:
                        fixed_name = info.filename

                # تخطي المجلدات الفارغة
                if info.is_dir():
                    continue

                # عرض اسم الملف الحالي في الواجهة
                file_basename = os.path.basename(fixed_name)
                status_var.set(f"جاري معالجة: {file_basename}")

                # 2. معالجة ملفات الإكسل فقط
                if fixed_name.lower().endswith('.xlsx'):
                    
                    # سحب الإكسل كملف مسطح لتجنب أخطاء الأسماء
                    with z_in.open(info) as source, open(temp_original_xlsx, "wb") as target:
                        shutil.copyfileobj(source, target)
                    
                    # فك الإكسل
                    os.makedirs(temp_unzip_dir, exist_ok=True)
                    with zipfile.ZipFile(temp_original_xlsx, 'r') as x_in:
                        x_in.extractall(temp_unzip_dir)
                        
                    # إزالة حماية الشيتات
                    sheet_path = os.path.join(temp_unzip_dir, "xl", "worksheets")
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
                                        
                    # إعادة التجميع
                    with zipfile.ZipFile(temp_fixed_xlsx, 'w', zipfile.ZIP_DEFLATED) as x_out:
                        for folder_name, subfolders, filenames in os.walk(temp_unzip_dir):
                            for filename in filenames:
                                file_path_inside = os.path.join(folder_name, filename)
                                arcname = os.path.relpath(file_path_inside, temp_unzip_dir)
                                x_out.write(file_path_inside, arcname)
                                
                    # الكتابة في الـ ZIP النهائي بالاسم العربي الصحيح!
                    z_out.write(temp_fixed_xlsx, fixed_name)
                    
                    # تنظيف سريع
                    shutil.rmtree(temp_unzip_dir, ignore_errors=True)
                    if os.path.exists(temp_original_xlsx): os.remove(temp_original_xlsx)
                    if os.path.exists(temp_fixed_xlsx): os.remove(temp_fixed_xlsx)
                    
                else:
                    # إذا كان ملفاً آخر (صورة أو وورد)، انقله باسمه العربي كما هو
                    with z_in.open(info) as source:
                        z_out.writestr(fixed_name, source.read())

        status_var.set("✅ اكتملت العملية بنجاح!")
        root.after(0, lambda: messagebox.showinfo("نجاح", "تم فك الحماية بنجاح! الأسماء العربية الآن سليمة 100% يا أبو رزان."))

    except Exception as e:
        status_var.set("❌ حدث خطأ أثناء المعالجة")
        root.after(0, lambda: messagebox.showerror("خطأ", f"حدث خطأ:\n{str(e)}"))

    finally:
        root.after(0, progress.stop)
        root.after(0, lambda: btn.config(state=tk.NORMAL))
        if os.path.exists(temp_unzip_dir): shutil.rmtree(temp_unzip_dir, ignore_errors=True)
        if os.path.exists(temp_original_xlsx): os.remove(temp_original_xlsx)
        if os.path.exists(temp_fixed_xlsx): os.remove(temp_fixed_xlsx)

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

    btn.config(state=tk.DISABLED)
    progress.start(15)
    
    threading.Thread(target=process_files_thread, args=(input_zip, output_zip), daemon=True).start()

# --- إعداد الواجهة الرسومية ---
root = tk.Tk()
root.title("أداة فك حماية الإكسل - أبو رزان")
root.geometry("450x300")
root.configure(bg="#f4f4f4")

# توسيط النافذة
window_width = 450
window_height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = int((screen_width / 2) - (window_width / 2))
y = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

title_label = tk.Label(root, text="برنامج لفك حماية شيتات الإكسل", font=("Arial", 14, "bold"), bg="#f4f4f4", fg="#333333")
title_label.pack(pady=(20, 5))

dev_label = tk.Label(root, text="تطوير: أبو رزان", font=("Arial", 11, "italic"), bg="#f4f4f4", fg="#0066cc")
dev_label.pack(pady=(0, 20))

btn = tk.Button(root, text="📂 اختر الملف وابدأ الفك", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", 
                padx=20, pady=5, cursor="hand2", command=start_process)
btn.pack(pady=10)

style = ttk.Style()
style.theme_use('default')
style.configure("TProgressbar", thickness=15)
progress = ttk.Progressbar(root, style="TProgressbar", orient="horizontal", length=300, mode="indeterminate")
progress.pack(pady=15)

status_var = tk.StringVar()
status_var.set("جاهز للعمل...")
status_label = tk.Label(root, textvariable=status_var, font=("Arial", 10), bg="#f4f4f4", fg="#555555")
status_label.pack(pady=5)

root.mainloop()
