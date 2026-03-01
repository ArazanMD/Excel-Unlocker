import os
import zipfile
import io
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

def process_files_thread(input_zip, output_zip):
    try:
        root.after(0, lambda: status_var.set("جاري التجهيز للعمل في الذاكرة..."))
        
        with zipfile.ZipFile(input_zip, 'r') as z_in, zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
            for info in z_in.infolist():
                # 1. قراءة بيانات الملف الأصلي
                try:
                    file_data = z_in.read(info)
                except:
                    continue
                
                # 2. فك تشفير الاسم العربي
                correct_name = info.filename
                if not (info.flag_bits & 0x800): # إذا كان الترميز قديماً
                    try:
                        correct_name = info.filename.encode('cp437').decode('cp1256')
                    except:
                        pass
                
                # تخطي المجلدات الفارغة
                if info.is_dir():
                    continue

                # إظهار الاسم على الشاشة
                file_basename = os.path.basename(correct_name)
                root.after(0, lambda name=file_basename: status_var.set(f"جاري معالجة: {name}"))

                # 3. السر هنا: إنشاء غلاف جديد تماماً للملف لإجبار الويندوز على قراءة العربي بشكل صحيح (UTF-8)
                new_info = zipfile.ZipInfo(correct_name, info.date_time)
                new_info.compress_type = zipfile.ZIP_DEFLATED

                # 4. معالجة الإكسل
                if correct_name.lower().endswith('.xlsx'):
                    xlsx_in_io = io.BytesIO(file_data)
                    xlsx_out_io = io.BytesIO()

                    with zipfile.ZipFile(xlsx_in_io, 'r') as x_in, zipfile.ZipFile(xlsx_out_io, 'w', zipfile.ZIP_DEFLATED) as x_out:
                        for x_info in x_in.infolist():
                            x_data = x_in.read(x_info)
                            
                            # فك حماية الشيتات
                            if x_info.filename.startswith("xl/worksheets/") and x_info.filename.endswith(".xml"):
                                try:
                                    xml_text = x_data.decode("utf-8")
                                    if "<sheetProtection" in xml_text:
                                        start_index = xml_text.find("<sheetProtection")
                                        end_index = xml_text.find("/>", start_index) + 2
                                        if start_index != -1 and end_index > start_index:
                                            xml_text = xml_text[:start_index] + xml_text[end_index:]
                                            x_data = xml_text.encode("utf-8")
                                except:
                                    pass
                            
                            x_out.writestr(x_info, x_data)
                    
                    # حفظ الإكسل المفتوح داخل الغلاف الجديد (بالاسم العربي)
                    modified_xlsx_bytes = xlsx_out_io.getvalue()
                    z_out.writestr(new_info, modified_xlsx_bytes)
                    
                else:
                    # حفظ أي ملف آخر داخل الغلاف الجديد
                    z_out.writestr(new_info, file_data)

        root.after(0, lambda: status_var.set("✅ اكتملت العملية بنجاح!"))
        root.after(0, lambda: messagebox.showinfo("نجاح 🏆", "تم فك الحماية وإصلاح الأسماء العربية! الآن الملفات تظهر بأسمائها الأصلية."))

    except Exception as e:
        root.after(0, lambda: status_var.set("❌ حدث خطأ أثناء المعالجة"))
        root.after(0, lambda: messagebox.showerror("خطأ", f"حدث خطأ:\n{str(e)}"))

    finally:
        root.after(0, progress.stop)
        root.after(0, lambda: btn.config(state=tk.NORMAL))

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
