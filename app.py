import os
import zipfile
import io
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

def process_files_thread(input_zip, output_zip):
    try:
        root.after(0, lambda: status_var.set("جاري التجهيز للعمل..."))
        
        # السحر هنا: استخدام ميزة (metadata_encoding) المتاحة في بايثون 3.11+ لقراءة العربي إجبارياً
        with zipfile.ZipFile(input_zip, 'r', metadata_encoding='cp1256') as z_in, \
             zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
            
            for info in z_in.infolist():
                try:
                    file_data = z_in.read(info)
                except:
                    continue
                
                # الاسم الآن يقرأ بشكل نقي وصحيح تماماً بفضل التحديث الجديد
                correct_name = info.filename
                
                if info.is_dir():
                    continue

                # إظهار الاسم على الشاشة
                file_basename = os.path.basename(correct_name)
                root.after(0, lambda name=file_basename: status_var.set(f"جاري معالجة: {name}"))

                # معالجة الإكسل
                if correct_name.lower().endswith('.xlsx'):
                    xlsx_in_io = io.BytesIO(file_data)
                    xlsx_out_io = io.BytesIO()

                    with zipfile.ZipFile(xlsx_in_io, 'r') as x_in, zipfile.ZipFile(xlsx_out_io, 'w', zipfile.ZIP_DEFLATED) as x_out:
                        for x_info in x_in.infolist():
                            x_data = x_in.read(x_info)
                            
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
                    
                    # حفظ الإكسل بالاسم النقي
                    z_out.writestr(correct_name, xlsx_out_io.getvalue())
                    
                else:
                    # حفظ الملفات الأخرى
                    z_out.writestr(correct_name, file_data)

        root.after(0, lambda: status_var.set("✅ اكتملت العملية بنجاح!"))
        root.after(0, lambda: messagebox.showinfo("نجاح 🏆", "ألف مبروك! تم فك الحماية والأسماء العربية الآن أصلية 100%."))

    except Exception as e:
        root.after(0, lambda: status_var.set("❌ حدث خطأ أثناء المعالجة"))
        root.after(0, lambda: messagebox.showerror("خطأ", f"حدث خطأ:\n{str(e)}"))

    finally:
        root.after(0, progress.stop)
        root.after(0, lambda: btn.config(state=tk.NORMAL))

def start_process():
    input_zip = filedialog.askopenfilename(
        title="اختر ملف ZIP", 
        filetypes=[("Zip files", "*.zip")]
    )
    if not input_zip: return

    output_zip = filedialog.asksaveasfilename(
        title="حفظ الملف النهائي", 
        defaultextension=".zip", 
        filetypes=[("Zip files", "*.zip")]
    )
    if not output_zip: return

    btn.config(state=tk.DISABLED)
    progress.start(15)
    
    threading.Thread(target=process_files_thread, args=(input_zip, output_zip), daemon=True).start()

# --- إعداد الواجهة الرسومية ---
root = tk.Tk()
root.title("أداة فك حماية الإكسل - أبو رزان")
root.geometry("450x300")
root.configure(bg="#f4f4f4")

window_width, window_height = 450, 300
screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()
x = int((screen_width / 2) - (window_width / 2))
y = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

tk.Label(root, text="برنامج لفك حماية شيتات الإكسل", font=("Arial", 14, "bold"), bg="#f4f4f4", fg="#333333").pack(pady=(20, 5))
tk.Label(root, text="تطوير: أبو رزان", font=("Arial", 11, "italic"), bg="#f4f4f4", fg="#0066cc").pack(pady=(0, 20))

btn = tk.Button(root, text="📂 اختر الملف وابدأ الفك", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", padx=20, pady=5, cursor="hand2", command=start_process)
btn.pack(pady=10)

style = ttk.Style()
style.theme_use('default')
style.configure("TProgressbar", thickness=15)
progress = ttk.Progressbar(root, style="TProgressbar", orient="horizontal", length=300, mode="indeterminate")
progress.pack(pady=15)

status_var = tk.StringVar(value="جاهز للعمل...")
tk.Label(root, textvariable=status_var, font=("Arial", 10), bg="#f4f4f4", fg="#555555").pack(pady=5)

root.mainloop()
