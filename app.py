import os
import zipfile
import io
import time
import threading
import subprocess
import customtkinter as ctk
from tkinter import filedialog, messagebox

# إعدادات الواجهة العصرية (دعم الوضع الليلي والنهاري)
ctk.set_appearance_mode("Dark")  
ctk.set_default_color_theme("blue")

class ExcelUnlockerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("أداة فك حماية الإكسل - إصدار أبو رزان المطور")
        self.geometry("650x500")
        
        # توسيط النافذة
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = int((screen_width / 2) - (650 / 2))
        y = int((screen_height / 2) - (500 / 2))
        self.geometry(f"650x500+{x}+{y}")

        # العنوان الرئيسي
        self.title_label = ctk.CTkLabel(self, text="⚡ برنامج فك حماية الإكسل الاحترافي ⚡", font=("Arial", 22, "bold"))
        self.title_label.pack(pady=(30, 5))

        self.dev_label = ctk.CTkLabel(self, text="تطوير: أبو رزان", font=("Arial", 14, "italic"), text_color="gray")
        self.dev_label.pack(pady=(0, 30))

        # إطار الأزرار (ملف ZIP أو ملفات فردية)
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(pady=10)

        self.btn_zip = ctk.CTkButton(self.btn_frame, text="📦 معالجة ملف ZIP", font=("Arial", 15, "bold"), 
                                     width=200, height=45, fg_color="#28a745", hover_color="#218838",
                                     command=self.start_zip_process)
        self.btn_zip.grid(row=0, column=1, padx=10)

        self.btn_files = ctk.CTkButton(self.btn_frame, text="📄 معالجة ملفات إكسل فردية", font=("Arial", 15, "bold"), 
                                       width=200, height=45, fg_color="#007bff", hover_color="#0069d9",
                                       command=self.start_individual_process)
        self.btn_files.grid(row=0, column=0, padx=10)

        # شريط التقدم
        self.progress = ctk.CTkProgressBar(self, width=450, height=15)
        self.progress.pack(pady=30)
        self.progress.set(0)

        # حالة العمل
        self.status_label = ctk.CTkLabel(self, text="في انتظار اختيار الملفات...", font=("Arial", 14))
        self.status_label.pack(pady=5)

        # إطار تقرير الإنجاز (مخفي افتراضياً)
        self.report_frame = ctk.CTkFrame(self, width=450, corner_radius=10, fg_color="#333333")
        self.report_label = ctk.CTkLabel(self.report_frame, text="", font=("Arial", 14, "bold"), justify="center")
        self.report_label.pack(pady=15, padx=20)
        
        self.open_folder_btn = ctk.CTkButton(self.report_frame, text="📂 فتح مجلد الحفظ", font=("Arial", 14, "bold"), 
                                             fg_color="#ffc107", text_color="black", hover_color="#e0a800",
                                             command=self.open_output_folder)
        
        self.output_path_to_open = ""

    # دالة فك الحماية الجوهرية (في الذاكرة)
    def unlock_excel_bytes(self, file_data):
        xlsx_in_io = io.BytesIO(file_data)
        xlsx_out_io = io.BytesIO()

        with zipfile.ZipFile(xlsx_in_io, 'r') as x_in, zipfile.ZipFile(xlsx_out_io, 'w', zipfile.ZIP_DEFLATED) as x_out:
            for x_info in x_in.infolist():
                x_data = x_in.read(x_info)
                if x_info.filename.startswith("xl/worksheets/") and x_info.filename.endswith(".xml"):
                    try:
                        xml_text = x_data.decode("utf-8")
                        if "<sheetProtection" in xml_text:
                            start_idx = xml_text.find("<sheetProtection")
                            end_idx = xml_text.find("/>", start_idx) + 2
                            if start_idx != -1 and end_idx > start_idx:
                                xml_text = xml_text[:start_idx] + xml_text[end_idx:]
                                x_data = xml_text.encode("utf-8")
                    except:
                        pass
                x_out.writestr(x_info, x_data)
        return xlsx_out_io.getvalue()

    # --- معالجة ملفات ZIP ---
    def start_zip_process(self):
        input_zip = filedialog.askopenfilename(title="اختر ملف ZIP", filetypes=[("Zip files", "*.zip")])
        if not input_zip: return
        output_zip = filedialog.asksaveasfilename(title="حفظ الملف النهائي", defaultextension=".zip", filetypes=[("Zip files", "*.zip")])
        if not output_zip: return

        self.output_path_to_open = os.path.dirname(output_zip)
        self.prepare_ui_for_processing()
        threading.Thread(target=self.process_zip_thread, args=(input_zip, output_zip), daemon=True).start()

    def process_zip_thread(self, input_zip, output_zip):
        start_time = time.time()
        success_count, fail_count = 0, 0

        try:
            with zipfile.ZipFile(input_zip, 'r', metadata_encoding='cp1256') as z_in, \
                 zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
                
                total_files = len(z_in.infolist())
                for index, info in enumerate(z_in.infolist()):
                    try:
                        file_data = z_in.read(info)
                    except:
                        continue
                    
                    correct_name = info.filename
                    if info.is_dir(): continue

                    file_basename = os.path.basename(correct_name)
                    self.update_status(f"معالجة: {file_basename}", (index + 1) / total_files)

                    if correct_name.lower().endswith('.xlsx'):
                        try:
                            unlocked_data = self.unlock_excel_bytes(file_data)
                            z_out.writestr(correct_name, unlocked_data)
                            success_count += 1
                        except:
                            z_out.writestr(correct_name, file_data)
                            fail_count += 1
                    else:
                        z_out.writestr(correct_name, file_data)

            self.finish_processing(start_time, success_count, fail_count)

        except Exception as e:
            self.show_error(str(e))

    # --- معالجة ملفات إكسل فردية ---
    def start_individual_process(self):
        input_files = filedialog.askopenfilenames(title="اختر ملفات الإكسل", filetypes=[("Excel files", "*.xlsx")])
        if not input_files: return
        output_folder = filedialog.askdirectory(title="اختر مجلد الحفظ للملفات المفتوحة")
        if not output_folder: return

        self.output_path_to_open = output_folder
        self.prepare_ui_for_processing()
        threading.Thread(target=self.process_individual_thread, args=(input_files, output_folder), daemon=True).start()

    def process_individual_thread(self, input_files, output_folder):
        start_time = time.time()
        success_count, fail_count = 0, 0
        total_files = len(input_files)

        try:
            for index, file_path in enumerate(input_files):
                file_basename = os.path.basename(file_path)
                self.update_status(f"معالجة: {file_basename}", (index + 1) / total_files)

                try:
                    with open(file_path, 'rb') as f:
                        file_data = f.read()
                    
                    unlocked_data = self.unlock_excel_bytes(file_data)
                    
                    # حفظ الملف في المجلد الجديد
                    output_path = os.path.join(output_folder, file_basename)
                    with open(output_path, 'wb') as f:
                        f.write(unlocked_data)
                    success_count += 1
                except:
                    fail_count += 1

            self.finish_processing(start_time, success_count, fail_count)

        except Exception as e:
            self.show_error(str(e))

    # --- دوال الواجهة المساعدة ---
    def prepare_ui_for_processing(self):
        self.btn_zip.configure(state="disabled")
        self.btn_files.configure(state="disabled")
        self.report_frame.pack_forget()
        self.progress.set(0)
        self.progress.start()

    def update_status(self, text, progress_val):
        self.after(0, lambda: self.status_label.configure(text=text))
        self.after(0, lambda: self.progress.set(progress_val))

    def finish_processing(self, start_time, success_count, fail_count):
        elapsed_time = round(time.time() - start_time, 2)
        
        self.after(0, self.progress.stop)
        self.after(0, lambda: self.progress.set(1))
        self.after(0, lambda: self.status_label.configure(text="✅ اكتملت المهمة بنجاح!", text_color="#28a745"))
        self.after(0, lambda: self.btn_zip.configure(state="normal"))
        self.after(0, lambda: self.btn_files.configure(state="normal"))
        
        # عرض تقرير الإنجاز
        report_text = f"📊 تقرير الإنجاز:\n\n" \
                      f"✅ تمت معالجة: {success_count} ملف بنجاح\n" \
                      f"⏱️ الوقت المستغرق: {elapsed_time} ثانية\n" \
                      f"❌ عدد الأخطاء: {fail_count} ملف"
        
        self.after(0, lambda: self.report_label.configure(text=report_text))
        self.after(0, lambda: self.report_frame.pack(pady=20))
        self.after(0, lambda: self.open_folder_btn.pack(pady=(0, 15)))

    def show_error(self, error_msg):
        self.after(0, self.progress.stop)
        self.after(0, lambda: self.status_label.configure(text="❌ حدث خطأ!", text_color="#dc3545"))
        self.after(0, lambda: self.btn_zip.configure(state="normal"))
        self.after(0, lambda: self.btn_files.configure(state="normal"))
        self.after(0, lambda: messagebox.showerror("خطأ", f"حدث خطأ غير متوقع:\n{error_msg}"))

    def open_output_folder(self):
        if os.path.exists(self.output_path_to_open):
            os.startfile(self.output_path_to_open)

if __name__ == "__main__":
    app = ExcelUnlockerApp()
    app.mainloop()
