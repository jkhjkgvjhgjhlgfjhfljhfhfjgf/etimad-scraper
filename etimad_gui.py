import requests
import time
import json
import os
import random
from datetime import datetime
import openpyxl
from openpyxl import Workbook, load_workbook
import configparser
import re
import csv
import sys  # <-- 1. Required for GUI output redirection
import threading  # <-- 2. Required for running scraper in background
import tkinter as tk  # <-- 3. The GUI library
from tkinter import scrolledtext, messagebox, END, font as tkFont # <-- إضافة للتحكم في الخط

# --- Global Constants ---
MASTER_FILE = "Etimad_Master_File.xlsx"
CONFIG_FILE = "config.ini"
TEMP_CSV_FILE = "new_projects.csv"

# --- Column Settings ---
ARABIC_HEADERS = [
    'إسم المشروع', 'الجهات', 'الربع السنوي', 'السنه',
    'مكان التنفيذ', 'طبيعة المشروع', 'وصف المشروع', 'الحالة',
    'مده التنفيذ المتوقعه (أيام)', 'مده التنفيذ المتوقعه (شهور)',
    'مده التنفيذ المتوقعه (سنين)', 'Encrypted_ID', 'تاريخ الإضافة'
]
UNIQUE_ID_COLUMN = 'Encrypted_ID'

# --- Scraper Settings ---
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36',
    'Referer': 'https://tenders.etimad.sa/Preplanning/PrePlaningForVisitor'
}
MAIN_PAGE_URL = "https://tenders.etimad.sa/Preplanning/PrePlaningForVisitor"
API_URL = "https://tenders.etimad.sa/Preplanning/PreplaningPagingForVisitorAsync"
PAGE_SIZE = 50

# --- Helper Functions (All text converted to Arabic) ---

def clean_text(text):
    if text is None:
        return ""
    text = str(text)
    illegal_chars_re = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')
    return illegal_chars_re.sub('', text)


def read_start_page():
    config = configparser.ConfigParser()
    start_page = 1
    if os.path.exists(CONFIG_FILE):
        try:
            config.read(CONFIG_FILE)
            start_page = int(config.get('Scraper', 'last_page', fallback=1))
            print(f" ✓ تم العثور على ملف الإعدادات: {CONFIG_FILE}")
            print(f" ✓ آخر صفحة ناجحة كانت: {start_page}. سنستكمل منها.")
        except Exception as e:
            print(f" [!] خطأ في قراءة {CONFIG_FILE}: {e}. سنبدأ من صفحة 1.")
            start_page = 1
    else:
        print(f" ✓ لم يتم العثور على ملف إعدادات.")
        print(f" ✓ سيبدأ الجلب الكامل من صفحة 1 (قد يستغرق وقتاً).")
    return start_page


def save_last_page(page_number):
    config = configparser.ConfigParser()
    config['Scraper'] = {'last_page': str(page_number)}
    try:
        with open(CONFIG_FILE, 'w') as f:
            config.write(f)
    except Exception as e:
        print(f"\n [!] تحذير: فشل حفظ رقم الصفحة {page_number} في {CONFIG_FILE}: {e}")


def append_csv_to_excel_and_delete(csv_file_path, excel_file_path):
    if not os.path.exists(csv_file_path):
        return 0  # لا يوجد شيء لدمجه

    print(f"\n 🔄 جاري دمج البيانات المؤقتة من {csv_file_path} إلى {excel_file_path}...")

    if not os.path.exists(excel_file_path):
        print(f"   ✓ ملف {excel_file_path} غير موجود. سيتم إنشاؤه بالترويسات العربية...")
        try:
            wb_new = Workbook()
            ws_new = wb_new.active
            ws_new.append(ARABIC_HEADERS)
            ws_new.sheet_view.rightToLeft = True # <-- مهم للإكسل العربي
            wb_new.save(excel_file_path)
            wb_new.close()
        except Exception as e:
            print(f" [!] فشل إنشاء ملف الإكسل الأولي: {e}. لا يمكن المتابعة.")
            return -1  # إشارة خطأ

    rows_to_add = []
    try:
        with open(csv_file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            for row in reader:
                rows_to_add.append(row)
    except Exception as e:
        print(f" [!] خطأ فادح أثناء قراءة {csv_file_path}: {e}")
        return -1

    if not rows_to_add:
        print("   ✓ ملف البيانات المؤقت (CSV) فارغ. سيتم حذفه.")
        try:
            os.remove(csv_file_path)
        except Exception as e:
            print(f" [!] فشل حذف ملف CSV الفارغ: {e}")
        return 0

    try:
        wb = load_workbook(excel_file_path)
        ws = wb.active
        for row in rows_to_add:
            ws.append(row)
        wb.save(excel_file_path)
        wb.close()
        print(f"   ✓ تم دمج {len(rows_to_add)} مشروع بنجاح.")
        
        try:
            os.remove(csv_file_path)
            print(f"   ✓ تم حذف ملف البيانات المؤقت {csv_file_path}.")
        except Exception as e:
            print(f" [!] تحذير: نجح الدمج، لكن فشل حذف {csv_file_path}: {e}")
            
        return len(rows_to_add)

    except PermissionError:
        print(f"\n{'!'*50}")
        print(f" [!] خطأ فادح (أثناء الدمج): الملف {excel_file_path} مفتوح حالياً.")
        print(f" [!] يرجى إغلاق ملف الإكسل وإعادة تشغيل الأداة.")
        print(f" [!] البيانات *لم* تضيع. هي لا تزال في {csv_file_path} وستتم محاولة دمجها في التشغيل القادم.")
        print(f"{'!'*50}")
        # إشعار لواجهة المستخدم
        messagebox.showerror("خطأ في الملف", 
                             f"الملف {excel_file_path} مفتوح حالياً.\nيرجى إغلاقه ثم إعادة تشغيل الأداة.")
        return -1
    except Exception as e:
        print(f" [!] خطأ فادح أثناء الدمج في الإكسل: {e}")
        return -1


def load_existing_project_ids_from_excel(filename):
    existing_ids = set()
    file_exists = os.path.exists(filename)
    total_existing = 0
    
    if file_exists:
        print(f"\n ✓ تم العثور على ملف الإكسل الرئيسي: {filename}")
        print(f" ✓ جاري قراءة IDs المشاريع الموجودة (لمنع التكرار)...")
        try:
            wb = load_workbook(filename, read_only=True)
            ws = wb.active
            headers = [cell.value for cell in ws[1]]
            if not headers:
                wb.close()
                return set(), False, 0
            try:
                id_index = headers.index(UNIQUE_ID_COLUMN)
            except ValueError:
                print(f" [!] خطأ: لم يتم العثور على عمود '{UNIQUE_ID_COLUMN}'. نفترض أنه ملف جديد.")
                wb.close()
                return set(), False, 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and len(row) > id_index and row[id_index]:
                    existing_ids.add(str(row[id_index]))
                    total_existing += 1
            wb.close()
            print(f" ✓ تم تحميل {total_existing} ID مشروع موجود مسبقاً.")
            return existing_ids, True, total_existing
        except Exception as e:
            print(f" [!] خطأ أثناء قراءة ملف الإكسل {filename}: {e}.")
            return set(), False, 0
    else:
        print(f"\n ✓ ملف الإكسل الرئيسي غير موجود. سيتم إنشاء ملف جديد: {filename}")
        return set(), False, 0

# --- Main Scraper Function (Modified for GUI) ---

def scrape_etimad_projects(on_complete_callback):
    """
    الدالة الرئيسية لجلب البيانات.
    تأخذ "دالة رد اتصال" (callback) لإشعار الواجهة عند الانتهاء.
    """
    try:
        start_time = time.time()
        print(f"{'='*50}")
        print(f"    بدء تشغيل أداة جلب مشاريع اعتماد")
        print(f"    تاريخ التشغيل: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"{'='*50}")
        
        # خطوة 1: دمج أي بيانات متبقية من المرة السابقة
        merge_result = append_csv_to_excel_and_delete(TEMP_CSV_FILE, MASTER_FILE)
        if merge_result == -1:
            print(" [!] تم الإلغاء بسبب فشل دمج البيانات. يرجى مراجعة السجل.")
            on_complete_callback() # إشعار الواجهة لإعادة تفعيل الزر
            return

        # خطوة 2: قراءة ملف الإعدادات والإكسل
        start_page = read_start_page()
        existing_ids, file_exists, total_projects_before = load_existing_project_ids_from_excel(MASTER_FILE)
        
        session = requests.Session()
        session.headers.update(HEADERS)
        
        print("\n ✓ جاري زيارة الصفحة الرئيسية (للحصول على كوكيز صالحة)...")
        try:
            session.get(MAIN_PAGE_URL, timeout=30).raise_for_status()
            print(f" ✓ نجحت زيارة الصفحة الرئيسية.")
        except requests.exceptions.RequestException as e:
            print(f" [!] فشل في زيارة الصفحة الرئيسية: {e}. سيتم الخروج.")
            on_complete_callback()
            return

        # خطوة 3: فتح ملف CSV المؤقت للكتابة
        try:
            csv_file = open(TEMP_CSV_FILE, 'a', encoding='utf-8-sig', newline='')
            csv_writer = csv.writer(csv_file)
        except Exception as e:
            print(f" [!] خطأ فادح: لا يمكن فتح {TEMP_CSV_FILE} للكتابة: {e}")
            on_complete_callback()
            return

        print(f"\n--- بدء عملية جلب البيانات من صفحة {start_page} ---")
        
        addition_date = datetime.now().strftime("%Y-%m-%d")
        current_page = start_page
        total_new_projects_added = 0
        start_page_for_summary = start_page
        last_page_processed = start_page - 1
        
        try:
            while True:
                timestamp = int(time.time() * 1000)
                params = {'pageSize': PAGE_SIZE, 'pageNumber': current_page, '_': timestamp}
                
                try:
                    response_api = session.get(API_URL, params=params, timeout=30)
                    response_api.raise_for_status()
                    data = response_api.json()
                    projects_list_json = data.get('data')
                    
                    new_projects_on_this_page = 0
                    
                    if not projects_list_json:
                        print(f"\n  [✓✓] تم الوصول إلى نهاية البيانات. لا توجد صفحات أخرى بعد صفحة {current_page-1}.")
                        last_page_processed = current_page - 1
                        break
                    
                    for project_json in projects_list_json:
                        project_id = project_json.get('encyptedPrePlanningId', '')
                        
                        if project_id and str(project_id) not in existing_ids:
                            new_projects_on_this_page += 1
                            total_new_projects_added += 1
                            
                            row_data = [
                                clean_text(project_json.get('projectName', '')),
                                clean_text(project_json.get('agencyName', '')),
                                clean_text(project_json.get('yearQuarterName', '')),
                                clean_text(project_json.get('year', '')),
                                clean_text(project_json.get('insideKSAString', '')),
                                clean_text(project_json.get('projectNature', '')),
                                clean_text(project_json.get('projectDescription', '')),
                                clean_text(project_json.get('statusName', '')),
                                project_json.get('durationInDays', 0),
                                project_json.get('durationInMonths', 0),
                                project_json.get('durationInYears', 0),
                                clean_text(project_id),
                                clean_text(addition_date)
                            ]
                            
                            csv_writer.writerow(row_data) # كتابة سريعة في CSV
                            existing_ids.add(str(project_id))
                    
                    if new_projects_on_this_page > 0:
                        print(f"  ✓ صفحة {current_page}: تمت إضافة {new_projects_on_this_page} مشاريع جديدة. (الإجمالي المضاف الآن: {total_new_projects_added})")
                    else:
                        print(f"  ✓ فحص صفحة {current_page}... (لا يوجد جديد)")

                    save_last_page(current_page) # حفظ رقم الصفحة سريع
                    last_page_processed = current_page
                    current_page += 1
                    
                    sleep_time = random.uniform(1.5, 4.0)
                    time.sleep(sleep_time)
                    
                except requests.exceptions.HTTPError as e:
                    if e.response.status_code == 429:
                        print(f"\n [!] خطأ 429 (ضغط طلبات) عند صفحة {current_page}. انتظار 60 ثانية...")
                        time.sleep(60)
                    else:
                        print(f"\n [!] خطأ HTTP {e.response.status_code} عند صفحة {current_page}. انتظار 10 ثواني...")
                        time.sleep(10)
                except requests.exceptions.RequestException as e:
                    print(f"\n [!] خطأ اتصال ({e.__class__.__name__}) عند صفحة {current_page}. انتظار 15 ثانية...")
                    time.sleep(15)
                except json.JSONDecodeError:
                    print(f"\n [!] فشل في قراءة JSON من صفحة {current_page}. محاولة مجددة...")
                    time.sleep(5)

        except KeyboardInterrupt:
            print(f"\n\n [!] تم استلام طلب إيقاف يدوي.")
            print(f" [!] آخر صفحة تم حفظها هي: {last_page_processed}")
            
        except Exception as e:
            print(f"\n [!] حدث خطأ فادح: {e}")

    finally:
        # خطوة 4: الإغلاق والدمج النهائي
        print("\n ⏳ جاري إغلاق ملف CSV المؤقت...")
        if 'csv_file' in locals() and not csv_file.closed:
            csv_file.close()
            
        print(" 🏁 انتهى الجلب. جاري دمج البيانات المجمعة في ملف الإكسل الرئيسي...")
        final_merge_count = append_csv_to_excel_and_delete(TEMP_CSV_FILE, MASTER_FILE)
        
        if final_merge_count == -1:
            print(" [!] فشل الدمج النهائي. البيانات محفوظة في new_projects.csv للتشغيل القادم.")
        
        # خطوة 5: الملخص النهائي
        print(f"\n\n{'='*50}")
        print(f"           📊 ملخص عملية التحديث 📊")
        print(f"{'='*50}")

        end_time = time.time()
        total_time_seconds = end_time - start_time
        total_time_str = time.strftime("%H:%M:%S", time.gmtime(total_time_seconds))
        print(f" ✓ الحالة: اكتملت العملية.")
        print(f" ✓ إجمالي الوقت المستغرق: {total_time_str} (ساعات:دقائق:ثواني)")
        
        print("-" * 50)

        total_pages_processed = max(0, (last_page_processed - start_page_for_summary + 1))
        print(f" ✓ إجمالي الصفحات التي تم فحصها: {total_pages_processed} صفحة")
        print(f"     (من صفحة {start_page_for_summary} إلى صفحة {last_page_processed})")
        print(f" ✓ التشغيل القادم سيبدأ من صفحة: {last_page_processed}")

        print("-" * 50)
        
        print(f" ✓ المشاريع الجديدة المضافة (في هذا التشغيل): {total_new_projects_added}")
        print(f" ✓ إجمالي المشاريع قبل التشغيل: {total_projects_before}")
        
        initial_merge_count = merge_result if merge_result > 0 else 0
        total_projects_after = total_projects_before + total_new_projects_added + initial_merge_count
        
        print(f" ✓ الإجمالي الجديد للمشاريع في الملف: {total_projects_after}")

        print("-" * 50)

        print(f" ✓ تم تحديث ملف الإكسل الرئيسي: {os.path.abspath(MASTER_FILE)}")
        print(f" ✓ تم تحديث ملف الإعدادات: {os.path.abspath(CONFIG_FILE)}")
        
        print(f"{'='*50}")
        
        # إشعار الواجهة بانتهاء العملية
        on_complete_callback()

# --- 4. New GUI Classes ---

class IORedirector(object):
    """كلاس لإعادة توجيه المخرجات (مثل 'print') إلى واجهة المستخدم."""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, str):
        # هذه الدالة يتم استدعاؤها بواسطة 'print'
        # نستخدم 'after' لجعلها آمنة للعمل مع الخيوط (threads)
        self.text_widget.after(0, self.insert_text, str)

    def insert_text(self, str):
        """تضيف النص وتمرر الواجهة للأسفل."""
        self.text_widget.config(state="normal") # فتح القفل للكتابة
        self.text_widget.insert(END, str)
        self.text_widget.see(END)
        self.text_widget.config(state="disabled") # إعادة القفل للقراءة فقط

    def flush(self):
        # مطلوب لإعادة توجيه stdout
        pass

class App:
    def __init__(self, root):
        self.root = root
        # --- تعديل: استخدام خط واضح ومقروء ---
        self.default_font = tkFont.nametofont("TkDefaultFont")
        self.default_font.configure(family="Arial", size=10) # يمكنك تغيير "Arial" إلى "Tahoma" أو أي خط واضح
        
        root.title("أداة جلب مشاريع اعتماد")
        root.geometry("700x500")

        # --- إنشاء الواجهة ---
        self.main_frame = tk.Frame(root, padx=10, pady=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- تعديل: تحديد خطوط عربية واضحة للواجهة ---
        self.button_font = ("Arial", 12, "bold")
        self.label_font = ("Arial", 10)
        self.log_font = ("Courier New", 10) # خط ثابت العرض للسجل

        self.start_button = tk.Button(self.main_frame, 
                                        text="بدء جلب البيانات", 
                                        command=self.start_scraping_thread,
                                        font=self.button_font,
                                        bg="#4CAF50", fg="white",
                                        padx=10, pady=5)
        self.start_button.pack(pady=(0, 10), fill=tk.X)

        self.log_label = tk.Label(self.main_frame, text="سجل العمليات:", font=self.label_font, anchor="e") # "e" = east (يمين)
        self.log_label.pack(fill=tk.X)

        self.log_area = scrolledtext.ScrolledText(self.main_frame, height=10, wrap=tk.WORD, font=self.log_font)
        self.log_area.config(state="disabled") # جعله للقراءة فقط
        self.log_area.pack(fill=tk.BOTH, expand=True)

        self.status_label = tk.Label(root, text="الحالة: جاهز", bd=1, relief=tk.SUNKEN, anchor="e", padx=10, font=self.label_font)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        # --- إعادة توجيه 'print' إلى الواجهة ---
        sys.stdout = IORedirector(self.log_area)
        sys.stderr = IORedirector(self.log_area)
        
        print("مرحباً بك في أداة جلب بيانات اعتماد.")
        print("اضغط على زر 'بدء جلب البيانات' لبدء التحديث.\n")
        
        # التعامل مع إغلاق النافذة
        root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.scraper_thread = None

    def start_scraping_thread(self):
        """يبدأ عملية الجلب في خيط (thread) منفصل لمنع تجميد الواجهة."""
        self.start_button.config(text="جاري العمل... يرجى الانتظار...", state="disabled", bg="#E0E0E0")
        self.status_label.config(text="الحالة: قيد التشغيل...")
        self.log_area.config(state="normal") # فتح القفل قبل المسح
        self.log_area.delete('1.0', END) # مسح السجل القديم
        self.log_area.config(state="disabled") # إعادة القفل
        
        # إنشاء وتشغيل الخيط
        # نمرر 'self.on_scraping_complete' كـ "دالة رد اتصال"
        self.scraper_thread = threading.Thread(target=scrape_etimad_projects, 
                                               args=(self.on_scraping_complete,),
                                               daemon=True)
        self.scraper_thread.start()

    def on_scraping_complete(self):
        """
        يتم استدعاؤها من الخيط عند الانتهاء.
        يجب استخدام 'root.after' لتحديث الواجهة بأمان.
        """
        self.root.after(0, self.update_gui_on_complete)

    def update_gui_on_complete(self):
        """تحديث عناصر الواجهة بأمان بعد الانتهاء."""
        self.start_button.config(text="بدء جلب البيانات", state="normal", bg="#4CAF50")
        self.status_label.config(text="الحالة: اكتملت. جاهز للتشغيل القادم.")
        messagebox.showinfo("اكتملت العملية", "اكتملت عملية جلب البيانات بنجاح!")

    def on_closing(self):
        """التعامل مع الضغط على زر إغلاق النافذة."""
        if self.scraper_thread and self.scraper_thread.is_alive():
            if messagebox.askyesno("تأكيد الخروج", "الأداة لا تزال تعمل في الخلفية.\nهل أنت متأكد أنك تريد الخروج؟"):
                # ملاحظة: هذا سيغلق الواجهة، لكن الخيط قد يستمر لثوانٍ
                self.root.destroy()
        else:
            self.root.destroy()

# --- 5. Main execution ---
if __name__ == "__main__":
    # لم نعد نستدعي الدالة مباشرة
    # بل نقوم بإنشاء وتشغيل الواجهة الرسومية
    root = tk.Tk()
    app = App(root)
    root.mainloop()
