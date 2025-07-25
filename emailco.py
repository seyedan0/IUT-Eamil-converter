import re
import sys
import os
from datetime import datetime, time
import pytz
from ics import Calendar, Event
from convertdate import persian as jalali
import imaplib
import email
from email.header import decode_header
import tkinter as tk
from tkinter import scrolledtext, messagebox

# --- Default Settings ---
# You can pre-fill these to save time
DEFAULT_IMAP_SERVER = "mail.iut.ac.ir"
DEFAULT_EMAIL_ADDRESS = "a.noorbakhsh@ec.iut.ac.ir"
DEFAULT_RECIPIENT_EMAIL = "seyed.a.n.83@gmail.com" # ایمیل گیرنده پیش‌فرض برای اطلاع‌رسانی

# --- Core Logic Functions ---

def get_email_body(msg):
    """Extracts the body of an email message, preferring plain text."""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and part.get("Content-Disposition") is None:
                try:
                    return part.get_payload(decode=True).decode('utf-8')
                except UnicodeDecodeError:
                    return part.get_payload(decode=True).decode('latin-1')
    return msg.get_payload(decode=True).decode('utf-8', errors='ignore')

def send_notification_with_outlook(recipient, subject, attachment_path):
    """Sends a notification email via the local Outlook app."""
    try:
        import win32com.client as win32
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.Body = "فایل تقویم برای این تکلیف در ضمیمه قرار دارد."
        absolute_path = os.path.abspath(attachment_path)
        mail.Attachments.Add(absolute_path)
        mail.Send()
        return True, f"📧 Notification sent to {recipient} via Outlook."
    except ImportError:
        return False, "کتابخانه 'pywin32' برای ارسال ایمیل یافت نشد."
    except Exception:
        return False, "خطا در ارسال ایمیل. آیا Outlook نصب و در حال اجرا است؟"

def extract_event_data(text):
    """Extracts event data from the email text. Returns (title, start, end) or raises ValueError."""
    task_match = re.search(r'^(.*?)\s+has been changed', text, re.DOTALL)
    course_match = re.search(r'course\s+([^\n.]+)', text)
    if not (task_match and course_match):
        raise ValueError("Not an assignment email.")

    task_name = task_match.group(1).strip()
    course_name = course_match.group(1).strip()
    if task_name.endswith(" درس"):
        task_name = task_name[:-5].strip()
    title = f"{task_name} - {course_name}"
    
    persian_months = {"فروردین": 1, "اردیبهشت": 2, "خرداد": 3, "تیر": 4, "مرداد": 5, "شهریور": 6, "مهر": 7, "آبان": 8, "آذر": 9, "دی": 10, "بهمن": 11, "اسفند": 12}
    date_pattern = r'(\d{1,2})\s+(' + '|'.join(persian_months.keys()) + r')\s+(\d{4})،\s+(\d{1,2}):(\d{2})\s+(صبح|شب|عصر)'
    matches = re.findall(date_pattern, text)

    if len(matches) < 2:
        raise ValueError("Could not find complete start and end date information.")

    def jalali_to_datetime(jy, jm, jd, time_obj):
        gy, gm, gd = jalali.to_gregorian(int(jy), int(jm), int(jd))
        return pytz.timezone("Asia/Tehran").localize(datetime.combine(datetime(gy, gm, gd), time_obj))

    def adjust_time_for_ampm(hour, minute, am_pm_str):
        hour = int(hour)
        minute = int(minute)
        if hour == 12 and am_pm_str == 'صبح': hour = 0
        elif hour < 12 and (am_pm_str == 'شب' or am_pm_str == 'عصر'): hour += 12
        return time(hour, minute)

    day1, m_name1, y1, h1, min1, ampm1 = matches[0]
    start_dt = jalali_to_datetime(y1, persian_months[m_name1], day1, adjust_time_for_ampm(h1, min1, ampm1))

    day2, m_name2, y2, h2, min2, ampm2 = matches[1]
    end_dt = jalali_to_datetime(y2, persian_months[m_name2], day2, adjust_time_for_ampm(h2, min2, ampm2))
    
    return title, start_dt, end_dt

# --- GUI Application Class ---

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("دستیار هوشمند ایمیل دانشگاه")
        self.root.geometry("550x400")

        main_frame = tk.Frame(root, padx=15, pady=15)
        main_frame.pack(fill="both", expand=True)

        tk.Label(main_frame, text="سرور IMAP:", font=("Tahoma", 10)).grid(row=0, column=0, sticky="w", pady=5)
        self.server_entry = tk.Entry(main_frame, font=("Tahoma", 10), width=30)
        self.server_entry.grid(row=0, column=1, sticky="ew")
        self.server_entry.insert(0, DEFAULT_IMAP_SERVER)

        tk.Label(main_frame, text="آدرس ایمیل (برای ورود):", font=("Tahoma", 10)).grid(row=1, column=0, sticky="w", pady=5)
        self.email_entry = tk.Entry(main_frame, font=("Tahoma", 10))
        self.email_entry.grid(row=1, column=1, sticky="ew")
        self.email_entry.insert(0, DEFAULT_EMAIL_ADDRESS)

        tk.Label(main_frame, text="رمز عبور برنامه (App Password):", font=("Tahoma", 10)).grid(row=2, column=0, sticky="w", pady=5)
        self.password_entry = tk.Entry(main_frame, font=("Tahoma", 10), show="*")
        self.password_entry.grid(row=2, column=1, sticky="ew")
        
        # --- NEW: Recipient Email Field ---
        tk.Label(main_frame, text="ایمیل گیرنده (برای اطلاع‌رسانی):", font=("Tahoma", 10)).grid(row=3, column=0, sticky="w", pady=5)
        self.recipient_email_entry = tk.Entry(main_frame, font=("Tahoma", 10))
        self.recipient_email_entry.grid(row=3, column=1, sticky="ew")
        self.recipient_email_entry.insert(0, DEFAULT_RECIPIENT_EMAIL)

        self.process_button = tk.Button(main_frame, text="بررسی ایمیل‌های خوانده نشده", font=("Tahoma", 12, "bold"), command=self.process_emails)
        self.process_button.grid(row=4, column=0, columnspan=2, pady=20)

        tk.Label(main_frame, text="وضعیت:", font=("Tahoma", 10, "bold")).grid(row=5, column=0, sticky="w", pady=(10,0))
        self.status_label = tk.Label(main_frame, text="آماده به کار...", font=("Tahoma", 10), wraplength=500, justify="left", anchor="nw")
        self.status_label.grid(row=6, column=0, columnspan=2, sticky="ew")
        
        main_frame.columnconfigure(1, weight=1)

    def process_emails(self):
        server = self.server_entry.get()
        username = self.email_entry.get()
        password = self.password_entry.get()
        recipient_email = self.recipient_email_entry.get() # گرفتن ایمیل گیرنده

        if not all([server, username, password]):
            messagebox.showerror("خطا", "لطفاً اطلاعات ورود را کامل کنید.")
            return
        
        if not recipient_email:
            messagebox.showerror("خطا", "لطفاً ایمیل گیرنده را برای ارسال اطلاع‌رسانی مشخص کنید.")
            return

        self.status_label.config(text="در حال اتصال به سرور...")
        self.root.update_idletasks()

        try:
            mail = imaplib.IMAP4_SSL(server)
            mail.login(username, password)
            mail.select('inbox')
            
            self.status_label.config(text="در حال جستجوی ایمیل‌های خوانده نشده...")
            self.root.update_idletasks()

            _, data = mail.search(None, '(UNSEEN)')
            unread_email_ids = data[0].split()
            
            if not unread_email_ids:
                self.status_label.config(text="هیچ ایمیل خوانده نشده‌ای یافت نشد.")
                mail.logout()
                return

            assignments_found = 0
            others_left_unread = 0
            notifications_sent = 0

            for email_id in unread_email_ids:
                self.status_label.config(text=f"در حال پردازش ایمیل {len(unread_email_ids) - (assignments_found + others_left_unread)} از {len(unread_email_ids)}...")
                self.root.update_idletasks()

                _, msg_data = mail.fetch(email_id, '(RFC822)')
                msg = email.message_from_bytes(msg_data[0][1])
                email_body = get_email_body(msg)

                try:
                    title, start, end = extract_event_data(email_body)
                    
                    cal = Calendar()
                    event = Event(name=title, begin=start, end=end, description=email_body)
                    cal.events.add(event)
                    
                    safe_title = re.sub(r'[<>:"/\\|?*]', '', title)
                    file_name = f"{safe_title}.ics"
                    with open(file_name, "w", encoding="utf-8") as f:
                        f.writelines(cal.serialize_iter())
                    
                    sent_ok, _ = send_notification_with_outlook(
                        recipient=recipient_email, # استفاده از ایمیل گیرنده جدید
                        subject=f"رویداد تقویم ایجاد شد: {title}",
                        attachment_path=file_name
                    )
                    if sent_ok:
                        notifications_sent += 1
                    
                    mail.store(email_id, '+FLAGS', '\\Seen')
                    assignments_found += 1

                except ValueError:
                    # --- CHANGED: Do nothing, leave the email as unread ---
                    others_left_unread += 1
            
            final_status = f"عملیات تمام شد.\n"
            final_status += f"✅ {assignments_found} تکلیف جدید پیدا و فایل .ics برای آنها ساخته شد.\n"
            if assignments_found > 0:
                 final_status += f"📧 {notifications_sent} مورد از طریق Outlook به {recipient_email} ارسال شد.\n"
            final_status += f"ℹ️ {others_left_unread} ایمیل دیگر خوانده نشده باقی ماندند."
            
            self.status_label.config(text=final_status)
            messagebox.showinfo("عملیات موفق", final_status)
            mail.logout()

        except imaplib.IMAP4.error as e:
            self.status_label.config(text="خطا در اتصال یا ورود.")
            messagebox.showerror("خطای IMAP", f"خطا در اتصال به ایمیل: {e}\n\n- آیا سرور IMAP صحیح است؟\n- آیا از 'رمز عبور برنامه' (App Password) استفاده کرده‌اید؟")
        except Exception as e:
            self.status_label.config(text="یک خطای پیش‌بینی نشده رخ داد.")
            messagebox.showerror("خطای ناشناخته", f"یک خطای کلی رخ داد: {e}")

if __name__ == "__main__":
    if sys.platform == "win32":
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    
    root = tk.Tk()
    app = App(root)
    root.mainloop()
