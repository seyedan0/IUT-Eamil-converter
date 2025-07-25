import re
import sys
import os
from datetime import datetime, time
import pytz
from ics import Calendar, Event
from convertdate import persian as jalali

# Import GUI library
import tkinter as tk
from tkinter import scrolledtext, messagebox

# --- Settings ---
# !!! Important: Enter your email address here
RECIPIENT_EMAIL = "seyed.a.n.83@gmail.com"

# --- Core Logic Functions ---

def send_email_with_outlook(recipient, subject, body, attachment_path):
    """Sends an email with an attachment using the Outlook desktop application."""
    try:
        import win32com.client as win32
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.Body = "The calendar event file is attached."
        absolute_path = os.path.abspath(attachment_path)
        mail.Attachments.Add(absolute_path)
        mail.Send()
        return True, f"📧 Email successfully sent to {recipient}."
    except ImportError:
        return False, "❌ 'pywin32' library not found. Email sending is disabled."
    except Exception as e:
        return False, f"❌ An error occurred while sending with Outlook: {e}"

def adjust_time_for_ampm(hour, minute, am_pm_str):
    """Adjusts the hour based on the AM/PM string ('صبح'/'شب'/'عصر')."""
    hour = int(hour)
    minute = int(minute)
    if hour == 12 and am_pm_str == 'صبح':
        hour = 0
    elif hour < 12 and (am_pm_str == 'شب' or am_pm_str == 'عصر'):
        hour += 12
    return time(hour, minute)

def jalali_to_datetime(jy, jm, jd, time_obj):
    """Converts a Jalali date to a Gregorian datetime with Tehran timezone."""
    gy, gm, gd = jalali.to_gregorian(int(jy), int(jm), int(jd))
    dt = datetime.combine(datetime(gy, gm, gd), time_obj)
    return pytz.timezone("Asia/Tehran").localize(dt)

def extract_event_data(text):
    """Extracts event data from the email text."""
    # --- NEW TITLE LOGIC ---
    # منطق جدید برای استخراج یک عنوان بهتر و کامل‌تر
    
    # بخش اول: پیدا کردن نام تکلیف (هر چیزی قبل از عبارت انگلیسی)
    task_match = re.search(r'^(.*?)\s+has been changed', text, re.DOTALL)
    # بخش دوم: پیدا کردن نام درس (هر چیزی بعد از کلمه course)
    course_match = re.search(r'course\s+([^\n.]+)', text)

    # استخراج و تمیز کردن نام‌ها
    task_name = task_match.group(1).strip() if task_match else "تکلیف"
    course_name = course_match.group(1).strip() if course_match else "درس بدون عنوان"

    # حذف کلمه "درس" از انتهای نام تکلیف برای جلوگیری از تکرار
    if task_name.endswith(" درس"):
        task_name = task_name[:-5].strip()

    # ترکیب دو بخش برای ساخت عنوان نهایی
    title = f"{task_name} - {course_name}"
    
    # --- End of New Title Logic ---

    persian_months = {
        "فروردین": 1, "اردیبهشت": 2, "خرداد": 3, "تیر": 4, "مرداد": 5, "شهریور": 6,
        "مهر": 7, "آبان": 8, "آذر": 9, "دی": 10, "بهمن": 11, "اسفند": 12
    }
    date_pattern = r'(\d{1,2})\s+(' + '|'.join(persian_months.keys()) + r')\s+(\d{4})،\s+(\d{1,2}):(\d{2})\s+(صبح|شب|عصر)'
    matches = re.findall(date_pattern, text)

    if len(matches) < 2:
        raise ValueError("Could not find complete start and end date information.")

    day1, month_name1, year1, hour1, minute1, am_pm1 = matches[0]
    time_obj1 = adjust_time_for_ampm(hour1, minute1, am_pm1)
    start_datetime = jalali_to_datetime(year1, persian_months[month_name1], day1, time_obj1)

    day2, month_name2, year2, hour2, minute2, am_pm2 = matches[1]
    time_obj2 = adjust_time_for_ampm(hour2, minute2, am_pm2)
    end_datetime = jalali_to_datetime(year2, persian_months[month_name2], day2, time_obj2)

    return title, start_datetime, end_datetime

# --- GUI Application Class ---

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Email to Calendar Event Converter")
        self.root.geometry("600x450")

        self.label = tk.Label(root, text="متن ایمیل را در کادر زیر الصاق کنید:", font=("Tahoma", 12))
        self.label.pack(pady=10)

        self.text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=("Tahoma", 10), height=15)
        self.text_area.pack(pady=5, padx=10, fill="both", expand=True)

        self.process_button = tk.Button(root, text="ساخت رویداد و ارسال ایمیل", font=("Tahoma", 12, "bold"), command=self.process_email)
        self.process_button.pack(pady=20)

    def process_email(self):
        """Main function called by the button."""
        email_text = self.text_area.get("1.0", tk.END)

        if not email_text.strip():
            messagebox.showwarning("ورودی خالی", "لطفاً متن ایمیل را وارد کنید.")
            return

        try:
            title, start, end = extract_event_data(email_text)

            cal = Calendar()
            event = Event()
            event.name = title
            event.begin = start
            event.end = end
            event.description = f"Event created automatically by Python script.\n\n---\n{email_text.strip()}"
            cal.events.add(event)

            file_name = "assignment_event.ics"
            with open(file_name, "w", encoding="utf-8") as f:
                f.writelines(cal.serialize_iter())
            
            success_message = f"✅ فایل '{file_name}' با موفقیت ساخته شد.\nعنوان: {title}"

            if RECIPIENT_EMAIL != "your-email@gmail.com":
                email_sent, email_status = send_email_with_outlook(RECIPIENT_EMAIL, f"New Event: {title}", "", file_name)
                final_message = f"{success_message}\n\n{email_status}"
                if email_sent:
                    messagebox.showinfo("موفقیت", final_message)
                else:
                    messagebox.showwarning("خطای ایمیل", final_message)
            else:
                messagebox.showinfo("موفقیت", success_message + "\n\n⚠️ هشدار: ایمیل گیرنده تنظیم نشده و ایمیل ارسال نشد.")

        except ValueError as e:
            messagebox.showerror("خطا در پردازش متن", f"خطا: {e}\n\nلطفاً مطمئن شوید که متن ایمیل کامل و صحیح است.")
        except Exception as e:
            messagebox.showerror("خطای پیش‌بینی نشده", f"یک خطای کلی رخ داد:\n{e}")


if __name__ == "__main__":
    main_window = tk.Tk()
    app = App(main_window)
    main_window.mainloop()
