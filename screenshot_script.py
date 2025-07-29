import tkinter as tk
from tkinter import messagebox
import keyboard
import mss
from PIL import Image, ImageDraw, ImageFont
import win32clipboard
from io import BytesIO
import threading
import sys
import win32com.client as win32
import pythoncom
import win32gui
import time
import re


class ScreenshotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Screenshot Helper")
        self.root.geometry("300x150")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.screenshot_count = 0

        self.status_label = tk.Label(root, text="Готов к работе...", font=("Arial", 12))
        self.status_label.pack(pady=20)

        self.count_label = tk.Label(root, text=f"Скриншотов сделано: {self.screenshot_count}", font=("Arial", 10))
        self.count_label.pack(pady=10)

        self.info_label = tk.Label(root, text="Caps Lock: отправить email и сделать скриншот\nF12: простой скриншот", font=("Arial", 10))
        self.info_label.pack(pady=10)

        self.setup_hotkey()

    def setup_hotkey(self):
        hotkey_thread = threading.Thread(target=self.listen_for_hotkey, daemon=True)
        hotkey_thread.start()

    def listen_for_hotkey(self):
        # Горячие клавиши теперь безопасно планируют работу в основном потоке
        keyboard.add_hotkey('f12', self.schedule_screenshot_with_number, suppress=True)
        keyboard.add_hotkey('z', self.schedule_outlook_flow, suppress=True)
        keyboard.wait()

    def schedule_outlook_flow(self):
        """Планирует задачу с Outlook в основном потоке GUI, чтобы избежать конфликтов."""
        self.root.after(0, self.outlook_and_screenshot_email_flow)

    def schedule_screenshot_with_number(self):
        """Планирует создание скриншота в основном потоке GUI."""
        self.root.after(0, self.take_screenshot_with_number)

    def outlook_and_screenshot_email_flow(self):
        """
        Копирует выделенный email, создает письмо, делает скриншот и отправляет.
        """
        pythoncom.CoInitialize()
        try:
            # Даем время пользователю отпустить клавишу и копируем выделенное
            time.sleep(0.2)
            keyboard.send('ctrl+c')
            time.sleep(0.2)

            recipient_text = self.root.clipboard_get()

            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = recipient_text
            mail.Subject = "Alabuga Start"
            mail.Body = """Hello!
You haven't responded to your HR yet. Please tell us, are you still interested in the Alabuga Start work programme? 
 
If you would like to participate, join the group by the link https://t.me/+Q7NbK8fysNRjYTZi

Please provide the following information in the chat-group so that you can be contacted again by our HRs: 
1. Full Name (Last Name, First Name)
2. Country 
3. Age 
4. Gender
5. Phone number 
6. TG username

Or just answer your responsible HR, we are waiting for your message!

Sincerely,
Suhanov Egor
Junior Specialist
SEZ «Alabuga»
EgSuhanov@alabuga.ru
 
CONFIDENTIALITY NOTICE: Attention! This email and any files attached to it are confidential. If you are not the intended recipient you are notified that using, copying, distributing or taking any action in reliance on the contents of this information is strictly prohibited. If you have received this email in error please notify the sender and delete this email. Entering into any correspondence with us you are considered to be informed on all that is stated above."""
            
            mail.Display()

            # Более надежный поиск окна
            hwnd = 0
            
            def find_window_callback(h, extra):
                window_list, subject = extra
                title = win32gui.GetWindowText(h)
                if subject in title and ("Сообщение" in title or "Message" in title):
                    window_list.append(h)

            timeout = 10 
            start_time = time.time()
            found_hwnds = []
            while not found_hwnds and time.time() - start_time < timeout:
                win32gui.EnumWindows(find_window_callback, [found_hwnds, mail.Subject])
                if not found_hwnds:
                    time.sleep(0.2)
            
            if not found_hwnds:
                mail.Send()
                self.root.after(0, lambda: self.show_error("Не удалось найти окно письма для скриншота. Письмо было отправлено."))
                return

            hwnd = found_hwnds[0]
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5)
            
            rect = win32gui.GetWindowRect(hwnd)
            
            with mss.mss() as sct:
                capture_area = {"top": rect[1], "left": rect[0], "width": rect[2] - rect[0], "height": rect[3] - rect[1]}
                sct_img = sct.grab(capture_area)
                img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")

                mail.Send()
                
                self.screenshot_count += 1
                draw = ImageDraw.Draw(img)
                try:
                    font = ImageFont.truetype("arial.ttf", 40)
                except IOError:
                    font = ImageFont.load_default()
                text = str(self.screenshot_count)
                draw.text((10, 10), text, font=font, fill=(255, 0, 0, 255))
                
                self.send_to_clipboard(img)
                self.root.after(0, self.update_gui_success)

        except Exception as e:
            error_message = f"Ошибка в процессе:\n{e}"
            self.root.after(0, self.show_error, error_message)

    def take_screenshot_with_number(self):
        self.screenshot_count += 1
        
        try:
            with mss.mss() as sct:
                monitor = sct.monitors[1]
                width = 959
                height = 540
                left = monitor["left"] + monitor["width"] - width
                top = monitor["top"] + monitor["height"] - height
                
                capture_area = {"top": top, "left": left, "width": width, "height": height}
                sct_img = sct.grab(capture_area)

                img = Image.frombytes("RGB", sct_img.size, sct_img.bgra, "raw", "BGRX")
                
                draw = ImageDraw.Draw(img)
                try:
                    font = ImageFont.truetype("arial.ttf", 40)
                except IOError:
                    font = ImageFont.load_default()
                
                text = str(self.screenshot_count)
                draw.text((10, 10), text, font=font, fill=(255, 0, 0, 255))

                self.send_to_clipboard(img)
                
                self.root.after(0, self.update_gui_success)

        except Exception as e:
            error_message = str(e)
            self.root.after(0, self.show_error, error_message)

    def update_gui_success(self):
        self.count_label.config(text=f"Скриншотов сделано: {self.screenshot_count}")
        self.status_label.config(text="Скриншот c номером скопирован!")

    def update_gui_for_email_find(self, email):
        self.status_label.config(text=f"Найден и скопирован:\n{email}")

    def update_gui_for_selection(self):
        self.status_label.config(text="Скриншот области скопирован!")

    def show_error(self, error_message):
        messagebox.showerror("Ошибка", f"Не удалось выполнить действие:\n{error_message}")

    def send_to_clipboard(self, img):
        output = BytesIO()
        img.save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()

        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()

    def on_closing(self):
        if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
            self.root.destroy()
            keyboard.unhook_all()
            sys.exit()

if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenshotApp(root)
    root.mainloop()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    