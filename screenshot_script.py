import tkinter as tk
from tkinter import messagebox
import keyboard
import mss
from PIL import Image, ImageDraw, ImageFont
import win32clipboard
from io import BytesIO
import threading
import sys

class ScreenshotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Screenshot Helper")
        self.root.geometry("300x150")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.screenshot_count = 0

        self.status_label = tk.Label(root, text="Нажмите Caps Lock для скриншота", font=("Arial", 12))
        self.status_label.pack(pady=20)

        self.count_label = tk.Label(root, text=f"Скриншотов сделано: {self.screenshot_count}", font=("Arial", 10))
        self.count_label.pack(pady=10)

        self.setup_hotkey()

    def setup_hotkey(self):
        """
        Используем threading, чтобы hotkey не блокировал GUI
        """
        hotkey_thread = threading.Thread(target=self.listen_for_hotkey, daemon=True)
        hotkey_thread.start()

    def listen_for_hotkey(self):
        keyboard.add_hotkey('caps lock', self.take_screenshot_with_number, suppress=True)
        keyboard.wait()

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
                
                # Добавляем номер на изображение
                draw = ImageDraw.Draw(img)
                try:
                    font = ImageFont.truetype("arial.ttf", 40)
                except IOError:
                    font = ImageFont.load_default()
                
                text = str(self.screenshot_count)
                # Используем textbbox для получения размеров текста
                if hasattr(draw, 'textbbox'):
                    bbox = draw.textbbox((10, 10), text, font=font)
                else: # для старых версий Pillow
                    bbox = (10, 10, 10 + font.getsize(text)[0], 10 + font.getsize(text)[1])

                draw.text((10, 10), text, font=font, fill=(255, 0, 0, 255))

                self.send_to_clipboard(img)
                
                # Обновляем GUI из основного потока
                self.root.after(0, self.update_gui_success)

        except Exception as e:
            # Показываем ошибку в GUI
            error_message = str(e)
            self.root.after(0, self.show_error, error_message)

    def update_gui_success(self):
        self.count_label.config(text=f"Скриншотов сделано: {self.screenshot_count}")
        self.status_label.config(text="Скриншот скопирован!")

    def show_error(self, error_message):
        messagebox.showerror("Ошибка", f"Не удалось сделать скриншот:\n{error_message}")

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
        """
        Корректно завершаем работу при закрытии окна
        """
        if messagebox.askokcancel("Выход", "Вы уверены, что хотите выйти?"):
            self.root.destroy()
            keyboard.unhook_all() # Снимаем все горячие клавиши
            # Это может потребовать более "жесткого" выхода, если поток не завершается
            sys.exit()

if __name__ == "__main__":
    root = tk.Tk()
    app = ScreenshotApp(root)
    root.mainloop()
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    