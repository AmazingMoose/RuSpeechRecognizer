from tkinter import StringVar, IntVar, Tk
from tkinter.ttk import Button, Label, Progressbar
from tkinter import filedialog, messagebox
import multiprocessing 
import speech_recognition as speech_r
import xlsxwriter
import os

def open_dialog(label_text):
    label_text.set(filedialog.askdirectory())

def recognize(to_recognize_path, to_save_path, progress):
    progress.set(0)
    recog_path = to_recognize_path.get() 
    if (recog_path == ""):
        messagebox.showerror(title="Error", message="Путь не указан")
        return
    save_path = to_save_path.get() 
    row_counter = 1
    if (save_path == ""):
        save_path = "."
    workbook = xlsxwriter.Workbook(f"{save_path}/result.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write(f"A{row_counter}", "Реплика")
    worksheet.write(f"B{row_counter}", "Текст")
    row_counter += 1
    results = []
    for file in os.listdir(recog_path):
        filename = os.fsdecode(file)
        if filename.endswith('.wav'):
            sample = speech_r.WavFile(f'{recog_path}\\{filename}')
            recognizer = speech_r.Recognizer()
            with sample as audio:
                content = recognizer.record(audio)
                recognizer.adjust_for_ambient_noise(audio)
                result = recognizer.recognize_google(content, language="ru-RU")    
                worksheet.write(f"A{row_counter}", filename.split('.')[0])
                worksheet.write(f"B{row_counter}", result)
                row_counter += 1
    workbook.close()




def init_screen(root):
    root.title("Распознаватель")
    root.minsize(200, 150)
    root.maxsize(1024, 450)
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    size = tuple(int(_) for _ in root.geometry().split('+')[0].split('x'))
    x = screen_width/2 - size[0]/2
    y = screen_height/2 - size[1]/2
    root.geometry("+%d+%d" % (x, y))
    root.resizable(0, 0)

def init_select_dir(info_text, row):
    label_info = Label(text=info_text)
    label_text = StringVar()
    label_path = Label(textvariable=label_text, wraplength=400)
    open_dialog_btn = Button(text="Обзор...", command=lambda: open_dialog(label_text))
    label_info.grid(row=row, column=0, padx=5, pady=10)
    label_path.grid(row=row, column=1, padx=10, pady=10)
    open_dialog_btn.grid(row=row, column=2, padx=10, pady=10)
    return label_text

def init_widgets(root):
    to_recognize_path = init_select_dir("Путь к файлам для распознавания:", 0)
    where_to_save_path = init_select_dir("Сохранить результат в:", 1)
    progress = IntVar(value=0)
    start_recognition_btn = Button(text="Начать распознавание", command=lambda: recognize(to_recognize_path, where_to_save_path, progress))
    start_recognition_btn.grid(row=2, column=0)
    Progressbar(orient="horizontal", length=400, variable=progress).grid(row=2, column=1)

def init_app():
    root = Tk()
    root.update_idletasks()
    init_screen(root)
    init_widgets(root)
    return root


def main():
    app = init_app()
    app.mainloop()

if __name__ == "__main__":
    main()

