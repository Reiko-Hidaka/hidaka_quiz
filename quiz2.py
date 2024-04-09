import random
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import PIL.Image
import PIL.ImageTk
from PIL import Image, ImageTk

def read_excel_file(file_path):  #エクセルを取り込む
    df = pd.read_excel(file_path)
    words = []
    for _,row in df.iterrows():
        first_bars = str(row.iloc[0]).strip()  #歌いだし
        song_title = str(row.iloc[1]).strip()  #曲名
        photo_file = str(row.iloc[2]).strip()  #写真
        if first_bars and song_title and photo_file:
            words.append((first_bars,song_title, photo_file))
    return words

class QuizApp(tk.Tk):
    def __init__(self, words):
        super().__init__()

        self.words = words
        self.score = 0
        self.current_question = 0

        self.title("ジャンヌダルクアプリ")
        self.geometry("800x600")

        title_font = ("Arial", 24, "bold")
        self.title_label = tk.Label(self, text="Janne Da Arcクイズ \n 正しい曲名を答えてください", font=title_font)
        self.title_label.pack(pady=20)

        question_font =("Arial",20)
        self.question_label = tk.Label(self, text="", font=question_font)
        self.question_label.pack(pady=10)

        counter_font = ("Arial", 16)
        self.counter_label = tk.Label(self, text="", font=counter_font)
        self.counter_label.pack()

        button_font = ("Arial", 14)
        self.choices =[]
        for _ in range(4):
            choice = tk.Button(self, text="", width=40, height=2, font=button_font)
            choice.pack(pady=10)
            self.choices.append(choice)
            
        self.next_question()
        
    def next_question(self):
        if self.current_question >= len(self.words):
            self.show_result()
            return

        first_bars, correct_song_title, photo_file = self.words[self.current_question]

        self.question_label.config(text=first_bars)

        meanings = [song_title for _, song_title, _ in self.words if song_title != correct_song_title]
        wrong_meanings = random.sample(meanings,3)
        choices = wrong_meanings + [correct_song_title]
        random.shuffle(choices)

        for i in range(4):
            self.choices[i].config(text=choices[i], command=lambda choice=choices[i], photo_file=photo_file: self.check_answer(choice, photo_file))

        self.counter_label.config(text=f"問題 ：{self.current_question + 1}/{len(self.words)}")

    def check_answer(self, choice, photo_file):
        first_bars, correct_song_title, _ = self.words[self.current_question]

        if choice == correct_song_title:
            self.score += 1
            messagebox.showinfo("正解", f"正解です！\n\歌いだし: {first_bars}\n曲名: {correct_song_title}")
            self.show_image_after_quiz(photo_file)
        else:
            messagebox.showinfo("不正解", f"不正解です...\n\n歌いだし: {first_bars}\n正解: {correct_song_title}")
            
        self.current_question += 1
        self.next_question()
    
    def show_result(self):
        photo_file = "janne.jpg"
        result_message = f"クイズ終了！\n\n正解数: {self.score}/{len(self.words)}"
        messagebox.showinfo("結果",result_message )
        self.show_image_after_quiz(photo_file)

    def show_image_after_quiz(self, photo_file=None):
        if photo_file:
            image = PIL.Image.open(photo_file)
            resized_image =  self.resize_image(image, 400, 400)
            photo = ImageTk.PhotoImage(self.resize_image(image, 400, 400))

            image_window = tk.Toplevel(self)
            image_window.title("ジャケット")
    
            label = tk.Label(image_window, image=photo)
            label.image = photo
            label.pack()

            def close_image_window():
                image_window.destroy()
                self.next_question()

            next_button = tk.Button(image_window,text="次のクイズに進む", command=self.next_question)
            next_button.pack()

            image_window.protocol("WM_DELETE_WINDOW", image_window.destroy)

    def resize_image(self, image, new_width, new_height):
        resized_image = image.resize((new_width, new_height), Image.BILINEAR)
        return resized_image
    
if __name__ == "__main__":
    file_path = "verd.xlsx"
    word_meanings = read_excel_file(file_path)

    app = QuizApp(word_meanings)
    app.mainloop()
