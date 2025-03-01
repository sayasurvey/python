import chardet
import pandas as pd
import io
from tkinter import Tk, Toplevel, ttk, filedialog, messagebox

class ProgressBarApp:
    def __init__(self, maximum):
        self.root = Tk()
        self.root.withdraw()

        self.progress_window = Toplevel(self.root)
        self.progress_window.title("処理中")
        self.progress_bar = ttk.Progressbar(self.progress_window, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=20)
        self.progress_bar["maximum"] = maximum
        self.progress_window.geometry("320x80")
        self.progress_window.deiconify()

        # 画面の中央に表示
        self.progress_window.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_coordinate = (screen_width // 2) - (width // 2)
        y_coordinate = (screen_height // 2) - (height // 2)
        self.progress_window.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")
        self.progress_window.deiconify()

    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.progress_window.update()

    def close(self):
        self.progress_window.destroy()
        self.root.quit()

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        return chardet.detect(f.read())["encoding"]

def read_csv_file(file_path, encoding):
    with open(file_path, 'r', encoding=encoding, errors='replace') as f:
        return pd.read_csv(io.StringIO(f.read()))

def process_csv_file(csv_files):
    progress_app = ProgressBarApp(len(csv_files))
    total_sum = 0
    
    for file_index, csv_file in enumerate(csv_files):
        encoding = detect_encoding(csv_file)
        data = read_csv_file(csv_file, encoding)

        for index, row in data.iterrows():
            total_sum += int(row.iloc[0])
            progress_app.update_progress(file_index + (index / len(data)))

    messagebox.showinfo("合計", f"合計数: {total_sum}")
    progress_app.close()

def main():
    csv_files = filedialog.askopenfilenames(title='csvファイルを選択してください', filetypes=[("CSV", "*.csv")])
    if csv_files:
        process_csv_file(csv_files)
    else:
        messagebox.showerror("エラー", "CSVファイルが選択されていません")

if __name__ == "__main__":
    main()
