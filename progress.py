import chardet
import pandas as pd
import io
from tkinter import Tk, Toplevel, ttk, filedialog, messagebox

class ProgressBarApp:
    def __init__(self, maximum):
        self.root = Tk()
        self.root.withdraw()  # Tkinterのメインウィンドウを非表示にする

        # プログレスバーのウィンドウを作成
        self.progress_window = Toplevel(self.root)
        self.progress_window.title("処理中")
        self.progress_bar = ttk.Progressbar(self.progress_window, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=20)
        self.progress_bar["maximum"] = maximum

        # 画面の中央に表示
        self.progress_window.update_idletasks()
        width = self.progress_window.winfo_width()
        height = self.progress_window.winfo_height()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.progress_window.geometry(f'{width}x{height}+{x}+{y}')
        self.progress_window.deiconify()

    # プログレスバーの進捗状況の表示を更新
    def update_progress(self, value):
        self.progress_bar["value"] = value
        self.progress_window.update()

    # プログレスバーアプリを終了
    def close(self):
        self.progress_window.destroy()
        self.root.quit()

def process_csv_file(csv_files):
    # プログレスバーアプリを開始
    progress_app = ProgressBarApp(len(csv_files))

    for file_index, csv_file in enumerate(csv_files):
        # ファイルのエンコーディングを検出
        with open(csv_file, 'rb') as f:
            result = chardet.detect(f.read())
            encoding = result['encoding']

        # CSVファイルの読み込み
        with open(csv_file, 'r', encoding=encoding, errors='replace') as f:
            content = f.read()
            data = pd.read_csv(io.StringIO(content))

        total_sum = 0
        for index, row in data.iterrows():
            total_sum += int(row.iloc[0])
            progress = index / len(data)
            progress_app.update_progress(file_index + progress)

    messagebox.showinfo("合計", f"合計数: {total_sum}")
    progress_app.close()

# CSVファイルを選択
csv_files = filedialog.askopenfilenames(title='csvファイルを選択してください', filetypes=[("CSV", "*.csv")])
if csv_files:
    process_csv_file(csv_files)
else:
    messagebox.showerror("エラー", "CSVファイルが選択されていません")
