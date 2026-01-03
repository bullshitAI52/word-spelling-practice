import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

# 可选依赖：处理 Excel/CSV
try:
    import pandas as pd
except Exception:
    pd = None


def _read_text_with_guess(path: str) -> str:
    encodings = [
        'utf-8', 'utf-8-sig',
        'gb18030', 'gbk', 'big5',
        'shift_jis', 'cp932',
        'windows-1252', 'iso-8859-1'
    ]
    with open(path, 'rb') as f:
        raw = f.read()
    for enc in encodings:
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    try:
        return raw.decode('utf-8', errors='strict')
    except UnicodeDecodeError:
        return raw.decode('latin-1', errors='replace')


def _write_text_utf8(path: str, text: str):
    with open(path, 'w', encoding='utf-8', newline='') as f:
        f.write(text)

def _remove_between_slashes(text: str) -> str:
    if not isinstance(text, str):
        return text
    start_index = text.find('/')
    end_index = text.rfind('/')
    if start_index != -1 and end_index > start_index:
        return text[:start_index] + text[end_index + 1:]
    return text


def remove_phonetics_from_file(input_file, output_file):
    """
    根据扩展名自动处理 TXT/CSV/Excel：
    - TXT/CSV：逐行/逐列去掉被斜杠包围的内容
    - XLSX/XLS：逐单元格处理，并以正确的 Excel 格式写出
    """
    try:
        in_path = Path(input_file)
        out_path = Path(output_file)
        ext_in = in_path.suffix.lower()
        ext_out = out_path.suffix.lower()

        # 优先走 pandas 分支处理结构化文件
        if ext_in in {'.xlsx', '.xls', '.csv'} or ext_out in {'.xlsx', '.xls'}:
            if pd is None:
                raise RuntimeError("需要 pandas 才能读写 Excel/CSV，请先安装：pip install pandas openpyxl")

            # 读取
            if ext_in in {'.xlsx', '.xls'}:
                df = pd.read_excel(in_path)
            elif ext_in == '.csv':
                # 尝试常见编码
                for enc in ['utf-8', 'utf-8-sig', 'gb18030', 'gbk']:
                    try:
                        df = pd.read_csv(in_path, encoding=enc)
                        break
                    except Exception:
                        df = None
                if df is None:
                    # 回退二进制读 + 手动解码
                    content = _read_text_with_guess(str(in_path))
                    lines = content.splitlines()
                    rows = [line.split(',') for line in lines]
                    df = pd.DataFrame(rows)
            else:
                # 对 txt 也支持表格导出
                content = _read_text_with_guess(str(in_path))
                lines = content.splitlines()
                df = pd.DataFrame({'text': lines})

            # 逐元素处理（仅对字符串）
            df = df.applymap(_remove_between_slashes)

            # 写出
            if ext_out in {'.xlsx', '.xls'}:
                engine = 'openpyxl' if ext_out == '.xlsx' else None
                with pd.ExcelWriter(out_path, engine=engine) as writer:
                    df.to_excel(writer, index=False)
            elif ext_out == '.csv':
                df.to_csv(out_path, index=False, encoding='utf-8-sig')
            else:
                # 写纯文本（按行合成）
                if df.shape[1] == 1:
                    text = '\n'.join(str(v) for v in df.iloc[:, 0].tolist())
                else:
                    text = '\n'.join(','.join(map(str, row)) for row in df.values.tolist())
                _write_text_utf8(str(out_path), text)
        else:
            # 普通文本文件逐行处理
            content = _read_text_with_guess(str(in_path))
            lines = content.splitlines(keepends=True)
            new_lines = [_remove_between_slashes(line) for line in lines]
            _write_text_utf8(str(out_path), ''.join(new_lines))

        messagebox.showinfo("成功", f"文件处理完成！已保存到：\n{output_file}")
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时发生错误：\n{e}")

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("音标去除工具")
        self.geometry("450x200")
        
        self.input_file_path = ""
        self.output_file_path = ""
        
        self.create_widgets()

    def create_widgets(self):
        # 1. 选取文件区域
        frame_input = tk.Frame(self, pady=10)
        frame_input.pack(fill="x", padx=10)
        
        tk.Label(frame_input, text="选择源文件:").pack(side="left")
        
        self.entry_input = tk.Entry(frame_input, width=30)
        self.entry_input.pack(side="left", padx=5, expand=True)
        
        tk.Button(frame_input, text="浏览...", command=self.select_input_file).pack(side="left")

        # 2. 保存路径区域
        frame_output = tk.Frame(self, pady=10)
        frame_output.pack(fill="x", padx=10)
        
        tk.Label(frame_output, text="选择保存路径:").pack(side="left")
        
        self.entry_output = tk.Entry(frame_output, width=30)
        self.entry_output.pack(side="left", padx=5, expand=True)
        
        tk.Button(frame_output, text="保存为...", command=self.select_output_file).pack(side="left")

        # 3. 开始按钮
        frame_start = tk.Frame(self, pady=20)
        frame_start.pack()
        
        tk.Button(frame_start, text="开始处理", font=("Arial", 12, "bold"), command=self.start_process).pack()

    def select_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择要处理的文本文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xls"),
                ("CSV 文件", "*.csv"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*"),
            ]
        )
        if file_path:
            self.input_file_path = file_path
            self.entry_input.delete(0, tk.END)
            self.entry_input.insert(0, file_path)
            
            # 根据输入文件自动生成输出文件名
            dir_name, file_name = os.path.split(file_path)
            name, ext = os.path.splitext(file_name)
            self.output_file_path = os.path.join(dir_name, f"{name}_no_phonetics{ext}")
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, self.output_file_path)

    def select_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="选择保存位置和文件名",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xls"),
                ("CSV 文件", "*.csv"),
                ("文本文件", "*.txt"),
                ("所有文件", "*.*"),
            ]
        )
        if file_path:
            self.output_file_path = file_path
            self.entry_output.delete(0, tk.END)
            self.entry_output.insert(0, file_path)

    def start_process(self):
        if not self.input_file_path:
            messagebox.showwarning("警告", "请先选择一个源文件！")
            return
        if not self.output_file_path:
            messagebox.showwarning("警告", "请选择一个保存路径！")
            return
            
        remove_phonetics_from_file(self.input_file_path, self.output_file_path)

if __name__ == "__main__":
    app = App()
    app.mainloop()
