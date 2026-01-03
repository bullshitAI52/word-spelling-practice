# Word Processing & Practice Tools Box (单词处理与练习工具箱)

This project contains a collection of Python scripts and web tools for English word practice and Word document processing.
本项目包含一系列用于英语单词练习和 Word 文档处理的 Python 脚本及网页工具。

## 📂 Tools List (工具列表)

### 🅰️ English Practice Tools (英语练习工具)

#### 1. `index.html` (Web Spelling App / 网页版拼写练习)
- **功能**: 一个基于网页的单词拼写练习应用。
- **特点**: 
  - 📝 **自动出题**: 从 CSV 文件加载单词和释义。
  - 🔊 **发音支持**: 支持 TTS 自动发音。
  - 📊 **智能练习**: 支持“生词本”和“易错题”模式，自动记录进度。
  - 📱 **手机适配**: 完美适配手机端使用。
- **用法**: 直接用浏览器打开 `index.html`，或访问在线演示地址。

#### 2. `word_typer.py` (CLI Practice / 命令行练习工具)
- **功能**: 在终端（命令行）里运行的互动式拼写练习工具。
- **特点**: 
  - 读取本地 `anki_words.csv` 词库。
  - 自动播放发音（Google TTS）并显示中文含义。
  - 互动指令：`s` 重听，`n` 跳过，`q` 退出。
- **用法**: `python word_typer.py`

#### 3. `anki_generator.py` (Anki Deck Creator / Anki 卡片生成器)
- **功能**: 将 CSV 单词表一键转换成 Anki 记忆库文件 (`.apkg`)。
- **特点**: 
  - 自动生成单词的 MP3 发音文件。
  - 制作“拼写题”类型的卡片（正面听音看意，背面拼写）。
- **用法**: 运行 `python anki_generator.py`，然后将生成的 `.apkg` 文件导入 Anki 软件。

---

### 🅱️ Office Automation Tools (办公自动化工具)

#### 4. `python word_table_converter_ui.py` (General Converter / 通用 Word 表格转换器)
- **功能**: 将 Word 文档里的表格提取出来，转换成其他格式。
- **特点**: 
  - 🖥️ **图形界面**: 操作简单直观。
  - 🔄 **多格式支持**: 支持转为 **Excel** (`.xlsx`)、**JSON** 或 **HTML** 网页表格。
  - **整取整存**: 适合一次性把文档里的所有表格都搬运出来。
- **用法**: 运行 `python "python word_table_converter_ui.py"`

#### 5. `提取Word表格写入到Excel.py` (Batch Pattern Extractor / 批量 Word 数据提取器)
- **功能**: 根据“模板”从大量 Word 文档中精准提取指定位置的数据，汇总到 Excel 表中。
- **特点**: 
  - 🎯 **模板定位**: 在模板 Word 的表格里写上 `{{姓名}}` 这样的标记，程序就能自动识别位置。
  - 📂 **批量处理**: 自动扫描 `Files` 文件夹下的所有 Word 文件。
  - 🧩 **智能识别**: 支持合并单元格，适合处理简历、报名表等格式固定的文档。
- **用法**: 
  1.  准备一个模板 `.docx`，在表格格子里填入标记（如 `{{Name}}`）。
  2.  把收集到的 Word 文件都放到 `Files` 文件夹里。
  3.  运行 `python 提取Word表格写入到Excel.py`。

#### 6. `phonetics_remover_gui.py` (Phonetics Remover / 音标去除工具)
- **功能**: 批量清除文本或表格中被斜杠 `/.*/` 包围的音标内容，并支持 Excel/CSV 格式转换。
- **特点**: 
  - 🧹 **一键净化**: 自动识别并删除 `/kæl.kjə.leɪ.tər/` 格式的音标。
  - 📊 **格式通用**: 支持 Excel (.xlsx), CSV, TXT 文件的导入和导出。
  - 🛡️ **智能处理**: 能够保留其他文本，只删除音标部分。
- **用法**: 运行 `python phonetics_remover_gui.py`，选择文件后点击处理。

---

## ⚙️ Installation (安装与配置)

1.  **Install Dependencies (安装依赖)**:
    Make sure you have Python installed, then run:
    ```bash
    pip install -r requirements.txt
    ```

2.  **Data Configuration (数据配置)**:
    - For English tools, ensure `anki_words.csv` exists in the root directory.
    - Format: `Word,Meaning` (e.g., `apple,苹果`).

## 🚀 Live Demo (在线演示)
[https://bullshitai52.github.io/word-spelling-practice/](https://bullshitai52.github.io/word-spelling-practice/)
