# Word Processing & Practice Tools Box (单词处理与练习工具箱)

This project contains a collection of Python scripts and web tools for English word practice and Word document processing.
本项目包含一系列用于英语单词练习和 Word 文档处理的 Python 脚本及网页工具。

## 📂 Tools List (工具列表)

### 🅰️ English Practice Tools (英语练习工具)

#### 1. `index.html` (Web Spelling App)
- **Function**: A web-based spelling practice application.
- **Features**: 
  - Loads word lists from CSV.
  - Interactive spelling check.
  - Text-to-Speech (TTS) pronunciation.
  - Mobile-responsive design.
- **Usage**: Open `index.html` in your browser.

#### 2. `word_typer.py` (CLI Practice)
- **Function**: An interactive **Command Line** spelling practice tool.
- **Features**: 
  - Reads from `anki_words.csv`.
  - Plays audio pronunciation (Google TTS) and shows Chinese meaning.
  - Interactive feedback loop (Speak, Next, Quit).
- **Usage**: `python word_typer.py`

#### 3. `anki_generator.py` (Anki Deck Creator)
- **Function**: Converts your CSV word list into an Anki Deck (`.apkg`).
- **Features**: 
  - Automatically generates audio files.
  - Creates "Typing Cards" for spelling practice.
- **Usage**: `python anki_generator.py` -> Import the generated `.apkg` into Anki.

---

### 🅱️ Office Automation Tools (办公自动化工具)

#### 4. `python word_table_converter_ui.py` (General Converter)
- **Function**: Converts Word tables to other formats.
- **Features**: 
  - Graphic User Interface (GUI).
  - Convert `.docx` tables to **Excel** (`.xlsx`), **JSON**, or **HTML**.
  - Best for simple, direct conversion of all tables in a document.
- **Usage**: `python "python word_table_converter_ui.py"`

#### 5. `提取Word表格写入到Excel.py` (Batch Pattern Extractor)
- **Function**: Batch extracts specific data from multiple Word documents into a single Excel sheet based on a template.
- **Features**: 
  - **Template System**: Use `{{tag}}` in a template Word doc to define what to extract.
  - **Batch Processing**: Automatically processes all `.docx` files in the `Files` directory.
  - **Smart Merge**: Handles merged cells correctly.
  - **Resume/Form Aggregation**: Ideal for collecting data from many identical forms.
- **Usage**: 
  1.  Prepare a template `.docx` with tags like `{{Name}}` in the table cells.
  2.  Place your data files in a `Files` folder.
  3.  Run `python 提取Word表格写入到Excel.py`.

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
