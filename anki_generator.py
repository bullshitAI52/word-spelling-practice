import pandas as pd
import genanki
import random
import os
from gtts import gTTS
from utils import clean_word_for_tts, DEFAULT_CSV_PATH

# --- 配置 --- #
CSV_FILE_PATH = DEFAULT_CSV_PATH  # 你的CSV文件路径
ANKI_DECK_NAME = "拼写练习：四上英语单词" # Anki牌组的名称
ANKI_OUTPUT_FILE = "拼写练习_四上英语单词.apkg" # 输出的Anki文件名称
MEDIA_DIR = "media_files" # 存放音频文件的文件夹

# --- 辅助函数 --- #

def generate_audio(word, audio_path):
    """生成单词的MP3发音文件。"""
    cleaned_word = clean_word_for_tts(word)
    if not cleaned_word: # 如果清理后单词为空，则不生成发音
        return False

    if not os.path.exists(MEDIA_DIR):
        os.makedirs(MEDIA_DIR)

    try:
        print(f"正在生成 '{cleaned_word}' 的发音...")
        tts = gTTS(text=cleaned_word, lang='en', slow=False)
        tts.save(audio_path)
        return True
    except Exception as e:
        print(f"生成 '{cleaned_word}' 的发音失败：{e}。请检查网络连接。")
        return False

# --- 主程序 --- #
def main():
    print("\n--- 正在准备生成 Anki 闪卡 ---")
    
    # 确保媒体文件夹存在
    if not os.path.exists(MEDIA_DIR):
        os.makedirs(MEDIA_DIR)

    # 读取CSV文件
    try:
        df = pd.read_csv(CSV_FILE_PATH)
        print(f"成功读取文件：{CSV_FILE_PATH}")
    except FileNotFoundError:
        print(f"错误：CSV文件未找到，请确保文件名为 '{CSV_FILE_PATH}' 且在脚本同目录下。")
        return
    except pd.errors.EmptyDataError:
        print(f"错误：CSV文件 '{CSV_FILE_PATH}' 是空的。")
        return
    except Exception as e:
        print(f"读取CSV文件时发生错误：{e}")
        return

    # 创建Anki模型
    # 字段：Word (英文单词), Meaning (中文释义), Audio (音频文件标签)
    # 模板：Typing Card (拼写练习卡片)
    model = genanki.Model(
        random.randrange(1 << 30, 1 << 31), # 随机生成Model ID
        'Typing-English-Word-Model',
        fields=[
            {'name': 'Word'},
            {'name': 'Meaning'},
            {'name': 'Audio'}, # 新增音频字段
        ],
        templates=[
            {
                'name': 'Typing Card',
                'qfmt': '{{Audio}}<br><br><span style="color: gray;">中文释义：{{Meaning}}</span><br><br>{{type:Word}}',
                'afmt': '{{FrontSide}}<hr id="answer">你输入的是：{{type:Word}}<br>正确拼写是：<b>{{Word}}</b>',
            },
        ],
        css="""
        .card {
          font-family: arial;
          font-size: 22px;
          text-align: center; /* 居中显示 */
          color: black;
          background-color: white;
        }
        input {
          font-size: 20px;
          text-align: center; /* 输入框文字居中 */
          border: 1px solid #ccc;
          padding: 5px;
          width: 80%;
          max-width: 300px;
        }
        .card .jp-audio { /* Anki音频播放器样式 */
            margin-top: 10px;
        }
        """,
    )

    # 创建Anki牌组
    deck = genanki.Deck(
        random.randrange(1 << 30, 1 << 31), # 随机生成Deck ID
        ANKI_DECK_NAME
    )

    # 存储媒体文件路径，用于genanki.Package
    media_files_list = []

    # 遍历CSV数据，添加卡片
    for index, row in df.iterrows():
        word_raw = str(row['Word']).strip() # 原始英文单词，可能带音标
        meaning = str(row['Meaning']).strip() # 中文释义

        # 过滤掉CSV中的标题行（如果它们被错误地当作数据行）
        if word_raw.lower() == 'word' or meaning.lower() == 'meaning':
            continue
        if word_raw.lower() == '英\u200b文' or meaning.lower() == '中\u200b文': # 过滤你提供的示例中的特殊标题
            continue

        # 清理单词，用于发音和比较
        english_word_clean = clean_word_for_tts(word_raw)

        if not english_word_clean or not meaning: # 跳过空单词或空释义
            print(f"警告：跳过空单词或空释义的行：Word='{word_raw}', Meaning='{meaning}'")
            continue

        # 生成音频文件路径
        audio_filename = f"{english_word_clean.lower().replace(' ', '_')}.mp3"
        audio_full_path = os.path.join(MEDIA_DIR, audio_filename)
        audio_tag = f"[sound:{audio_filename}]"

        # 生成音频（如果不存在）
        if not os.path.exists(audio_full_path):
            if generate_audio(english_word_clean, audio_full_path):
                media_files_list.append(audio_full_path)
        else:
            media_files_list.append(audio_full_path) # 如果已存在，也加入列表，确保打包

        # 创建Anki Note
        note = genanki.Note(
            model=model,
            fields=[english_word_clean, meaning, audio_tag] # 传递清理后的单词、释义和音频标签
        )
        deck.add_note(note)

    # 导出Anki牌组
    try:
        genanki.Package(deck, media_files=media_files_list).write_to_file(ANKI_OUTPUT_FILE)
        print(f"\n导出成功！已生成：{ANKI_OUTPUT_FILE}")
        print(f"音频文件保存在：{MEDIA_DIR} 文件夹中。")
        print("请将生成的 .apkg 文件导入到 Anki 中使用。")
    except Exception as e:
        print(f"导出Anki文件失败：{e}")

if __name__ == "__main__":
    main()
