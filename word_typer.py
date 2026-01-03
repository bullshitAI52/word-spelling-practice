import os
import random
import time
import pandas as pd
from gtts import gTTS
from playsound import playsound
from utils import clean_word_for_tts, DEFAULT_CSV_PATH

# --- 配置 --- #
WORD_FILE_PATH = DEFAULT_CSV_PATH  # 你的单词文件路径
AUDIO_DIR = "audio_cache"     # 存放单词发音的文件夹

# --- 辅助函数 --- #

def load_words(file_path):
    """从CSV文件中加载单词列表。"""
    words = []
    try:
        df = pd.read_csv(file_path)
        for index, row in df.iterrows():
            word = str(row['Word']).strip()
            meaning = str(row['Meaning']).strip()
            
            # 过滤掉CSV中的标题行（如果它们被错误地当作数据行）
            if word.lower() == 'word' or meaning.lower() == 'meaning':
                continue
            if word.lower() == '英⽂' or meaning.lower() == '中⽂': # 过滤你提供的示例中的特殊标题
                continue

            if word and meaning:
                words.append({'english': word, 'chinese': meaning})
            else:
                print(f"警告：跳过空单词或空释义的行：Word='{word}', Meaning='{meaning}'")
    except FileNotFoundError:
        print(f"错误：单词文件未找到，请检查路径：{file_path}")
    except pd.errors.EmptyDataError:
        print(f"错误：单词文件 {file_path} 是空的。")
    except Exception as e:
        print(f"加载单词时发生错误：{e}")
    return words

def get_audio_path(word):
    """获取单词发音文件的路径。"""
    if not os.path.exists(AUDIO_DIR):
        os.makedirs(AUDIO_DIR)
    # 使用清理后的单词作为文件名，避免特殊字符
    safe_word = clean_word_for_tts(word).lower().replace(' ', '_')
    return os.path.join(AUDIO_DIR, f"{safe_word}.mp3")

def speak_word(word):
    """生成并播放单词发音。"""
    cleaned_word = clean_word_for_tts(word)
    if not cleaned_word: # 如果清理后单词为空，则不发音
        return

    audio_file = get_audio_path(word)
    if not os.path.exists(audio_file):
        try:
            print(f"正在生成 '{cleaned_word}' 的发音...")
            tts = gTTS(text=cleaned_word, lang='en', slow=False)
            tts.save(audio_file)
        except Exception as e:
            print(f"生成发音失败：{e}。请检查网络连接。")
            return
    try:
        playsound(audio_file)
    except Exception as e:
        print(f"播放声音失败：{e}。请确保安装了playsound所需的音频播放器（如macOS上的afplay，Windows上的mpv）。")

# --- 主程序 --- #
def main():
    print("\n--- 欢迎来到单词打字背诵小助手！---")
    print(f"正在从文件 {WORD_FILE_PATH} 加载单词...")

    words = load_words(WORD_FILE_PATH)
    if not words:
        print("没有加载到任何单词，程序退出。")
        return

    print(f"成功加载 {len(words)} 个单词。")
    input("按回车键开始练习...")

    random.shuffle(words) # 打乱单词顺序

    correct_count = 0
    total_attempts = 0

    while True:
        for word_data in words:
            total_attempts += 1
            english_word_raw = word_data['english'] # 原始英文单词，可能带音标
            english_word_clean = clean_word_for_tts(english_word_raw) # 清理后的英文单词，用于比较和发音
            chinese_meaning = word_data['chinese']

            os.system('cls' if os.name == 'nt' else 'clear') # 清屏
            print(f"\n--- 第 {total_attempts} 题 ---")
            print(f"中文意思：{chinese_meaning}")
            
            # 播放发音
            speak_word(english_word_raw)

            user_input = input("请拼写英文单词 (输入 'q' 退出，'s' 听发音，'n' 跳过): ").strip()

            if user_input.lower() == 'q':
                print("\n--- 练习结束 ---")
                print(f"你一共练习了 {total_attempts-1} 个单词，正确 {correct_count} 个。")
                return
            elif user_input.lower() == 's':
                speak_word(english_word_raw)
                user_input = input("请拼写英文单词 (输入 'q' 退出，'s' 听发音，'n' 跳过): ").strip()
                if user_input.lower() == 'q':
                    print("\n--- 练习结束 ---")
                    print(f"你一共练习了 {total_attempts-1} 个单词，正确 {correct_count} 个。")
                    return
                elif user_input.lower() == 'n':
                    print(f"跳过。正确答案是：{english_word_clean}")
                    time.sleep(1.5)
                    continue
            elif user_input.lower() == 'n':
                print(f"跳过。正确答案是：{english_word_clean}")
                time.sleep(1.5)
                continue

            if user_input.lower() == english_word_clean.lower():
                print("太棒了！拼写正确！")
                correct_count += 1
            else:
                print(f"不对哦。正确答案是：{english_word_clean}")
            
            time.sleep(1.5) # 暂停一下，让用户看到反馈

        print("\n--- 这一轮单词练习完成！---")
        print(f"你一共练习了 {total_attempts} 个单词，正确 {correct_count} 个。")
        if input("是否继续下一轮练习？(y/n): ").lower() != 'y':
            print("\n--- 练习结束 ---")
            return
        random.shuffle(words) # 重新打乱顺序

if __name__ == "__main__":
    main()