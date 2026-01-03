# English Word Practice Tools

This project contains tools to help you practice English words using Anki and an interactive command-line typing game.

## Files

- **anki_generator.py**: Generates an Anki deck (`.apkg` file) from a CSV word list. Includes audio pronunciation properly generated via Google TTS.
- **word_typer.py**: An interactive command-line tool to practice spelling words from the CSV list.
- **utils.py**: Shared utility functions and configuration.
- **anki_words.csv**: The source data file containing words and meanings.

## Dependencies

Install the required Python packages:

```bash
pip install -r requirements.txt
```

*Note: For `playsound` on macOS, you might need `PyObjC` if it's not automatically handled, or ensure you have a compatible player.*

## Usage

### 1. General Configuration
The default CSV file is expected to be named `anki_words.csv` and located in the same directory.
Format:
```csv
Word,Meaning
"apple /.../",苹果
...
```

### 2. Generate Anki Deck
Run the generator to create an `.apkg` file:
```bash
python anki_generator.py
```
This will create `拼写练习_四上英语单词.apkg` and a `media_files` directory. Import the `.apkg` file into Anki.

### 3. Interactive Spelling Practice
Run the typing practice tool:
```bash
python word_typer.py
```
Follow the on-screen instructions to listen to the word and type it out.
