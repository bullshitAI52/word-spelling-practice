import re
import os

# Default configuration
DEFAULT_CSV_PATH = "anki_words.csv"

def clean_word_for_tts(word_str):
    """
    Remove phonetic symbols and non-word characters from the word string for TTS.
    
    Args:
        word_str (str): The raw word string (e.g., "apple /.../").
        
    Returns:
        str: The cleaned word string suitable for text-to-speech.
    """
    if not isinstance(word_str, str):
        return ""
    
    # Remove phonetic symbols in /.../ format
    cleaned_word = re.sub(r'/.*?/', '', word_str)
    
    # Remove content within parentheses
    cleaned_word = re.sub(r'\(.*\)', '', cleaned_word)
    
    # Remove other non-alphanumeric, non-space characters
    # Keeping numbers just in case, though usually words are just letters
    cleaned_word = re.sub(r'[^a-zA-Z0-9\s]', '', cleaned_word)
    
    return cleaned_word.strip()
