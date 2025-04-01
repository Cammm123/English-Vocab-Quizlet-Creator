import sys
import subprocess
from docx import Document

# Install required packages if not already installed
try:
    from docx import Document
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    from docx import Document

# File path to the local .docx document
local_file_path = "/Users/cameronkani/Desktop/STEM Projects/Python Projects/src/English Vocab/Handmaid's Tale Vocabulary.docx"

# Read the document
def read_local_file(file_path):
    try:
        doc = Document(file_path)
        return doc
    except Exception as e:
        print(f"Error reading the document: {e}")
        return None

# Extract vocabulary terms and definitions
def extract_vocab_and_definitions(doc):
    vocab_and_definitions = []
    vocab = []
    definitions = []
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if text and "(" in text and "):" in text:
            try:
                # Split the word, part of speech, and definition
                word_part, definition = text.split("):", 1)
                word, pos = word_part.split("(", 1)

                clean_word = word.strip()
                clean_pos = pos.strip()
                clean_definition = definition.strip()

                # Ensure the definition is not duplicated
                if clean_definition not in definitions:
                    vocab_and_definitions.append((clean_word, clean_pos, clean_definition))
                    vocab.append(clean_word)
                    definitions.append(clean_definition)

            except ValueError:
                continue
    
    return vocab_and_definitions, vocab, definitions

# Load and process the document
doc = read_local_file(local_file_path)
if doc:
    vocab_and_definitions, vocab, definitions = extract_vocab_and_definitions(doc)
    
    if vocab_and_definitions:
        print("\nVocabulary terms and definitions found:\n")
        for word, pos, definition in vocab_and_definitions:
            print(f"{word} ({pos}): {definition}")

        print("\nVocabulary Terms:\n")
        for word in vocab:
            print(f"{word}")

        print("\nDefinitions Found:\n")
        for definition in definitions:
            print(f"{definition}")
    else:
        print("No vocabulary terms found.")
else:
    print("Failed to process the document.")
