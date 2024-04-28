import spacy
import pandas as pd
import textacy
import syllables
import os

# Load English language model for SpaCy
nlp = spacy.load("en_core_web_sm")

# Function to read words from a file
def read_words(file_path):
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
        words = [line.strip() for line in file]
    return words

# Function to compute variables using SpaCy
def compute_variables_spacy(article_text, positive_words, negative_words, stop_words):
    # Process the article text using SpaCy
    doc = nlp(article_text)

    # Calculate sentiment score using SpaCy's sentiment analysis
    sentiment_score = doc.sentiment

    # Calculate positive and negative scores based on the sentiment score
    positive_score = max(sentiment_score, 0)
    negative_score = -min(sentiment_score, 0)

    # Calculate the subjectivity score
    non_stopword_tokens = [token for token in doc if token.is_alpha and token.text.lower() not in stop_words]
    subjectivity_score = (positive_score + negative_score) / (len(non_stopword_tokens) + 0.000001)

    # Include positive and negative words for sentiment analysis
    positive_word_count = len([token for token in doc if token.is_alpha and token.text.lower() in positive_words])
    negative_word_count = len([token for token in doc if token.is_alpha and token.text.lower() in negative_words])
    positive_score += positive_word_count  # Adjust positive score based on the count of positive words
    negative_score += negative_word_count  # Adjust negative score based on the count of negative words

    # Calculate average sentence length
    sent_lengths = [len(sent) for sent in doc.sents]
    avg_sentence_length = sum(sent_lengths) / len(sent_lengths)
    
    # Calculate percentage of complex words
    complex_words = [token for token in doc if token.is_alpha and syllables.estimate(token.text) >= 3]
    total_words = [token for token in doc if token.is_alpha]
    percentage_complex_words = len(complex_words) / len(total_words)

    # Calculate fog index
    fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)

    # Calculate average words per sentence
    avg_words_per_sentence = len(total_words) / len(sent_lengths)

    # Additional variables
    complex_word_count = len(complex_words)
    word_count = len(total_words)
    syllable_per_word = sum([syllables.estimate(token.text) for token in doc]) / len(total_words)
    personal_pronouns = sum(1 for token in doc if token.text.lower() in ['i', 'me', 'my', 'mine', 'myself', 'you', 'your', 'yours', 'yourself', 'yourselves', 'he', 'him', 'his', 'himself', 'she', 'her', 'hers', 'herself', 'it', 'its', 'itself', 'we', 'us', 'our', 'ours', 'ourselves', 'they', 'them', 'their', 'theirs', 'themselves'])
    avg_word_length = sum([len(token) for token in total_words]) / len(total_words)
    
    # Calculate the subjectivity score
    non_stopword_tokens = [token for token in doc if token.is_alpha and token.text.lower() not in stop_words]
    subjectivity_score = (positive_score + negative_score) / (len(non_stopword_tokens) + 0.000001)


    return {
        "POSITIVE SCORE": positive_score,
        "NEGATIVE SCORE": negative_score,
        "POLARITY SCORE": positive_score + negative_score,
        "SUBJECTIVITY SCORE": subjectivity_score,
        "AVG SENTENCE LENGTH": avg_sentence_length,
        "PERCENTAGE OF COMPLEX WORDS": percentage_complex_words,
        "FOG INDEX": fog_index,
        "AVG NUMBER OF WORDS PER SENTENCE": avg_words_per_sentence,
        "COMPLEX WORD COUNT": complex_word_count,
        "WORD COUNT": word_count,
        "SYLLABLE PER WORD": syllable_per_word,
        "PERSONAL PRONOUNS": personal_pronouns,
        "AVG WORD LENGTH": avg_word_length
    }

# Directory containing article files
articles_directory = 'extracted_articles'  

# Files containing positive and negative wordstxt
positive_words_file = r"C:\Users\user\Desktop\MasterDictionary-20240305T084627Z-001\MasterDictionary\positive-words.txt"  
negative_words_file = r"C:\Users\user\Desktop\MasterDictionary-20240305T084627Z-001\MasterDictionary\negative-words.txt"   
stop_words_directory = r"C:\Users\user\Desktop\StopWords-20240305T140220Z-001\StopWords"  

# Read positive and negative words
positive_words = read_words(positive_words_file)
negative_words = read_words(negative_words_file)

# Initialize an empty set for stop words
stop_words = set()

# Iterate over files in the stop words directory
for filename in os.listdir(stop_words_directory):
    if filename.endswith(".txt"):  # Assuming all files have a .txt extension
        file_path = os.path.join(stop_words_directory, filename)
        stop_words.update(read_words(file_path))

# List to store results for all articles
results = []

# Iterate over files in the directory
for filename in os.listdir(articles_directory):
    if filename.endswith(".txt"):  # Assuming all files have a .txt extension
        file_path = os.path.join(articles_directory, filename)
        with open(file_path, 'r', encoding='utf-8') as file:
            article_text = file.read()
            variables = compute_variables_spacy(article_text, positive_words, negative_words, stop_words)
            results.append(variables)

# Create a DataFrame with the computed variables for all articles
df_result = pd.DataFrame(results)

# Output DataFrame to Excel file
output_excel_path = r"C:\Users\user\Desktop\output.xlsx"  
df_result.to_excel(output_excel_path, index=False)

print(f"Results exported to {output_excel_path}")

