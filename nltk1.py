import requests
from bs4 import BeautifulSoup
import nltk
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords, cmudict
from nltk.sentiment.vader import SentimentIntensityAnalyzer
import spacy
import pandas as pd
import subprocess

# Download necessary resources
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('vader_lexicon')
nltk.download('cmudict')
stop_words = set(stopwords.words('english'))
sid = SentimentIntensityAnalyzer()

# Download the spaCy model
subprocess.run(["python", "-m", "spacy", "download", "en_core_web_sm"])

nlp = spacy.load('en_core_web_sm')

# Initialize CMU Pronouncing Dictionary
d = cmudict.dict()

# Function to extract text from URL with error handling
def extract_text_from_url(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Attempt to find the title and paragraphs
        title = soup.find('h1')
        paragraphs = soup.find_all('p')
        
        # Check if title and paragraphs are found
        if title is None or paragraphs is None:
            raise ValueError(f"Title or paragraphs not found in {url}")
        
        # Extract text from title and paragraphs
        title_text = title.get_text() if title else ""
        article_text = ' '.join([para.get_text() for para in paragraphs])
        
        return title_text, article_text
    
    except Exception as e:
        print(f"Error extracting text from {url}: {e}")
        return "", ""

# Function to count syllables using CMU Pronouncing Dictionary
def count_syllables(word):
    try:
        syllables = [len(list(y for y in x if y[-1].isdigit())) for x in d[word.lower()]][0]
    except KeyError:
        # If word not found in dictionary, estimate syllables based on length
        syllables = max(1, len(word) / 3)  # A rough estimate based on word length
    return syllables

# Read input Excel file
input_df = pd.read_excel('Input.xlsx')

# Initialize an empty DataFrame for output
output_structure_df = pd.read_excel('Output Data Structure.xlsx')
output_df = pd.DataFrame(columns=output_structure_df.columns)

# Process each row in the input Excel file
for index, row in input_df.iterrows():
    url = row['URL']
    url_id = row['URL_ID']
    
    # Extract text from URL
    title, article_text = extract_text_from_url(url)
    
    # Save the extracted text to a text file
    with open(f'{url_id}.txt', 'w', encoding='utf-8') as file:
        file.write(f"{title}\n{article_text}")
    
    # Perform text analysis
    if article_text:
        words = word_tokenize(article_text)
        sentences = sent_tokenize(article_text)
        
        # Check if words list is empty
        if len(words) == 0:
            # Handle case where no words are extracted
            positive_score = 0
            negative_score = 0
            polarity_score = 0
            subjectivity_score = 0
            avg_sentence_length = 0
            percentage_complex_words = 0
            fog_index = 0
            avg_words_per_sentence = 0
            complex_word_count = 0
            word_count = 0
            syllable_per_word = 0
            personal_pronoun_count = 0
            avg_word_length = 0
        else:
            positive_score = sum([sid.polarity_scores(word)['pos'] for word in words])
            negative_score = sum([sid.polarity_scores(word)['neg'] for word in words])
            polarity_score = positive_score - negative_score
            subjectivity_score = sum([sid.polarity_scores(word)['compound'] for word in words]) / len(words)
            avg_sentence_length = len(words) / len(sentences)
            complex_words = [word for word in words if count_syllables(word) > 2]
            percentage_complex_words = len(complex_words) / len(words) * 100
            fog_index = 0.4 * (avg_sentence_length + percentage_complex_words)
            avg_words_per_sentence = avg_sentence_length
            complex_word_count = len(complex_words)
            word_count = len(words)
            syllable_per_word = sum([count_syllables(word) for word in words]) / len(words)
            doc = nlp(article_text)
            personal_pronouns = [token.text for token in doc if token.pos_ == 'PRON']
            personal_pronoun_count = len(personal_pronouns)
            avg_word_length = sum(len(word) for word in words) / len(words)
        
    else:
        # Handle case where article text extraction fails
        positive_score = 0
        negative_score = 0
        polarity_score = 0
        subjectivity_score = 0
        avg_sentence_length = 0
        percentage_complex_words = 0
        fog_index = 0
        avg_words_per_sentence = 0
        complex_word_count = 0
        word_count = 0
        syllable_per_word = 0
        personal_pronoun_count = 0
        avg_word_length = 0
    
    # Prepare dictionary with analysis results
    analysis_results = {
        'URL_ID': url_id,
        'URL': url,
        'TITLE': title,
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG SENTENCE LENGTH': avg_sentence_length,
        'PERCENTAGE OF COMPLEX WORDS': percentage_complex_words,
        'FOG INDEX': fog_index,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
        'COMPLEX WORD COUNT': complex_word_count,
        'WORD COUNT': word_count,
        'SYLLABLE PER WORD': syllable_per_word,
        'PERSONAL PRONOUNS': personal_pronoun_count,
        'AVG WORD LENGTH': avg_word_length
    }
    
    # Append results to output DataFrame
    output_df = pd.concat([output_df, pd.DataFrame([analysis_results])], ignore_index=True)

# Save the output DataFrame to Excel
output_df.to_excel('Output Data Structure.xlsx', index=False)
