import requests
from bs4 import BeautifulSoup
import pandas as pd
from nltk.tokenize import sent_tokenize, word_tokenize
from nltk.corpus import stopwords, opinion_lexicon
import string

# Download NLTK resources (run this if you haven't downloaded NLTK's resources)
import nltk
nltk.download('punkt')
nltk.download('stopwords')

# Function to extract article content from a URL
def extract_content(url):
    try:
        # Send HTTP GET request to the URL
        response = requests.get(url)
        response.raise_for_status()  # Raise HTTPError for bad responses
        soup = BeautifulSoup(response.content, 'html.parser')

        # Extract the article title and text
        article_title = soup.find('title').get_text()  # Get the title of the webpage
        article_text = soup.get_text()  # Get all text from the page

        return article_title, article_text
    
    except requests.exceptions.RequestException as e:
        # Handle exceptions related to fetching URL content
        raise Exception(f"Error fetching content from URL: {url} ({e})")

# Function to calculate text metrics with stop words removal
def calculate_metrics(text):
    sentences = sent_tokenize(text)  # Tokenize text into sentences
    words = word_tokenize(text.lower())  # Tokenize text into words and convert to lowercase

    # Remove punctuation and stopwords
    stop_words = set(stopwords.words('english'))  # Load English stopwords
    words = [word for word in words if word not in stop_words and word not in string.punctuation]

    word_count = len(words)  # Count the number of words
    sentence_count = len(sentences)  # Count the number of sentences
    
    # Load positive and negative words from opinion lexicon
    positive_words = set(opinion_lexicon.positive())
    negative_words = set(opinion_lexicon.negative())

    # Calculate sentiment-related metrics
    positive_score = sum(1 for word in words if word in positive_words)
    negative_score = sum(1 for word in words if word in negative_words)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (word_count + 0.000001)
    avg_words_per_sentence = word_count / sentence_count if sentence_count > 0 else 0
    avg_word_length = sum(len(word) for word in words) / word_count if word_count > 0 else 0
    personal_pronouns = ['I', 'we', 'my', 'ours', 'us']
    personal_pronoun_count = sum(1 for word in words if word in personal_pronouns)
    
    return {
        'POSITIVE SCORE': positive_score,
        'NEGATIVE SCORE': negative_score,
        'POLARITY SCORE': polarity_score,
        'SUBJECTIVITY SCORE': subjectivity_score,
        'AVG NUMBER OF WORDS PER SENTENCE': avg_words_per_sentence,
        'WORD COUNT': word_count,
        'PERSONAL PRONOUNS': personal_pronoun_count,
        'AVG WORD LENGTH': avg_word_length,
    }

# Load input data from Excel file
input_file = pd.read_excel('Input.xlsx')

# List to store calculated metrics
metrics_list = []

# Loop through each URL in the input data
for index, row in input_file.iterrows():
    url_id = row['URL_ID']
    url = row['URL']
    
    try:
        # Extract article content from the URL
        article_title, article_text = extract_content(url)
        
        # Calculate text metrics with stop words removal
        metrics = calculate_metrics(article_text)
        metrics['URL_ID'] = url_id
        metrics['URL'] = url
        
        metrics_list.append(metrics)
        print(f'Article {url_id} extracted and processed successfully.')
    
    except Exception as e:
        # Handle exceptions occurred during processing
        print(f'Error processing article {url_id} from URL: {url}')
        print(f'Detailed error: {e}')

# Create DataFrame from metrics list
df = pd.DataFrame(metrics_list)

# Reorder DataFrame columns
df = df[['URL_ID', 'URL', 'WORD COUNT', 'AVG WORD LENGTH', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE', 'AVG NUMBER OF WORDS PER SENTENCE', 'PERSONAL PRONOUNS']]

# Save DataFrame to Excel file
output_file = 'Output.xlsx'
df.to_excel(output_file, index=False)

print(f"Results saved to {output_file}")
