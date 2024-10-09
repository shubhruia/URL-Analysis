import time
import string
import requests
import pandas as pd
import concurrent.futures
from bs4 import BeautifulSoup
from nltk.corpus import stopwords, opinion_lexicon
from nltk.tokenize import sent_tokenize, word_tokenize

# This is used to check the code runtime
start_time = time.time()

# Download NLTK resources (run this if you haven't downloaded NLTK's resources)
# import nltk
# nltk.download('punkt_tab')
# nltk.download('stopwords')

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
    personal_pronouns = ['I', 'they', 'them', 'theirs', 'we', 'my', 'ours', 'us']
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

# List to store the results
metrics_list = []

# Function to extract and process each article
def process_url(row):
    url_id = row['URL_ID']
    url = row['URL']
    
    try:
        article_title, article_text = extract_content(url)
        metrics = calculate_metrics(article_text)
        metrics['URL_ID'] = url_id
        metrics['URL'] = url
        print(f'Article {url_id} processed successfully.')
        return metrics
    except Exception as e:
        print(f'Error processing article {url_id}: {e}')
        return None

# Use ThreadPoolExecutor to process URLs in parallel
with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
    futures = [executor.submit(process_url, row) for index, row in input_file.iterrows()]
    
    # Collect results as they complete
    for future in concurrent.futures.as_completed(futures):
        result = future.result()
        if result:
            metrics_list.append(result)

# Create DataFrame from metrics list
df = pd.DataFrame(metrics_list)

# Sort the DataFrame by URL_ID to maintain input order
df = df.sort_values(by='URL_ID').reset_index(drop=True)

# Reorder DataFrame columns
df = df[['URL_ID', 'URL', 'WORD COUNT', 'AVG WORD LENGTH', 'AVG NUMBER OF WORDS PER SENTENCE', 'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE', 'PERSONAL PRONOUNS']]

# Save results to Excel
df.to_excel('Output.xlsx', index=False)
print("Results saved to Output.xlsx")
print("Processing finished in %s seconds" % (time.time() - start_time))
