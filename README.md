# Article Text Analysis Tool

This Python script is designed to extract article content from given URLs, analyze the text, and calculate various metrics related to sentiment, readability, and word usage. The extracted metrics are then saved into an Excel file for further analysis and reporting.

## Requirements

- Python 3.x
- Required Python libraries:
  - `requests`
  - `beautifulsoup4`
  - `pandas`
  - `nltk`

To install the required libraries, run the following command:
```bash
pip install requests beautifulsoup4 pandas nltk
```

Additionally, NLTK resources for tokenization and stopwords need to be downloaded. Run the following commands once to download the required resources:
```python
import nltk
nltk.download('punkt')
nltk.download('stopwords')
```

## Usage

1. **Input Data Preparation**: Prepare an input Excel file (`Input.xlsx`) with the following columns:
   - `URL_ID`: Unique identifier for each URL.
   - `URL`: URL of the article to be analyzed.

2. **Running the Script**:
   - Ensure that the input Excel file (`Input.xlsx`) is in the same directory as the script.
   - Execute the script `main.py`.

```bash
python article_text_analysis.py
```

3. **Output**:
   - After execution, the script will process each URL, extract article content, calculate metrics, and save the results into an Excel file named `Output.xlsx`.

## Text Metrics Calculated

The script calculates the following text metrics for each article:
- Word count
- Average word length
- Positive score
- Negative score
- Polarity score
- Subjectivity score
- Average number of words per sentence
- Count of personal pronouns

## Error Handling

The script handles exceptions related to fetching content from URLs and processing articles. If an error occurs during processing, the script will print detailed error messages for troubleshooting.

## License

[MIT License](LICENSE)
