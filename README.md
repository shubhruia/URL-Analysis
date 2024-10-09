# Article Text Analysis Tool

A Python application for extracting article content from URLs and performing text analysis, including sentiment metrics. This app allows users to input a list of URLs from an Excel file, fetch the content of each article, and calculate various text metrics.

## Features

- **URL Content Extraction**: Fetch article titles and text from provided URLs.
- **Text Metrics Calculation**: Analyze extracted text to obtain:
  - Word count
  - Average word length
  - Average number of words per sentence
  - Positive and negative scores
  - Polarity and subjectivity scores
  - Count of personal pronouns
- **Parallel Processing**: Utilize multithreading for efficient URL processing.
- **Excel Input/Output**: Load input URLs from an Excel file and save results to a new Excel file.

## Demo

Try the application with your own set of URLs. Simply prepare an Excel file with a column labeled URL containing the URLs you wish to analyze.

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/shubhruia/article-text-analysis.git
   ```

2. **Install the required Python packages**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Prepare an Excel file named Input.xlsx with a column labeled URL_ID and another labeled URL containing the URLs to analyze.
2. Run the script:
   ```bash
   python main.py
   ```
3. The results will be saved in an Excel file named Output.xlsx in the same directory.

## Code Explanation

- **extract_content(url)**: Fetches the article title and text from the specified URL using requests and BeautifulSoup.
- **calculate_metrics(text)**: Computes various metrics on the provided text, including sentiment analysis using NLTK's opinion lexicon.
- **process_url(row)**: Extracts and processes each URL, storing metrics for further analysis.
- **Multithreading**: Uses ThreadPoolExecutor to speed up the processing of multiple URLs simultaneously.

## Contribution

Feel free to submit issues or pull requests if you find bugs or have suggestions for improvements.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Resources

- **[pandas](https://pandas.pydata.org/)**: For data manipulation and analysis.
- **[requests](https://requests.readthedocs.io/en/latest/)**: For making HTTP requests to fetch URL content.
- **[BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/)**: For parsing HTML and extracting data.
- **[NLTK](https://www.nltk.org/)**: For text processing and sentiment analysis.
