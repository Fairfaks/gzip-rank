import requests
from bs4 import BeautifulSoup
import gzip
import openpyxl
from collections import Counter
from itertools import islice
import re

# Список стоп-слов (можно расширить при необходимости)
STOP_WORDS = {
    'и', 'в', 'на', 'с', 'по', 'а', 'что', 'как', 'из', 'для', 'у', 'за',
    'от', 'же', 'но', 'к', 'или', 'о', 'про', 'под', 'то', 'это', 'без',
    'до', 'после', 'да', 'тоже', 'ни', 'также', 'бы', 'там', 'тут',
    'вот', 'еще', 'то', 'так', 'чтобы', 'когда', 'где', 'кто', 'ну', 'ли', 'не', 'вы', 'через'
}

# Function to fetch and parse a webpage with headers to mimic a browser
def fetch_and_parse(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }

    with requests.Session() as session:
        response = session.get(url, headers=headers)
        soup = BeautifulSoup(response.content, 'html.parser')

    # Сохраняем данные из <head>, <title>, и <meta> перед их удалением
    title = soup.title.string.strip() if soup.title else ""
    description_tag = soup.find('meta', attrs={'name': 'description'})
    description = description_tag['content'].strip() if description_tag else ""

    # Удаляем ненужные теги
    for tag in soup(['header', 'footer', 'script', 'style']):
        tag.decompose()

    return soup, title, description


# Extract meta information: Title, H1, and Description
def extract_meta_info(soup):
    title = soup.title.string.strip() if soup.title else ""
    h1 = soup.find('h1').get_text(strip=True) if soup.find('h1') else ""
    description = soup.find('meta', attrs={'name': 'description'})
    description = description['content'].strip() if description else ""
    return title, h1, description


# Extract all anchor texts (link anchors)
def extract_anchors(soup):
    anchors = [a.get_text(strip=True) for a in soup.find_all('a') if a.get_text(strip=True)]
    return anchors


# Extract text and calculate bigrams, trigrams, and word frequency
def extract_text_and_ngrams(soup, title):
    individual_tags = {'p', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'table', 'tr'}
    text_lines = []

    # Extract text from specified tags
    for element in soup.find_all(individual_tags):
        line = element.get_text(separator=' ', strip=True)
        if line:
            text_lines.append(line)

    # Combine text and title into one
    combined_text = ' '.join(text_lines)
    full_text = f"{title} {combined_text}".strip()  # Include Title for analysis

    # Split the text into words
    words = full_text.split()

    # Filter words: remove digits, stop-words, and short words
    filtered_words = [
        word.lower() for word in words
        if word.lower() not in STOP_WORDS and not re.match(r'^\d+$', word) and len(word) > 1
    ]

    # Calculate bigrams, trigrams, and word frequency
    bigrams = zip(filtered_words, islice(filtered_words, 1, None))
    trigrams = zip(filtered_words, islice(filtered_words, 1, None), islice(filtered_words, 2, None))

    bigram_counts = Counter(bigrams)
    trigram_counts = Counter(trigrams)
    word_counts = Counter(filtered_words)

    # Get the top 20 bigrams, trigrams, and words
    top_bigrams = bigram_counts.most_common(20)
    top_trigrams = trigram_counts.most_common(20)
    top_words = word_counts.most_common(20)

    # Format bigrams, trigrams, and words as strings
    formatted_bigrams = [f"{' '.join(bigram)} - {count}" for bigram, count in top_bigrams]
    formatted_trigrams = [f"{' '.join(trigram)} - {count}" for trigram, count in top_trigrams]
    formatted_words = [f"{word} - {count}" for word, count in top_words]

    return full_text, formatted_bigrams, formatted_trigrams, formatted_words


# Function to calculate compression ratio
def calculate_compression_ratio(text):
    original_size = len(text.encode('utf-8'))
    compressed_data = gzip.compress(text.encode('utf-8'))
    compressed_size = len(compressed_data)
    return original_size / compressed_size if compressed_size > 0 else 0


# Function to check for low content
def is_low_content(text, threshold=500):
    return len(text.replace(' ', '')) < threshold


# Function to read URLs from a file
def read_urls_from_file(file_name):
    with open(file_name, 'r', encoding='utf-8') as file:
        return [line.strip() for line in file if line.strip()]


# Function to write results to an Excel file
def write_results_to_excel(results, output_file='output.xlsx'):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Page Analysis'

    # Write headers
    ws.append([
        'URL', 'Compression Ratio', 'Low Content', 'Title', 'H1', 'Meta Description',
        'Top Bigrams', 'Top Trigrams', 'Top Words', 'Top Link Anchors'
    ])

    # Write data
    for row in results:
        ws.append([
            row['url'], row['compression_ratio'], row['low_content'],
            row['title'], row['h1'], row['description'],
            '\n'.join(row['top_bigrams']), '\n'.join(row['top_trigrams']),
            '\n'.join(row['top_words']), '\n'.join(row['top_anchors'])
        ])

    # Save the file
    wb.save(output_file)
    print(f"Results saved to {output_file}")


# Main function
def main():
    input_file = "страницы для анализа.txt"  # Input file with URLs
    output_file = "output.xlsx"  # Output Excel file

    urls = read_urls_from_file(input_file)
    if not urls:
        print("No URLs found in the input file.")
        return

    results = []
    for url in urls:
        print(f"Processing {url}...")
        try:
            soup, title, description = fetch_and_parse(url)
            h1 = soup.find('h1').get_text(strip=True) if soup.find('h1') else ""
            anchors = extract_anchors(soup)
            combined_text, top_bigrams, top_trigrams, top_words = extract_text_and_ngrams(soup, title)
            compression_ratio = calculate_compression_ratio(combined_text)
            low_content = is_low_content(combined_text)

            # Collect top anchors by frequency
            anchor_counts = Counter(anchors)
            top_anchors = [f"{anchor} - {count}" for anchor, count in anchor_counts.most_common(20)]

            # Collect results
            results.append({
                'url': url,
                'compression_ratio': compression_ratio,
                'low_content': 'Yes' if low_content else 'No',
                'title': title,
                'h1': h1,
                'description': description,
                'top_bigrams': top_bigrams,
                'top_trigrams': top_trigrams,
                'top_words': top_words,
                'top_anchors': top_anchors
            })
        except Exception as e:
            print(f"Error processing {url}: {e}")

    # Save results to Excel
    write_results_to_excel(results, output_file)


if __name__ == "__main__":
    main()