import requests
from bs4 import BeautifulSoup
import re
from datetime import datetime
import pandas as pd
import time

def construct_url(page_num, state):
    base_url = 'https://www.xetra.com/xetra-en/newsroom/xetra-newsboard/4442!search'
    params = {
        'state': state,
        'sort': 'sDate desc',
        'hitsPerPage': 100,
        'pageNum': page_num
    }
    response = requests.get(base_url, params=params, timeout=10)
    response.raise_for_status()
    return response

def parse_date(date_str):
    date_str = re.sub(r'[./]', '-', date_str)
    date_formats = [
        "%d-%m-%Y", "%d-%m-%y", "%m-%d-%Y", "%m-%d-%y", "%Y-%m-%d"
    ]
    for date_format in date_formats:
        try:
            parsed_date = datetime.strptime(date_str, date_format)
            return parsed_date.strftime('%Y-%m-%d')
        except ValueError:
            continue
    return 'NA'

def extract_date(text, details):
    ex_date_pattern = r'ex-dividend/interest day on (\d{2}.\d{2}.\d{4})'
    match = re.search(ex_date_pattern, text)
    if match:
        ex_date_raw = match.group(1)
        ex_date = parse_date(ex_date_raw)
        details['Ex-date'] = ex_date
    else:
        details['Ex-date'] = 'NA'
    return details

def extract_details(detail_url):
    response = requests.get(detail_url)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    details = {}
    title = soup.find('h2', class_='main-title').get_text(separator=' ').strip()
    text = soup.find('div', class_='detailText').get_text(separator=' ').strip()
    isins = re.findall(r'\b[A-Z]{2}(?=[A-Z0-9]*\d)[A-Z0-9]{10}\b', title)
    if isins:
        details['ISIN'] = isins[0]
    details = extract_date(text, details)
    return details

def extract_state(soup):
    state_input = soup.find('input', {'name': 'state'})
    if state_input and state_input.get('value'):
        return state_input['value']
    return None

def scrape_xetra_newsboard(max_pages):
    results = []
    page_num = 0
    max_retries = 3
    get_state_response = requests.get('https://www.xetra.com/xetra-en/newsroom/xetra-newsboard', timeout=10)
    get_state_response.raise_for_status()
    get_state_soup = BeautifulSoup(get_state_response.text, 'html.parser')
    state = extract_state(get_state_soup)

    while page_num < max_pages:
        for attempt in range(max_retries):
            try:
                response = construct_url(page_num, state)
                print("Page: ", page_num)
                soup = BeautifulSoup(response.text, 'html.parser')
                break
            except Exception as e:
                print(f"Failed to fetch page {page_num}: {e}")
                if attempt < max_retries - 1:
                    print(f"Retrying... ({attempt + 1}/{max_retries})")
                    time.sleep(5)
                else:
                    print(f"Giving up on page {page_num}")
                    results.append({
                        'ISIN': 'Page gives up',
                        'Ex-date': 'Page gives up',
                        'detail_url': f'Page {page_num} gives up'
                    })
                    break

        if 'soup' in locals():
            items = soup.select('ol.list.search-results li')
            if not items:
                break

            for item in items:
                title_element = item.select_one('h3 a')
                if title_element:
                    title = title_element.text.strip()
                    if any(keyword in title.lower() for keyword in ['xfra : dividend/interest information -']):
                        detail_url = 'https://www.xetra.com' + title_element['href']
                        try:
                            details = extract_details(detail_url)
                            if details:
                                details['detail_url'] = detail_url  # Correctly add the 'detail_url'
                                results.append(details)
                        except Exception as e:
                            print(f"Failed to extract details from {detail_url}: {e}")

        page_num += 1

    return results

def save_to_excel(data, filename='xetra_newsboard.xlsx'):
    df = pd.DataFrame(data)
    df = df.where(pd.notnull(df), 'NA')
    df_no_duplicates = df.drop_duplicates(subset=['ISIN', 'Ex-date'])
    df_no_duplicates.to_excel(filename, index=False)

if __name__ == '__main__':
    data = scrape_xetra_newsboard(max_pages=100)
    save_to_excel(data, 'Xetra_dividends.xlsx')
    print(f"Data saved to 'Xetra_dividends.xlsx'")