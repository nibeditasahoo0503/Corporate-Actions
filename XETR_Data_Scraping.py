import requests
from bs4 import BeautifulSoup
import re
from datetime import datetime
import pandas as pd
import time

def construct_url(page_num):
    base_url = 'https://www.xetra.com/xetra-en/newsroom/xetra-newsboard/4442!search'
    params = {
        'state': 'H4sIAAAAAAAAAFWOQQvCMAyF_4rkvMO89igiCB4mE--lzbTQtZikyBj775a5su2W996X5I1gteCFYg8qJO-rWT9iUZ02KAxqnPLsiOWGIkglfjvhBqnRLwR1rOsKXDA-WWydIBcqBj80tgPVac9YwSchDaAAKiDk5OXp8FtgjiQ543PucbDIJlMmscT-lERiWMrMN652u7WqhdeM9_-rrd1m9IHU790Z3H1Y6Y09_QDywux2MQEAAA',
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

def safe_convert(value):
    try:
        return float(value.replace(',', '.'))
    except ValueError:
        return None

def extract_details_from_segment(text_segment, isin_old, isin_new):
    details = {}
    details['ISIN_old'] = isin_old
    details['ISIN_new'] = isin_new
    details['security_name'] = 'NA'
    details['effective_date'] = 'NA'
    details['corporate_action_type'] = 'NA'
    details['corporate_action_terms'] = 'NA'

    parts = re.split(r'(\d{1,2}[./-]\d{1,2}[./-]\d{4})', text_segment)
    if len(parts) >= 2:
        details['security_name'] = parts[0].strip()
        effective_date_str = parts[1].strip()
        details['effective_date'] = parse_date(effective_date_str)
        ratio_pattern = r'Tausch\s+(\d+(?:[.,]\d+)?:(?:\d+(?:[.,]\d+)?(?:,\d+(?:[.,]\d+)?)?))'
        ratio_match = re.search(ratio_pattern, text_segment)
        if ratio_match:
            split_ratio = ratio_match.group(1)
            ratio = split_ratio.split(':')
            if len(ratio) == 2:
                ratio_0 = safe_convert(ratio[0])
                ratio_1 = safe_convert(ratio[1])
                if ratio_0 is not None and ratio_1 is not None:
                    details['corporate_action_type'] = 'Stock Split' if ratio_0 <= ratio_1 else 'Reverse Stock Split'
                    details['corporate_action_terms'] = split_ratio

    return details

def extract_details_for_isins(text, isins):
    details_list = []
    num_isins = len(isins)
    for i in range(0, num_isins - 1, 2):
        isin_old = isins[i]
        isin_new = isins[i+1]
        text_after_second_isin = text.split(isin_new, 1)[1]
        details = extract_details_from_segment(text_after_second_isin, isin_old, isin_new)
        details_list.append(details)
    return details_list

def extract_details(detail_url):
    response = requests.get(detail_url)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, 'html.parser')
    details = []
    title = soup.find('h2', class_='main-title').text.strip()
    if 'ISIN Change' not in title:
        return []
    text = soup.find('div', class_='detailText').get_text(separator=' ').strip()
    isins = re.findall(r'\b[A-Z]{2}(?=[A-Z0-9]*\d)[A-Z0-9]{10}\b', text)
    if len(isins) < 2:
        return []
    details = extract_details_for_isins(text, isins)
    return details

def scrape_xetra_newsboard(max_pages):
    results = []
    page_num = 0
    max_retries = 3
    while page_num < max_pages:
        for attempt in range(max_retries):
            try:
                response = construct_url(page_num)
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
                        'ISIN_old': 'Page gives up',
                        'ISIN_new': 'Page gives up',
                        'security_name': 'Page gives up',
                        'effective_date': 'Page gives up',
                        'corporate_action_type': 'Page gives up',
                        'corporate_action_terms': 'Page gives up',
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
                    if any(keyword in title.lower() for keyword in ['xfra : isin change', 'isin change', 'isin-change']):
                        detail_url = 'https://www.xetra.com' + title_element['href']
                        try:
                            details = extract_details(detail_url)
                            if details:
                                for detail in details:
                                    detail['detail_url'] = detail_url
                                    results.append(detail)
                        except Exception as e:
                            print(f"Failed to extract details from {detail_url}: {e}")

        page_num += 1

    return results

def save_to_excel(data, filename='xetra_newsboard.xlsx'):
    df = pd.DataFrame(data)
    df = df.where(pd.notnull(df), 'NA')
    df_no_duplicates = df.drop_duplicates(subset=['ISIN_old', 'ISIN_new', 'effective_date', 'corporate_action_terms'])
    df_no_duplicates.to_excel(filename, index=False)

if __name__ == '__main__':
    data = scrape_xetra_newsboard(max_pages=1000)
    save_to_excel(data, 'xetra_newsboard.xlsx')
    print(f"Data saved to 'xetra_newsboard.xlsx'")