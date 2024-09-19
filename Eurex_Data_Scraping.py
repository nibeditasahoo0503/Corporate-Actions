import requests
from bs4 import BeautifulSoup
import pdfplumber
import re
import os
from datetime import datetime
import pandas as pd
from text2digits import text2digits

def get_corporate_action_terms(soup):
    terms_div = soup.select_one('.dbx-richtext')
    if terms_div:
        corporate_action_terms = terms_div.get_text(strip=True)
        terms_match = re.search(r'(\d+(\.\d+)?\s*[:]\s*\d+(\.\d+)?|\d+(\.\d+)?\s*-for-\s*\d+(\.\d+)?|(\w+)-for-(\w+)|\b\w+\b\s+ratio(?:\s+\w+)*\s+(\d+)(?:\s+\w+)*\s+(\d+)(?:\s+\w+)*\b|\s*(\d+(?:[.,]\d+)?:(?:\d+(?:[.,]\d+)?(?:,\d+(?:[.,]\d+)?)?)))', corporate_action_terms)
        if terms_match:
            original_terms = terms_match.group(0)
            if '-' in original_terms:
                return convert_terms_to_ratio(original_terms)
            elif 'ratio' in original_terms:
                return convert_to_ratio(original_terms)
            else:
                return original_terms
    return 'N/A'

def convert_to_ratio(original_terms):
  words_to_check = {"for", "per", "into"}
  words_in_sentence = set(original_terms.lower().split())
  if any(word in words_in_sentence for word in words_to_check):
    parts = re.search(r'.*ratio\s+.*?(\d+)\s+.*?(\d+)', original_terms)
    if parts:
      number1 = parts.group(1)
      number2 = parts.group(2)
      return f"{number1}:{number2}"
  return "N/A"

def convert_terms_to_ratio(terms):
    parts = re.split('-for-|-', terms)
    if len(parts) == 2:
        return f"{parts[0]}:{parts[1]}"
    return terms

def extract_isins_from_pdf(soup, base_url):
    pdf_links = soup.select('.dbx-linklist__item a')
    if pdf_links:
        pdf_url = base_url + pdf_links[-1]['href']
        pdf_response = requests.get(pdf_url)
        pdf_path = 'temp.pdf'
        with open(pdf_path, 'wb') as f:
            f.write(pdf_response.content)
        try:
            with pdfplumber.open(pdf_path) as pdf:
              all_isins = []
              pdf_text = ""
              for page in pdf.pages:
                text = page.extract_text()
                pdf_text += text
                isin_match = re.findall(r'\b[A-Z]{2}[A-Z0-9]{10}\b', text)
                all_isins.extend(isin_match)
              isin_old, isin_new = 'N/A', 'N/A'
              text_combined = ", ".join(all_isins)
              if text_combined in pdf_text:
                isin_old = text_combined
                isin_new = text_combined
              elif len(all_isins) == 1:
                  isin_old = all_isins[0]
                  isin_new = all_isins[0]
              elif len(all_isins) == 2:
                isin_old = all_isins[0]
                isin_new = all_isins[1]
              elif len(all_isins) > 2:
                if all_isins[0] == all_isins[1]:
                  isin_old = all_isins[1]
                  isin_new = all_isins[2]
                elif all_isins[0] != all_isins[1]:
                  isin_old = all_isins[0]
                  isin_new = all_isins[1]
              else:
                  isin_old, isin_new = 'N/A', 'N/A'
        except Exception as e:
            print(f"Error extracting ISINs from PDF: {e}")
            isin_old, isin_new = 'N/A', 'N/A'
        os.remove(pdf_path)
    else:
        isin_old, isin_new = 'N/A', 'N/A'
    return isin_old, isin_new

def extract_isins_from_webpage(soup):
    isin_pattern = re.compile(r'\b[A-Z]{2}[A-Z0-9]{10}\b')
    tables = soup.select('.tableWrapper .dataTable')
    for table in tables:
        rows = table.select('tr')
        if len(rows) > 1:
            cells = rows[1].select('td')
            if len(cells) >= 3:
                isin_old = cells[1].get_text(strip=True)
                isin_new = cells[2].get_text(strip=True)
                if isin_pattern.fullmatch(isin_old) and isin_pattern.fullmatch(isin_new):
                    return isin_old, isin_new
    return 'N/A', 'N/A'

def extract_isins(soup, base_url):
    isin_old, isin_new = extract_isins_from_webpage(soup)
    if isin_old == 'N/A' or isin_new == 'N/A':
        isin_old, isin_new = extract_isins_from_pdf(soup, base_url)
    return isin_old, isin_new

def get_effective_date(soup):
    date_span = soup.select_one('.dbx-tagline-date__topline span')
    if date_span:
        date_str = date_span.text.strip()
        date_obj = datetime.strptime(date_str, '%d %b %Y')
        return date_obj.strftime('%Y-%m-%d')
    return 'N/A'

def get_corporate_action_type(title):
    if 'share consolidation' in title.lower():
        return 'Reverse Stock Split'
    elif 'reverse stock split' in title.lower():
        return 'Reverse Stock Split'
    elif 'stock split' in title.lower():
        return 'Stock Split'
    elif 'reverse split' in title.lower():
        return 'Reverse Stock Split'
    else:
        return 'N/A'

def adjust_ratio_based_on_action_type(ratio, corporate_action_type):
    left, right = map(float, ratio.split(':'))
    if corporate_action_type == 'Stock Split':
        if left > right:
            left, right = right, left
    elif corporate_action_type == 'Reverse Stock Split':
        if left < right:
            left, right = right, left
    left = int(left) if left.is_integer() else left
    right = int(right) if right.is_integer() else right
    return f"{left}:{right}"

def extract_data(title, soup, base_url, full_url):
    security_name = title.split(':')[0].strip() or title.split('-')[0].strip() or title.split(',')[0].strip() or title
    corporate_action_type = get_corporate_action_type(title)
    corporate_action_terms = get_corporate_action_terms(soup)
    if corporate_action_terms != "N/A":
      t2d = text2digits.Text2Digits()
      spaces = corporate_action_terms.replace(":", " : ")
      result = t2d.convert(spaces)
      result = result.replace(" : ", ":")
      final_corporate_action_terms = adjust_ratio_based_on_action_type(result, corporate_action_type)
    else:
      final_corporate_action_terms = corporate_action_terms
    effective_date = get_effective_date(soup)
    isin_old, isin_new = extract_isins(soup, base_url)
    return {
        'ISIN_old': isin_old,
        'ISIN_new': isin_new,
        'security_name': security_name,
        'effective_date': effective_date,
        'corporate_action_type': corporate_action_type,
        'corporate_action_terms': final_corporate_action_terms,
        'detail_url': full_url
    }

def filter_and_extract_action_data(start_url, total_pages):
    data_list = []
    base_url = "https://www.eurex.com"
    for page_num in range(total_pages):
        print(page_num)
        page_url = f"{start_url}?query=&state=&pageNum={page_num}"
        response = requests.get(page_url)
        soup = BeautifulSoup(response.text, 'html.parser')
        action_links = soup.select('.teasable-search-result-link')
        for link in action_links:
            full_url = base_url + link['href']
            title = link.find('h1', class_='search-result-description').text.lower()
            if any(keyword in title for keyword in ['share consolidation', 'reverse stock split', 'stock split', 'reverse split']):
                response = requests.get(full_url)
                soup = BeautifulSoup(response.text, 'html.parser')
                data = extract_data(title, soup, base_url, full_url)
                data_list.append(data)
    return data_list

def get_total_pages(start_url):
    response = requests.get(start_url)
    soup = BeautifulSoup(response.text, 'html.parser')
    pagination_container = soup.select_one('.pagination-list')
    if not pagination_container:
        print("Pagination container not found.")
        return 0
    pagination_buttons = pagination_container.select('li.pagination-list-element button.pagination-element')
    page_numbers = []
    for btn in pagination_buttons:
        text = btn.get_text(strip=True)
        if text.isdigit():
            page_numbers.append(int(text))
    if page_numbers:
        return max(page_numbers)
    else:
        print("No page numbers found in pagination.")
        return 0

def main():
    start_url = "https://www.eurex.com/ex-en/rules-regs/corporate-actions/corporate-action-information/3656!search"
    start_url2 = "https://www.eurex.com/ex-en/rules-regs/corporate-actions/corporate-action-information"

    total_pages = get_total_pages(start_url2)

    total_pages = 20

    data_list = filter_and_extract_action_data(start_url, total_pages)

    df = pd.DataFrame(data_list)
    df.to_excel('eurex_newsboard.xlsx', index=False)
    print("Data saved to eurex_newsboard.xlsx")

if __name__ == "__main__":
    main()