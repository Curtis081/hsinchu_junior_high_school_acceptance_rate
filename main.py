import os
import requests
from bs4 import BeautifulSoup
from xlsxwriter import Workbook
import pandas as pd
from styleframe import StyleFrame, Styler
import excel2img
import sys
import logging
import warnings


# Constants
ACCEPTANCE_RATE_COL_ORIGINAL = '錄取率'
ACCEPTANCE_RATE_COL = '錄取率%'

# Setup logging
logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("app.log"),
                        logging.StreamHandler(sys.stdout)
                    ])


def fetch_local_content(
        html_file_path='./111_academic_year_hsinchu_junior_high_school_acceptance_rate.html'):
    # Opening the html file
    html_file = open(html_file_path, "r", encoding='utf-8')
    # Reading the file
    response_text = html_file.read()
    # response_text = response.text
    soup = BeautifulSoup(response_text, "html.parser")
    return soup


def fetch_webpage_content(url):
    try:
        response = requests.get(url)
        soup = BeautifulSoup(response.text, "html.parser")

        post_title_html = soup.find('h1')
        if post_title_html is None:
            logging.error('404: Page not found. Could not find the title <h1> tag in the HTML.')
            return None, None

        title = post_title_html.text
        return soup, title
    except requests.RequestException as e:
        logging.error(f'Failed to fetch webpage content due to: {e}')
        return None, None


def create_and_save_excel(soup, title, output_folder):
    # Ignore DeprecationWarning
    warnings.filterwarnings("ignore", category=DeprecationWarning, module='openpyxl')

    try:
        excel_output_file_name = os.path.join(output_folder, title + '.xlsx')

        workbook = Workbook(excel_output_file_name)
        worksheet = workbook.add_worksheet()

        table_html = soup.find('tbody')
        td_htmls = table_html.find_all('td')

        for td_html in td_htmls:
            data_cell_id = td_html.get('data-cell-id')
            data_original_value = td_html.get('data-original-value')
            worksheet.write(data_cell_id, data_original_value)

        workbook.close()
        logging.info(f'Successfully created and saved Excel file: {excel_output_file_name}')
        return excel_output_file_name
    except Exception as e:
        logging.error(f'Failed to create and save Excel file due to: {e}')
        return None


def convert_acceptance_rate(excel_output_file_name):
    try:
        df = pd.read_excel(excel_output_file_name)

        acceptance_rate_slice_percentage = df[ACCEPTANCE_RATE_COL_ORIGINAL].str.slice(stop=-1)
        acceptance_rate = pd.to_numeric(acceptance_rate_slice_percentage, errors='coerce')
        df[ACCEPTANCE_RATE_COL] = acceptance_rate

        df = df.sort_values(by=ACCEPTANCE_RATE_COL, ascending=False)
        df = df.drop([ACCEPTANCE_RATE_COL_ORIGINAL], axis=1)

        logging.info(f'Successfully converted acceptance rates in the DataFrame.')
        return df, excel_output_file_name
    except Exception as e:
        logging.error(f'Failed to convert acceptance rates due to: {e}')
        return None


def style_and_write_to_excel(df, excel_output_file_name, sheet_name="Sheet1"):
    try:
        sf = StyleFrame(df, styler_obj=Styler(bg_color=None, bold=False, font='Arial', font_size=10.0, font_color=None,
                                              number_format='General', protection=False, underline=None,
                                              border_type='thin',
                                              horizontal_alignment='left', vertical_alignment='center', wrap_text=True,
                                              shrink_to_fit=True, fill_pattern_type='solid', indent=0.0,
                                              comment_author=None,
                                              comment_text=None, text_rotation=0))

        with StyleFrame.ExcelWriter(excel_output_file_name) as writer:
            sf.to_excel(writer, index=False, sheet_name=sheet_name, best_fit=list(df.columns.values))

        logging.info(f'Successfully styled and write DataFrame to Excel: {excel_output_file_name}')
        return sheet_name
    except Exception as e:
        logging.error(f'Failed to style and write DataFrame to Excel due to: {e}')


def convert_excel_to_png(excel_file_name, post_title_html, sheet_name, output_folder):
    output_png_file_name = os.path.join(output_folder, post_title_html + '.png')
    try:
        excel2img.export_img(excel_file_name, output_png_file_name, sheet_name, None)
        logging.info(f'Successfully converted {excel_file_name} to {output_png_file_name}')
    except Exception as e:
        logging.error(f'Error converting {excel_file_name} to {output_png_file_name}: {e}')


def create_output_folder():
    output_folder = 'output_files'
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    return output_folder


def hsinchu_junior_high_school_acceptance_rate(academic_year):
    url = f'https://shirley.tw/{academic_year}y-hsinchu-exam/'

    soup, post_title_html = fetch_webpage_content(url)
    if soup is None or post_title_html is None:
        return

    output_folder = create_output_folder()

    excel_file_name = create_and_save_excel(soup, post_title_html, output_folder)
    df, excel_output_file_name = convert_acceptance_rate(excel_file_name)

    sheet_name = style_and_write_to_excel(df, excel_output_file_name)
    convert_excel_to_png(excel_output_file_name, post_title_html, sheet_name, output_folder)

    print('Process completed successfully.')


if __name__ == '__main__':
    for academic_year in range(110, 113):
        try:
            hsinchu_junior_high_school_acceptance_rate(str(academic_year))
        except Exception as e:
            logging.error(f'Error processing academic year {academic_year}: {e}')
            continue
