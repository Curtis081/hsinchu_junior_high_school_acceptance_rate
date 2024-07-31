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


class AcceptanceRateProcessor:
    ACCEPTANCE_RATE_COL_ORIGINAL = '錄取率'
    ACCEPTANCE_RATE_COL = '錄取率%'

    def __init__(self, academic_year, output_folder='output_files'):
        self.academic_year = academic_year
        self.output_folder = output_folder
        self.url = f'https://shirley.tw/{self.academic_year}y-hsinchu-exam/'
        self.excel_file_name = None
        self.post_title_html = None
        self.df = None

        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s',
                            handlers=[
                                logging.FileHandler("app.log"),
                                logging.StreamHandler(sys.stdout)
                            ])
        self._create_output_folder()

    def _create_output_folder(self):
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)

    def fetch_local_content(self, html_file_path='./111_academic_year_hsinchu_junior_high_school_acceptance_rate.html'):
        try:
            with open(html_file_path, "r", encoding='utf-8') as html_file:
                response_text = html_file.read()
            soup = BeautifulSoup(response_text, "html.parser")
            return soup
        except Exception as e:
            logging.error(f'Failed to fetch local content due to: {e}')
            return None

    def fetch_webpage_content(self):
        try:
            response = requests.get(self.url)
            soup = BeautifulSoup(response.text, "html.parser")

            post_title_html = soup.find('h1')
            if post_title_html is None:
                logging.error('404: Page not found. Could not find the title <h1> tag in the HTML.')
                return None, None

            self.post_title_html = post_title_html.text
            return soup, self.post_title_html
        except requests.RequestException as e:
            logging.error(f'Failed to fetch webpage content due to: {e}')
            return None, None

    def create_and_save_excel(self, soup):
        # Ignore DeprecationWarning
        warnings.filterwarnings("ignore", category=DeprecationWarning, module='openpyxl')

        try:
            self.excel_file_name = os.path.join(self.output_folder, self.post_title_html + '.xlsx')

            workbook = Workbook(self.excel_file_name)
            worksheet = workbook.add_worksheet()

            table_html = soup.find('tbody')
            td_htmls = table_html.find_all('td')

            for td_html in td_htmls:
                data_cell_id = td_html.get('data-cell-id')
                data_original_value = td_html.get('data-original-value')
                worksheet.write(data_cell_id, data_original_value)

            workbook.close()
            logging.info(f'Successfully created and saved Excel file: {self.excel_file_name}')
            return self.excel_file_name
        except Exception as e:
            logging.error(f'Failed to create and save Excel file due to: {e}')
            return None

    def convert_acceptance_rate(self):
        try:
            self.df = pd.read_excel(self.excel_file_name)

            acceptance_rate_slice_percentage = self.df[self.ACCEPTANCE_RATE_COL_ORIGINAL].str.slice(stop=-1)
            acceptance_rate = pd.to_numeric(acceptance_rate_slice_percentage, errors='coerce')
            self.df[self.ACCEPTANCE_RATE_COL] = acceptance_rate

            self.df = self.df.sort_values(by=self.ACCEPTANCE_RATE_COL, ascending=False)
            self.df = self.df.drop([self.ACCEPTANCE_RATE_COL_ORIGINAL], axis=1)

            logging.info('Successfully converted acceptance rates in the DataFrame.')
            return self.df, self.excel_file_name
        except Exception as e:
            logging.error(f'Failed to convert acceptance rates due to: {e}')
            return None, None

    def style_and_write_to_excel(self, sheet_name="Sheet1"):
        try:
            sf = StyleFrame(self.df, styler_obj=Styler(bg_color=None, bold=False, font='Arial', font_size=10.0, font_color=None,
                                                      number_format='General', protection=False, underline=None, border_type='thin',
                                                      horizontal_alignment='left', vertical_alignment='center', wrap_text=True,
                                                      shrink_to_fit=True, fill_pattern_type='solid', indent=0.0, comment_author=None,
                                                      comment_text=None, text_rotation=0))

            with StyleFrame.ExcelWriter(self.excel_file_name) as writer:
                sf.to_excel(writer, index=False, sheet_name=sheet_name, best_fit=list(self.df.columns.values))

            logging.info(f'Successfully styled and wrote DataFrame to Excel: {self.excel_file_name}')
            return sheet_name
        except Exception as e:
            logging.error(f'Failed to style and write DataFrame to Excel due to: {e}')
            return None

    def convert_excel_to_png(self, sheet_name):
        output_png_file_name = os.path.join(self.output_folder, self.post_title_html + '.png')
        try:
            excel2img.export_img(self.excel_file_name, output_png_file_name, sheet_name, None)
            logging.info(f'Successfully converted {self.excel_file_name} to {output_png_file_name}')
        except Exception as e:
            logging.error(f'Error converting {self.excel_file_name} to {output_png_file_name}: {e}')

    def process(self):
        soup, post_title_html = self.fetch_webpage_content()
        if soup is None or post_title_html is None:
            return

        excel_file_name = self.create_and_save_excel(soup)
        if excel_file_name is None:
            return

        df, excel_output_file_name = self.convert_acceptance_rate()
        if df is None or excel_output_file_name is None:
            return

        sheet_name = self.style_and_write_to_excel()
        if sheet_name is None:
            return

        self.convert_excel_to_png(sheet_name)
        print('Process completed successfully.')
