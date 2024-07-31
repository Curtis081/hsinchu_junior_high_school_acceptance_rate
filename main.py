import logging
from acceptance_rate_processor import AcceptanceRateProcessor

if __name__ == '__main__':
    for academic_year in range(110, 113):
        try:
            processor = AcceptanceRateProcessor(str(academic_year))
            processor.process()
        except Exception as e:
            logging.error(f'Error processing academic year {academic_year}: {e}')
            continue
