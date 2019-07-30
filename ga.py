# -*- coding: utf-8 -*-
import sys

from googleads import adwords


# billboadbot tokens
# access_token: ya29.GltVB1CZMI5EeKtudN4U1Q7UkyLkO0sowci47y3BsGXhPRVB7mvX3dN4QhMNdTjaEktaeP0VuSD91O3Q2wZYhnZ6egTr1uDFQ3OyoRVxy9tXBM6RbfOFedhi73OP
# refresh_token: 1/Jp1a51XGYcGLy3L_klNsDUU_iZhK6xd8-r3SARfIGEQ


# billboadbottest tokens
# Access token: ya29.GltVBwf6YydBqXGapPlRWzK9Ohw0l1pxEgGtWCuV2X88e6pcqE3PPrCzqJtNfqPcnx5RW3rBMO74-dS6DwdAIAVvDD2ZcwYwaJT-GEAdYTtw_loepaIV8xA1l_u_

def main(client):
    # Initialize appropriate service.
    report_downloader = client.GetReportDownloader(version='v201809')

    # Create report query.
    report_query = (adwords.ReportQueryBuilder()
                    .Select('CampaignId', 'Cost')
                    .From('CRITERIA_PERFORMANCE_REPORT')
                    .Where('Status').In('ENABLED', 'PAUSED')
                    .During('LAST_7_DAYS')
                    .Build())

    # You can provide a file object to write the output to. For this
    # demonstration we use sys.stdout to write the report to the screen.
    report_downloader.DownloadReportWithAwql(
        report_query, 'CSV', sys.stdout, skip_report_header=False,
        skip_column_header=False, skip_report_summary=False,
        include_zero_impressions=True)


if __name__ == '__main__':
    # Initialize client object.
    adwords_client = adwords.AdWordsClient.LoadFromStorage()

    main(adwords_client)
