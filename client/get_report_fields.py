from googleads import adwords

REPORT_TYPE = 'CAMPAIGN_PERFORMANCE_REPORT'


def main(client, report_type):
    # Initialize appropriate service.
    report_definition_service = client.GetService(
        'ReportDefinitionService', version='v201809')

    # Get report fields.
    fields = report_definition_service.getReportFields(report_type)

    # Display results.
    print('Report type "%s" contains the following fields:' % report_type)
    for field in fields:
        print(' - %s (%s)' % (field['fieldName'], field['fieldType']))
        if 'enumValues' in field:
            print('  := [%s]' % ', '.join(field['enumValues']))


if __name__ == '__main__':
    # Initialize client object.
    adwords_client = adwords.AdWordsClient.LoadFromStorage()

    main(adwords_client, REPORT_TYPE)
