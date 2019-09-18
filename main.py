import time

import costsDownload
import gsheets
import metrika


def main():
    start_time = time.time()
    costsDownload.getCosts()
    metrika.saveConversions()
    metrika.connectConversionsWithCampaigns()
    gsheets.uploadData()
    print("\n--- %s seconds ---" % (time.time() - start_time))


main()
