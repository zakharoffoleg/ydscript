import time

import clientSum
import costsDownload
import gsheets
import metrika


def export():
    start_time = time.time()
    costsDownload.getCosts()
    metrika.saveConversions()
    metrika.connectConversionsWithCampaigns()
    # time.sleep(3)
    clientSum.importValues()
    # time.sleep(3)
    gsheets.uploadData()
    print("\n--- %s seconds ---" % (time.time() - start_time))
    return


export()
