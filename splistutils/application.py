
import logging
import pandas as pd
import requests.exceptions

from pprint import pprint
from itertools import islice
from openpyxl import load_workbook


# Define the local logger
log = logging.getLogger(__name__)

######################################################################################
# !!!!! REDICULOUS KLUDGE WARNING !!!!!
# THIS NEEDS TO MOVE TO WHEREVER THE SHAREPOINT CODE WILL LIVE

# --- Monkey Patch #1 
import importlib
import shareplum.errors

# More robust exception handling -- add nested exception data
# for additional processing details
class _ShareplumError(Exception):
    def __init__(self, msg, details=None):
        self.details = details
        if details:
            super().__init__(f"{msg} : {details}")
        else:
            super().__init__(msg)

# Monkey patch the exception class with a better handler
shareplum.errors.ShareplumError = _ShareplumError
ShareplumError = _ShareplumError

class _ShareplumRequestError(ShareplumError):
    pass

# Monkey patch the actual Request exception so it properly picks
# up the modified base class. 
shareplum.errors.ShareplumRequestError = _ShareplumRequestError
ShareplumRequestError = _ShareplumRequestError

importlib.reload(shareplum.request_helper)

# --- Monkey Patch #2
import requests
import shareplum.request_helper

# Excluded the raise_for_status() call to work around a spurious
# 403 error that is thrown by Purdue's tenant. No idea why, but
# allows me to continue beyond it if I ignore it. Unfortunately,
# I am now ignoring ANY HTTP error situations... 
def _post(session, url, **kwargs):
    try:
        response = session.post(url, **kwargs)
        return response
    except requests.exceptions.RequestException as err:
        raise ShareplumRequestError("Shareplum HTTP Post Failed", err)

# Monkey patch the post method inside the request_helper module
shareplum.request_helper.post = _post

# Reload the office365 module to recognize the patched method
importlib.reload(shareplum.office365)

######################################################################################

from shareplum import Site
from shareplum import Office365
from shareplum.site import Version



class SharePointListUtils:


    @staticmethod
    def run():

        wb = load_workbook(filename='data.xlsx', read_only=True)
        ws = wb['Raw Data']

        data = ws.values
        columns = next(data)[1:]
        data = list(data)
        idx = [r[0] for r in data]
        data = (islice(r, 1, None) for r in data)
        df = pd.DataFrame(data, index=idx, columns=columns)


        # print(df.loc[[1130359]])

        authcookie = Office365('https://purdue0.sharepoint.com', username='', password='').GetCookies()
        site = Site('https://purdue0.sharepoint.com/sites/HRIS', version=Version.o365, authcookie=authcookie)


        test_list = site.List('ELT-Test')

        # while True:
        #     list_items = test_list.GetListItems('All Items', fields=['ID'], row_limit=500)

        #     if len(list_items) == 0:
        #         break

        #     id_list = [x['ID'] for x in list_items]

        #     log.info('Starting deletion of {} records'.format(str(len(list_items))))
        #     test_list.UpdateListItems(data=id_list, kind='Delete')
        #     log.info('Deletion complete.')


        # print (len(list_items))

        list_items = []
        try:
            list_items = test_list.GetListItems('All Items', 
                fields=['ID', 'Employee Name'], 
                # query={'Where': ['Or', ('Eq', 'Employee Name', 'Mark Holmes'), ('Eq', 'Employee Name', 'Patricia Prince')]},
                query={'Where': [('Eq', 'Employee Name', 'Mark Holmes')]},
                )
        except shareplum.errors.ShareplumRequestError as err:
            log.error(err)
            if err.details and type(err.details) == requests.exceptions.HTTPError:
                if err.details.response.status_code in [429, 503]:
                    # TODO: Sleep for Retry-After to prevent further throttling
                    pass
                elif err.details.response.status_code in [500]:
                    log.error(err.details.response.request.body)
                    log.error(err.details.response.content)

        for list_item in list_items:
            log.info(list_item)

     