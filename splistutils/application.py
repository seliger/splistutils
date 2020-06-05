
import logging

import pandas as pd

from openpyxl import load_workbook
from pprint import pprint
from itertools import islice


log = logging.getLogger(__name__)

class SharePointListUtils:


    @staticmethod
    def run():
        log.info("SharePoint List Utilities - Starting up...")

        wb = load_workbook(filename='data.xlsx', read_only=True)
        ws = wb['Raw Data']

        data = ws.values
        columns = next(data)[1:]
        data = list(data)
        idx = [r[0] for r in data]
        data = (islice(r, 1, None) for r in data)
        df = pd.DataFrame(data, index=idx, columns=columns)


        print(df.loc[[1130359]])


     