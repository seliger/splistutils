import logging


# Define the local logger
log = logging.getLogger(__name__)


######################################################################################
# !!!!! REDICULOUS KLUDGE WARNING !!!!!
# Several required patches for Shareplum

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


class SharePointSite():
    def __init__(self, sharepoint_url, site_name, username, password):
        self.sharepoint_url = sharepoint_url
        self.username = username

        self.authcookie = self.__login(sharepoint_url, username, password)
        self.site = Site("{}/sites/{}".format(sharepoint_url, site_name), version=Version.v365, authcookie=self.authcookie)


    def __login(self, sharepoint_url, username, password):
        authcookie = Office365(sharepoint_url, username=username, password=password).GetCookies()
        return authcookie


class SharePointList():
    def __init__(self, sharepoint_site, list_name):
        self.sharepoint_site = sharepoint_site
        self.list_name = list_name

        self.contents = sharepoint_site.site.List(list_name)
        
