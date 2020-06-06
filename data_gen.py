import logging
import sys

from faker import Faker
from random import seed
from random import randint

from shareplum import Site
from shareplum import Office365
from shareplum.site import Version

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] [%(funcName)s] [%(levelname)s]  %(message)s",
    handlers=[
        logging.StreamHandler(sys.stderr),
    ]
)


######################################################################################
# !!!!! REDICULOUS KLUDGE WARNING !!!!!
# THIS NEEDS TO MOVE TO WHEREVER THE SHAREPOINT CODE WILL LIVE
import importlib
import requests
import shareplum.request_helper
from shareplum.errors import ShareplumRequestError

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


buildings = ['ABE', 'ADDL', 'AERO', 'AEV', 'AF01', 'AF02', 'AF08', 'AGAD', 'AHF', 'AQUA', 'AR', 'ARST', 'ASA', 'ASB', 
    'ASB1', 'ASTL', 'BBAL', 'BCC', 'BRK', 'BMSE', 'BMSN', 'BMSW', 'BOWN', 'BRNG', 'BRWN', 'BSG', 'BTC', 'CARL', 'CDFS', 
    'CHAF', 'FRNY', 'CIND', 'CL50', 'CMBR', 'COMP', 'CQE', 'CQNE', 'CQNW', 'CQS', 'CQW', 'DANL', 'DAUC', 'DMNT', 'DOYL', 
    'DUHM', 'EE', 'EEL', 'EHSB', 'ELLT', 'ENAD', 'ERHT', 'FBSS', 'FOPN', 'FWLR', 'GAS', 'GCMB', 'GDRH', 'GRIS', 'GRSB', 
    'GRVL', 'GSMB', 'HA01', 'HA02', 'HA03', 'HA04', 'HA05', 'HA06', 'HA07', 'HA08', 'HA09', 'HA10', 'HA11', 'HA12', 
    'HA13', 'HA14', 'HA15', 'HA16', 'HA17', 'HA18', 'HA19', 'HA20', 'HA21', 'HA22', 'HA23', 'HA24', 'HA25', 'HA26', 
    'HA27', 'HA28', 'HA29', 'HA30', 'HA31', 'HA32', 'HANS', 'HARR', 'HAWK', 'HEAV', 'HGR4', 'HGR5', 'HGR6', 'HGRH', 
    'HIKS', 'HILL', 'HMMT', 'HOVD', 'HPN', 'HRTP', 'INSS', 'JNSN', 'KCTR', 'KNOY', 'KRAN', 'LAMB', 'LILY', 'LINS', 
    'LMSA', 'LMSB', 'LMST', 'LSA', 'LSPS', 'LSR', 'LYNN', 'MACK', 'MATH', 'MCUT', 'ME', 'MGL', 'MLAB', 'MMDC', 'MMS1', 
    'MMS2', 'MMS3', 'MOLL', 'MRDH', 'MRGN', 'MSEE', 'MTHW', 'NUCL', 'NWCP', 'NWSS', 'OLMN', 'OWEN', 'PEST', 'PFEN', 
    'PFSB', 'PGG', 'PGGH', 'PGM', 'PGMD', 'PGNW', 'PGU', 'PIT1', 'PIT2', 'PIT3', 'PMU', 'PMUC', 'POAN', 'POTR', 'PRCE', 
    'PSYC', 'PUSH', 'PVAB', 'PVCC', 'RAIL', 'RALR', 'RAP', 'RAWL', 'REC', 'REMS', 'RHPH', 'SALT', 'SAT', 'SC', 'SCCA', 
    'SCCB', 'SCCC', 'SCCD', 'SCCE', 'SCHL', 'SCPA', 'SEAN', 'SHLY', 'SHRV', 'SIML', 'SMLY', 'SMTH', 'SOCR', 'SRR1', 
    'SSA1', 'SSOF', 'STEW', 'STON', 'TARK', 'TEL', 'TER1', 'TERM',  
    'TMB', 'TPB', 'TRAM', 'TRFM', 'TTS', 'UNIV', 'UPOB', 'UPOF', 'UPSB', 'VA1', 'VA2', 'VAWT', 'VBPB', 'VCPR', 'VIC', 
    'VLAB', 'VOIN', 'VPRB', 'VPTH', 'VSPB', 'VTSB', 'WADE', 'WARN', 'WILY', 'WOOD', 'WSLR', 'WTHR', 'WWG', 'ZL1', 
    'ZL2', 'ZL3', 'ZL4', 'ZL5', 'ZS01', 'ZS04', 'ZS05', 'ZS09', 'ZS10', 'VMIF', 'RAPP', 'HGR8', 'SPUR', 'TURF', 'SOIL', 
    'CHAO', 'EHSA', 'KENT', 'LCC', 'FORD', 'HAAS', 'MANN', 'LWSN', 'PGW', 'SCHW', 'ARMS', 'NISW', 'COAL', 'CSF', 'MJIS', 
    'PAO', 'YONG', 'FSTC', 'FSTE', 'HOCK', 'TRNR', 'MRRT', 'HNLY', 'TREC', 'DLR', 'GMGF', 'GMPS', 'NLSN', 'GMF', 'ADM', 
    'BRES', 'VFH', 'CREC', 'HAMP', 'DRUG', 'LYLE', 'HLAB', 'HERL', 'BALY', 'KRCH', 'WANG', 'AACC', 'NACC', 'PAGE', 'ECEC', 
    'PTCA', 'WALC', 'PGH', 'CRTN', 'LOLC', 'FPC', 'CEPF', 'FLEX', 'BIDC', 'CAI']

seed(10101)

emp_mgt_classes = ['M/P Management', 'Executive', 'Faculty', 'Management']
emp_classes = ['Resident', 'Professional', 'Police/Fire Hourly', 'Fellowship Pre Doc', 'Post Doc', 
    'LTD', 'PAA/PRF/PolyTech', 'Police/Fire BW Sal', 'Police/Fire Admin', 'Non-Pay', 'Continuing Lecturer', 
    'Visiting Faculty', 'Graduate Student', 'Temporary', 'Police/Fire Mgmt', 'Limited Term Lecturer', 'M/P Professional', 
    'Service', 'Support', 'Clinical/Research', 'Intern', 'Residence Hall Counselor', 'Student']


logging.info('Starting up...')

fake = Faker()


authcookie = Office365('https://purdue0.sharepoint.com', username='', password='').GetCookies()
site = Site('https://purdue0.sharepoint.com/sites/HRIS', version=Version.o365, authcookie=authcookie)

test_list = site.List('ELT-Test')


employees = []
supervisor = None
count = 0
for x in range(1, 10001):
    if supervisor == None or x % randint(3,7) == 0:
        # print ("SUPER: {}\t COUNT: {}".format(supervisor, str(count)))
        count = 1
        supervisor = fake.name()
    else:
        count += 1

    employee = {}
    employee['Title'] = str(x)
    employee['Employee PERNR'] = str(x)
    employee['Employee Name'] = fake.name()
    employee['Supervisor Name'] = supervisor
    employee['Employee Class'] =  emp_classes[randint(0, len(emp_classes) - 1)]
    employee['Editor Name'] = supervisor
    employee['Division Code'] = 'DIVCODE'
    employee['Division'] = 'Some Division Name'
    employee['Department Code'] = 'DEPTCODE'
    employee['Department Name'] = 'Some Department Name'
    employee['Building Code'] = buildings[randint(0, len(buildings) - 1)]

    employees.append(employee)

    # print('{}, {}, {}, {}, {}, {}, {}, {}, {}, {}'.format(*list(employee.values())))

    if x % 500 == 0:
        logging.info ('Flushing records to SharePoint at count {}.'.format(str(x)))
        test_list.UpdateListItems(data=employees, kind='New')
        employees = []
        logging.info ('Flush complete.')
    

logging.info('Shutting down...')

