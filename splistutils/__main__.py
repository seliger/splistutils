
import logging
import atexit
import sys

from splistutils.application import SharePointListUtils


# Helper function to ensure we capture stack traces via logging
# (Note: This method does not work with threads until Python 3.8)
def log_excepthook(excType, excValue, traceback, logger=logging.getLogger()):
    logging.error("Logging an uncaught exception",
                 exc_info=(excType, excValue, traceback))

# Helper function to clean up and alert that we are ending runtime
@atexit.register
def shutdown():
    logging.info("SharePoint List Utilities - Shutting down...")

# Stub code to bootstrap environment and launch into main application
def run():
    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(name)s] [%(funcName)s] [%(levelname)s]  %(message)s",
        handlers=[
            logging.StreamHandler(sys.stderr),
            logging.FileHandler("{0}/{1}.log".format("./", "splistutils"))
        ]
    )
    
    # Ensure we capture stack traces in our log
    sys.excepthook = log_excepthook 

    # Announce that we are starting up.
    logging.info('SharePoint List Utilities - Starting up...')

    # Launch into the main application run code
    SharePointListUtils.run()

if __name__ == '__main__':
    run()
