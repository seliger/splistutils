
import logging
import sys

from .application import SharePointListUtils

if __name__ == '__main__':

    # Configure logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(name)s] [%(funcName)s] [%(levelname)s]  %(message)s",
        handlers=[
            logging.StreamHandler(sys.stderr),
            logging.FileHandler("{0}/{1}.log".format("./", "splistutils"))
        ]
    )
    

    # Launch into the main application run code
    SharePointListUtils.run()
    