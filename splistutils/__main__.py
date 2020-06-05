
import logging
import sys

from splistutils.application import SharePointListUtils


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
    

    # Launch into the main application run code
    SharePointListUtils.run()

if __name__ == '__main__':
    run()
