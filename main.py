import glob
import os

from utils import utils
from utils import LoadExcelFile


def main():
    print('### LoadExcelFile start ###')
    LoadExcelFile.execute()
    print('### LoadExcelFile end ###')

if __name__ == "__main__":
    # calling main function
    main()