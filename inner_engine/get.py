from openpyxl import Workbook, load_workbook, styles
from hashlib import sha256
from random import choice
from string import ascii_lowercase, ascii_uppercase, digits
from datetime import datetime
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

standard_name: str = 'DataBase'
file_name: str = ''
DEFAULT_TEXT_SIZE = None
DEFAULT_TEXT_FONT = None
DEFAULT_TEXT_COLOR = None
DEFAULT_STARTUP_BORDER = False
DEFAULT_ORGANIZE = True
DEFAULT_AUTOMATIC_FINISH = True
