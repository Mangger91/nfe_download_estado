import os
import requests
from datetime import datetime
from base64 import b64decode
import pandas as pd
from zeep import Client
from zeep.transports import Transports
from zeep.plugins import HistoryPlugin

# Var
