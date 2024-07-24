import pandas as pd
import glob
import numpy as np
import time
import datetime as dt
import re


uploaded_file = st.file_uploader("Upload File", type="xlsx")
df_3207  = pd.read_excel(uploaded_file,header=1).fillna('')
