from asi import *
import asue

# Reading Excel File
df = pd.read_excel(asue.EXCEL_FILE_NAME)

# List of Column Headers
df_headers = df.columns

# Storing Column Data in Lists
# Not Generalised Them
mName = list(df[df_headers[0]])
mCompany = list(df[df_headers[1]])
mColor = list(df[df_headers[2]])
mMovie = list(df[df_headers[3]])
mCurrency = list(df[df_headers[4]])
mHobby = list(df[df_headers[5]])
mOccupation = list(df[df_headers[6]])
mQuote = list(df[df_headers[7]])

# Number of Entries
mLength = len(mName)
