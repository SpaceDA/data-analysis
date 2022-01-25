from datetime import datetime
import pandas as pd
import functions
import database
from tempfile import NamedTemporaryFile
import webbrowser
from datetime import datetime
import os

base_html = """
<!doctype html>
<html><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/css/bootstrap.min.css" integrity="sha384-MCw98/SFnGE8fJT3GXwEOngsV7Zt27NXFoaoApmYm81iuXoPkFOJwJ8ERdknLPMO" crossorigin="anonymous">
</head><body><h2 class="text-center mt-4"> %s </h2>%s<script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" integrity="sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo" crossorigin="anonymous"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.3/umd/popper.min.js" integrity="sha384-ZMP7rVo3mIykV+2+9J3UJ46jBk0WLaUAdn689aCwoqbBJiSnjAK/l8WvCWPIPm49" crossorigin="anonymous"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.1.3/js/bootstrap.min.js" integrity="sha384-ChfqqxuZUCnJSK3+MXmPNIyE6ZbWh2IMqE241rYiqJxyMiZ6OW/JmZQ5stwEULTy" crossorigin="anonymous"></script>
</body></html>
"""

dir_path = os.path.dirname(os.path.realpath(__file__))

def df_html(df, title):
    """HTML table with pagination and other goodies"""
    df_html = df.to_html(classes=classes, index=False)
    return base_html %(title, df_html)

def df_window(dict):
    """Open dataframe in browser window using a temporary file"""
    dir_name = f"{today.strftime('%d%b%y')} COVID Report"
    with open(f"{dir_path}/{dir_name}/{datetime.now().strftime('%d%b%y')} COVID Summary.html", 'w') as f:
        f.write(html_fig)
        for key, value in dict.items():
            f.write(df_html(value, key))
    webbrowser.open(f.name)

# bootstrap classes
classes = 'table table-striped table-bordered table-hover table-sm thead-dark w-50 text-left mx-auto'


pd.set_option('precision', 0)

# read excel files

mrrs = pd.read_excel("MRRS_EXCEL.xlsx", header=2)
by_name = pd.read_excel("MRRS_byname.xlsx", header=3)
civ_df = pd.read_excel("pos_civ.xlsx", sheet_name="civ_vax", nrows=7)
pos_by_MSC_df = pd.read_excel("pos_civ.xlsx", sheet_name="total_pos", skiprows=13, nrows=7, usecols="A:E")

pd.set_option('colheader_justify', 'center')
civ_df["% Complete Vaccine"] = round(100 * civ_df["Vaccinated"]/ civ_df["On Hand"], 1)
civ_df["% In Progress (Includes Pending Exemptions)"] = round(100 * (civ_df["Vaccinated"] + civ_df["Medical Exemption Pending"] + civ_df["Religious Exemption Pending"] + civ_df["Temporary Admin Exemption Pending"])/ civ_df["On Hand"], 1)
civ_df = civ_df.drop(civ_df.columns[8], axis=1)
civ_df["On Hand"].astype(int)

# Clean data
cleaned_df = functions.clean(mrrs)



#unit summary
summary_df = functions.unit_summary(cleaned_df)

studs, perm = functions.student_perm_summary(cleaned_df)

#by name
by_name = functions.clean(by_name)

#excel outputs

functions.excel_outputs(by_name)



#summary tables

by_name_summary, stud_perm_summary = functions.by_name_summary(by_name)

#to sqlite
database.to_db(summary_df, "main")
database.to_db(pos_by_MSC_df, "msc_pos")


today = datetime.now()
date_stamp = today.strftime("%H%M %d%b%y")

b64_bytes = functions.pos_cases()
b64_string = b64_bytes.decode()


dir_name = f"{today.strftime('%d%b%y')} COVID Report/my_plot.png"
img_path = os.path.join(dir_path, dir_name)




html_fig = '<div style="text-align:center"><img src="data:image/png;base64,%s"/></div>' %b64_string

#summary_df.drop('date', axis=1, inplace=True) 
df_window({f"MSC Summary as of {date_stamp}": summary_df, f"Students as of {date_stamp}":studs, f"Permanent Personnel as of {date_stamp}": perm, f"Total Deferral Summary as of {date_stamp}":by_name_summary, f"Student / Permanent Personnel as of {date_stamp}": stud_perm_summary, f"Civilian Vax Status as of {date_stamp}":civ_df, } )
