import sqlite3
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime
from datetime import timedelta



conn = sqlite3.connect("COVID_data_working.sqlite")
cur = conn.cursor()

df = pd.read_excel("pos_civ.xlsx", sheet_name="pos_cases")
df["date"] = df["date"].dt.strftime("%Y-%m-%d")
df.to_sql("positives", conn, if_exists='replace')

date_stamp = datetime.now().strftime("%Y-%m-%d")

def to_db(df, table_name):
    
    df['date'] = date_stamp
    

    #df.to_sql(table_name, conn, if_exists="append")
    
    df.drop('date', axis=1, inplace=True)

    return None



def generate_email():
    current_date = datetime.now().strftime("%Y-%m-%d")
    
    if datetime.today().weekday() == 0:
        last = (datetime.now() - timedelta(days=5)).strftime("%Y-%m-%d")
        


    if datetime.today().weekday() == 2:
        last = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%d")

    
    #last report's positives
    
    try:
        global positives_change
        global current_positives
        global prev_positives

        cur.execute("SELECT active_cases FROM positives WHERE date=?", (last,))
        prev_positives = cur.fetchone()[0]
    
        #current report's positives
        cur.execute("SELECT active_cases FROM positives WHERE date=?", (current_date,))
        current_positives = cur.fetchone()[0]
    
        
        positives_change = current_positives - prev_positives

        c = {}
        if positives_change < 0:
            c["positives_change_dir"] = "down by"
        elif positives_change > 0:
            c["positives_change_dir"] = "up by"
        else:
            c["positives_change_dir"] = "steady at"
    
    except:
        positives = None

    #TECOM total personnel
    #last report personnel numbers
    cur.execute("SELECT [Personnel] from main WHERE date = ? AND MSC = ?", (last, "TECOM Total"))
    prev_personnel = cur.fetchone()[0]
    
    #current report's personnel numbers
    cur.execute("SELECT [Personnel] from main WHERE date = ? AND MSC = ?", (current_date, "TECOM Total"))
    current_personnel = cur.fetchone()[0]

    personnel_change = current_personnel - prev_personnel

    #current report's overall vaccination rate

    cur.execute("SELECT [% Complete] from main WHERE date = ? AND MSC = ?", (current_date, "TECOM Total"))
    current_vacc_rate = cur.fetchone()[0]

    cur.execute("SELECT [% Complete] from main WHERE date = ? AND MSC = ?", (last, "TECOM Total"))
    prev_vacc_rate = cur.fetchone()[0]
    

    vaccination_rate_change = current_vacc_rate - prev_vacc_rate
    

    #last report's perm vacc rate
    cur.execute("SELECT [% Complete] from main WHERE date = ? AND MSC = ?", (last, "Perm"))
    prev_perm_vacc_rate = cur.fetchone()[0]

    #current report's perm vacc rate
    cur.execute("SELECT [% Complete] from main WHERE date = ? AND MSC = ?", (current_date, "Perm"))
    current_perm_vacc_rate = cur.fetchone()[0]

    perm_vacc_rate_change = current_perm_vacc_rate - prev_perm_vacc_rate

    #last report's stud vacc rate
    cur.execute("SELECT [% Complete] from main WHERE date = ? AND MSC = ?", (last, "Student"))
    prev_stud_vacc_rate = cur.fetchone()[0]

    #current report's perm vacc rate
    cur.execute("SELECT [% Complete] from main WHERE date = ? AND MSC = ?", (current_date, "Student"))
    current_stud_vacc_rate = cur.fetchone()[0]

    stud_vacc_rate_change =  current_stud_vacc_rate - prev_stud_vacc_rate

    #last report no action 
    cur.execute("SELECT [% No Action] from main WHERE date = ? AND MSC = ?", (last, "TECOM Total"))
    prev_noaction_rate = cur.fetchone()[0]

    #current report's no action
    cur.execute("SELECT [% No Action] from main WHERE date = ? AND MSC = ?", (current_date, "TECOM Total"))
    current_noaction_rate = cur.fetchone()[0]

    noaction_change = current_noaction_rate - prev_noaction_rate
    
    d = {}
   
    if personnel_change < 0:
        d["personnel_change_dir"] = "down by"
    elif personnel_change > 0:
        d["personnel_change_dir"] = "up by"
    else:
        d["personnel_change_dir"] = "steady at"

    
    if vaccination_rate_change < 0:
        d["vaccination_rate_change_dir"] = "down by"
    elif vaccination_rate_change > 0:
        d["vaccination_rate_change_dir"] = "up by"
    else:
        d["vaccination_rate_change_dir"] = "steady at"

    if perm_vacc_rate_change < 0:
        d["perm_vacc_rate_change_dir"] = "down by"
    elif perm_vacc_rate_change > 0:
        d["perm_vacc_rate_change_dir"] = "up by"
    else:
        d["perm_vacc_rate_change_dir"] = "steady at"

    if stud_vacc_rate_change < 0:
        d["stud_vacc_rate_change_dir"] = "down by"
    elif stud_vacc_rate_change > 0:
        d["stud_vacc_rate_change_dir"] = "up by"
    else:
        d["stud_vacc_rate_change_dir"] = "steady at"

    if noaction_change < 0:
        d["noaction_change_dir"] = "down by"
    elif noaction_change > 0:
        d["noaction_change_dir"] = "up by"
    else:
        d["noaction_change_dir"] = "steady at"

    print(d)
    
    if positives:
        print(f"Good afternoon gentlemen,\n\nOverall, active positive cases are {d['positives_change_dir']} {abs(positives_change) if positives_change != 0 else ''}"
        f"{' Marines' if positives_change != 0 else ''} {'to' if positives_change != 0 else ''} {current_positives}.\n\n")
    
    print(f"TECOM's personnel numbers are {d['personnel_change_dir']} {abs(int(personnel_change)) if personnel_change != 0 else ''}"
    f"{'' if personnel_change != 0 else ''}{' to' if personnel_change != 0 else ''} {int(current_personnel)}.\n\n" 
    f"TECOM's overall vaccination rate is {d['vaccination_rate_change_dir']} {abs(vaccination_rate_change) if vaccination_rate_change != 0 else ''}"
    f"{'%' if vaccination_rate_change != 0 else ''}{'to' if vaccination_rate_change != 0 else ''}{current_vacc_rate}%.\n\n"
    f"Permanent personnel vaccination rate is {d['perm_vacc_rate_change_dir']} {abs(perm_vacc_rate_change) if perm_vacc_rate_change != 0 else ''}"
    f"{'%' if perm_vacc_rate_change != 0 else ''}{'to' if perm_vacc_rate_change != 0 else ''}{current_perm_vacc_rate}%.\n\n"
    f"Student personnel vaccination rate is {d['stud_vacc_rate_change_dir']} {abs(stud_vacc_rate_change) if stud_vacc_rate_change != 0 else ''}"
    f"{'%' if stud_vacc_rate_change != 0 else ''}{'to' if stud_vacc_rate_change != 0 else ''}{current_stud_vacc_rate}%.\n\n"
    f"TECOM 'no action' rate is {d['noaction_change_dir']} {abs(noaction_change) if noaction_change != 0 else ''}"
    f"{'%' if noaction_change != 0 else ''}{'to' if noaction_change != 0 else ''}{current_noaction_rate}%.\n\n")






    
if __name__ == "__main__":

    generate_email()



cur.close()




#MSC_only = ["TECOM_HQ", "EDCOM", "MCRD_PI", "MCRD_SD", "MAGTFTC", "TRAINING_CMD"]
#cur.execute('''SELECT MSC, [1 Dose], [No Action Taken #], [Admin Refusal], [Relig. Accom. Req], [Total Expired], ([Admin PCS] + [Admin Temporary] + [Admin Separation] + [Medical Temporary]) AS exemptions
#FROM main
#WHERE date = '2022-01-12' ''')
#summary_data = cur.fetchall()
