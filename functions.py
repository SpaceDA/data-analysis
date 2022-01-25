import numpy as np
import pandas as pd
from datetime import datetime
import os
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import base64
import sqlite3

pd.options.mode.chained_assignment = None
pd.set_option('display.float_format', str)

dir_path = os.path.dirname(os.path.realpath(__file__))

stud_list = [
    "30010J62",
    "30010J64",
    "31401K79",
    "85217K84",
    "54077K27",
    "81269K63",
    "83303K75",
    "85217TKP",
    "32092016",
    "34001J9F",
    "34022017",
    "30903K49",
    "31340JBT",
    "31350JA4",
    "06015G9N",
    "06015J9M",
    "06015JB4",
    "06015JB5",
    "00014J9X",
    "06080K73",
    "06050KB6",
    "53103JA1",
    "012441A5",
    "06050K03",
    "06050KB9",
    "06080K71",
    "06115K74",
    "30303J61",
    "30382068",
    "31301K78",
    "31316K78",
    "31318K78",
    "31360K78",
    "33353J9Y",
    "33808JA0",
    "35102K9R",
    "38445088",
    "53103K66",
    "53103K82",
    "53103KB2",
    "54042U30",
    "54060H99",
    "54060K9J",
    "54061K80",
    "54063K68",
    "54066K91",
    "54069K89",
    "54071K88",
    "54076JAU",
    "54076KAZ",
    "54078K99",
    "54079KBM",
    "54081K9U",
    "54090KBB",
]
tecom_hq_list = [
    "3000201E",
    "30002084",
    "30002086",
    "54100M6A",
    "88601084", ]
edcom_list = [
    "30002068",
    "30002TZ8",
    "30010J62",
    "30010J64",
    "30010J65",
    "30010J66",
    "85217K84",
    "81269K63",
    "83303K75",
    "85217TKP",
    "02301JBH",
    "20230JBG",
    "54028Q40",
    "54060H99",
]
magtftc_list = [
    "012431A5",
    "88624UKT",
    "29600U18",
    "33610028",
    "35010015",
    "35010025",
    "35010050",
    "35010JBJ",
    "35010UKT",
    "35010UKU",
    "38445067",
    "88624015",
    "88635U18",
    "88673028",
    "88723067",

]
mcrd_pi_list = [
    "32001016",
    "32001021",
    "32090016",
    "32090040",
    "32090J9G",
    "32091016",
    "32091040",
    "32092016",
    "32100016",
    "32100040",
    "32111016",
    "32111040",
    "32121016",
    "32121040",
    "32170016",
    "32170040",
    "32300060",
    "88684016",

]
mcrd_sd_list = [
    "33710038",
    "34001017",
    "34001461",
    "34001J9D",
    "34001J9F",
    "34001JAA",
    "34001MB5",
    "34022017",
    "34027017",
    "34027041",
    "34027J9E",
    "34028017",
    "34028041",
    "34100017",
    "34100041",
    "34110017",
    "34110041",
    "34120017",
    "34120041",
    "54028Q07",
    "88637017",

]
training_command_list = [
    "30002087",
    "31401K79",
    "06080T0F",
    "54028Q39",
    "54077K27",
    "88601086",
    "88601UKA",
    "88667078",
    "88686K18",
    "06050KB8",
    "31319JAX",
    "31319K61",
    "31319KA4",
    "31319KA9",
    "31319KE2",
    "53103J34",
    "53103K26",
    "53103K55",
    "54028Q38",
    "54042825",
    "54069J44",
    "54078J78",
    "54079JAL",
    "54079JB2",
    "54079TAR",
    "54081J88",
    "54091JB3",
    "886361GF",
    "88662JAL",
    "02301KBG",
    "54060UCJ",
    "54060UCL",
    "54060W52",
    "54071J36",
    "88626KAV",
    "54063K46",
    "54066J54",
    "54076JAH",
    "54077M75",
    "54090JAD",
    "30374078",
    "30381069",
    "30370074",
    "30370078",
    "30903095",
    "30903J71",
    "30903JAV",
    "30903K49",
    "30903K95",
    "30903UAN",
    "88630095",
    "88630JAV",
    "31301J15",
    "31316J15",
    "31318J15",
    "31319J15",
    "31340JBF",
    "31340JBT",
    "31350KA2",
    "31350KAA",
    "31351KAB",
    "31352KAD",
    "31353KAF",
    "31354KAH",
    "31360J15",
    "33350JAB",
    "31350JA4",
    "31401J33",
    "33350KA0",
    "33354KAP",
    "88626KAM",
    "33950K81",
    "33350JBS",
    "33350KA3",
    "33351KAK",
    "33352KAM",
    "33354KAQ",
    "33354KBJ",
    "33355KAR",
    "33808KA1",
    "88703KA1",
    "35101K18",
    "35101KEB",
    "35101KES",
    "54081JAN",
    "88663JAN",
    "06015G77",
    "06015G91",
    "06015G9K",
    "06015G9N",
    "06015J9M",
    "06015JB4",
    "06015JB5",
    "06050G78",
    "06050KD6",
    "06050NKZ",
    "06050TAN",
    "06050TAV",
    "00014J9X",
    "06080G71",
    "06080K73",
    "06080L18",
    "06015G90",
    "06015K09",
    "06050G9H",
    "06050KB6",
    "06115G81",
    "54081G97",
    "06050KBC",
    "06050MCF",
    "31319M68",
    "53027T1B",
    "53103JA1",
    "53103JAS",
    "54060J08",
    "54061J18",
    "54061MD8",
    "54066JAW",
    "54066K96",
    "56001030",
    "56001JAR",
    "56011029",
    "84243J25",
    "84243J41",
    "012441A5",
    "06050K03",
    "06050KB9",
    "06080K71",
    "06115K74",
    "30303J61",
    "30382068",
    "31301K78",
    "31316K78",
    "31318K78",
    "31360K78",
    "33353J9Y",
    "33808JA0",
    "35102K9R",
    "38445088",
    "53103K66",
    "53103K82",
    "53103KB2",
    "54042U30",
    "54060K9J",
    "54061K80",
    "54063K68",
    "54066K91",
    "54069K89",
    "54071K88",
    "54076JAU",
    "54076KAZ",
    "54078K99",
    "54079KBM",
    "54081K9U",
    "54090KBB",

]

tecom_total_list = tecom_hq_list + edcom_list + magtftc_list + mcrd_pi_list + mcrd_sd_list + training_command_list


def pos_cases():
    today = datetime.now()
    dir_name = f"{today.strftime('%d%b%y')} COVID Report"
    date_stamp = today.strftime("%H%M %d%b%y")
    img_path = f"{dir_path}/{dir_name}/my_plot.png"

    #Get total personnel in TECOM from SQLITE DB
    conn = sqlite3.connect("COVID_data_working.sqlite")
    cur = conn.cursor()
    cur.execute("SELECT Personnel FROM main WHERE date = ?", (datetime.now().strftime("%Y-%m-%d"),))
    TECOM_total_pers = int(cur.fetchone()[0])
    cur.close()

    #convert to average daily cases per capita in TECOM
    
    pos_df = pd.read_excel("pos_civ.xlsx", sheet_name="pos_cases")
    pd.to_datetime(pos_df.date)
    pos_df["Active 7-Day Avg Per Capita"] = (pos_df["active_cases"]/7)/TECOM_total_pers
    
    # clean and prep CDC data
    daily_df = pd.read_csv("daily.csv", header=2)
    daily_df.Date = pd.to_datetime(daily_df.Date)
    daily_df["Moving Average Per Capita"] = daily_df["7-Day Moving Avg"]/334000000

    #plot TECOM vs CDC US Data
    fig, ax = plt.subplots()
    plt.figure(figsize=(8,6))
    plt.plot(pos_df.date, pos_df["Active 7-Day Avg Per Capita"], label = "TECOM")
    plt.plot(daily_df.Date, daily_df["Moving Average Per Capita"], label = "US")
    date_format = mdates.DateFormatter("%b")
    date_format = mdates.DateFormatter("%b")
    ax.xaxis.set_major_formatter(date_format)
    plt.title(f"COVID 19 Active Cases")
    plt.xlabel("Time")
    plt.ylabel("Active Cases Per Capita")
    plt.legend()
    plt.savefig(img_path)

    #save positive case data to SQL
    conn = sqlite3.connect("COVID_data_working.sqlite")
    cur = conn.cursor()
    pos_df["date"] = pos_df["date"].dt.strftime("%Y-%m-%d")
    pos_df.to_sql("positives", conn, if_exists="replace")
    cur.close()

    #export graph as b64 string:
    with open(img_path, "rb") as img_file:
        b64_string = base64.b64encode(img_file.read())



    return b64_string

def clean(df):

    # filter
    df = df[df["Unit"].isin(tecom_total_list)]

    

    # Assign units to MSC's
    conditions = [
        (df["Unit"].isin(tecom_hq_list)),
        (df["Unit"].isin(edcom_list)),
        (df["Unit"].isin(mcrd_pi_list)),
        (df["Unit"].isin(mcrd_sd_list)),
        (df["Unit"].isin(magtftc_list)),
        (df["Unit"].isin(training_command_list)),
    ]

    values = ["TECOM_HQ", "EDCOM", "MCRD_PI", "MCRD_SD", "MAGTFTC", "TRAINING_CMD"]

    df["MSC"] = np.select(conditions, values)

    df["Student"] = np.where(df.Unit.isin(stud_list), "Student", "Perm")

    df.reset_index(drop=True, inplace=True)

    return df

def unit_summary(df):

    df["Deferred & Exp"] = df["Deferred #"] + df["Expired Deferral #"]

    summary_df = df.groupby(["MSC"]).agg(
        {"Personnel Count": sum, "Took 1st Dose #": sum, "All Completed #": sum, "Deferred & Exp": sum,
         "No Action Taken #": sum, "Error Record #": sum})

    perm_stud = df.groupby("Student").agg(
        {"Personnel Count": sum, "Took 1st Dose #": sum, "All Completed #": sum, "Deferred & Exp": sum,
         "No Action Taken #": sum, "Error Record #": sum})

    summary_df.loc['TECOM Total', :] = round(summary_df.sum(axis=0),0)

    summary_df = pd.concat([summary_df, perm_stud])
    #
    summary_df["% In progress"] = round(100 * (
            summary_df["Took 1st Dose #"] + summary_df["Error Record #"] + summary_df["Deferred & Exp"] + summary_df[
        "All Completed #"]) / summary_df["Personnel Count"], 1)

    summary_df["% Vaccinated"] = round(100 * summary_df["All Completed #"] / summary_df["Personnel Count"], 0)
    summary_df["No Action Taken %"] = round(100 * summary_df["No Action Taken #"] / summary_df["Personnel Count"], 0)
    summary_df = summary_df.reindex(
        ["TECOM Total", "Perm", "Student", "TECOM_HQ", "EDCOM", "MAGTFTC", "MCRD_PI", "MCRD_SD", "TRAINING_CMD"])

    summary_df.reset_index(inplace=True)
    summary_df = summary_df.rename(columns={'index': 'MSC', 'All Completed #': 'Complete', 'Took 1st Dose #': '1 Dose', 'No Action Taken %': '% No Action', 'Personnel Count': 'Personnel', 'Error Record #': 'Errors', "% Vaccinated": "% Complete", })
    summary_df.round()
    return summary_df

def student_perm_summary(df):
    """# Perm/Stud breakdown"""

    studs = df[df["Student"] == "Student"]

    studs = studs.groupby(["MSC"]).agg(
        {"Personnel Count": sum, "Took 1st Dose #": sum, "All Completed #": sum, "Deferred & Exp": sum,
         "No Action Taken #": sum, "Error Record #": sum})

    """# Students"""

    studs["% In progress"] = round(100 * (
                studs["Took 1st Dose #"] + studs["Error Record #"] + studs["Deferred & Exp"] + studs[
            "All Completed #"]) / studs["Personnel Count"], 1)
    studs["% Vaccinated"] = round(100 * studs["All Completed #"] / studs["Personnel Count"], 1)
    studs["No Action Taken %"] = round(100 * studs["No Action Taken #"] / studs["Personnel Count"], 1)
    studs.reset_index(inplace=True)
    studs.rename(columns={'index': 'MSC'})

    """# Perm"""


    perm = df[df["Student"] == "Perm"]

    perm = perm.groupby(["MSC"]).agg(
        {"Personnel Count": sum, "Took 1st Dose #": sum, "All Completed #": sum, "Deferred & Exp": sum,
         "No Action Taken #": sum, "Error Record #": sum})


    perm["% In progress"] = round(
        100 * (perm["Took 1st Dose #"] + perm["Error Record #"] + perm["Deferred & Exp"] + perm["All Completed #"]) /
        perm["Personnel Count"], 1)
    perm["% Vaccinated"] = round(100 * perm["All Completed #"] / perm["Personnel Count"], 1)
    perm["No Action Taken %"] = round(100 * perm["No Action Taken #"] / perm["Personnel Count"], 2)
    perm.reset_index(inplace=True)
    perm.rename(columns={'index': 'MSC'})





    studs = studs.rename(columns={'All Completed #': 'Complete'})
    studs = studs.rename(columns={'Took 1st Dose #': '1st Dose'})
    studs = studs.rename(columns={'No Action Taken %': '% No Action'})
    studs = studs.rename(columns={'Personnel Count': 'Personnel'})
    perm = perm.rename(columns={'All Completed #': 'Complete'})
    perm = perm.rename(columns={'Took 1st Dose #': '1st Dose'})
    perm = perm.rename(columns={'No Action Taken %': '% No Action'})
    perm = perm.rename(columns={'Personnel Count': 'Personnel'})

    return studs, perm

def save_file(df_dict):
    today = datetime.now()

    dir_name = f"{today.strftime('%d%b%y')} COVID Report"
    os.mkdir(os.path.join(dir_path, dir_name))
    date_stamp = today.strftime("%H%M %d%b%y")
    with pd.ExcelWriter(f"{dir_path}/{dir_name}/{date_stamp} MRRS_rosters.xlsx") as writer:
        for sheetname, df in df_dict.items():
            df.to_excel(writer, sheet_name=sheetname)



# saves filtered excel files by criteria below
def excel_outputs(df):

    # No action
    df_no_action = df[(df["Series"].isnull()) & (df["Deferral Type"].isnull())]


    # Errors
    df_errors = df[df.Completed == "Immune Error"]


    # HQ Only
    df_HQ = df[df.Unit.isin(tecom_hq_list)]


    # Deferral DF
    df_deferrals = df[~df["Deferral Type"].isnull()]

    df_dict = {"No Action":df_no_action, "Errors":df_errors,  "HQ Marines": df_HQ, "Deferrals":df_deferrals}

    save_file(df_dict)

    return None

def by_name_summary(df):

    # assign deferral types
    conditions = [
        df["Deferral Type"] == "Admin PCS",
        df["Deferral Type"] == "Admin Refusal",
        df["Deferral Type"] == "Admin Separation",
        df["Deferral Type"] == "Admin Temporary",
        df["Deferral Type"] == "Expired Deferral - AP",
        df["Deferral Type"] == "Expired Deferral - AT",
        df["Deferral Type"] == "Medical Temporary",
        df["Deferral Type"] == "Relig. Accom. Req",
        df["Deferral Type"] == "Expired Deferral - MT",
        df["Deferral Type"] == "Expired Deferral - RA",
        df["Deferral Type"] == "Expired Deferral - AR",
    ]

    df["Admin PCS"] = np.where(df["Deferral Type"] == "Admin PCS", 1, 0)
    df["Admin Refusal"] = np.where(df["Deferral Type"] == "Admin Refusal", 1, 0)
    df["Admin Separation"] = np.where(df["Deferral Type"] == "Admin Separation", 1, 0)
    df["Admin Temporary"] = np.where(df["Deferral Type"] == "Admin Temporary", 1, 0)
    df["Expired Deferral - AP"] = np.where(df["Deferral Type"] == "Expired Deferral - AP", 1, 0)
    df["Expired Deferral - AT"] = np.where(df["Deferral Type"] == "Expired Deferral - AT", 1, 0)
    df["Medical Temporary"] = np.where(df["Deferral Type"] == "Medical Temporary", 1, 0)
    df["Relig. Accom. Req"] = np.where(df["Deferral Type"] == "Relig. Accom. Req", 1, 0)
    df["Expired Deferral - MT"] = np.where(df["Deferral Type"] == "Expired Deferral - MT", 1, 0)
    df["Expired Deferral - RA"] = np.where(df["Deferral Type"] == "Expired Deferral - RA", 1, 0)
    df["Expired Deferral - AR"] = np.where(df["Deferral Type"] == "Expired Deferral - AR", 1, 0)

    df["Total Expired"] = round(
        df["Expired Deferral - AR"] + df["Expired Deferral - RA"] + df["Expired Deferral - MT"] + df[
            "Expired Deferral - AT"] + df["Expired Deferral - AP"], 0)

    summary_df = df.groupby(["MSC"]).agg({"Admin Refusal": sum, "Admin PCS": sum, "Admin Separation": sum,
                                          "Admin Temporary": sum, "Medical Temporary": sum, "Relig. Accom. Req": sum,
                                          "Total Expired": sum})

    summary_df["Total Deferrals"] = summary_df.sum(axis=1)
    summary_df = summary_df[["Total Deferrals", "Total Expired", "Relig. Accom. Req", "Admin Refusal",
                             "Admin PCS", "Admin Separation", "Admin Temporary", "Medical Temporary"]]

    summary_df.loc["TECOM Total"] = summary_df.sum(axis=0)

    summary_df.reset_index(inplace=True)
    summary_df = summary_df.rename(columns={'index': 'MSC'})

    stud_perm_df = df.groupby(["MSC", "Student"]).agg({"Admin Refusal": sum, "Admin PCS": sum, "Admin Separation": sum,
                                                       "Admin Temporary": sum, "Medical Temporary": sum,
                                                       "Relig. Accom. Req": sum, "Total Expired": sum})

    stud_perm_df["Total Deferrals"] = stud_perm_df.sum(axis=1)
    stud_perm_df = stud_perm_df[["Total Deferrals", "Total Expired", "Relig. Accom. Req", "Admin Refusal",
                                 "Admin PCS", "Admin Separation", "Admin Temporary", "Medical Temporary"]]

    stud_perm_df.reset_index(inplace=True)
    stud_perm_df = stud_perm_df.rename(columns={'index': 'MSC'})

    return summary_df, stud_perm_df

