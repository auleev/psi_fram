import json
import sys

import pygsheets
from google.oauth2 import service_account

sys.path.append('../')
import pandas as pd
from pygsheets.datarange import DataRange


name_sheet = "Historisation_temps_de_chargement_Fram"
with open('searchconsole-334412-82a453177feb.json') as source:
    info = json.load(source)
    

credentials = service_account.Credentials.from_service_account_info(info)

VERBOSE = 1
def clean_google_sheet(name_sheet):
     
    client = pygsheets.authorize(service_account_file="searchconsole-334412-82a453177feb.json")
    sh = client.open(name_sheet)
    #Onglet crawl cleané
    wks = sh.worksheet('title','Data HP_mobile')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet crawl édito cleané
    wks = sh.worksheet('title','Data SL_mobile')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet majestic cleané
    wks = sh.worksheet('title','Data FP_mobile')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet confuguration cleané
    wks = sh.worksheet('title','Data LD_mobile')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet confuguration cleané
    wks = sh.worksheet('title','Data HPT_mobile')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet confuguration cleané
    wks = sh.worksheet('title','Data QDP_mobile')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet confuguration cleané
    wks = sh.worksheet('title','Data HP')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet confuguration cleané
    wks = sh.worksheet('title','Data SL')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet confuguration cleané
    wks = sh.worksheet('title','Data FP')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet confuguration cleané
    wks = sh.worksheet('title','Data HPT')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
    #Onglet confuguration cleané
    wks = sh.worksheet('title','Data QDP')
    crawl=DataRange("A2","M1000",wks)
    crawl.clear()
 

def excel_to_googlesheet(name_sheet):
    #Import dataframe 

    client = pygsheets.authorize(service_account_file="searchconsole-334412-82a453177feb.json")
    sh = client.open(name_sheet)
    if VERBOSE:
        print("excel2google-local sheetName")
        print(pd.ExcelFile('suivi_perf_mep.xlsx').sheet_names)
        print("excel2google-remote sheetName")
        for sheet in sh.worksheets():
            print('sheetName: {}, sheetId(GID): {}'.format(sheet.title, sheet.id))
    # import des excels
    int_all=pd.read_excel("suivi_perf_mep.xlsx","Data HP_mobile")
    int_all=int_all.fillna(" ")

    int_all_2=pd.read_excel("suivi_perf_mep.xlsx","Data SL_mobile")
    int_all_2=int_all_2.fillna(" ")
    
    
    int_all_3=pd.read_excel("suivi_perf_mep.xlsx","Data FP_mobile")
    int_all_3=int_all_3.fillna(" ")

    int_all_4=pd.read_excel("suivi_perf_mep.xlsx","Data LD_mobile")
    int_all_4=int_all_4.fillna(" ")

    int_all_5=pd.read_excel("suivi_perf_mep.xlsx","Data HPT_mobile")
    int_all_5=int_all_5.fillna(" ")

    int_all_6=pd.read_excel("suivi_perf_mep.xlsx","Data QDP_mobile")
    int_all_6=int_all_6.fillna(" ")

    int_all_7=pd.read_excel("suivi_perf_mep.xlsx","Data HP")
    int_all_7=int_all_7.fillna(" ")

    int_all_8=pd.read_excel("suivi_perf_mep.xlsx","Data SL")
    int_all_8=int_all_8.fillna(" ")

    int_all_9=pd.read_excel("suivi_perf_mep.xlsx","Data FP")
    int_all_9=int_all_9.fillna(" ")

    int_all_10=pd.read_excel("suivi_perf_mep.xlsx","Data LD")
    int_all_10=int_all_10.fillna(" ")

    int_all_11=pd.read_excel("suivi_perf_mep.xlsx","Data HPT")
    int_all_11=int_all_11.fillna(" ")

    int_all_12=pd.read_excel("suivi_perf_mep.xlsx","Data QDP")
    int_all_12=int_all_12.fillna(" ")

    #int_html=pd.read_excel(f"../exports/{url_name}/{url_name}-internal_html.xlsx",skiprows=11)
    #int_html=int_html.fillna(" ")

    #exportation des données vers googlesheet 
    '''
    fonction : update_col
    index : le numero de colonne (A=1,B=2...)
    values : la liste des valeurs qu'on souhaite remplir
    row_offset: le nombre de ligne qu'on passe avant de remplir
    '''
   
    #exportation des données vers le googlesheet Data HP_mobile
    wks = sh.worksheet_by_title('Data HP_mobile')
    wks.update_col(index=1,values=list(int_all["URL"]),row_offset=1)
    wks.update_col(2,list(int_all["Date"]),row_offset=1)
    wks.update_col(3,list(int_all["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all["Score Global"]),row_offset=1)
   

    #exportation des données vers le googlesheet Data SL_mobile
    wks = sh.worksheet_by_title('Data SL_mobile')
    wks.update_col(index=1,values=list(int_all_2["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_2["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_2["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_2["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_2["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_2["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_2["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_2["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_2["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_2["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_2["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_2["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_2["Score Global"]),row_offset=1)
    
    #exportation des données vers le googlesheet Data FP_mobile
    wks = sh.worksheet_by_title('Data FP_mobile')
    wks.update_col(index=1,values=list(int_all_3["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_3["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_3["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_3["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_3["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_3["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_3["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_3["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_3["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_3["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_3["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_3["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_3["Score Global"]),row_offset=1)

    #exportation des données vers le googlesheet Data LD_mobile
    wks = sh.worksheet_by_title('Data LD_mobile')
    wks.update_col(index=1,values=list(int_all_4["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_4["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_4["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_4["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_4["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_4["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_4["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_4["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_4["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_4["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_4["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_4["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_4["Score Global"]),row_offset=1)

    #exportation des données vers le googlesheet Data HPT_mobile
    wks = sh.worksheet_by_title('Data HPT_mobile')
    wks.update_col(index=1,values=list(int_all_5["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_5["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_5["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_5["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_5["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_5["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_5["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_5["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_5["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_5["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_5["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_5["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_5["Score Global"]),row_offset=1)

    #exportation des données vers le googlesheet Data QDP_mobile
    wks = sh.worksheet_by_title('Data QDP_mobile')
    wks.update_col(index=1,values=list(int_all_6["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_6["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_6["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_6["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_6["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_6["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_6["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_6["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_6["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_6["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_6["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_6["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_6["Score Global"]),row_offset=1)

    #exportation des données vers le googlesheet Data HP
    wks = sh.worksheet_by_title('Data HP')
    wks.update_col(index=1,values=list(int_all_7["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_7["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_7["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_7["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_7["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_7["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_7["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_7["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_7["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_7["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_7["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_7["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_7["Score Global"]),row_offset=1)


    #exportation des données vers le googlesheet Data SL
    wks = sh.worksheet_by_title('Data SL')
    wks.update_col(index=1,values=list(int_all_8["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_8["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_8["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_8["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_8["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_8["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_8["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_8["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_8["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_8["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_8["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_8["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_8["Score Global"]),row_offset=1)

    #exportation des données vers le googlesheet Data FP
    wks = sh.worksheet_by_title('Data FP')
    wks.update_col(index=1,values=list(int_all_9["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_9["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_9["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_9["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_9["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_9["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_9["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_9["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_9["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_9["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_9["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_9["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_9["Score Global"]),row_offset=1)

    #exportation des données vers le googlesheet Data LD
    wks = sh.worksheet_by_title('Data LD')
    wks.update_col(index=1,values=list(int_all_10["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_10["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_10["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_10["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_10["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_10["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_10["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_10["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_10["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_10["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_10["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_10["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_10["Score Global"]),row_offset=1)


    #exportation des données vers le googlesheet Data HPT
    wks = sh.worksheet_by_title('Data HPT')
    wks.update_col(index=1,values=list(int_all_11["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_11["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_11["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_11["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_11["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_11["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_11["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_11["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_11["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_11["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_11["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_11["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_11["Score Global"]),row_offset=1)

    #exportation des données vers le googlesheet Data QDP
    wks = sh.worksheet_by_title('Data QDP')
    wks.update_col(index=1,values=list(int_all_12["URL"]),row_offset=1)
    wks.update_col(2,list(int_all_12["Date"]),row_offset=1)
    wks.update_col(3,list(int_all_12["Time to Interactive"]),row_offset=1)
    wks.update_col(4,list(int_all_12["Server Reponse Time"]),row_offset=1)
    wks.update_col(5,list(int_all_12["Speed Index"]),row_offset=1)
    wks.update_col(6,list(int_all_12["Total Blocking Time"]),row_offset=1)
    wks.update_col(7,list(int_all_12["FCP"]),row_offset=1)
    wks.update_col(8,list(int_all_12["LCP"]),row_offset=1)
    wks.update_col(9,list(int_all_12["CLS"]),row_offset=1)
    wks.update_col(10,list(int_all_12["Score FCP"]),row_offset=1)
    wks.update_col(11,list(int_all_12["LCP Score"]),row_offset=1)
    wks.update_col(12,list(int_all_12["CLS Score"]),row_offset=1)
    wks.update_col(13,list(int_all_12["Score Global"]),row_offset=1)


#clean_google_sheet(name_sheet)
#excel_to_googlesheet(name_sheet)