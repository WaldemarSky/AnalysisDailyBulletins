import os
import subprocess
import time
import sys

Zip7 = r'"C:\Program Files\7-Zip\7z.exe"'
PDFconverter = r'"C:\Program Files (x86)\CoolUtils\Total PDF Converter\PDFConverter.exe"'

directory = r"E:\VladimirFiles\PPr\DailyBulletin\oi"
done_directory = r"E:\VladimirFiles\PPr\DailyBulletin\done_oi"
final_directory = r"E:\VladimirFiles\PPr\DailyBulletin\final"
files = os.listdir(directory)
filesDone = os.listdir(done_directory)

TP = "TargetPosition"
name = "name"

GeneralTP = 92
GeneralTPC = 97
GeneralTPE = 112

pdflist = ["Section03_Agricultural_Futures",
           "Section04_Agricultural_Soft_AltInvestment_Futures",
           "Section07_Currency_Futures",
           "Section08_Currency_Futures_Continued",
           "Section09_Interest_Rate_Futures",
           "Section10_Interest_Rate_Futures_Continued",
           "Section11_Equity_And_Index_Futures",
           "Section12_Equity_And_Index_Futures_Continued",
           "Section61_Energy_Futures_Products",
           "Section62_Metals_Futures_Products"
           ]

MonthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
LegalTotalWords = ['UNCH', '-', '+']
fut_list = {
    # "   CORN FUT": {name: "corn_cbot.txt", TP: GeneralTP},
    # "   SOYBEAN FUT": {name: "soybean_cbot.txt", TP: GeneralTP},
    # "   ROUGH RICE FUT": {name: "rough_rice_cbot.txt", TP: GeneralTP},
    # "   SOY MEAL FUT": {name: "soybean_meal_cbot.txt", TP: GeneralTP},
    # "   SOY OIL FUT": {name: "soybean_oil_cbot.txt", TP: GeneralTP},
    # "   WHEAT FUT": {name: "chicago_srw_wheat_cbot.txt", TP: GeneralTP},
    # "   LV CATTLE FUT": {name: "live_cattle_cme.txt", TP: GeneralTP},
    # "   LEAN HOGS FUT": {name: "lean_hog_cme.txt", TP: GeneralTP},
    # "   KEF CBT FUT": {name: "kc_hrw_wheat_cbot.txt", TP: 112},
    # "   MILK FUT": {name: "class_iii_milk_cme.txt", TP: GeneralTP},
    # "   CHEESE CSC FUT": {name: "cash-settled_cheese_cme.txt", TP: GeneralTP},
    # "   BUTTER CS FUT": {name: "cash-settled_butter_cme.txt", TP: GeneralTP},
    # "   FDR CATTLE FUT": {name: "feeder_cattle_cme.txt", TP: GeneralTP},
    # "   BRIT PND FUT": {name: "british_pound_cme.txt", TP: GeneralTPC},
    # "   CANADA DLR FUT": {name: "canadian_dollar_cme.txt", TP: GeneralTPC},
    # "   EURO FX FUT": {name: "euro_fx_cme.txt", TP: GeneralTPC},
    # "   JAPAN YEN FUT": {name: "japanese_yen_cme.txt", TP: GeneralTPC},
    # "   SWISS FRNC FUT": {name: "swiss_franc_cme.txt", TP: GeneralTPC},
    # "   AUST DLR FUT": {name: "australian_dollar_cme.txt", TP: GeneralTPC},
    # "   NEW ZEALND FUT": {name: "new_zealand_dollar_cme.txt", TP: GeneralTPC},
    # "   RUS RUBLE FUT": {name: "russian_ruble_cme.txt", TP: GeneralTPC},
    # "   EURO DLR FUT": {name: "eurodollar_cme.txt", TP: GeneralTPC},
    # "   10-YR NOTE FUTURES": {name: "10_year_t-note_cbot.txt", TP: GeneralTPC},
    # "   5-YR NOTE FUTURES": {name: "5_year_t-note_cbot.txt", TP: GeneralTPC},
    # "   2-YR NOTE FUTURES": {name: "2_year_t-note_cbot.txt", TP: GeneralTPC},
    # "   30Y BOND FUT": {name: "us_treasury_bond_cbot.txt", TP: GeneralTPC},
    # "   ULTRA T-BND FUT": {name: "ultra_us_treasury_bond_cbot.txt", TP: GeneralTPC},
    # "   TN FUT": {name: "ultra_10-year_us_treasury_note_cbot.txt", TP: GeneralTPC},
    # "   30D FED FD FUT": {name: "30_day_federal_funds_cbot.txt", TP: GeneralTPC},
    # "   EMINI S&P FUT": {name: "e-mini_s&p500_cme.txt", TP: GeneralTPC},
    # "   S&P 500 FUT": {name: "s&p500_cme.txt", TP: GeneralTPC},
    # "   E-400 MIDCAP F": {name: "e-mini_s&p_midcap_400_cme.txt", TP: GeneralTPC},
    # "   EMINI NASD FUT": {name: "e-mini_nasdaq-100_cme.txt", TP: GeneralTPC},
    # "   EMINI RUSSELL 2000 INDEX FUTURES": {name: "e-mini_russell_2000_index_cme.txt", TP: GeneralTPC},
    # "   BITCOIN FUTURES": {name: "bitcoin_cme.txt", TP: GeneralTPC},
    # "   CL FUT": {name: "wti_crude_oil_nymex.txt", TP: GeneralTPE},
    # "   NG FUT": {name: "henry_hub_natural_gas_nymex.txt", TP: GeneralTPE},
    "   AUP FUT": {name: "aluminum_mw_us_transaction_premium_platts(25mt)_comex.txt", TP: GeneralTPE},
    "   COMEX GOLD FUTURES": {name: "gold_comex.txt", TP: GeneralTPE},
    "   HG FUT": {name: "copper_comex.txt", TP: GeneralTPE},
    "   HRC FUT": {name: "us_midwest_domestic_hot-rolled_coil_steel_(cru)_index_comex.txt", TP: GeneralTPE},
    "   SI FUT": {name: "silver_comex.txt", TP: GeneralTPE},
    "   PA FUT": {name: "palladium_nymex.txt", TP: GeneralTPE},
    "   PL FUT": {name: "platinum_nymex.txt", TP: GeneralTPE}
}

def make_date(str):
    list_str = str.split('_')
    datestr = list_str[2]
    year = datestr[:4]
    month = datestr[4:6]
    day = datestr[6:8]
    return f"{year}.{month}.{day}"

def GetValueFromLine(line, pointer):
    CurP = pointer
    print(line[CurP])
    if line[CurP].isdigit():
        while True:
            if line[CurP - 1].isdigit():
                CurP -= 1
            else:
                break
    else:
        if '--' in line[CurP:CurP+5]:
            return '0'
        while True:
            CurP += 1
            if line[CurP].isdigit():
                break
            else:
                if line[CurP] == 'U' or line[CurP] == '+' or line[CurP] == '-':
                    while True:
                        CurP -= 1
                        if line[CurP] == 'U' or line[CurP] == '+' or line[CurP] == '-':
                            return '0'
                        elif line[CurP].isdigit():
                            while True:
                                if line[CurP - 1].isdigit():
                                    CurP -= 1
                                else:
                                    break
                            break
                    break

    finalvalue = ''
    while line[CurP].isdigit():
        finalvalue +=line[CurP]
        CurP += 1
    return finalvalue

def IsTotalLine(line):
    for month in MonthNames:
        if month in line:
            return False

    if "TOTAL" in line:
        return True

    ListLine = line.split()
    counter = 0
    for word in ListLine:
        if (not word.isdigit()) and (word not in LegalTotalWords):
            return False
        if word.isdigit():
            counter +=1
    if counter >= 3:
        return True

    return False



def GeneralProcessor(InstrName, date, file):
    try:
        print(InstrName)
        print(date)
        done_fut_list.append(InstrName)
        DoneFileLine = f"{date}\tFINAL\tTOTAL\t"
        while True:
            FileLine = file.readline()
            if len(FileLine.split()) < 2:
                continue
            if IsTotalLine(FileLine):
                OIvalue = GetValueFromLine(FileLine, fut_list[InstrName][TP])
                DoneFileLine = DoneFileLine[:23] + OIvalue + '\t' + DoneFileLine[23:]
                break
            else:
                for i in range(len(FileLine)):
                    if FileLine[i].isupper():
                        MonthName = FileLine[i:i + 3]
                        if MonthName not in MonthNames:
                            break
                        DoneFileLine += MonthName
                        DoneFileLine += ' '
                        DoneFileLine += FileLine[i+3:i + 5]
                        DoneFileLine += '\t'
                        break
                if MonthName not in MonthNames:
                    continue
                OIvalue = GetValueFromLine(FileLine, fut_list[InstrName][TP])
                DoneFileLine += OIvalue
                DoneFileLine += '\t'
        DoneFileLine += '\n'

        with open(fr"{final_directory}\{fut_list[InstrName][name]}", 'a+') as finalfile:
            finalfile.write(DoneFileLine)

    except Exception as err:
        print(InstrName)
        print(date)
        print(file)
        sys.exit()


def writefile(filename, subdirectory, date):
    try:
        with open(fr'{subdirectory}\txts\{filename}.txt', 'r') as filetxt:
            while True:
                FileLine = filetxt.readline()
                if FileLine == "":
                    break
                for key, value in fut_list.items():
                    if key == "   COMEX GOLD FUTURES" and (date == "2019.08.26" or date == "2019.12.16"):
                        continue
                    if key == "   MILK FUT" and "   MILK FUT" in FileLine and filename == "Section04_Agricultural_Soft_AltInvestment_Futures" and '/' not in FileLine and key not in done_fut_list:
                        GeneralProcessor(key, date, filetxt)

                    if key in FileLine and "   MILK FUT" not in FileLine and '/' not in FileLine and key not in done_fut_list:
                        GeneralProcessor(key, date, filetxt)




    except FileNotFoundError as err:
        print(err)
        subprocess.run(fr'{PDFconverter} "{subdirectory}\{filename}.pdf" "{subdirectory}\txts\{filename}.txt" -c TXT')
        writefile(pdf, subdirectory, date)
        time.sleep(2)

    # except Exception as err:
    #     print(err)
    #     sys.exit()



# for file in files:
#     listfile = file.split('.')
#     subprocess.run(fr'{Zip7} x "{directory}\{file}" -o"{done_directory}\{listfile[0]}"')
#     subdirectory = "E:\VladimirFiles\PPr\DailyBulletin\done_oi\{}".format(listfile[0])
#     date = make_date(listfile[0])
#     done_fut_list = []
#     for pdf in pdflist:
#         subprocess.run(fr'{PDFconverter} "{subdirectory}\{pdf}.pdf" "{subdirectory}\txts\{pdf}.txt" -c TXT')
#         writefile(pdf, subdirectory, date)
#         time.sleep(0)


for file in filesDone:
    listfile = file.split('.')
    # subprocess.run(fr'{Zip7} x "{directory}\{file}" -o"{done_directory}\{listfile[0]}"')
    subdirectory = "E:\VladimirFiles\PPr\DailyBulletin\done_oi\{}".format(listfile[0])
    date = make_date(listfile[0])
    done_fut_list = []
    for pdf in pdflist:
        # subprocess.run(fr'{PDFconverter} "{subdirectory}\{pdf}.pdf" "{subdirectory}\txts\{pdf}.txt" -c TXT')
        writefile(pdf, subdirectory, date)
        time.sleep(0)
