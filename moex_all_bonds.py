import requests
import xlsxwriter
import datetime
import sqlite3
import csv  

#get all bonds from MOEX
all_moex_bonds={}

def get_all_bonds_from_moex():
    global all_moex_bonds
    d = datetime.datetime.today()
    today_str=d.strftime("%Y%m%d")       
                        
    req_str='https://iss.moex.com/iss/engines/stock/markets/bonds/securities.json?marketprice_board=1'
    SECID=0
    ISIN=0
    BOARDID=0
    SHORTNAME=0
    SECNAME=0
    COUPONPERCENT=0
    FACEVALUE=0
    MATDATE=0
    COUPONPERIOD=0
    
    j=requests.get(req_str).json()  #'https://iss.moex.com/iss/engines/stock/markets/bonds/securities/RU000A106Z38.json?marketprice_board=1'
    
    data=j['securities']['data']
    
    for i in data:    
        for f, b in zip(j['securities']['columns'], i):
            if f=="SECID":
                SECID=b
            if f=="ISIN":
                ISIN=b        
            if f=="BOARDID":
                if b is not None:
                    BOARDID=b
            if f=="SHORTNAME":
                if b is not None:
                    SHORTNAME=b                
            if f=="SECNAME":
                if b is not None:
                    SECNAME=b   
            if f=="COUPONPERCENT":
                if b is not None:
                    COUPONPERCENT=b   
            if f=="FACEVALUE":
                if b is not None:
                    FACEVALUE=b  
            if f=="MATDATE":
                if b is not None:
                    MATDATE=b    
            if f=="COUPONPERIOD":
                if b is not None:
                    COUPONPERIOD=b                   
            
        all_moex_bonds[SECID]={"ISIN":ISIN, "BOARDID":BOARDID, "SHORTNAME":SHORTNAME, "SECNAME":SECNAME, "COUPONPERCENT":COUPONPERCENT, "FACEVALUE":FACEVALUE, "MATDATE":MATDATE, "COUPONPERIOD":COUPONPERIOD}                

    SECID1=0
    last_price=0
    market_price=0
    YIELD=0
    LAST=0	
    DURATION=0   
    
    data=j['marketdata']['data']
    for i in data:  
        for f, b in zip(j['marketdata']['columns'], i):
            if f=="SECID":
                if b is not None:
                    SECID1=b        
            if f=="LAST":
                if b is not None:
                    last_price=b
            if f=="MARKETPRICE":
                if b is not None:
                    market_price=b  
            if f=="YIELD":
                if b is not None:
                    YIELD=b      
            if f=="DURATION":
                if b is not None:
                    DURATION=b                   
        
            if last_price==0 and market_price>0:
                last_price=market_price
        
        if SECID1 in all_moex_bonds:
            values=all_moex_bonds[SECID1]
            values["LAST"]=last_price
            values["YIELD"]=YIELD
            values["DURATION"]=DURATION
            all_moex_bonds[SECID1]=values
        else:
            all_moex_bonds[SECID1]={"LAST":last_price, "YIELD":YIELD, "DURATION":DURATION}                        
                
    SECID2=0
    BOARDID=0
    PRICE=0
    YIELDDATE=0
    EFFECTIVEYIELD=0
    DURATION=0
    ZSPREADBP=0
    GSPREADBP=0

    data=j['marketdata_yields']['data']
    for i in data:          
        for f, b in zip(j['marketdata_yields']['columns'], i):
            if f=="SECID":
                if b is not None:
                    SECID2=b   
            if f=="BOARDID":
                if b is not None:
                    BOARDID=b  
            if f=="PRICE":
                if b is not None:
                    PRICE=b  
            if f=="YIELDDATE":
                if b is not None:
                    YIELDDATE=b  
            if f=="EFFECTIVEYIELD":
                if b is not None:
                    EFFECTIVEYIELD=b  
            if f=="DURATION":
                if b is not None:
                    DURATION=b    
            if f=="ZSPREADBP":
                if b is not None:
                    ZSPREADBP=b       
            if f=="GSPREADBP":
                if b is not None:
                    GSPREADBP=b  
                    
        if SECID2 in all_moex_bonds:
            values=all_moex_bonds[SECID2]
            values["BOARDID"]=BOARDID
            values["PRICE"]=PRICE
            values["YIELDDATE"]=YIELDDATE
            values["EFFECTIVEYIELD"]=EFFECTIVEYIELD
            values["DURATION"]=DURATION
            values["ZSPREADBP"]=ZSPREADBP
            values["GSPREADBP"]=GSPREADBP
            
            all_moex_bonds[SECID2]=values
        else:
            all_moex_bonds[SECID2]={"BOARDID":BOARDID, "PRICE":PRICE, "YIELDDATE":YIELDDATE, "EFFECTIVEYIELD":EFFECTIVEYIELD, "DURATION":DURATION, "ZSPREADBP":ZSPREADBP, "GSPREADBP":GSPREADBP}                                    
                                

    f_name=f'Export_files\\all_moex_bonds-{today_str}.xlsx'
    workbook = xlsxwriter.Workbook(f_name)
    worksheet = workbook.add_worksheet("all_moex")
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    #date_format = workbook.add_format({'num_format': 'dd/mm/yy'})
    
    worksheet.write('A1', 'SECID', bold)
    worksheet.write('C1', 'ISIN', bold)
    worksheet.write('B1', 'BOARDID', bold)
    worksheet.write('D1', 'SHORTNAME', bold)
    worksheet.write('E1', 'SECNAME', bold)
    worksheet.write('F1', 'YIELD', bold)
    worksheet.write('G1', 'LAST', bold)    
    worksheet.write('H1', 'COUPONPERCENT', bold)
    worksheet.write('I1', 'Current coupon', bold)    
    worksheet.write('J1', 'FACEVALUE', bold)
    worksheet.write('K1', 'MATDATE', bold)
    worksheet.write('L1', 'COUPONPERIOD', bold)    
    worksheet.write('M1', 'DURATION', bold)
    worksheet.write('N1', 'PRICE', bold)
    worksheet.write('O1', 'YIELDDATE', bold)    
    worksheet.write('P1', 'EFFECTIVEYIELD', bold)
    worksheet.write('Q1', 'ZSPREADBP', bold)
    worksheet.write('R1', 'GSPREADBP', bold)
        
    row = 1
    col = 0           
    for item in all_moex_bonds:
        worksheet.write(row, col,     item)
        worksheet.write(row, col+1,     all_moex_bonds[item]["BOARDID"])
        worksheet.write(row, col+2,     all_moex_bonds[item]["ISIN"])        
        worksheet.write(row, col+3,     all_moex_bonds[item]["SHORTNAME"])
        worksheet.write(row, col+4,     all_moex_bonds[item]["SECNAME"])
        worksheet.write(row, col+5,     all_moex_bonds[item]["YIELD"])
        worksheet.write(row, col+6,     all_moex_bonds[item]["LAST"])
        worksheet.write(row, col+7,     all_moex_bonds[item]["COUPONPERCENT"])
        try:
            worksheet.write(row, col+8,     all_moex_bonds[item]["COUPONPERCENT"]/all_moex_bonds[item]["LAST"])
        except:
            worksheet.write(row, col+8,     "")
        worksheet.write(row, col+9,     all_moex_bonds[item]["FACEVALUE"])
        worksheet.write(row, col+10,     all_moex_bonds[item]["MATDATE"])
        worksheet.write(row, col+11,     all_moex_bonds[item]["COUPONPERIOD"])
        worksheet.write(row, col+12,     all_moex_bonds[item]["DURATION"])
        worksheet.write(row, col+13,     all_moex_bonds[item].get("PRICE"))
        worksheet.write(row, col+14,     all_moex_bonds[item].get("YIELDDATE"))
        worksheet.write(row, col+15,     all_moex_bonds[item].get("EFFECTIVEYIELD"))
        worksheet.write(row, col+16,     all_moex_bonds[item].get("ZSPREADBP"))
        worksheet.write(row, col+17,     all_moex_bonds[item].get("GSPREADBP"))
        
        row+=1
    
    workbook.close()
    
    
    print("Bonds screening completed!")
    return

def get_excluded_bonds_from_file(filename="Export_files\\bonds_exclude.csv"):
    bonds_exclude={}
    
    with open(filename, 'r', encoding='latin-1') as csvfile:  
        # Use csv.reader to read the file  
        csvreader = csv.reader(csvfile, delimiter=';')  
          
        # Iterate over each row in the CSV file  
        for row in csvreader:  
            if len(row) == 2:  
                key, value = row
                key=key.strip()
                value=value.strip()
                bonds_exclude[key] = value
    
    return bonds_exclude
    

def make_recommendations(test_isin="RU000A106EP1"):
    global all_moex_bonds
    
    bonds_exclude=get_excluded_bonds_from_file()
    
    d = datetime.datetime.today()
    today_str=d.strftime("%Y%m%d") 
    
    connection = sqlite3.connect('portfolio_database.db')
    cursor = connection.cursor() 
    sql_str=f'select p.isin from portfolio p join trading_instruments ti on ti.isin=p.isin join bonds_static bs on bs.isin=p.isin where qty>0 and ti.instrument_type="bond" and p.portfolio_id="Alexey" and (bs.percent_type is null or bs.percent_type not in ("linker", "float"))'
    cursor.execute(sql_str)
    tbl = cursor.fetchall()
    
    f_name=f'Export_files\\recomendations_moex_bonds-{today_str}.xlsx'
    workbook = xlsxwriter.Workbook(f_name)
    
   
    
    for item in tbl:
        row=1
        col=0         
        isin=item[0]
        #if isin!=test_isin:
            #continue
        
        instr=0
        for instrid in all_moex_bonds:
            if all_moex_bonds[instrid]["ISIN"]==isin:
                instr=instrid
                break
        
        sheet_name=f'{all_moex_bonds[instr]["ISIN"]}-{all_moex_bonds[instr]["SHORTNAME"]}'
        worksheet = workbook.add_worksheet(sheet_name)
        bold = workbook.add_format({'bold': True})
        worksheet.write('A1', 'SECID', bold)
        worksheet.write('C1', 'ISIN', bold)
        worksheet.write('B1', 'BOARDID', bold)
        worksheet.write('D1', 'SHORTNAME', bold)
        worksheet.write('E1', 'SECNAME', bold)
        worksheet.write('F1', 'YIELD', bold)
        worksheet.write('G1', 'LAST', bold)    
        worksheet.write('H1', 'COUPONPERCENT', bold)
        worksheet.write('I1', 'Current coupon', bold)    
        worksheet.write('J1', 'FACEVALUE', bold)
        worksheet.write('K1', 'MATDATE', bold)
        worksheet.write('L1', 'COUPONPERIOD', bold)    
        worksheet.write('M1', 'DURATION', bold)
        worksheet.write('N1', 'PRICE', bold)
        worksheet.write('O1', 'YIELDDATE', bold)    
        worksheet.write('P1', 'EFFECTIVEYIELD', bold)
        worksheet.write('Q1', 'ZSPREADBP', bold)
        worksheet.write('R1', 'GSPREADBP', bold)
        worksheet.write('S1', 'Reco_type', bold)
        
        worksheet.write(row, col,     instr)
        worksheet.write(row, col+1,     all_moex_bonds[instr]["BOARDID"])
        worksheet.write(row, col+2,     all_moex_bonds[instr]["ISIN"])        
        worksheet.write(row, col+3,     all_moex_bonds[instr]["SHORTNAME"])
        worksheet.write(row, col+4,     all_moex_bonds[instr]["SECNAME"])
        worksheet.write(row, col+5,     all_moex_bonds[instr]["YIELD"])
        worksheet.write(row, col+6,     all_moex_bonds[instr]["LAST"])
        worksheet.write(row, col+7,     all_moex_bonds[instr]["COUPONPERCENT"])
        try:
            worksheet.write(row, col+8,     all_moex_bonds[instr]["COUPONPERCENT"]/all_moex_bonds[instr]["LAST"])
        except:
            worksheet.write(row, col+8,     "")
        worksheet.write(row, col+9,     all_moex_bonds[instr]["FACEVALUE"])
        worksheet.write(row, col+10,     all_moex_bonds[instr]["MATDATE"])
        worksheet.write(row, col+11,     all_moex_bonds[instr]["COUPONPERIOD"])
        worksheet.write(row, col+12,     all_moex_bonds[instr]["DURATION"])
        worksheet.write(row, col+13,     all_moex_bonds[instr].get("PRICE"))
        worksheet.write(row, col+14,     all_moex_bonds[instr].get("YIELDDATE"))
        worksheet.write(row, col+15,     all_moex_bonds[instr].get("EFFECTIVEYIELD"))
        worksheet.write(row, col+16,     all_moex_bonds[instr].get("ZSPREADBP"))
        worksheet.write(row, col+17,     all_moex_bonds[instr].get("GSPREADBP"))        
        
        row+=1
        col=0
        
        duration0=all_moex_bonds[instr]["DURATION"]
        price=all_moex_bonds[instr]["LAST"]
        
        for item2 in all_moex_bonds:
            instrid2=item2
            isin2=all_moex_bonds[instrid2]["ISIN"]
            if isin2==isin:
                continue
            
            skip=0
            for exclude in bonds_exclude:
                #print(f'{exclude}, {all_moex_bonds[instrid2]["SHORTNAME"]}')
                if exclude in all_moex_bonds[instrid2]["ISIN"]:
                    #print(f'{exclude}, {all_moex_bonds[instrid2]["ISIN"]}')
                    skip+=1
            
            if skip>0:
                continue
                        
            duration=all_moex_bonds[instrid2]["DURATION"]
            duration_low=duration0*(1.0-30.0/100)
            duration_top=duration0*(1.0+20.0/100)
            price2=all_moex_bonds[instrid2]["LAST"]
            price2_low=price*(1-50/100)
            price2_top=price*(1+3/100)
            #print(f'duration target: {duration0}. Duration search: {duration_low}, {duration_top}')
            if duration<=duration_top and duration>=duration_low and price2>=price2_low and price2<=price2_top:
                worksheet.write(row, col,     instrid2)
                worksheet.write(row, col+1,     all_moex_bonds[instrid2]["BOARDID"])
                worksheet.write(row, col+2,     all_moex_bonds[instrid2]["ISIN"])                
                worksheet.write(row, col+3,     all_moex_bonds[instrid2]["SHORTNAME"])
                worksheet.write(row, col+4,     all_moex_bonds[instrid2]["SECNAME"])
                worksheet.write(row, col+5,     all_moex_bonds[instrid2]["YIELD"])
                worksheet.write(row, col+6,     all_moex_bonds[instrid2]["LAST"])
                worksheet.write(row, col+7,     all_moex_bonds[instrid2]["COUPONPERCENT"])
                try:
                    worksheet.write(row, col+8,     all_moex_bonds[instrid2]["COUPONPERCENT"]/all_moex_bonds[instrid2]["LAST"])
                except:
                    worksheet.write(row, col+8,     "")
                worksheet.write(row, col+9,     all_moex_bonds[instrid2]["FACEVALUE"])
                worksheet.write(row, col+10,     all_moex_bonds[instrid2]["MATDATE"])
                worksheet.write(row, col+11,     all_moex_bonds[instrid2]["COUPONPERIOD"])
                worksheet.write(row, col+12,     all_moex_bonds[instrid2]["DURATION"])
                worksheet.write(row, col+13,     all_moex_bonds[instrid2].get("PRICE"))
                worksheet.write(row, col+14,     all_moex_bonds[instrid2].get("YIELDDATE"))
                worksheet.write(row, col+15,     all_moex_bonds[instrid2].get("EFFECTIVEYIELD"))
                worksheet.write(row, col+16,     all_moex_bonds[instrid2].get("ZSPREADBP"))
                worksheet.write(row, col+17,     all_moex_bonds[instrid2].get("GSPREADBP"))   
                worksheet.write(row, col+18,     "1")  # Recommendation #1 - by duration
                
                row+=1
                col=0                
                

    
    workbook.close()
    print("Recommendations completed!")


if __name__=="__main__":
    get_all_bonds_from_moex()
    make_recommendations()