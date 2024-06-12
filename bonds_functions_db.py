import shutil
import datetime
import requests
import pandas as pd
import plotly.subplots as ps
import plotly.graph_objs as go
from sortedcontainers import SortedDict
import settings
import xlsxwriter
import re
import sqlite3

portfolio_ext = SortedDict()

ratings={'Gov':27, 'AAA':26, 'AAA-':25, 'AA+':24, 'AA':23, 'AA-':22, 'A+':21, 'A':20, 'A-':19, 'BBB+':18, 'BBB':17, 'BBB-':16, 'BB+':15, 'BB':14, 'BB-':13, 'B+':12, 'B':11, 'B-':10 ,'CCC+':9, 'CCC':8, 'CCC-':7, 'CC+':6, 'CC':5, 'CC-':4, 'C+':3, 'C':2, 'C-':1, 'DDD':0}
cross_rates={'USD':92.1}

def get_bond_maturity(cursor, isin):
    d = datetime.datetime(1900, 1, 1, 0, 1)
    #cursor = connection.cursor()
    
    sql_str=f'select max(date) from bonds_schedule WHERE ISIN like "{isin}"'
    cursor.execute(sql_str)
    maturity_date = cursor.fetchone()[0]
    d = datetime.datetime.strptime(maturity_date, '%Y%m%d')
    
    return d

def get_bond_rating(cursor, isin):
    #cursor = connection.cursor()
    
    sql_str=f'select rating, ifnull(percent_type,"fixed") as type_ from bonds_static WHERE ISIN like "{isin}"'
    cursor.execute(sql_str)
    sql_res = cursor.fetchone()    
    results={'rating':sql_res[0], 'type': sql_res[1]}
    
    sql_str=f'select distinct ifnull(nominal_currency,"RUB") as currency from bonds_schedule where isin like "{isin}"'
    cursor.execute(sql_str)
    sql_res = cursor.fetchone()
    results['currency']=sql_res[0]
       
    return results

def update_fcy_rates():
    req_str='https://iss.moex.com/iss/statistics/engines/currency/markets/fixing.json'
    j=requests.get(req_str).json()
    for b in zip(j['history']['data']):    
        if b[0][1]=='USDFIXME':
            cross_rates['USD']=float(b[0][2])
            print(cross_rates)    
    

def get_bond_type_by_rating(cursor, isin):
    r=get_bond_rating(cursor, isin)
    rating=r['rating']
    type_=r['type']    
    currency=r['currency']
    
    num=ratings.get(rating, -1)
    
    if num == 27:
        return 'Gov'
    elif num<=26 and num>=16:
        if type_=='float':
            return 'Corp-fl'
        elif currency!="RUB":
            return f'Corp-{currency}'
        else:
            return 'Corp'
    elif num<=15 and num>=0:
        if type_=='float':
            return 'VDO-fl'    
        else:
            return 'VDO'
    elif num==-1:
        return 'wrong rating!'
    
    return 'none'

def get_bond_nearest_coupon_date(cursor, isin):
    d = datetime.datetime.today()
    #cursor = connection.cursor()
    today_str=d.strftime("%Y%m%d")
    
    sql_str=f'select min(date) from bonds_schedule where isin like "{isin}" and date>"{today_str}"'
    cursor.execute(sql_str)
    nearest_coupon = cursor.fetchone()[0]
    if nearest_coupon is not None:
        d = datetime.datetime.strptime(nearest_coupon, '%Y%m%d')
    
    return d


def get_bond_nearest_coupon(cursor, isin):
    d = datetime.datetime.today()
    #cursor = connection.cursor()
    today_str=d.strftime("%Y%m%d")
    
    sql_str=f'select pct_value     from bonds_schedule     where isin like "{isin}" and date = (select min(date) from bonds_schedule where isin like "{isin}" and date>"{today_str}") '
    cursor.execute(sql_str)
    val = cursor.fetchone()
    if val is None:
        return 0
    return val[0]



def get_bond_info_moex(isin):
    secid=""
    shortname=""
    req_str='https://iss.moex.com/iss/securities.json?q='+isin+"'"
    j=requests.get(req_str).json() #Получить инструмент по isin коду  #'https://iss.moex.com/iss/securities.json?q=RU000A105XV1'
    if len(j['securities']['data'])<1:
        print('Security ID for ',isin,' isnt found on MOEX API!')
        return {}
    
    for f, b in zip(j['securities']['columns'], j['securities']['data'][0]):
        if f=="secid":
            secid=b
        if f=="shortname":
            shortname=b                            
                
    req_str='https://iss.moex.com/iss/engines/stock/markets/bonds/securities/'+secid+'.json?marketprice_board=1'
    nkd=0
    nominal=0
    last_price=0
    j=requests.get(req_str).json()  #'https://iss.moex.com/iss/engines/stock/markets/bonds/securities/RU000A106Z38.json?marketprice_board=1'
    for f, b in zip(j['securities']['columns'], j['securities']['data'][0]):
        if f=="ACCRUEDINT":
            nkd=b
        if f=="FACEVALUE":
            nominal=b        
    
    last_price=0
    for f, b in zip(j['marketdata']['columns'], j['marketdata']['data'][0]):
        if f=="LAST":
            if b is not None:
                last_price=b
    
    market_price=0
    for f, b in zip(j['marketdata']['columns'], j['marketdata']['data'][0]):
        if f=="MARKETPRICE":
            if b is not None:
                market_price=b
    
    if last_price==0 and market_price>0:
        last_price=market_price
        
    bond_yield=0
    for f, b in zip(j['marketdata']['columns'], j['marketdata']['data'][0]):
        if f=="YIELD":
            if b is not None:
                bond_yield=b    
                
    bond_duration=0
    for f, b in zip(j['marketdata']['columns'], j['marketdata']['data'][0]):
        if f=="DURATION":
            if b is not None:
                bond_duration=b     
                
    coupon_period=0
    issue_size=0
    current_coupon=0  
    bond_currency='RUB'  
    for f, b in zip(j['securities']['columns'], j['securities']['data'][0]):
        if f=="COUPONPERIOD":
            if b is not None:
                coupon_period=b    
        if f=="ISSUESIZEPLACED":
            if b is not None:
                issue_size=b    
        if f=="COUPONPERCENT":
            if b is not None:
                current_coupon=b                    
        if f=="FACEUNIT":
            if b is not None:
                bond_currency=b                 
                            
    bond_info={}
    bond_info["isin"]=isin
    bond_info["secid"]=secid
    bond_info["shortname"]=shortname
    bond_info["nkd"]=nkd
    bond_info["nominal"]=nominal
    bond_info["last_price"]=last_price
    full_price=0.0
    if bond_currency in ["USD"]:
        full_price=nominal*last_price*cross_rates.get(bond_currency)/100+nkd
    else:
        full_price=nominal*last_price/100+nkd
    bond_info["full_price"]=full_price
    bond_info["yield"]=bond_yield
    bond_info["duration"]=bond_duration
    bond_info["coupon_period"]=coupon_period
    bond_info["issue_size"]=issue_size
    bond_info["current_coupon"]=current_coupon
    bond_info["bond_currency"]=bond_currency

    return bond_info

def calc_portfolio_pct_days(days=365):
    # calculate payments pcts in portfolio betwen current date and DAYS 
    accrual_pct=0
    start_date=datetime.datetime.today()
    end_date=start_date + datetime.timedelta(days=days)
    

    for i in portfolio_ext:
        if portfolio_ext[i].get("cf",0)==0:
            continue
        cf=portfolio_ext[i]["cf"]
        count=portfolio_ext[i]["count"]
        
        for j in cf:
            date_=j["date"]
            coupon=j["coupon"] 
            amo=j["amortization"]            
            
            if date_>=start_date and date_<=end_date:
                    accrual_pct=accrual_pct+coupon*count

    return accrual_pct


def calc_bond_portfolio_value(cursor):
    total_val=0
    d = datetime.datetime.today()
    today_str=d.strftime("%Y%m%d")
    
    sql_str=f'SELECT isin, qty FROM bond_portfolio WHERE qty>0 '
    cursor.execute(sql_str)
    tbl = cursor.fetchall()
        
    for item in tbl:
        data=get_bond_info_moex(item[0])
        total_val=total_val+data["full_price"]*item[1]
        
    sql_str=f'select count(price) from market_data where id="my_portfolio" and date="{today_str}"'
    cursor.execute(sql_str)
    fetch_cnt = cursor.fetchone()[0]
    
    if fetch_cnt==0:
        sql_str=f'insert into market_data(id, date, price, price_nominal) values ("my_portfolio", "{today_str}", {total_val}, "RUB")'
        cursor.execute(sql_str)        
    else:
        sql_str=f'update market_data set price={total_val} where id="my_portfolio" and date="{today_str}"'
        cursor.execute(sql_str)
        
        
    return total_val

def get_current_bond_nominal(cursor, isin, on_date=datetime.datetime.today()):
    initial=1000
    #cursor = connection.cursor()
    
    d = datetime.datetime.today()
    today_str=d.strftime("%Y%m%d")
    
    sql_str=f'select sum(nominal_value) from bonds_schedule where isin like "{isin}" and date<"{today_str}" '
    cursor.execute(sql_str)
    sum_amortizations = cursor.fetchone()[0]
    sum_amortizations=(0 if sum_amortizations is None else sum_amortizations)

    return initial-sum_amortizations

def create_allocation_pie_chart():
    values=[]
    labels=[]
    for i in portfolio_ext:
        labels.append(portfolio_ext[i]['moex_code'])
        values.append(portfolio_ext[i]['count']*get_current_bond_nominal(i))
        
    fig2= go.Figure(data=go.Pie(
        labels=labels,
        values=values))
    fig2.show()    
    
    return 0


def create_cash_flows_graph(cursor, calc_type=1):
    # calc_type == 1 -> multiply by qty in portfolio
    # calc_type == 2 -> DON't multiply by qty in portfolio  
    # calc_type == 3 -> only coupon * qty  
    # calc_type == 4 -> only amortization * qty
    
    cf_all = SortedDict()
    for i in portfolio_ext:
        if portfolio_ext[i].get("cf",0)==0:
            continue
        cf=portfolio_ext[i]["cf"]
        count=portfolio_ext[i]["count"]
        
        for j in cf:
            date_=j["date"]
            coupon=j["coupon"] 
            amo=j["amortization"]            
            
            if date_>=datetime.datetime.today():
                if date_ in cf_all:
                    if calc_type==1:
                        cf_all[date_]=cf_all[date_]+(amo+coupon)*count
                    if calc_type==2:
                        cf_all[date_]=cf_all[date_]+(amo+coupon)                   
                    if calc_type==3:
                        cf_all[date_]=cf_all[date_]+coupon*count                        
                    if calc_type==4:
                        cf_all[date_]=cf_all[date_]+amo*count                              
                    
                else:
                    if calc_type==1:
                        cf_all[date_]=(amo+coupon)*count
                    if calc_type==2:
                        cf_all[date_]=(amo+coupon)    
                    if calc_type==3:
                        cf_all[date_]=coupon*count
                    if calc_type==4:
                        cf_all[date_]=amo*count                        
            
    
    p=pd.DataFrame.from_dict(cf_all.items())
    p.columns=['date', 'amount']
    p['date'] = p['date'].dt.strftime('%Y-%m')
    r=p.groupby('date')['amount'].sum().reset_index()  
    
    fig1 = go.Figure()    
    fig1.add_trace(go.Bar(x=r['date'], y=r['amount'], text=round(r['amount']), texttemplate="%{y:,.0f}"))
    fig1.layout = dict(xaxis=dict(type="category"))  
    fig1.show()    
    
    return 0

def create_cash_flows_graph4(calc_type=1):
    # calc_type == 1 -> multiply by qty in portfolio
    # calc_type == 2 -> DON't multiply by qty in portfolio  
    # calc_type == 3 -> only coupon * qty  
    # calc_type == 4 -> only amortization * qty
    
    cf_all = SortedDict()
    xs_dates=set()
    #min_date=datetime.datetime.today().replace(day=1)
    #max_date=datetime.datetime.today()+relativedelta(day=32)
    
    for i in portfolio_ext:
        #if 'RU000A102RU2'==portfolio_ext[i]["isin"]:
            #print('RU000A102RU2')
        if portfolio_ext[i].get("cf",0)==0:
            continue
        cf=portfolio_ext[i]["cf"]
        
        for j in cf:
            if j["amortization"]>0 and j["date"]>=datetime.datetime.today():
                date_=j["date"]
                xs_dates.add(date_.replace(day=1))
        
    xl_dates=list(xs_dates)
    xl_dates.sort()
           
    fig1 = go.Figure()
    
    for i in portfolio_ext:
        if portfolio_ext[i].get("cf",0)==0:
            continue
        cf=portfolio_ext[i]["cf"]
        
        yl_values=[]
        
        for k in range(0, len(xl_dates)):
            value=0
            for j in cf:                
                if j["amortization"]>0 and j["date"]>=datetime.datetime.today() and j["date"].replace(day=1)==xl_dates[k]:
                    value=j["amortization"]*portfolio_ext[i]["count"]
            yl_values.append(value)
        fig1.add_trace(go.Bar(x=xl_dates, y=yl_values, name=portfolio_ext[i]["moex_code"]))  
            
    fig1.update_layout(barmode='stack')
    #fig1.layout = dict(xaxis=dict(type="category"))
    fig1.show()       
    
    return 0

def create_cash_flows_graph4_1(calc_type=1):
    # calc_type == 1 -> multiply by qty in portfolio
    # calc_type == 2 -> DON't multiply by qty in portfolio  
    # calc_type == 3 -> only coupon * qty  
    # calc_type == 4 -> only amortization * qty
    
    cf_all = SortedDict()
    xs_dates=set()
    #min_date=datetime.datetime.today().replace(day=1)
    #max_date=datetime.datetime.today()+relativedelta(day=32)
    
    for i in portfolio_ext:
        if 'RU000A102RU2'==portfolio_ext[i]["isin"]:
            print('RU000A102RU2')
        if portfolio_ext[i].get("cf",0)==0:
            continue
        cf=portfolio_ext[i]["cf"]
        
        for j in cf:
            if j["amortization"]>0 and j["date"]>=datetime.datetime.today():
                date_=j["date"]
                xs_dates.add(date_.replace(day=1))
        
    xl_dates=list(xs_dates)
    xl_dates.sort()
           
    fig1 = go.Figure()
    
    for i in portfolio_ext:
        if portfolio_ext[i].get("cf",0)==0:
            continue
        cf=portfolio_ext[i]["cf"]
        
        yl_values=[]
        
        for k in range(0, len(xl_dates)):
            value=0
            for j in cf:                
                if j["amortization"]>0 and j["date"]>=datetime.datetime.today() and j["date"].replace(day=1)==xl_dates[k]:
                    value=j["amortization"]*portfolio_ext[i]["count"]
            yl_values.append(value)
        fig1.add_trace(go.Bar(x=xl_dates, y=yl_values, name=portfolio_ext[i]["moex_code"]))  
            
    fig1.update_layout(barmode='stack')
    #fig1.layout = dict(xaxis=dict(type="category"))
    fig1.show()       
    
    return 0

def print_portfolio():
    tickers=[]
    isins=[]
    qty=[]
    matty=[]
    next_coupons=[]
    
    for i in portfolio_ext:
        tickers.append(portfolio_ext[i]["moex_code"])
        isins.append(portfolio_ext[i]["isin"])
        qty.append(portfolio_ext[i]["count"])
        matty.append(get_bond_maturity(portfolio_ext[i]["isin"]).strftime('%Y-%m-%d'))
        next_coupons.append(get_bond_nearest_coupon(portfolio_ext[i]["isin"]).strftime('%Y-%m-%d'))
    
    fig = go.Figure(data=[go.Table(header=dict(values=['Ticker', 'Isin', 'Quantity', 'Maturity', 'Next coupon']),
                     cells=dict(values=[tickers, isins, qty, matty, next_coupons] ))
                         ])
    fig.show() 
    

def print_portfolio_console():
    
    for i in portfolio_ext:
        ticker=portfolio_ext[i]["moex_code"]
        isin=portfolio_ext[i]["isin"]
        qty=portfolio_ext[i]["count"]
        matty=get_bond_maturity(portfolio_ext[i]["isin"]).strftime('%Y-%m-%d')
        print(f'{isin};{qty};{ticker};{matty}')

def check_cfs_portfolio():
        
    for i in portfolio_ext:
        cf=portfolio_ext[i].get('cf',0)
        ticker=portfolio_ext[i]["moex_code"]
        isin=portfolio_ext[i]["isin"]
        qty=portfolio_ext[i]["count"]
        
        if cf==0:
            print(f'No cash-flows for {isin}. Remove line from bonds_portfolio file ( {isin}, {qty}, {ticker} ) OR add cash-flows for the bond !!!')


def calc_full_fair_value(quontity, moex_data):
    fv=0.0

    moex_full_price=moex_data["full_price"]
    fv=quontity*moex_full_price
    
    if moex_data.get("bond_currency") not in ['SUR', 'RUB']:
        fv=fv*cross_rates.get(moex_data.get("bond_currency", 1))
    
    return fv


def print_portfolio_excel():

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('my_portfolio.xlsx')
    worksheet = workbook.add_worksheet()
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    date_format = workbook.add_format({'num_format': 'dd/mm/yy'})
    
    # Write some data headers.
    worksheet.write('A1', 'Ticker', bold)
    worksheet.write('B1', 'Isin', bold)
    worksheet.write('C1', 'Quantity', bold)
    worksheet.write('D1', 'Maturity', bold)
    worksheet.write('E1', 'Next coupon date', bold)
    worksheet.write('F1', 'Next coupon', bold)
    worksheet.write('G1', 'Current nominal', bold)
    worksheet.write('H1', 'Rating', bold)
    worksheet.write('I1', 'Yield', bold)
    worksheet.write('J1', 'Fair value', bold)
    worksheet.write('K1', 'Duration', bold)
    worksheet.write('L1', 'Duration_years', bold)
    worksheet.write('M1', 'Bond_Type', bold)
    worksheet.write('N1', 'Last_Price', bold)
    worksheet.write('O1', 'Coupon_yield', bold)
    worksheet.write('P1', 'Coupon_period', bold)
    worksheet.write('Q1', 'Issue_size', bold)
     
    # Start from the first cell below the headers.
    row = 1
    col = 0    
    for i in portfolio_ext:
        moex_data=get_bond_info_moex(portfolio_ext[i]["isin"])
        
        worksheet.write(row, col,     portfolio_ext[i]["moex_code"])
        worksheet.write(row, col + 1, portfolio_ext[i]["isin"])
        worksheet.write(row, col + 2, portfolio_ext[i]["count"])
        worksheet.write_datetime(row, col + 3, get_bond_maturity(portfolio_ext[i]["isin"]), date_format)
        worksheet.write_datetime(row, col + 4, get_bond_nearest_coupon_date(portfolio_ext[i]["isin"]), date_format)
        worksheet.write(row, col + 5, portfolio_ext[i]["count"]*get_bond_nearest_coupon(portfolio_ext[i]["isin"]))
        worksheet.write(row, col + 6, get_current_bond_nominal(portfolio_ext[i]["isin"]) )
        worksheet.write(row, col + 7, portfolio_ext[i].get("rating") )        
        worksheet.write(row, col + 8, moex_data["yield"] )
        fullFV=calc_full_fair_value(portfolio_ext[i]["count"], moex_data)
        #worksheet.write(row, col + 9, portfolio_ext[i]["count"]*moex_data["full_price"] )
        worksheet.write(row, col + 9, fullFV)
        
        worksheet.write(row, col + 10, moex_data["duration"] )        
        worksheet.write(row, col + 11, moex_data["duration"]/365 )
        worksheet.write(row, col + 12, get_bond_type_by_rating(ratings, portfolio_ext[i].get("rating")) )
        worksheet.write(row, col + 13, moex_data["last_price"])
        worksheet.write(row, col + 14, moex_data["current_coupon"]/moex_data["last_price"])
        worksheet.write(row, col + 15, moex_data["coupon_period"])
        worksheet.write(row, col + 16, moex_data["issue_size"])
                        
        row += 1

    # Write a total using a formula.
    workbook.close()
    

def read_portfolio_from_txt(fname):
    global portfolio_ext    
    read_pos=open(fname, 'r', encoding='utf-8').read().splitlines() 
    #print("Reading a file %s..." % (fname))    
    
    for line in read_pos:
        line.rstrip('\n').replace("\n", "")
        l1=line.split(';')
        if (len(l1))<2:
            continue

        #if l1[0]!="RU000A106HB4":
            #continue
        
        elems={"count":float(l1[1]), "moex_code":l1[2], "isin":l1[0]}
        if l1[0] in portfolio_ext:
            portfolio_ext[l1[0]]["count"]=portfolio_ext[l1[0]]["count"]+float(l1[1])
        else:
            portfolio_ext[l1[0]]=elems
    
    #print(portfolio_ext)
    print(str(len(portfolio_ext))+" instruments in portfolio")    
    return 0


def read_bond_from_txt(cursor, fname):
    
    if fname=="bonds_portfolio.txt":
        return 0
    
    read_rates=open(fname, 'r').read().splitlines() 
    #print("Reading a file %s..." % (fname))
    isin=""
    rating=""
     
    for i in range(0, len(read_rates)):
        line=read_rates[i]
        line.rstrip('\n').replace("\n", "")
        l1=line.split(';')
        
        if l1[0].startswith('isin') or l1[0].startswith('Isin') or l1[0].startswith('ISIN'):
            l2=str(l1[0]).strip()
            l2=line.split(':')
            isin=str(l2[1])
            if len(isin)!=12:
                print(f'Isin code {isin} has length not equl 12. Error, processing this file {fname} stoped!')
                break
            continue            
        
        if i==0 and len(str(l1[0]).strip())==12: 
            isin=str(l1[0]).strip()
            continue
                
        if l1[0].startswith('rating') or l1[0].startswith('Rating') or l1[0].startswith('RATING'):
            l2=str(l1[0]).strip()
            l2=line.split(':')
            rating=str(l2[1])
            continue          
        
        sql_str=f'SELECT count(*) FROM bonds_static WHERE 1=1 and ISIN like "{isin}"'
        cursor.execute(sql_str)
        fetch_cnt = cursor.fetchone()[0]
        
        if fetch_cnt==0:
            sql_str=f'insert into bonds_static(isin, rating) values("{isin}", "{rating}")'
            cursor.execute(sql_str)
            print(f'Inserted: isin={isin}, rating={rating}')
       
        #connection.commit()
                
        if len(l1)>1:            
            sd=str(l1[0])
            coupon=0
            if re.match(r'^-?\d+(?:\.\d+)$', l1[1]) is not None or l1[1].isnumeric():
                coupon=float(l1[1])
                
            amortization=0            
            if re.match(r'^-?\d+(?:\.\d+)$', l1[2]) is not None or l1[2].isnumeric():
                amortization=float(l1[2])
                
            if len(sd)==8:
                date_=datetime.datetime.strptime(sd, '%d.%m.%y')
            if len(sd)==10:
                date_=datetime.datetime.strptime(sd, '%d.%m.%Y')                
            cf_element={"date":date_, "coupon":coupon, "amortization":amortization}
            db_date_insert=date_.strftime("%Y%m%d")
            
            sql_str=f'SELECT count(*) FROM bonds_schedule WHERE 1=1 and ISIN like "{isin}" and date="{db_date_insert}"'
            cursor.execute(sql_str)
            fetch_cnt = cursor.fetchone()[0]
            if fetch_cnt==0:
                sql_str=f'insert into bonds_schedule(isin, date, pct_value, nominal_value) values("{isin}", "{db_date_insert}", {coupon}, {amortization})'
                cursor.execute(sql_str)
                print(f'Inserted: isin={isin}, date={db_date_insert}')
           
            #connection.commit()            
            
    return 0



def calc_months_return():
    mcash_flow=SortedDict()
    
    return 0

def calc_bond_duration(isin):
    d=0  # duratiuon
    today_= datetime.datetime.today()
    
    cf=portfolio_ext[isin].get('cf', 0)
    
    if cf == 0:
        print(f"No cash-flows for {isin}")
        return 0

    nom=0
    for i in range(0, len(cf)):
        if cf[i]["date"]>today_:
            days_between=(cf[i]["date"]-today_).days
            nom=nom+cf[i]["coupon"]*days_between+cf[i]["amortization"]*days_between
    
    bond_data=get_bond_info_moex(isin)
    denominator=bond_data["full_price"]
    
    d=nom/denominator
    
    return d

def calc_last_day_of_month(date):
    d=date.replace(day=1).replace(hour=0, minute=0, second=0, microsecond=0)
    fist_day_next_month=(d + datetime.timedelta(days=33)).replace(day=1)
    end_date=(fist_day_next_month-datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    return end_date




