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
cross_rates={'USD':99}

def get_bond_maturity(cursor, isin):
    d = datetime.datetime(1900, 1, 1, 0, 1)
    #cursor = connection.cursor()
    
    sql_str=f'select count(*) from bonds_schedule WHERE ISIN like "{isin}"'
    cursor.execute(sql_str)
    maturity_date = cursor.fetchone()[0]
    
    if maturity_date>0:    
        sql_str=f'select max(date) from bonds_schedule WHERE ISIN like "{isin}"'
        cursor.execute(sql_str)
        maturity_date = cursor.fetchone()[0]
        d = datetime.datetime.strptime(maturity_date, '%Y%m%d')
    else:
        sql_str=f'select maturity_date from bonds_static WHERE ISIN like "{isin}"'
        cursor.execute(sql_str)
        maturity_date = cursor.fetchone()[0]
        d = datetime.datetime.strptime(maturity_date, '%Y%m%d')        
    
    return d

def get_EntityUTI_by_Name(cursor, name):
    
    sql_str=f'select count(1) from entity WHERE short_name like "{name}" '
    cursor.execute(sql_str)
    tbl = cursor.fetchone()
    #print(tbl[0])
    if tbl[0]>0:
        sql_str=f'select ifnull(uti, "not_found") from entity WHERE short_name like "{name}"'
        cursor.execute(sql_str)
        uti = cursor.fetchone()[0] 
        return uti
    else:
        return "not_found"
    
def get_EntityName_by_UTI(cursor, uti):
    
    sql_str=f'select count(1) from entity WHERE uti like "{uti}" '
    cursor.execute(sql_str)
    tbl = cursor.fetchone()
    #print(tbl[0])
    if tbl[0]>0:
        sql_str=f'select ifnull(short_name, "not_found") from entity WHERE uti like "{uti}"'
        cursor.execute(sql_str)
        short_name = cursor.fetchone()[0] 
        return short_name
    else:
        return "not_found"    
        
def get_EntityUTI_by_isin(cursor, isin):
    
    sql_str=f'select count(1) from bonds_static WHERE isin like "{isin}" '
    cursor.execute(sql_str)
    tbl = cursor.fetchone()
    #print(tbl[0])
    if tbl[0]>0:
        sql_str=f'select ifnull(issuer_uti, "not_found") from bonds_static WHERE isin like "{isin}"'
        cursor.execute(sql_str)
        uti = cursor.fetchone()[0] 
        return uti
    else:
        return "not_found"


def get_bond_amortization(cursor, isin):
    amo={}
    d = datetime.datetime.today()
    today_str=d.strftime("%Y%m%d")    
    
    sql_str=f'select ifnull(min(date), "na") from bonds_schedule where isin="{isin}" and date>"{today_str}" and nominal_value>0 and date<>(select max(date) from bonds_schedule where isin="{isin}" and nominal_value>0) '
    cursor.execute(sql_str)
    amo_date = cursor.fetchone()[0]    
    
    if amo_date=="na":
        return {"date":"", "value":0}
    
    sql_str=f'select ifnull(nominal_value, 0) from bonds_schedule where isin="{isin}" and date="{amo_date}"'
    cursor.execute(sql_str)
    amo_value = cursor.fetchone()[0]
    
    return {"date":amo_date, "value":amo_value}

def get_bond_issuer(cursor, isin):
    #cursor = connection.cursor()
    
    sql_str=f'select ifnull(issuer_uti,"") from bonds_static WHERE ISIN like "{isin}"'
    cursor.execute(sql_str)
    sql_res = cursor.fetchone()
    uti=sql_res[0]
    results={'issuer_uti':uti}
    
    sql_str=f'select ifnull(short_name,"") from entity where uti like "{uti}"'
    cursor.execute(sql_str)
    sql_res = cursor.fetchone()
    if sql_res is not None:
        results['issuer_short_name']=sql_res[0]
    else:
        results['issuer_short_name']=""
       
    return results

def update_fcy_rates():
    req_str='https://iss.moex.com/iss/statistics/engines/currency/markets/fixing.json'
    j=requests.get(req_str).json()
    for b in zip(j['history']['data']):    
        if b[0][1]=='USDFIXME':
            cross_rates['USD']=float(b[0][2])
            print(cross_rates)    
    
def get_bond_credit_rating(cursor, isin):
    #cursor = connection.cursor()
    
    sql_str=f'select rating, ifnull(percent_type,"fixed") as type_ from bonds_static WHERE ISIN like "{isin}"'
    cursor.execute(sql_str)
    sql_res = cursor.fetchone()    
    results={'rating':sql_res[0], 'type': sql_res[1]}
    
    sql_str=f'select distinct ifnull(instrument_currency,"RUB") as currency from trading_instruments where isin like "{isin}"'
    cursor.execute(sql_str)
    sql_res = cursor.fetchone()
    results['currency']=sql_res[0]
       
    return results

def get_bond_type_by_rating(cursor, isin):
    r=get_bond_credit_rating(cursor, isin)
    rating=r['rating']
    type_=r['type']    
    currency=r['currency']
    
    num=ratings.get(rating, -1)
    
    if num == 27:
        return 'Gov'
    elif num<=26 and num>16:
        if type_=='float':
            return 'Corp-fl'
        elif currency!="RUB":
            return f'Corp-{currency}'
        else:
            return 'Corp'
    elif num<=16 and num>=0:
        if type_=='float':
            return 'VDO-fl'    
        else:
            return 'VDO'
    elif num==-1:
        return 'wrong rating!'
    
    return 'none'

def get_instrument_type_extended(cursor, isin):
    instr_type=get_instrument_type(cursor, isin)
    
    ret_string=f'{instr_type}'
    if instr_type != 'bond':
        return ret_string
    else:
        r=get_bond_credit_rating(cursor, isin)
        rating=r['rating']
        type_=r['type']    
        currency=r['currency']    
        num=ratings.get(rating, -1)    
        if num == 27:
            bond_credit_type='Gov'
            return f'{instr_type}/{bond_credit_type}/{type_}/{currency}'
        elif num<=26 and num>16:
            bond_credit_type='Corp'
            return f'{instr_type}/{bond_credit_type}/{type_}/{currency}'
        elif num<=16 and num>=0:
            bond_credit_type='VDO'
            return f'{instr_type}/{bond_credit_type}/{type_}/{currency}'            
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

def get_instrument_type(cursor, isin):
    sql_str=f'select instrument_type from trading_instruments where isin like "{isin}"  '
    cursor.execute(sql_str)
    val = cursor.fetchone()
    if val is None:
        return 0
    return val[0]    


def get_bond_info_moex(isin):
    secid=""
    shortname=""
    inn=""
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
        if f=="emitent_inn":
            inn=b         
                
    req_str='https://iss.moex.com/iss/engines/stock/markets/bonds/securities/'+secid+'.json?marketprice_board=1'
    nkd=0
    nominal=0
    last_price=0
    fixed_coupon=0
    settle_date=0
    j=requests.get(req_str).json()  #'https://iss.moex.com/iss/engines/stock/markets/bonds/securities/RU000A106Z38.json?marketprice_board=1'
    for f, b in zip(j['securities']['columns'], j['securities']['data'][0]):
        if f=="ACCRUEDINT":
            nkd=b
        if f=="FACEVALUE":
            nominal=b        
        if f=="COUPONPERCENT":
            if b is not None:
                fixed_coupon=b
        if f=="SETTLEDATE":
            if b is not None:
                settle_date=b                
    
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
    bond_info["emitent_inn"]=inn
    bond_info["fixed_coupon"]=fixed_coupon
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
    bond_info["settle_date"]=settle_date

    return bond_info


def get_equity_info_moex(isin):
    secid=""
    shortname=""
    inn=""
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
        if f=="emitent_inn":
            inn=b
                
    req_str='https://iss.moex.com/iss/engines/stock/markets/shares/securities/'+secid+'.json?marketprice_board=1'
    nkd=0
    nominal=0
    last_price=0
    fixed_coupon=0
    j=requests.get(req_str).json()  #'https://iss.moex.com/iss/engines/stock/markets/bonds/securities/RU000A106Z38.json?marketprice_board=1'
    
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
                            
    eq_info={}
    eq_info["isin"]=isin
    eq_info["secid"]=secid
    eq_info["shortname"]=shortname
    eq_info["last_price"]=last_price
    eq_info["emitent_inn"]=inn
    full_price=last_price
    eq_info["full_price"]=full_price

    return eq_info

def get_cash_info(isin):
    secid=""
    shortname=""
    inn=""    
                            
    cash_info={}
    cash_info["isin"]=isin
    cash_info["secid"]=isin
    cash_info["shortname"]="cash"
    cash_info["last_price"]=1
    cash_info["emitent_inn"]="my cash"
    full_price=1
    cash_info["full_price"]=1

    return cash_info


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


def calc_portfolio_value(cursor):
    total_val=0
    d = datetime.datetime.today()
    today_str=d.strftime("%Y%m%d")
    
    sql_str=f'SELECT isin, qty, portfolio_id FROM portfolio bp WHERE qty>0 '    
    cursor.execute(sql_str)
    tbl = cursor.fetchall()
    
    portfolios={}
        
    for item in tbl:
        sql_str=f'SELECT instrument_type FROM trading_instruments ti WHERE ti.isin="{item[0]}" '
        cursor.execute(sql_str)
        instrument_type=cursor.fetchone()[0]
        
        data={}
        if instrument_type=='bond':
            data=get_bond_info_moex(item[0])
        elif instrument_type=='equity':
            data=get_equity_info_moex(item[0])
        elif instrument_type=='etf':
            data=get_equity_info_moex(item[0])
        elif instrument_type=='cash':
            data=get_cash_info(item[0])            
            
        portfolios[item[2]]=portfolios.get(item[2],0)+data["full_price"]*item[1]
        total_val=total_val+data["full_price"]*item[1]
        #print(f'{item[2]};{item[1]};{data["full_price"]}')

        #save price to DB
        post_market_data(cursor, item[0], f'{instrument_type}_price', today_str, data["last_price"])
        if instrument_type=='bond':
            sql_str=f'SELECT isin, ifnull(percent_type, "fixed") FROM bonds_static WHERE isin="{item[0]}" '
            cursor.execute(sql_str)
            pct_type = cursor.fetchone()[1]
            if pct_type=="linker":
                post_market_data(cursor, item[0], "bond_nominal", today_str, data["nominal"])
        
        
    sql_str=f'select count(price) from market_data where id="my_portfolio" and date="{today_str}"'
    cursor.execute(sql_str)
    fetch_cnt = cursor.fetchone()[0]
    
    if fetch_cnt==0:
        sql_str=f'insert into market_data(id, date, price, price_nominal) values ("my_portfolio", "{today_str}", {total_val}, "RUB")'
        cursor.execute(sql_str)        
    else:
        sql_str=f'update market_data set price={total_val} where id="my_portfolio" and date="{today_str}"'
        cursor.execute(sql_str)
        
    portfolios["total"]=total_val
    return portfolios

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
                print(f'Inserted: isin={isin}, date={db_date_insert}, pct_value={coupon}, nominal_value={amortization}')
           
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


def post_market_data(cursor, isin, post_type, post_date, value):
    sql_str=f'select price from market_data where id="{isin}" and date="{post_date}" and id_type="{post_type}"'
    cursor.execute(sql_str)
    fetch_cnt = cursor.fetchone()
    
    price_nominal="pct"
    if post_type=="bond_nominal":
        price_nominal="RUB"
    if post_type=="etf_price":
        price_nominal="RUB"        
        
    if cursor.rowcount==-1:
        sql_str=f'insert into market_data(id, id_type, date, price, price_nominal) values ("{isin}", "{post_type}", "{post_date}", {value}, "{price_nominal}")'
        cursor.execute(sql_str)        
    else:
        sql_str=f'delete from market_data where id="{isin}" and date="{post_date}" and id_type="{post_type}"'
        cursor.execute(sql_str)   
        sql_str=f'insert into market_data(id, id_type, date, price, price_nominal) values ("{isin}", "{post_type}", "{post_date}", {value}, "{price_nominal}")'
        cursor.execute(sql_str)          
        

def get_credit_rating_for_uti(cursor, uti):
    rating='n/d'
    sql_str=f'select count(1) from credit_ratings where rating_owner_uti="{uti}" and date=(select max(date) from credit_ratings where rating_owner_uti="{uti}") '
    cursor.execute(sql_str)
    tbl = cursor.fetchone()
    if tbl[0]>0:
        sql_str=f'select rating from credit_ratings where rating_owner_uti="{uti}" and date=(select max(date) from credit_ratings where rating_owner_uti="{uti}") '
        cursor.execute(sql_str)
        tbl = cursor.fetchone()
        rating=tbl[0]
        return rating
    else:
        return 'n/d'
    
def get_credit_rating_for_isin(cursor, isin):
    rating='n/d'
    sql_str=f'select count(1) from bonds_static where isin="{isin}" '
    cursor.execute(sql_str)
    tbl = cursor.fetchone()
    if tbl[0]>0:
        sql_str=f'select rating from bonds_static where isin="{isin}" '
        cursor.execute(sql_str)
        tbl = cursor.fetchone()
        rating=tbl[0]
        return rating
    else:
        return 'n/d'        

def get_bond_rating(cursor, uti, isin):    
    # uti - uti of bond issuer, isin - isin of bond
    cr1=get_credit_rating_for_uti(cursor, uti)
    cr2=get_credit_rating_for_isin(cursor, isin)
    
    if cr1!='n/d':
        return cr1
    elif cr2!='n/d':
        return cr2
    else:
        return 'n/d'
       
    return results    

def get_bond_static_data(cursor, isin):    
    # isin - isin of bond
    
    bond_static={}
    
    sql_str=f'SELECT count(*) FROM bonds_static bs WHERE 1=1 and bs.isin="{isin}"'
    cursor.execute(sql_str)
    result=cursor.fetchone()    
    
    if result[0]>0:
        sql_str=f'SELECT isin, tiker, percent_type, percent_base, call_opt_date, put_opt_dates, issue_date, maturity_date FROM bonds_static bs WHERE 1=1 and bs.isin="{isin}"'
        cursor.execute(sql_str)
        result=cursor.fetchone()    
        bond_static["isin"]=result[0]
        bond_static["tiker"]=result[1]
        bond_static["percent_type"]=result[2]
        bond_static["percent_base"]=result[3]
        bond_static["call_opt_date"]=result[4]
        bond_static["put_opt_dates"]=result[5]
        bond_static["issue_date"]=result[6]
        bond_static["maturity_date"]=result[7]    
       
    return bond_static    

def days_between_dates(date_str1, date_str2):
    # первый аргумет - строка в формате YYYYMMDD
    # второй аргумет - дата
    # Шаг 1: Конвертируем строку в объект datetime
    date_format = '%Y%m%d'  # Указываем формат строки
    date_object1 = datetime.datetime.strptime(date_str1, date_format)
    date_object2 = datetime.datetime.strptime(date_str2, date_format)
    
    # Шаг 2: Вычисляем разницу в днях
    delta = date_object2 - date_object1
    return abs(delta.days)

def calc_bond_YTM(cursor, isin='RU000A1074Q1'):
    # calculate discounted margine for bonds with floating interest rate.
    # We replicate current coupoun for all rest payments of the bond and discount them like we calculate YTM.
    
    d = datetime.datetime.today()
    #cursor = connection.cursor()
    today_str=d.strftime("%Y%m%d")
    
    bond_data=get_bond_info_moex(isin)
    bond_full_price=bond_data["full_price"]
    #print(bond_full_price)
    bond_settle_date=bond_data["settle_date"]
    bond_settle_date2 = datetime.datetime.strptime(bond_settle_date, "%Y-%m-%d")
    bond_settle_date2 = bond_settle_date2.strftime("%Y%m%d")    
    
    sql_str=f'select count(*) from bonds_schedule where isin="{isin}" and date>="{bond_settle_date2}"'
    #print(sql_str)
    cursor.execute(sql_str)
    result=cursor.fetchone() 
    track_calc=[]
    
    if result[0]>0:
        sql_str=f'select date, pct_value, ifnull(nominal_value,0) as nominal from bonds_schedule where isin="{isin}" and date>="{bond_settle_date2}"'
        cursor.execute(sql_str)
        tbl=cursor.fetchall()          
        
        for i in tbl:
            date1=i[0]
            pct_value1=i[1]
            nominal1=i[2]
            days_between=days_between_dates(bond_settle_date2, date1)
            ti_365=days_between/365
            
            #print(f'{date1}, {pct_value1}, {nominal1}, {days_between}, {ti_365}')            
            track_calc.append({"pct_value":pct_value1, "nominal_value":nominal1, "days_between":days_between, "ti_365":ti_365})
        
    else:
        print(f'No payment schedule for bond {isin}')
        return None
    
    #print(track_calc)
    
    start = -1000
    end = 1000
    step = 10 #0.0001
    min_diff=10000000000
    YTM=0
    phase=0
    
    while step>=0.0001 and phase<1000:
        ytm_discount = start
        while ytm_discount <= end:
            #print(ytm_discount)
            
            summ=0
            for i in track_calc:
                discount1=(1+ytm_discount/100)
                if discount1==0:
                    break
                discount_pwr=discount1 ** i["ti_365"]
                summ=summ+i["pct_value"]/discount_pwr+i["nominal_value"]/discount_pwr
            
            diff=bond_full_price-summ
            #print(summ)        
            #print(diff)
            #print(diff.real)
            if abs(diff.real)<min_diff:
                YTM=ytm_discount
                #print(YTM)
                min_diff=abs(diff.real)
                
                start1=YTM-3*step
                end1=YTM+3*step
                step1=step/10
                
                #print(min_diff, YTM, start1, end1, step1)
                
                
            ytm_discount += step
            phase=phase+1
        start=start1
        end=end1
        step=step1
        min_diff=10000000000
    
    print(f'finally calculated YTM: {YTM}')
    print('fin')
    
    
def calc_bond_discounted_margine(cursor, isin='RU000A108777'):
    # calculate discounted margine for bonds with floating interest rate.
    # We replicate current coupoun for all rest payments of the bond and discount them like we calculate YTM.
    
    #Check if bond if float
    sql_str=f'select count(*) from bonds_static where isin="{isin}"'
    cursor.execute(sql_str)
    result=cursor.fetchone() 
    if result==0:
        print(f'Bond {isin} is not in bonds_static table')
        return None
    
    sql_str=f'select ifnull(percent_type, "fix") from bonds_static where isin="{isin}"'
    cursor.execute(sql_str)
    result=cursor.fetchone()     
    if result[0]!='float':
        print(f'Bond {isin} is not float')
        return None   
    
    sql_str=f'select count(*) from bonds_schedule where isin="{isin}" and pct_value>0'
    cursor.execute(sql_str)
    result=cursor.fetchone() 
    if result==0:
        print(f'Bond {isin} doesnt have caluclated payments in payments schedule!')
        return None
    
    
    d = datetime.datetime.today()
    #cursor = connection.cursor()
    today_str=d.strftime("%Y%m%d")
    
    bond_data=get_bond_info_moex(isin)
    bond_full_price=bond_data["full_price"]
    print(bond_full_price)
    bond_settle_date=bond_data["settle_date"]
    bond_settle_date2 = datetime.datetime.strptime(bond_settle_date, "%Y-%m-%d")
    bond_settle_date2 = bond_settle_date2.strftime("%Y%m%d")    
    
    sql_str=f'select count(*) from bonds_schedule where isin="{isin}" and date>="{bond_settle_date2}"'
    #print(sql_str)
    cursor.execute(sql_str)
    result=cursor.fetchone() 
    track_calc=[]
    
    if result[0]>0:
        sql_str=f'select date, pct_value, ifnull(nominal_value,0) as nominal from bonds_schedule where isin="{isin}" and date>="{bond_settle_date2}"'
        cursor.execute(sql_str)
        tbl=cursor.fetchall()          
        
        for i in tbl:
            date1=i[0]
            pct_value1=i[1]
            if not (pct_value1>0):
                sql_str=f'select pct_value from bonds_schedule where isin="{isin}" and date=(select max(date) from bonds_schedule where isin="{isin}" and pct_value>0) '
                cursor.execute(sql_str)
                pct=cursor.fetchone() 
                pct_value1=pct[0]
            nominal1=i[2]
            days_between=days_between_dates(bond_settle_date2, date1)
            ti_365=days_between/365
            
            print(f'{date1}, {pct_value1}, {nominal1}, {days_between}, {ti_365}')            
            track_calc.append({"pct_value":pct_value1, "nominal_value":nominal1, "days_between":days_between, "ti_365":ti_365})
        
    else:
        print(f'No payment schedule for bond {isin}')
        return None
    
    #print(track_calc)
    
    start = 0
    end = 1000
    step = 10 #0.0001
    min_diff=10000000000
    YTM=0
    phase=0
    
    while step>=0.0001 and phase<1000:
        ytm_discount = start
        while ytm_discount <= end:
            print(ytm_discount)
            if ytm_discount==-190:
                print(1)
            
            summ=0
            for i in track_calc:
                discount1=(1+ytm_discount/100)
                if discount1==0:
                    break
                discount_pwr=discount1 ** i["ti_365"]
                #if isinstance(discount_pwr, complex):
                    #discount_pwr=abs(discount_pwr)
                summ=summ+i["pct_value"]/discount_pwr+i["nominal_value"]/discount_pwr
            
            if isinstance(summ, complex):
                summ=summ.real
            diff=bond_full_price-summ
            print(summ)        
            print(diff)
            print(f'diff.real: {diff.real}')
            if abs(diff)<min_diff:
                YTM=ytm_discount
                print(YTM)
                min_diff=abs(diff.real)
                
                start1=YTM-3*step
                end1=YTM+3*step
                step1=step/10
                
                print(min_diff, YTM, start1, end1, step1)
                
                
            print(f'min diff: {min_diff}, YTM: {YTM}, {start1}, {end1}, {step1}')
            ytm_discount += step
            phase=phase+1
        start=start1
        end=end1
        step=step1
        min_diff=10000000000
    
    print(f'finally calculated Discounted Margine: {YTM}')
    print('fin')
    return YTM
    