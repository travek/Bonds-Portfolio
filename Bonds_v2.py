import wx
from base_ui_bonds_portfolio import Bonds_portfolio
from base_ui_bonds_portfolio import Portfolio_add_bond
from base_ui_bonds_portfolio import update_position
from base_ui_bonds_portfolio import Add_Instrument
from base_ui_bonds_portfolio import Entity
from base_ui_bonds_portfolio import CreditRatings
import sqlite3
import bonds_functions_db
import xlsxwriter
import datetime
from sortedcontainers import SortedDict
import pandas as pd
import plotly.subplots as ps
import plotly.graph_objs as go


class CEntity(Entity):
    def __init__(self, db_connection):
        super(CEntity, self).__init__(parent=None)
        self.connection=db_connection        
                
    def f_add_entity(self, event):
        uti=self.m_textCtrl17.GetValue().strip()
        short_name=self.m_textCtrl18.GetValue().strip()
        short_name=short_name.translate({ord(i): None for i in '"'})
        full_name=self.m_textCtrl19.GetValue().strip()
        full_name=full_name.translate({ord(i): None for i in '"'})
        cursor = self.connection.cursor()
        
        sql_str=f'select count(*) from entity where UTI like "{uti}" '
        cursor.execute(sql_str)
        tbl = cursor.fetchone()
        
        if tbl[0]>0:
            print(f'Entity UTI={uti} already in DB')
        else:        
            if len(uti)>5 and len(short_name)>3:
                sql_str=f'insert into entity(UTI, short_name, long_name) values("{uti}", "{short_name}", "{full_name}" )'
                cursor.execute(sql_str)
                self.connection.commit()         
        
        self.Close()
        
    def f_cancel_entity(self, event):
        self.Close()
        
class CCreditRatings(CreditRatings):
    def __init__(self, db_connection):
        super(CCreditRatings, self).__init__(parent=None)
        self.connection=db_connection
        self.m_choice8.SetItems(["Create", "Update", "Delete"])
        self.m_choice8.SetSelection( 0 )        
        self.m_choice9.SetItems([ u"Gov", u"AAA", u"AAA-", u"AA+", u"AA", u"AA-", u"A+", u"A", u"A-", u"BBB+", u"BBB", u"BBB-", u"BB+", u"BB", u"BB-", u"B+", u"B", u"B-", u"CCC", u"CC", u"C", u"D" ])
        self.m_choice9.SetSelection( 0 )
        self.m_choice14.SetItems(["АО Эксперт РА", "АКРА (АО)", "ООО НРА", "ООО НКР"])
        
        
        sql_str=f'select short_name, uti from entity order by short_name '
        cursor = self.connection.cursor()
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        
        lb_lst=[]
        for res in tbl:
            lb_item=f'{res[0]} / {res[1]}'
            lb_lst.append(lb_item)
            
        if len(lb_lst)>0:
            self.m_choice13.SetItems(lb_lst) 
            
        
    def CreditRatings_onCancel( self, event ):
        event.Skip()   
        self.Close()
        
    def onAction_Selected( self, event ):
        action=self.m_choice8.GetString(self.m_choice8.GetCurrentSelection())
        
        if action=="Create":
            self.m_button13.Label="Create"
        elif action=="Update":
            self.m_button13.Label="Update"
            rating_owner=self.m_choice13.GetString(self.m_choice13.GetCurrentSelection())
            rating_owner_uti=rating_owner.split('/')[1].strip()
            
            sql_str=f'select count(1) from credit_ratings where rating_owner_uti like "{rating_owner_uti}" '
            cursor = self.connection.cursor()
            cursor.execute(sql_str)
            tbl = cursor.fetchone()
            if tbl[0]>0:
                selected=self.m_listBox2.GetStringSelection()
                cursor = self.connection.cursor()
                selected=selected.split('/')
                date_=selected[0].strip()
                rating_=selected[1].strip()
                rating_agency_=selected[2].strip()
                forecast_=selected[3].strip()
                                
                self.m_textCtrl37.SetValue(date_)
                self.m_choice9.SetStringSelection(rating_)
                self.m_choice14.SetStringSelection(rating_agency_)
                self.m_textCtrl40.SetValue(forecast_)                                                                

            
        elif action=="Delete":
            self.m_button13.Label="Delete"
                    
        
    def CreditRating_OnEntity( self, event ):
        rating_owner=self.m_choice13.GetString(self.m_choice13.GetCurrentSelection())
        rating_owner_uti=rating_owner.split('/')[1].strip()
        
        sql_str=f'select count(1) from credit_ratings where rating_owner_uti like "{rating_owner_uti}" '
        cursor = self.connection.cursor()
        cursor.execute(sql_str)
        tbl = cursor.fetchone()
        #print(tbl[0])
        if tbl[0]>0:
            sql_str=f'select date, rating, rating_issuer_uti, rating_forecast from credit_ratings where rating_owner_uti like "{rating_owner_uti}" order by date '
            cursor.execute(sql_str)
            tbl = cursor.fetchall()
            
            lb_lst=[]
            for res in tbl:
                short_name=bonds_functions_db.get_EntityName_by_UTI(cursor, res[2])
                lb_lst.append(f'{res[0]} / {res[1]} / {short_name} / {res[3]}\n')
            
            self.m_listBox2.InsertItems(lb_lst, 0)
                #lb_lst.append(lb_item)      
                
    def CreditRating_onAction( self, event ):       
        cursor = self.connection.cursor()
        action=self.m_choice8.GetString(self.m_choice8.GetCurrentSelection())
        
        if action=="Create":            
            rating_owner=self.m_choice13.GetString(self.m_choice13.GetCurrentSelection())
            rating_owner_uti=rating_owner.split('/')[1].strip()
            
            date=self.m_textCtrl37.GetValue().strip()
            rating=self.m_choice9.GetString(self.m_choice9.GetCurrentSelection())
            rating_agency=self.m_choice14.GetString(self.m_choice14.GetCurrentSelection())
            forecast=self.m_textCtrl40.GetValue().strip()
            rating_agency_uti=bonds_functions_db.get_EntityUTI_by_Name(self.connection.cursor(), rating_agency)
            
            sql_str=f'insert into credit_ratings(date, rating_owner_uti, rating, rating_issuer_uti, rating_forecast) values("{date}", "{rating_owner_uti}", "{rating}", "{rating_agency_uti}", "{forecast}" ) '
            cursor.execute(sql_str)
            self.connection.commit()              
            
        elif action=="Update":
            self.m_button13.Label="Update"
            
            rating_owner=self.m_choice13.GetString(self.m_choice13.GetCurrentSelection())
            rating_owner_uti=rating_owner.split('/')[1].strip()
            
            sql_str=f'select count(1) from credit_ratings where rating_owner_uti like "{rating_owner_uti}" '
            cursor = self.connection.cursor()
            cursor.execute(sql_str)
            tbl = cursor.fetchone()
            if tbl[0]>0:
                selected=self.m_listBox2.GetStringSelection()
                cursor = self.connection.cursor()
                selected=selected.split('/')
                date_=selected[0].strip()
                rating_=selected[1].strip()
                rating_agency_=selected[2].strip()
                forecast_=selected[3].strip()
                rating_agency_uti_=bonds_functions_db.get_EntityUTI_by_Name(self.connection.cursor(), rating_agency_)
                            
                new_date_=self.m_textCtrl37.GetValue().strip()
                new_rating_=self.m_choice9.GetString(self.m_choice9.GetCurrentSelection())
                new_rating_agency=self.m_choice14.GetString(self.m_choice14.GetCurrentSelection())
                new_rating_agency_uti=bonds_functions_db.get_EntityUTI_by_Name(self.connection.cursor(), new_rating_agency)                 
                new_forecast_=self.m_textCtrl40.GetValue().strip()                          
                
                sql_str_delete=f'delete from credit_ratings where date="{date_}" and rating_owner_uti="{rating_owner_uti}" and rating_issuer_uti="{rating_agency_uti_}"  '
                sql_str_insert=f'insert into credit_ratings(date, rating_owner_uti, rating, rating_issuer_uti, rating_forecast) values("{new_date_}", "{rating_owner_uti}", "{new_rating_}", "{new_rating_agency_uti}", "{new_forecast_}") '
                cursor.execute(sql_str_delete)
                self.connection.commit()  
                cursor.execute(sql_str_insert)
                self.connection.commit()  
            
                
        elif action=="Delete":
            self.m_button13.Label="Delete"   
            
        self.Close()        
        
    
class Add_to_portfolio(Portfolio_add_bond):
    def __init__(self, db_connection):
        super(Add_to_portfolio, self).__init__(parent=None)
        self.connection=db_connection
        
    def f_add_to_portfolio( self, event ):
        isin_=self.m_ISIN_input.GetValue().strip()
        qty_=float(self.m_quantity_input.GetValue().strip())
        tiker_=self.m_tiker_input.GetValue().strip()
        portfolio_id=self.m_choice5.GetString(self.m_choice5.GetCurrentSelection())
        print(f'{isin_}, {qty_}, {tiker_}, {portfolio_id}')
        cursor = self.connection.cursor()
        
        if len(isin_)>5 and qty_>0 and len(portfolio_id)>1: # len(tiker_)>0 and
            sql_str=f'select count(*) from portfolio where isin like "{isin_}%" and portfolio_id="{portfolio_id}"'
            cursor.execute(sql_str)
            tbl = cursor.fetchone()
            if tbl is not None and tbl[0]>0:                
                sql_str=f'insert into portfolio(isin, qty, short_name, portfolio_id) values("{isin_}", {qty_}, "{tiker_}", "{portfolio_id}")'
                cursor.execute(sql_str)
                self.connection.commit()        
        
        self.Close()
        
    def PortfolioAddBond_ISINEnter( self, event ):
        isin_=self.m_ISIN_input.GetValue().strip()
        cursor = self.connection.cursor()
        sql_str=f'select count(*) from bonds_static where isin like "{isin_}%" '
        cursor.execute(sql_str)
        tbl = cursor.fetchone()
        if tbl is not None and tbl[0]>0:
            sql_str=f'select tiker from bonds_static where isin like "{isin_}%" '
            cursor.execute(sql_str)
            tbl = cursor.fetchone()
            tiker_=tbl[0]
            self.m_tiker_input.SetValue(tiker_)
        event.Skip()
        

        
    def f_Cancel_button_pushed( self, event ):
        self.Close()
        
class my_Add_Instrument(Add_Instrument):
    def __init__(self, db_connection):
        super(my_Add_Instrument, self).__init__(parent=None)
        self.connection=db_connection
        
        sql_str=f'select short_name, uti from entity order by short_name'
        cursor = self.connection.cursor()
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        
        lb_lst=[]
        for res in tbl:
            lb_item=f'{res[0]} / {res[1]}'
            lb_lst.append(lb_item)
            
        if len(lb_lst)>0:
            self.m_choice3.SetItems(lb_lst)              
                

    def fAdd_instrument( self, event ):
        isin_=self.m_textCtrl11.GetValue().strip()
        tiker_=self.m_textCtrl12.GetValue().strip()
        inst_type=self.m_choice10.GetString(self.m_choice10.GetCurrentSelection())
        trading_place=self.m_choice11.GetString(self.m_choice11.GetCurrentSelection())
        coupon_type=self.m_choice1.GetString(self.m_choice1.GetCurrentSelection())
        coupon_base=self.m_choice4.GetString(self.m_choice4.GetCurrentSelection())
        coupon_addon=self.m_textCtrl18.GetValue().strip()
        issue_date=self.m_textCtrl13.GetValue().strip()
        maturity_date=self.m_textCtrl14.GetValue().strip()
        calls=self.m_textCtrl15.GetValue().strip()
        puts=self.m_textCtrl16.GetValue().strip()
        credit_rating=self.m_choice2.GetString(self.m_choice2.GetCurrentSelection())
        instrument_currency=self.m_choice12.GetString(self.m_choice12.GetCurrentSelection())
        if coupon_type=="fixed":
            percent_base=""
        else:
            percent_base=f'{coupon_base}+{coupon_addon}'
        
        le=""
        cs=self.m_choice3.GetCurrentSelection()
        if cs>=0:            
            le=self.m_choice3.GetString(cs)
        elif inst_type!="cash":
            wx.MessageBox("Legal entity not selected","Error",wx.OK)            
            return
        
        if inst_type=="cash":
            isin_=inst_type+"-"+instrument_currency+"-"+trading_place
            
        if len(le)>5:
            le_details=le.split('/')
            li_uti=le_details[1].strip()
        else:
            li_uti=""
        
        print(f'{isin_}, {tiker_}, {coupon_type}, {issue_date}, {maturity_date}, {calls}, {puts}, {credit_rating}, {li_uti} ')
        
        cursor = self.connection.cursor()
        sql_str=f'select count(*) from bonds_static where isin="{isin_}" '
        cursor.execute(sql_str)
        tbl = cursor.fetchone()
        sql_str=f'select count(*) from trading_instruments where isin="{isin_}" '
        cursor.execute(sql_str)
        tbl2 = cursor.fetchone()        
        
        if tbl[0]>0 or tbl2[0]>0:
            print(f'Bond {isin_} already in DB')
        elif len(li_uti)<5 and inst_type!="cash":
            print(f'Legal entity is NOT selected!!!')
        else:
            if inst_type=="bond":
                sql_str=f'insert into bonds_static(isin, rating, issue_date, percent_type, percent_base, maturity_date, call_opt_date, put_opt_dates, tiker, issuer_uti ) values ("{isin_}",  "{credit_rating}",  "{issue_date}", "{coupon_type}", "{percent_base}", "{maturity_date}", "{calls}", "{puts}", "{tiker_}", "{li_uti}")'
                print(sql_str)
                cursor.execute(sql_str)
                print(sql_str)
                sql_str=f'insert into trading_instruments(isin, instrument_type, trading_place, trading_code, instrument_currency) values ("{isin_}",  "{inst_type}",  "{trading_place}", "{tiker_}",  "{instrument_currency}" )'
                print(sql_str)
                cursor.execute(sql_str)            
                self.connection.commit() 
                self.Close()
                
            if inst_type=="equity":
                sql_str=f'insert into trading_instruments(isin, instrument_type, trading_place, trading_code, instrument_currency) values ("{isin_}",  "{inst_type}",  "{trading_place}", "{tiker_}",  "{instrument_currency}" )'
                print(sql_str)
                cursor.execute(sql_str)            
                self.connection.commit() 
                self.Close()  
                
            if inst_type=="etf":
                sql_str=f'insert into trading_instruments(isin, instrument_type, trading_place, trading_code, instrument_currency) values ("{isin_}",  "{inst_type}",  "{trading_place}", "{tiker_}",  "{instrument_currency}" )'
                print(sql_str)
                cursor.execute(sql_str)            
                self.connection.commit() 
                self.Close()     

            if inst_type=="cash":
                sql_str=f'insert into trading_instruments(isin, instrument_type, trading_place, trading_code, instrument_currency) values ("{isin_}",  "{inst_type}",  "{trading_place}", "cash",  "{instrument_currency}")'
                print(sql_str)
                cursor.execute(sql_str)            
                self.connection.commit() 
                self.Close()                 
                
        

    def Add_bond_cancel( self, event ):
        self.Close()
        
    
        
class Upd_Position(update_position):
    def __init__(self, db_connection):
        super(Upd_Position, self).__init__(parent=None)
        self.connection=db_connection    
        
        
    def ISIN_char_entered( self, event ):
        self.m_listBox1.Clear()
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")           
        
        cursor = self.connection.cursor()
        isin_template=self.m_textCtrl6.GetValue()
        
        sql_str=f'select bp.ISIN, bp.short_name, bp.portfolio_id from portfolio bp left join bonds_static bs on bs.isin=bp.isin left join trading_instruments ti on ti.isin=bp.isin where (bp.isin like "{isin_template}%" or bs.tiker like "{isin_template}%" or ti.trading_code like "{isin_template}%") and (bs.maturity_date>"{today_str}" or bs.maturity_date is null)'
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
                
        lb_lst=[]
        for res in tbl:
            lb_item=f'{res[0]} / {res[1]} / {res[2]}'
            lb_lst.append(lb_item)
        if len(lb_lst)>0:
            self.m_listBox1.InsertItems(lb_lst, 0)     
    
    def f_lb_ISIN_selected( self, event ):
        selected=self.m_listBox1.GetStringSelection()
        cursor = self.connection.cursor()
        selected=selected.split('/')
        isin=selected[0].strip()
        portfolio_id=selected[2].strip()
        
        
        sql_str=f'select ISIN, qty, short_name, portfolio_id from portfolio where isin like "{isin}" and portfolio_id like "{portfolio_id}" '
        cursor.execute(sql_str)
        tbl = cursor.fetchone()                  
        self.m_textCtrl9.SetValue(tbl[0])
        self.m_textCtrl10.SetValue(str(tbl[1]))
        self.m_textCtrl11.SetValue(tbl[2])
        self.m_textCtrl12.SetValue(tbl[3])
        
        
    def f_update_position( self, event ):
        cursor = self.connection.cursor()
        isin_=self.m_textCtrl9.GetValue()
        qty_=float(self.m_textCtrl10.GetValue())
        tiker_=self.m_textCtrl11.GetValue()
        portfolio_id=self.m_textCtrl12.GetValue()
        
        if len(isin_)>3 and qty_>=0 and len(tiker_)>0 and len(portfolio_id)>0:
            sql_str=f'update portfolio set qty={qty_}, short_name="{tiker_}" where isin="{isin_}" and portfolio_id="{portfolio_id}"'
            print(sql_str)
            cursor.execute(sql_str)
            self.connection.commit()              
            
        self.Close()
        
        
    def f_cancel_button( self, event ):
        self.Close()
    
    

class Portfolio_UI(Bonds_portfolio):
    def __init__(self, db_connection):
        super(Portfolio_UI, self).__init__(parent=None)
        self.connection=db_connection
        #self.connection = sqlite3.connect('portfolio_database.db')
        self.run_data_checks(None)
        
        
    def Exit_app( self, event ):
        event.Skip()   
        wx.Exit()
        self.connection.close() 
        
    def run_data_checks(self, event):
        cursor = self.connection.cursor()        
        #Check Bonds without payment schedule in DB
        self.m_textCtrl3.Clear()        
        
        sql_str='select short_name, isin from portfolio p where isin not in (select isin from bonds_schedule) and isin in (select isin from trading_instruments ti where ti.isin=p.isin and ti.instrument_type="bond")'
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        for item in tbl:
            str=f'Bond {item[1]} with short name {item[0]} doesnt have payment schedule in bonds_schedule table in DB \n'
            self.m_textCtrl3.AppendText(str)
        
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")
        #sql_str=f'select bs.isin, pct_value from bonds_schedule bs join (select isin, min(date) as md from bonds_schedule where date>="{today_str}" group by isin) as bs2 on bs.isin=bs2.isin and bs.date=bs2.md where pct_value is null or pct_value = 0'
        sql_str=f'select bp.isin, bs3.pct_value from portfolio bp join ( select bs.isin, pct_value from bonds_schedule bs join (select isin, min(date) as md from bonds_schedule where date>="{today_str}" group by isin) as bs2 on bs.isin=bs2.isin and bs.date=bs2.md where pct_value is null or pct_value = 0 ) bs3 on bp.isin=bs3.isin where bp.qty>0'
        
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        for item in tbl:
            str=f'Bond {item[0]} doesnt have payment amount in bonds_schedule table in DB for current time period! \n'
            self.m_textCtrl3.AppendText(str)        
            
        sql_str=f'select count(*) from bonds_static bs where issuer_uti is null and (select max(date) as maturity from bonds_schedule where isin=bs.isin)>"{today_str}" and (select qty from portfolio where isin=bs.isin)>0'
        cursor.execute(sql_str)
        res_inn=cursor.fetchone()[0]
        if res_inn>0:
            sql_str=f'select isin from bonds_static bs where issuer_uti is null and (select max(date) as maturity from bonds_schedule where isin=bs.isin)>"{today_str}" and (select qty from portfolio where isin=bs.isin)>0'
            cursor.execute(sql_str)
            tbl = cursor.fetchall()
            for item in tbl:
                bond_info=bonds_functions_db.get_bond_info_moex(item[0])
                emitent_inn=bond_info["emitent_inn"]
                sql_str=f'update bonds_static set issuer_uti="{emitent_inn}" where isin like "{item[0]}" '
                cursor.execute(sql_str)
                self.connection.commit() 
                
        sql_str=f'SELECT count(*) FROM portfolio bp WHERE qty>0 and isin not in (select isin from trading_instruments)'
        cursor.execute(sql_str)
        res=cursor.fetchone()[0]
        if res>0:
            sql_str=f'SELECT isin FROM portfolio bp WHERE qty>0 and isin not in (select isin from trading_instruments)'
            cursor.execute(sql_str)
            tbl = cursor.fetchall()
            for item in tbl:
                str=f'Instrument with ISIN {item[0]} is in Portfolio but not in table trading_instruments! \n'
                self.m_textCtrl3.AppendText(str)                  
        
        
        
        self.m_textCtrl3.AppendText('Data checks completed!\n')
        
        

    def upload_portfolio_from_file2DB( self, event ):
        event.Skip()
        
        if wx.MessageBox("Are You Sure you want to upload portfolio from file into DB?","Checking",wx.YES_NO) == wx.YES:
            read_pos=open("bonds_portfolio.txt", 'r', encoding='utf-8').read().splitlines() 
            
            cursor = self.connection.cursor()
            sql_str=f'delete from portfolio'
            cursor.execute(sql_str)
            self.connection.commit()        
            
            for line in read_pos:
                line.rstrip('\n').replace("\n", "")
                l1=line.split(';')
                if (len(l1))<2:
                    continue
        
                #elems={"count":float(l1[1]), "moex_code":l1[2], "isin":l1[0]}
                
                sql_str=f'SELECT count(*) FROM portfolio WHERE 1=1 and ISIN like "{l1[0]}"'
                cursor.execute(sql_str)
                cnt = cursor.fetchone()[0]
                if cnt==0:
                    sql_str=f'insert into portfolio values("{l1[0]}", {float(l1[1])}, "{l1[2]}")'
                    cursor.execute(sql_str)
                    print(f'Inserted: isin={l1[0]}, count={l1[1]}, short_name={l1[2]}')
                else:
                    sql_str=f'delete from portfolio where isin = "{l1[0]}"'
                    cursor.execute(sql_str)
                    print(f'Deleted: isin={l1[0]}, count={l1[1]}, short_name={l1[1]}')
                    sql_str=f'insert into portfolio values("{l1[0]}", {float(l1[1])}, "{l1[2]}")'
                    cursor.execute(sql_str)
                    print(f'Inserted: isin={l1[0]}, count={l1[1]}, short_name={l1[1]}')
                
                    
                self.connection.commit()
            
            sql_str=f'SELECT count(*) FROM portfolio'
            cursor.execute(sql_str)
            cnt = cursor.fetchone()[0]
            
            print(f'There are {cnt} posions in portfolio')
        

    def f_print_portfolio_excel( self, event ):
        event.Skip()

        #connection = sqlite3.connect('portfolio_database.db')
        cursor = self.connection.cursor()
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")           
        
        # Create a workbook and add a worksheet.
        f_name=f'Export_files\my_portfolio2-{today_str}.xlsx'
        workbook = xlsxwriter.Workbook(f_name)
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
        worksheet.write('M1', 'Instrument_type', bold)
        worksheet.write('N1', 'Last_Price', bold)
        worksheet.write('O1', 'Coupon_yield', bold)
        worksheet.write('P1', 'Coupon_period', bold)
        worksheet.write('Q1', 'Portfolio ID', bold)

        worksheet.write('R1', 'Issuer entity', bold) 
        
        worksheet.write('S1', 'amo date', bold)
        worksheet.write('T1', 'amo value', bold)
        
        worksheet.write('U1', 'Coupon_type', bold)
        worksheet.write('V1', 'Coupon_base', bold)        
        
        worksheet.write('W1', 'Call option dates', bold) 
        worksheet.write('X1', 'Put option dates', bold) 
        worksheet.write('Y1', 'Current coupon', bold) 
        
                
        #sql_str=f'SELECT bp.isin, qty, short_name, percent_type, percent_base, portfolio_id, call_opt_date, put_opt_dates FROM portfolio bp join bonds_static bs on bs.isin=bp.isin WHERE 1=1 and qty>0 and exists (select * from bonds_schedule bsc where bsc.isin=bp.isin and bsc.date>"{today_str}")'
        sql_str=f'select bp.isin, bp.qty, bp.portfolio_id FROM portfolio bp WHERE 1=1 and qty>0'
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        
        row = 1
        col = 0          
        for item in tbl:
            isin=item[0]
            instr_type=bonds_functions_db.get_instrument_type(self.connection.cursor(), item[0])
            if instr_type=='bond':
                bond=bonds_functions_db.get_bond_static_data(self.connection.cursor(),isin)
                
                #sql_str=f'SELECT bp.isin, qty, short_name, percent_type, percent_base, portfolio_id, call_opt_date, put_opt_dates FROM portfolio bp join bonds_static bs on bs.isin=bp.isin WHERE 1=1 and qty>0 and bp.isin="{isin}"'
                #cursor.execute(sql_str)
                #result=cursor.fetchone()
                
                moex_data=bonds_functions_db.get_bond_info_moex(isin)
                bond_rating=bonds_functions_db.get_bond_rating(self.connection.cursor(), bonds_functions_db.get_EntityUTI_by_isin(cursor, isin), isin)
                worksheet.write(row, col,     bond["tiker"])
                worksheet.write(row, col + 1, isin)
                worksheet.write(row, col + 2, item[1])
                worksheet.write_datetime(row, col + 3, bonds_functions_db.get_bond_maturity(self.connection.cursor(), isin), date_format)
                worksheet.write_datetime(row, col + 4, bonds_functions_db.get_bond_nearest_coupon_date(self.connection.cursor(), isin), date_format)
                worksheet.write(row, col + 5, item[1]*bonds_functions_db.get_bond_nearest_coupon(self.connection.cursor(), isin))
                #worksheet.write(row, col + 6, bonds_functions_db.get_current_bond_nominal(self.connection.cursor(), item[0]) )
                worksheet.write(row, col + 6, moex_data["nominal"])
                worksheet.write(row, col + 7, bond_rating)     
                worksheet.write(row, col + 8, moex_data["yield"] )
    
                worksheet.write(row, col + 9, item[1]*moex_data["full_price"])
                            
                worksheet.write(row, col + 10, moex_data["duration"] )        
                worksheet.write(row, col + 11, moex_data["duration"]/365 )
                worksheet.write(row, col + 12, bonds_functions_db.get_instrument_type_extended(self.connection.cursor(), isin) )
                worksheet.write(row, col + 13, moex_data["last_price"])
                try:
                    worksheet.write(row, col + 14, moex_data["current_coupon"]/moex_data["last_price"])
                except ZeroDivisionError:
                        worksheet.write(row, col + 14, moex_data["current_coupon"])
                worksheet.write(row, col + 15, moex_data["coupon_period"])
                worksheet.write(row, col + 16, item[2])
                
                issuer=bonds_functions_db.get_bond_issuer(self.connection.cursor(), isin)
                worksheet.write(row, col + 17, issuer["issuer_short_name"])              
    
                amo_value=bonds_functions_db.get_bond_amortization(self.connection.cursor(), isin)
                if amo_value.get("date")!='':
                    d = datetime.datetime.strptime(amo_value.get("date", None), '%Y%m%d')
                    worksheet.write_datetime(row, col + 18, d, date_format)            
                    worksheet.write(row, col + 19, amo_value.get("value"))
                
                worksheet.write(row, col + 20, bond["percent_type"])
                worksheet.write(row, col + 21, bond["percent_base"]) 
                
                worksheet.write(row, col + 22, bond["call_opt_date"])  
                worksheet.write(row, col + 23, bond["put_opt_dates"])
                
                worksheet.write(row, col + 24, moex_data["fixed_coupon"])
                
            if instr_type in ['equity', 'etf']:
                if instr_type=='equity':
                    moex_data=bonds_functions_db.get_equity_info_moex( isin)
                if instr_type=='etf':
                    moex_data=bonds_functions_db.get_equity_info_moex( isin)
                    
                #sql_str=f'SELECT bp.isin, qty, trading_code, 0, 0, portfolio_id, 0, 0 FROM portfolio bp join trading_instruments ti on ti.isin=bp.isin WHERE 1=1 and qty>0 and bp.isin="{isin}"'
                #cursor.execute(sql_str)
                #result=cursor.fetchone()    
                sql_str=f'SELECT trading_code FROM trading_instruments ti WHERE 1=1 and isin="{isin}"'
                cursor.execute(sql_str)
                result=cursor.fetchone()                
                
                worksheet.write(row, col,     result[0])
                worksheet.write(row, col + 1, isin)
                worksheet.write(row, col + 2, item[1])
                #worksheet.write_datetime(row, col + 3, bonds_functions_db.get_bond_maturity(self.connection.cursor(), isin), date_format)
                #worksheet.write_datetime(row, col + 4, bonds_functions_db.get_bond_nearest_coupon_date(self.connection.cursor(), isin), date_format)
                #worksheet.write(row, col + 5, item[1]*bonds_functions_db.get_bond_nearest_coupon(self.connection.cursor(), isin))
                #worksheet.write(row, col + 6, bonds_functions_db.get_current_bond_nominal(self.connection.cursor(), item[0]) )
                #worksheet.write(row, col + 6, moex_data["nominal"])
                #worksheet.write(row, col + 7, bond_rating)     
                #worksheet.write(row, col + 8, moex_data["yield"] )
    
                worksheet.write(row, col + 9, item[1]*moex_data["full_price"])
                            
                #worksheet.write(row, col + 10, moex_data["duration"] )        
                #worksheet.write(row, col + 11, moex_data["duration"]/365 )
                worksheet.write(row, col + 12, instr_type )
                worksheet.write(row, col + 13, moex_data["last_price"])
                #try:
                    #worksheet.write(row, col + 14, moex_data["current_coupon"]/moex_data["last_price"])
                #except ZeroDivisionError:
                        #worksheet.write(row, col + 14, moex_data["current_coupon"])
                #worksheet.write(row, col + 15, moex_data["coupon_period"])
                worksheet.write(row, col + 16, item[2])
                
                #issuer=bonds_functions_db.get_bond_issuer(self.connection.cursor(), isin)
                #worksheet.write(row, col + 17, issuer["issuer_short_name"])              
    
                #amo_value=bonds_functions_db.get_bond_amortization(self.connection.cursor(), isin)
                #if amo_value.get("date")!='':
                    #d = datetime.datetime.strptime(amo_value.get("date", None), '%Y%m%d')
                    #worksheet.write_datetime(row, col + 18, d, date_format)            
                    #worksheet.write(row, col + 19, amo_value.get("value"))
                
                #worksheet.write(row, col + 20, item[3])
                #worksheet.write(row, col + 21, item[4]) 
                
                #worksheet.write(row, col + 22, item[6])  
                #worksheet.write(row, col + 23, item[7])
                
                #worksheet.write(row, col + 24, moex_data["fixed_coupon"])                

            if instr_type in ['cash']:
                #if instr_type=='equity':
                    #moex_data=bonds_functions_db.get_equity_info_moex( isin)
                #if instr_type=='etf':
                    #moex_data=bonds_functions_db.get_equity_info_moex( isin)
                    
                #sql_str=f'SELECT bp.isin, qty, instrument_currency, 0, 0, portfolio_id, 0, 0 FROM portfolio bp join trading_instruments ti on ti.isin=bp.isin WHERE 1=1 and qty>0 and bp.isin="{isin}"'
                #sql_str=f'SELECT bp.isin, qty, instrument_currency, 0, 0, portfolio_id, 0, 0 FROM portfolio bp join trading_instruments ti on ti.isin=bp.isin WHERE 1=1 and qty>0 and bp.isin="{isin}"'
                #cursor.execute(sql_str)
                #result=cursor.fetchone()                                        
                
                worksheet.write(row, col,     item[0])
                worksheet.write(row, col + 1, isin)
                worksheet.write(row, col + 2, item[1])
                #worksheet.write_datetime(row, col + 3, bonds_functions_db.get_bond_maturity(self.connection.cursor(), isin), date_format)
                #worksheet.write_datetime(row, col + 4, bonds_functions_db.get_bond_nearest_coupon_date(self.connection.cursor(), isin), date_format)
                worksheet.write(row, col + 5, item[1])
                #worksheet.write(row, col + 6, bonds_functions_db.get_current_bond_nominal(self.connection.cursor(), item[0]) )
                #worksheet.write(row, col + 6, moex_data["nominal"])
                #worksheet.write(row, col + 7, bond_rating)     
                #worksheet.write(row, col + 8, moex_data["yield"] )
    
                worksheet.write(row, col + 9, item[1])
                            
                #worksheet.write(row, col + 10, moex_data["duration"] )        
                #worksheet.write(row, col + 11, moex_data["duration"]/365 )
                worksheet.write(row, col + 12, instr_type )
                worksheet.write(row, col + 13, 1)
                #try:
                    #worksheet.write(row, col + 14, moex_data["current_coupon"]/moex_data["last_price"])
                #except ZeroDivisionError:
                        #worksheet.write(row, col + 14, moex_data["current_coupon"])
                #worksheet.write(row, col + 15, moex_data["coupon_period"])
                worksheet.write(row, col + 16, item[2])
                
                #issuer=bonds_functions_db.get_bond_issuer(self.connection.cursor(), isin)
                #worksheet.write(row, col + 17, issuer["issuer_short_name"])              
    
                #amo_value=bonds_functions_db.get_bond_amortization(self.connection.cursor(), isin)
                #if amo_value.get("date")!='':
                    #d = datetime.datetime.strptime(amo_value.get("date", None), '%Y%m%d')
                    #worksheet.write_datetime(row, col + 18, d, date_format)            
                    #worksheet.write(row, col + 19, amo_value.get("value"))
                
                #worksheet.write(row, col + 20, item[3])
                #worksheet.write(row, col + 21, item[4]) 
                
                #worksheet.write(row, col + 22, item[6])  
                #worksheet.write(row, col + 23, item[7])
                
                #worksheet.write(row, col + 24, moex_data["fixed_coupon"])                                
            
                            
            row += 1            
                
        # Write a total using a formula.
        workbook.close()                
        self.m_textCtrl3.AppendText(f'Excel file exported! \n')
        
        #connection.close()
        
    def f_load_bond_from_file( self, event ):        
        self.m_textCtrl3.Clear()
        fd = wx.FileDialog(self, message="Choose a file", style=wx.FD_OPEN|wx.FD_FILE_MUST_EXIST, wildcard="Text files (*.txt)|*.txt", defaultDir="d:\\Alexey\\Python Projects\\Bonds CF\\Data\\", defaultFile="example.txt")
        if fd.ShowModal() == wx.ID_OK:
            path = fd.GetPath()
        fd.Destroy()
        cursor = self.connection.cursor()
        
        read_rates=open(path, 'r').read().splitlines() 
        isin=""
         
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
            
            if i==0 and len(str(l1[0]).strip())==12: 
                isin=str(l1[0]).strip()
                break
        
        #read_rates.close()
        
        sql_str=f'SELECT count(*) FROM bonds_static WHERE 1=1 and ISIN like "{isin}"'
        cursor.execute(sql_str)
        fetch_cnt = cursor.fetchone()[0]
        
        if fetch_cnt==0:
            self.m_textCtrl3.AppendText(f'Error! Add bond to the dictionary first, no static data for bond with ISIN {isin} \n')
            return -1   
        
        res=bonds_functions_db.read_bond_from_txt(self.connection.cursor(), path)
        if res==0:
            self.connection.commit()
            self.m_textCtrl3.AppendText(f'Bond schedule with ISIN={isin} uploaded into DB! \n')              
        elif res!=0:
            self.m_textCtrl3.AppendText(f'Error! return code={res}\n')
        
        return 0
    
    
    def graph_cashflows_old( self, calc_type=1 ):
        #Посчитать поток в каждом месяце и вывести график гисторграммой
        
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")        
        cursor = self.connection.cursor()
        
        sql_str=f'select max(date) from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0'
        cursor.execute(sql_str)
        max_pay_date = cursor.fetchone()[0]
        d = datetime.datetime.strptime(max_pay_date, '%Y%m%d')
        
        start_date=datetime.datetime.today().replace(day=1).replace(hour=0, minute=0, second=0, microsecond=0)
        fist_day_next_month=(start_date + datetime.timedelta(days=33)).replace(day=1)
        end_date=(fist_day_next_month- datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        
        cf_all = SortedDict()

        while end_date<=bonds_functions_db.calc_last_day_of_month(d):
            start_date_str=start_date.strftime("%Y%m%d")
            end_date_str=end_date.strftime("%Y%m%d")
            
            if calc_type==1:
                sql_str=f'select sum(pct_value*bp.qty + nominal_value*bp.qty) from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}"'
            if calc_type==2:
                sql_str=f'select sum(pct_value*bp.qty) from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}"'
            if calc_type==3:
                sql_str=f'select sum(nominal_value*bp.qty) from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}"'
                                
            cursor.execute(sql_str)
            month_cash_flow = cursor.fetchone()[0]
            
            if month_cash_flow is not None:
                cf_all[start_date]=month_cash_flow
            
            start_date=fist_day_next_month
            fist_day_next_month=(start_date + datetime.timedelta(days=33)).replace(day=1)
            end_date=(fist_day_next_month- datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)            
            
        p=pd.DataFrame.from_dict(cf_all.items())
        p.columns=['date', 'amount']
        p['date'] = p['date'].dt.strftime('%Y-%m')
        r=p.groupby('date')['amount'].sum().reset_index()
        
        fig1 = go.Figure()    
        fig1.add_trace(go.Bar(x=r['date'], y=r['amount'], text=round(r['amount']), texttemplate="%{y:,.0f}"))
        fig1.layout = dict(xaxis=dict(type="category"))  
        fig1.show()         
            
        
        return 0
    
    def graph_cashflows( self, calc_type=1 ):
        #Посчитать поток в каждом месяце и вывести график гисторграммой
        
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")        
        cursor = self.connection.cursor()
        
        sql_str=f'select max(date) from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0'
        cursor.execute(sql_str)
        max_pay_date = cursor.fetchone()[0]
        d = datetime.datetime.strptime(max_pay_date, '%Y%m%d')
        
        start_date=datetime.datetime.today().replace(day=1).replace(hour=0, minute=0, second=0, microsecond=0)
        fist_day_next_month=(start_date + datetime.timedelta(days=33)).replace(day=1)
        end_date=(fist_day_next_month- datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
                
        bonds_functions_db.update_fcy_rates()
                        
        cf_all = SortedDict()

        while end_date<=bonds_functions_db.calc_last_day_of_month(d):
            start_date_str=start_date.strftime("%Y%m%d")
            end_date_str=end_date.strftime("%Y%m%d")
            
            if calc_type==1:
                #sql_str=f'select sum(pct_value*bp.qty + nominal_value*bp.qty) from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}"'                
                sql_str=f'select sum(pct_value*bp.qty + nominal_value*bp.qty) as val, ifnull(bs.nominal_currency, "RUB") from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}" group by nominal_currency'
                
            if calc_type==2:
                sql_str=f'select sum(pct_value*bp.qty) from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}"'
            if calc_type==3:
                sql_str=f'select sum(nominal_value*bp.qty) from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}"'
                                
            cursor.execute(sql_str)
            results = cursor.fetchall()
            
            for row in results:
                if row[0] is not None:
                    cfy=row[1]
                    if cfy=='RUB':
                        if start_date not in cf_all:
                            cf_all[start_date]= row[0]
                        else:
                            cf_all[start_date]+= row[0]
                    else:
                        if start_date not in cf_all:
                            cf_all[start_date]= row[0]*bonds_functions_db.cross_rates['USD']
                        else:
                            cf_all[start_date]+= row[0]*bonds_functions_db.cross_rates['USD']
            
            start_date=fist_day_next_month
            fist_day_next_month=(start_date + datetime.timedelta(days=33)).replace(day=1)
            end_date=(fist_day_next_month- datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)            
            
        p=pd.DataFrame.from_dict(cf_all.items())
        p.columns=['date', 'amount']
        p['date'] = p['date'].dt.strftime('%Y-%m')
        r=p.groupby('date')['amount'].sum().reset_index()
        
        fig1 = go.Figure()    
        fig1.add_trace(go.Bar(x=r['date'], y=r['amount'], text=round(r['amount']), texttemplate="%{y:,.0f}"))
        fig1.layout = dict(xaxis=dict(type="category"))  
        fig1.show()         
            
        
        return 0
    
    
    def graph_cashflows2( self ):
        #Посчитать поток в каждом месяце и вывести график гисторграммой
        
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")        
        cursor = self.connection.cursor()
        
        sql_str=f'select max(date) from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0'
        cursor.execute(sql_str)
        max_pay_date = cursor.fetchone()[0]
        d = datetime.datetime.strptime(max_pay_date, '%Y%m%d')
        
        start_date=datetime.datetime.today().replace(day=1).replace(hour=0, minute=0, second=0, microsecond=0)
        fist_day_next_month=(start_date + datetime.timedelta(days=33)).replace(day=1)
        end_date=(fist_day_next_month- datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        
        cf_all = SortedDict()
        xs_dates=set()
        
        sql_str=f'select distinct(substr(date,1,6)) from bonds_schedule bs join portfolio bp on bp.isin=bs.isin where bp.qty>0 and bs.nominal_value>0 and date>="{today_str}"'
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        
        for item in tbl:
            date_=item[0]+'01'
            d = datetime.datetime.strptime(date_, '%Y%m%d')            
            xs_dates.add(d)
            
        xl_dates=list(xs_dates)
        xl_dates.sort()
        
        fig1 = go.Figure()
        
        sql_str=f'select bp.isin, bp.short_name, bp.qty from portfolio bp where bp.qty>0'
        cursor.execute(sql_str)
        tbl1 = cursor.fetchall()        
        for item in tbl1:
            
            yl_values=[]
            
            for k in range(0, len(xl_dates)):
                d_str=xl_dates[k].strftime("%Y%m")
                sql_str=f'select bs.nominal_value from bonds_schedule bs join portfolio bp on bp.isin=bs.isin where bp.qty>0 and bs.nominal_value>0 and bs.isin="{item[0]}" and substr(date,1,6)="{d_str}"'
                cursor.execute(sql_str)
                tbl2 = cursor.fetchone()
                
                if tbl2 is not None:
                    yl_values.append(tbl2[0]*item[2])    
                else:
                    yl_values.append(0)
            
            fig1.add_trace(go.Bar(x=xl_dates, y=yl_values, name=item[1]))
            
        fig1.update_layout(barmode='stack')
        #fig1.layout = dict(xaxis=dict(type="category"))
        fig1.show()          
    
        
        return 0
    

    def calc_cashflows1( self, event ):
        self.graph_cashflows( calc_type=1)
        
    def calc_cashflows2( self, event ):
        self.graph_cashflows( calc_type=2)
        
    def calc_cashflows3( self, event ):
        self.graph_cashflows2()
    
    def w_calc_bond_portfolio_value( self, event ):
        event.Skip()
        cursor = self.connection.cursor()
        self.m_textCtrl3.AppendText(f'Start portfolio fair value calculation...')
        bond_portfolio_value=bonds_functions_db.calc_portfolio_value(cursor)
        self.connection.commit()
        self.m_textCtrl3.AppendText(f'Completed! \n')
        for key in bond_portfolio_value:
            self.m_textCtrl3.AppendText(f'Bond portfolio "{key}" value: {bond_portfolio_value[key]:,.2f}\n')          
            
        
        sql_str='select date, price FROM market_data where id="my_portfolio" order by date asc'
        cursor.execute(sql_str)
        tbl1 = cursor.fetchall()        
        x_dates=[]
        y_port_value=[]
        for item in tbl1:
            x_dates.append(item[0])
            y_port_value.append(item[1])
        fig = go.Figure(data=go.Scatter(x = x_dates, y = y_port_value))
        fig.show()
            
                
    def portfolio_export2CVS(self, event):
        file = open("Export_files\portfolio_exportDB.txt", "w")
        cursor = self.connection.cursor()
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")        
        
        sql_str=f'select isin, qty, short_name, portfolio_id from portfolio bp where qty>0 and exists (select * from bonds_schedule bs where bs.isin=bp.isin and bs.date>"{today_str}")'
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        
        for item in tbl:
            q=str(round(item[1]))
            str_=f'{item[0]},{q},{item[2]}, {item[3]}\n'
            file.write(str_)
        file.close()
        self.m_textCtrl3.AppendText(f'CSV file exported! \n')
        
    def f_export_cash_flow_Excel( self, event ):
        event.Skip()

        #connection = sqlite3.connect('portfolio_database.db')
        cursor = self.connection.cursor()
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")        
        
        # Create a workbook and add a worksheet.
        f_name=f'Export_files\cash_flows-{today_str}.xlsx'
        workbook = xlsxwriter.Workbook(f_name)
        worksheet = workbook.add_worksheet()
        # Add a bold format to use to highlight cells.
        bold = workbook.add_format({'bold': True})
        date_format = workbook.add_format({'num_format': 'dd/mm/yy'})
        
        # Write some data headers.
        worksheet.write('A1', 'Ticker', bold)
        worksheet.write('B1', 'Isin', bold)
        worksheet.write('C1', 'Date', bold)
        worksheet.write('D1', 'Value', bold)
        worksheet.write('E1', 'currency', bold)
        worksheet.write('F1', 'type', bold)
        worksheet.write('G1', 'Portfolio_ID', bold)
        worksheet.write('H1', 'year-month', bold)
        worksheet.write('I1', 'bond_type', bold)
                
        sql_str=f'select * from (select short_name, bp.isin, date, qty*pct_value as value, ifnull(pct_currency, "RUB") as currency, "percentage", bp.portfolio_id from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{today_str}" union all select short_name, bp.isin, date, qty*nominal_value as value, ifnull(nominal_currency, "RUB") as currency, "nominal", bp.portfolio_id from portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{today_str}" and nominal_value>0 ) order by date '
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        
        row = 1
        col = 0          
        for item in tbl:
            worksheet.write(row, col,     item[0])
            worksheet.write(row, col + 1, item[1])
            worksheet.write(row, col + 2, item[2])
            if item[4]=="USD":
                worksheet.write(row, col + 3, item[3]*bonds_functions_db.cross_rates["USD"])
            else:
                worksheet.write(row, col + 3, item[3])            
            worksheet.write(row, col + 4, item[4])
            worksheet.write(row, col + 5, item[5])
            worksheet.write(row, col + 6, item[6])
            worksheet.write(row, col + 7, item[2][0:4]+"-"+item[2][4:6])
            worksheet.write(row, col + 8, bonds_functions_db.get_bond_type_by_rating(self.connection.cursor(), item[1]))
                            
            row += 1            
                
        # Write a total using a formula.
        workbook.close()                
        self.m_textCtrl3.AppendText(f'Excel file exported! \n')
    
    def f_add_to_portfolio_selected( self, event ):
        frame_add=Add_to_portfolio(db_connection=self.connection)
        frame_add.Show()
        
    def f_update_portfolio_selected( self, event ):
        frame_upd=Upd_Position(db_connection=self.connection)
        frame_upd.Show()        
        
    def f_add_bond_static_data( self, event ):
        frame_add_bond=my_Add_Instrument(db_connection=self.connection)
        frame_add_bond.Show()          
        
    def f_Add_Entity_Action( self, event ):
        frame_Entity=CEntity(db_connection=self.connection)
        frame_Entity.Show()   
        
    def OnCreditRatings_Manage( self, event ):
        frame_Credit_ratings=CCreditRatings(db_connection=self.connection)
        frame_Credit_ratings.Show()         
        

if __name__ == "__main__":
    connection = sqlite3.connect('portfolio_database.db')    
    
    app = wx.App(False)
    frame = Portfolio_UI(db_connection=connection)
    frame.Show()
    app.MainLoop()
    
    