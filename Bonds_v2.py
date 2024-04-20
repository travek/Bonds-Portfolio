import wx
from base_ui_bonds_portfolio import Bonds_portfolio
import sqlite3
import bonds_functions_db
import xlsxwriter
import datetime
from sortedcontainers import SortedDict
import pandas as pd
import plotly.subplots as ps
import plotly.graph_objs as go


class Bonds_UI(Bonds_portfolio):
    def __init__(self, db_connection):
        super(Bonds_UI, self).__init__(parent=None)
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
        
        sql_str='select short_name, isin from bond_portfolio where isin not in (select isin from bonds_schedule)'
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        for item in tbl:
            str=f'Bond {item[1]} with short name {item[0]} doesnt have payment schedule in bonds_schedule table in DB \n'
            self.m_textCtrl3.AppendText(str)
        
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")
        sql_str=f'select bs.isin, pct_value from bonds_schedule bs join (select isin, min(date) as md from bonds_schedule where date>="{today_str}" group by isin) as bs2 on bs.isin=bs2.isin and bs.date=bs2.md where pct_value is null or pct_value = 0'
        
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        for item in tbl:
            str=f'Bond {item[0]} doesnt have payment amount in bonds_schedule table in DB for current time period! \n'
            self.m_textCtrl3.AppendText(str)        
        
        self.m_textCtrl3.AppendText('Data checks completed!\n')
        
        

    def upload_portfolio_from_file2DB( self, event ):
        event.Skip()
        
        read_pos=open("bonds_portfolio.txt", 'r', encoding='utf-8').read().splitlines() 
        
        cursor = self.connection.cursor()
        sql_str=f'delete from bond_portfolio'
        cursor.execute(sql_str)
        self.connection.commit()        
        
        for line in read_pos:
            line.rstrip('\n').replace("\n", "")
            l1=line.split(';')
            if (len(l1))<2:
                continue
    
            #elems={"count":float(l1[1]), "moex_code":l1[2], "isin":l1[0]}
            
            sql_str=f'SELECT count(*) FROM bond_portfolio WHERE 1=1 and ISIN like "{l1[0]}"'
            cursor.execute(sql_str)
            cnt = cursor.fetchone()[0]
            if cnt==0:
                sql_str=f'insert into bond_portfolio values("{l1[0]}", {float(l1[1])}, "{l1[2]}")'
                cursor.execute(sql_str)
                print(f'Inserted: isin={l1[0]}, count={l1[1]}, short_name={l1[2]}')
            else:
                sql_str=f'delete from bond_portfolio where isin = "{l1[0]}"'
                cursor.execute(sql_str)
                print(f'Deleted: isin={l1[0]}, count={l1[1]}, short_name={l1[1]}')
                sql_str=f'insert into bond_portfolio values("{l1[0]}", {float(l1[1])}, "{l1[2]}")'
                cursor.execute(sql_str)
                print(f'Inserted: isin={l1[0]}, count={l1[1]}, short_name={l1[1]}')
            
                
            self.connection.commit()
        
        sql_str=f'SELECT count(*) FROM bond_portfolio'
        cursor.execute(sql_str)
        cnt = cursor.fetchone()[0]
        
        print(f'There are {cnt} posions in portfolio')
          

    def f_print_portfolio_excel( self, event ):
        event.Skip()

        #connection = sqlite3.connect('portfolio_database.db')
        cursor = self.connection.cursor()
        
        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook('my_portfolio2.xlsx')
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
                
        sql_str=f'SELECT isin, qty, short_name FROM bond_portfolio WHERE 1=1 and qty>0 '
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        
        row = 1
        col = 0          
        for item in tbl:
            moex_data=bonds_functions_db.get_bond_info_moex(item[0])
            worksheet.write(row, col,     item[2])
            worksheet.write(row, col + 1, item[0])
            worksheet.write(row, col + 2, item[1])
            worksheet.write_datetime(row, col + 3, bonds_functions_db.get_bond_maturity(self.connection.cursor(), item[0]), date_format)
            worksheet.write_datetime(row, col + 4, bonds_functions_db.get_bond_nearest_coupon_date(self.connection.cursor(), item[0]), date_format)
            worksheet.write(row, col + 5, item[1]*bonds_functions_db.get_bond_nearest_coupon(self.connection.cursor(), item[0]))
            worksheet.write(row, col + 6, bonds_functions_db.get_current_bond_nominal(self.connection.cursor(), item[0]) )
            worksheet.write(row, col + 7, bonds_functions_db.get_bond_rating(self.connection.cursor(), item[0]) )        
            worksheet.write(row, col + 8, moex_data["yield"] )

            worksheet.write(row, col + 9, item[1]*moex_data["full_price"])
                        
            worksheet.write(row, col + 10, moex_data["duration"] )        
            worksheet.write(row, col + 11, moex_data["duration"]/365 )
            worksheet.write(row, col + 12, bonds_functions_db.get_bond_type_by_rating(self.connection.cursor(), item[0]) )
            worksheet.write(row, col + 13, moex_data["last_price"])
            worksheet.write(row, col + 14, moex_data["current_coupon"]/moex_data["last_price"])
            worksheet.write(row, col + 15, moex_data["coupon_period"])
            worksheet.write(row, col + 16, moex_data["issue_size"])
                            
            row += 1            
                
        # Write a total using a formula.
        workbook.close()
        
        #connection.close()
        
    def f_load_bond_from_file( self, event ):        
        fd = wx.FileDialog(self, message="Choose a file", style=wx.FD_OPEN|wx.FD_FILE_MUST_EXIST, wildcard="Text files (*.txt)|*.txt", defaultDir="d:\\Alexey\\Python Projects\\Bonds CF\\Data\\", defaultFile="example.txt")
        if fd.ShowModal() == wx.ID_OK:
            path = fd.GetPath()
        fd.Destroy()
        
        bonds_functions_db.read_bond_from_txt(self.connection.cursor(), path)
        self.connection.commit()
        self.m_textCtrl3.AppendText(f'Bond schedule uploaded into DB! \n')              
        
        return 0
    
    
    def graph_cashflows( self, calc_type=1 ):
        #Посчитать поток в каждом месяце и вывести график гисторграммой
        
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")        
        cursor = self.connection.cursor()
        
        sql_str=f'select max(date) from bond_portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0'
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
                sql_str=f'select sum(pct_value*bp.qty + nominal_value*bp.qty) from bond_portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}"'
            if calc_type==2:
                sql_str=f'select sum(pct_value*bp.qty) from bond_portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}"'
            if calc_type==3:
                sql_str=f'select sum(nominal_value*bp.qty) from bond_portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0 and date>="{start_date_str}" and date<="{end_date_str}" and date>="{today_str}"'
                                
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
    
    def graph_cashflows2( self ):
        #Посчитать поток в каждом месяце и вывести график гисторграммой
        
        d = datetime.datetime.today()
        today_str=d.strftime("%Y%m%d")        
        cursor = self.connection.cursor()
        
        sql_str=f'select max(date) from bond_portfolio bp join bonds_schedule bs on bp.isin=bs.isin where bp.qty>0'
        cursor.execute(sql_str)
        max_pay_date = cursor.fetchone()[0]
        d = datetime.datetime.strptime(max_pay_date, '%Y%m%d')
        
        start_date=datetime.datetime.today().replace(day=1).replace(hour=0, minute=0, second=0, microsecond=0)
        fist_day_next_month=(start_date + datetime.timedelta(days=33)).replace(day=1)
        end_date=(fist_day_next_month- datetime.timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        
        cf_all = SortedDict()
        xs_dates=set()
        
        sql_str=f'select distinct(substr(date,1,6)) from bonds_schedule bs join bond_portfolio bp on bp.isin=bs.isin where bp.qty>0 and bs.nominal_value>0 and date>="{today_str}"'
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        
        for item in tbl:
            date_=item[0]+'01'
            d = datetime.datetime.strptime(date_, '%Y%m%d')            
            xs_dates.add(d)
            
        xl_dates=list(xs_dates)
        xl_dates.sort()
        
        fig1 = go.Figure()
        
        sql_str=f'select bp.isin, bp.short_name, bp.qty from bond_portfolio bp where bp.qty>0'
        cursor.execute(sql_str)
        tbl1 = cursor.fetchall()        
        for item in tbl1:
            
            yl_values=[]
            
            for k in range(0, len(xl_dates)):
                d_str=xl_dates[k].strftime("%Y%m")
                sql_str=f'select bs.nominal_value from bonds_schedule bs join bond_portfolio bp on bp.isin=bs.isin where bp.qty>0 and bs.nominal_value>0 and bs.isin="{item[0]}" and substr(date,1,6)="{d_str}"'
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
        bond_portfolio_value=bonds_functions_db.calc_bond_portfolio_value(cursor)
        self.connection.commit()
        self.m_textCtrl3.AppendText(f'Completed! \n') 
        self.m_textCtrl3.AppendText(f'Bond portfolio value: {bond_portfolio_value:,.2f}\n')          
    
    def portfolio_export2CVS(self, event):
        file = open("portfolio_exportDB.txt", "w")
        cursor = self.connection.cursor()
        
        sql_str=f'select isin, qty, short_name from bond_portfolio where qty>0'
        cursor.execute(sql_str)
        tbl = cursor.fetchall()
        
        for item in tbl:
            q=str(item[1])
            str_=f'{item[0]},{q},{item[2]}\n'
            file.write(str_)
        file.close()
            
        

if __name__ == "__main__":
    connection = sqlite3.connect('portfolio_database.db')    
    
    app = wx.App(False)
    frame = Bonds_UI(db_connection=connection)
    frame.Show()
    app.MainLoop()
    
    