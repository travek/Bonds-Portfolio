

import wx
import wx.xrc

ID_EXIT = 1000
ID_PORTFOLIO_MARKET_VALUE = 1001
ID_INSERT_POSITION = 1002
ID_UPDATE_POSITION = 1003
ID_UPLOAD_FROM_FILE = 1004
ID_LOAD_BOND_FROM_FILE = 1005

###########################################################################
## Class Bonds_portfolio
###########################################################################

class Bonds_portfolio ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Bonds analysis", pos = wx.DefaultPosition, size = wx.Size( 857,342 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        self.m_menubar1 = wx.MenuBar( 0 )
        self.m_menu1 = wx.Menu()
        self.m_menu2 = wx.Menu()
        self.m_menuItem3 = wx.MenuItem( self.m_menu2, wx.ID_ANY, u"interest and amortization", wx.EmptyString, wx.ITEM_NORMAL )
        self.m_menu2.Append( self.m_menuItem3 )

        self.m_menuItem4 = wx.MenuItem( self.m_menu2, wx.ID_ANY, u"only interest", wx.EmptyString, wx.ITEM_NORMAL )
        self.m_menu2.Append( self.m_menuItem4 )

        self.m_menuItem5 = wx.MenuItem( self.m_menu2, wx.ID_ANY, u"only amortization", wx.EmptyString, wx.ITEM_NORMAL )
        self.m_menu2.Append( self.m_menuItem5 )

        self.m_menu1.AppendSubMenu( self.m_menu2, u"Calc cash-flows" )

        self.Run_data_checks = wx.MenuItem( self.m_menu1, wx.ID_ANY, u"Run data checks", wx.EmptyString, wx.ITEM_NORMAL )
        self.m_menu1.Append( self.Run_data_checks )

        self.m_menu1.AppendSeparator()

        self.exit = wx.MenuItem( self.m_menu1, ID_EXIT, u"Exit", wx.EmptyString, wx.ITEM_NORMAL )
        self.m_menu1.Append( self.exit )

        self.m_menubar1.Append( self.m_menu1, u"Main" )

        self.portfolio = wx.Menu()
        self.portfolioMarketValue = wx.MenuItem( self.portfolio, ID_PORTFOLIO_MARKET_VALUE, u"Portfolio market value", wx.EmptyString, wx.ITEM_NORMAL )
        self.portfolio.Append( self.portfolioMarketValue )

        self.m_PortfolioValueGraph = wx.MenuItem( self.portfolio, wx.ID_ANY, u"Portfolio value graph", wx.EmptyString, wx.ITEM_NORMAL )
        self.portfolio.Append( self.m_PortfolioValueGraph )

        self.portfolio.AppendSeparator()

        self.insertPosition = wx.MenuItem( self.portfolio, ID_INSERT_POSITION, u"Insert position", wx.EmptyString, wx.ITEM_NORMAL )
        self.portfolio.Append( self.insertPosition )

        self.updatePosition = wx.MenuItem( self.portfolio, ID_UPDATE_POSITION, u"Update position", wx.EmptyString, wx.ITEM_NORMAL )
        self.portfolio.Append( self.updatePosition )

        self.portfolio.AppendSeparator()

        self.m_export2Excel = wx.MenuItem( self.portfolio, wx.ID_ANY, u"Export to Excel", wx.EmptyString, wx.ITEM_NORMAL )
        self.portfolio.Append( self.m_export2Excel )

        self.m_ExportCSV = wx.MenuItem( self.portfolio, wx.ID_ANY, u"Expot to .cvs", wx.EmptyString, wx.ITEM_NORMAL )
        self.portfolio.Append( self.m_ExportCSV )

        self.m_ExportPayment = wx.MenuItem( self.portfolio, wx.ID_ANY, u"Export cash-flow to Excel", wx.EmptyString, wx.ITEM_NORMAL )
        self.portfolio.Append( self.m_ExportPayment )

        self.portfolio.AppendSeparator()

        self.uploadFromFile = wx.MenuItem( self.portfolio, ID_UPLOAD_FROM_FILE, u"Upload from file", wx.EmptyString, wx.ITEM_NORMAL )
        self.portfolio.Append( self.uploadFromFile )

        self.m_menubar1.Append( self.portfolio, u"Portfolio" )

        self.staticData = wx.Menu()
        self.loadBondFromFile = wx.MenuItem( self.staticData, ID_LOAD_BOND_FROM_FILE, u"Load bond from file", wx.EmptyString, wx.ITEM_NORMAL )
        self.staticData.Append( self.loadBondFromFile )

        self.m_menubar1.Append( self.staticData, u"Static data" )

        self.SetMenuBar( self.m_menubar1 )

        fgSizer1 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer1.AddGrowableCol( 0 )
        fgSizer1.AddGrowableRow( 0 )
        fgSizer1.SetFlexibleDirection( wx.BOTH )
        fgSizer1.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_ALL )

        self.m_textCtrl3 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_BESTWRAP|wx.TE_CHARWRAP|wx.TE_MULTILINE|wx.TE_WORDWRAP )
        fgSizer1.Add( self.m_textCtrl3, 0, wx.ALL|wx.EXPAND, 5 )


        self.SetSizer( fgSizer1 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.Bind( wx.EVT_MENU, self.calc_cashflows1, id = self.m_menuItem3.GetId() )
        self.Bind( wx.EVT_MENU, self.calc_cashflows2, id = self.m_menuItem4.GetId() )
        self.Bind( wx.EVT_MENU, self.calc_cashflows3, id = self.m_menuItem5.GetId() )
        self.Bind( wx.EVT_MENU, self.run_data_checks, id = self.Run_data_checks.GetId() )
        self.Bind( wx.EVT_MENU, self.Exit_app, id = self.exit.GetId() )
        self.Bind( wx.EVT_MENU, self.w_calc_bond_portfolio_value, id = self.portfolioMarketValue.GetId() )
        self.Bind( wx.EVT_MENU, self.f_portfolio_value_graph, id = self.m_PortfolioValueGraph.GetId() )
        self.Bind( wx.EVT_MENU, self.f_add_to_portfolio_selected, id = self.insertPosition.GetId() )
        self.Bind( wx.EVT_MENU, self.f_update_portfolio_selected, id = self.updatePosition.GetId() )
        self.Bind( wx.EVT_MENU, self.f_print_portfolio_excel, id = self.m_export2Excel.GetId() )
        self.Bind( wx.EVT_MENU, self.portfolio_export2CVS, id = self.m_ExportCSV.GetId() )
        self.Bind( wx.EVT_MENU, self.f_export_cash_flow_Excel, id = self.m_ExportPayment.GetId() )
        self.Bind( wx.EVT_MENU, self.upload_portfolio_from_file2DB, id = self.uploadFromFile.GetId() )
        self.Bind( wx.EVT_MENU, self.f_load_bond_from_file, id = self.loadBondFromFile.GetId() )

    def __del__( self ):
        pass


    # Virtual event handlers, override them in your derived class
    def calc_cashflows1( self, event ):
        event.Skip()

    def calc_cashflows2( self, event ):
        event.Skip()

    def calc_cashflows3( self, event ):
        event.Skip()

    def run_data_checks( self, event ):
        event.Skip()

    def Exit_app( self, event ):
        event.Skip()

    def w_calc_bond_portfolio_value( self, event ):
        event.Skip()

    def f_portfolio_value_graph( self, event ):
        event.Skip()

    def f_add_to_portfolio_selected( self, event ):
        event.Skip()

    def f_update_portfolio_selected( self, event ):
        event.Skip()

    def f_print_portfolio_excel( self, event ):
        event.Skip()

    def portfolio_export2CVS( self, event ):
        event.Skip()

    def f_export_cash_flow_Excel( self, event ):
        event.Skip()

    def upload_portfolio_from_file2DB( self, event ):
        event.Skip()

    def f_load_bond_from_file( self, event ):
        event.Skip()


###########################################################################
## Class Portfolio_add_bond
###########################################################################

class Portfolio_add_bond ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Add bond to portfolio", pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        fgSizer2 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer2.AddGrowableCol( 0 )
        fgSizer2.AddGrowableCol( 1 )
        fgSizer2.SetFlexibleDirection( wx.BOTH )
        fgSizer2.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_ALL )

        self.m_staticText1 = wx.StaticText( self, wx.ID_ANY, u"ISIN", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
        self.m_staticText1.Wrap( -1 )

        fgSizer2.Add( self.m_staticText1, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_ISIN_input = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer2.Add( self.m_ISIN_input, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText2 = wx.StaticText( self, wx.ID_ANY, u"Quantity in portfolio", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
        self.m_staticText2.Wrap( -1 )

        fgSizer2.Add( self.m_staticText2, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_quantity_input = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer2.Add( self.m_quantity_input, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText3 = wx.StaticText( self, wx.ID_ANY, u"Tiker", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
        self.m_staticText3.Wrap( -1 )

        fgSizer2.Add( self.m_staticText3, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_tiker_input = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer2.Add( self.m_tiker_input, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText4 = wx.StaticText( self, wx.ID_ANY, u"Portfolio name", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
        self.m_staticText4.Wrap( -1 )

        fgSizer2.Add( self.m_staticText4, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_portfolio_id = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer2.Add( self.m_portfolio_id, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText5 = wx.StaticText( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText5.Wrap( -1 )

        fgSizer2.Add( self.m_staticText5, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText6 = wx.StaticText( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText6.Wrap( -1 )

        fgSizer2.Add( self.m_staticText6, 0, wx.ALL, 5 )

        self.m_Cancel_button = wx.Button( self, wx.ID_ANY, u"Cancel", wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer2.Add( self.m_Cancel_button, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_add_to_portfolio = wx.Button( self, wx.ID_ANY, u"Add to portfolio", wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer2.Add( self.m_add_to_portfolio, 0, wx.ALL|wx.EXPAND, 5 )


        self.SetSizer( fgSizer2 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.m_Cancel_button.Bind( wx.EVT_BUTTON, self.f_Cancel_button_pushed )
        self.m_add_to_portfolio.Bind( wx.EVT_BUTTON, self.f_add_to_portfolio )

    def __del__( self ):
        pass


    # Virtual event handlers, override them in your derived class
    def f_Cancel_button_pushed( self, event ):
        event.Skip()

    def f_add_to_portfolio( self, event ):
        event.Skip()


###########################################################################
## Class update_position
###########################################################################

class update_position ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 662,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        fgSizer3 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer3.AddGrowableCol( 0 )
        fgSizer3.AddGrowableCol( 1 )
        fgSizer3.AddGrowableRow( 1 )
        fgSizer3.SetFlexibleDirection( wx.BOTH )
        fgSizer3.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_ALL )

        self.m_textCtrl6 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer3.Add( self.m_textCtrl6, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText7 = wx.StaticText( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText7.Wrap( -1 )

        fgSizer3.Add( self.m_staticText7, 0, wx.ALL, 5 )

        m_listBox1Choices = []
        self.m_listBox1 = wx.ListBox( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_listBox1Choices, 0 )
        fgSizer3.Add( self.m_listBox1, 0, wx.ALL|wx.EXPAND, 5 )

        gSizer1 = wx.GridSizer( 5, 2, 0, 0 )

        self.m_staticText8 = wx.StaticText( self, wx.ID_ANY, u"ISIN", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText8.Wrap( -1 )

        gSizer1.Add( self.m_staticText8, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl9 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer1.Add( self.m_textCtrl9, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText9 = wx.StaticText( self, wx.ID_ANY, u"QTY", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText9.Wrap( -1 )

        gSizer1.Add( self.m_staticText9, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl10 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer1.Add( self.m_textCtrl10, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText10 = wx.StaticText( self, wx.ID_ANY, u"Tiker", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText10.Wrap( -1 )

        gSizer1.Add( self.m_staticText10, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl11 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer1.Add( self.m_textCtrl11, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText11 = wx.StaticText( self, wx.ID_ANY, u"Portfolio_ID", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText11.Wrap( -1 )

        gSizer1.Add( self.m_staticText11, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl12 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer1.Add( self.m_textCtrl12, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_button3 = wx.Button( self, wx.ID_ANY, u"Update", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer1.Add( self.m_button3, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_button4 = wx.Button( self, wx.ID_ANY, u"CANCEL", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer1.Add( self.m_button4, 0, wx.ALL|wx.EXPAND, 5 )


        fgSizer3.Add( gSizer1, 1, wx.EXPAND, 5 )


        self.SetSizer( fgSizer3 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.m_textCtrl6.Bind( wx.EVT_TEXT, self.ISIN_char_entered )
        self.m_listBox1.Bind( wx.EVT_LISTBOX, self.f_lb_ISIN_selected )
        self.m_button3.Bind( wx.EVT_BUTTON, self.f_update_position )
        self.m_button4.Bind( wx.EVT_BUTTON, self.f_cancel_button )

    def __del__( self ):
        pass


    # Virtual event handlers, override them in your derived class
    def ISIN_char_entered( self, event ):
        event.Skip()

    def f_lb_ISIN_selected( self, event ):
        event.Skip()

    def f_update_position( self, event ):
        event.Skip()

    def f_cancel_button( self, event ):
        event.Skip()


