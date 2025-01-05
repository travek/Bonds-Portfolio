import wx
import wx.xrc
import wx.grid

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
        self.AddBondStaticData = wx.MenuItem( self.staticData, wx.ID_ANY, u"Add instrument", wx.EmptyString, wx.ITEM_NORMAL )
        self.staticData.Append( self.AddBondStaticData )

        self.staticData.AppendSeparator()

        self.loadBondFromFile = wx.MenuItem( self.staticData, ID_LOAD_BOND_FROM_FILE, u"Load bond from file", wx.EmptyString, wx.ITEM_NORMAL )
        self.staticData.Append( self.loadBondFromFile )

        self.m_menuItem18 = wx.MenuItem( self.staticData, wx.ID_ANY, u"Bond schedules", wx.EmptyString, wx.ITEM_NORMAL )
        self.staticData.Append( self.m_menuItem18 )

        self.staticData.AppendSeparator()

        self.m_menuEntity = wx.MenuItem( self.staticData, wx.ID_ANY, u"Add Entity", wx.EmptyString, wx.ITEM_NORMAL )
        self.staticData.Append( self.m_menuEntity )

        self.m_menubar1.Append( self.staticData, u"Static data" )

        self.m_CreditRatings = wx.Menu()
        self.m_menuItem17 = wx.MenuItem( self.m_CreditRatings, wx.ID_ANY, u"Manage", wx.EmptyString, wx.ITEM_NORMAL )
        self.m_CreditRatings.Append( self.m_menuItem17 )

        self.m_menubar1.Append( self.m_CreditRatings, u"Credit Ratings" )

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
        self.Bind( wx.EVT_MENU, self.f_add_bond_static_data, id = self.AddBondStaticData.GetId() )
        self.Bind( wx.EVT_MENU, self.f_load_bond_from_file, id = self.loadBondFromFile.GetId() )
        self.Bind( wx.EVT_MENU, self.fMenuBondSchedule, id = self.m_menuItem18.GetId() )
        self.Bind( wx.EVT_MENU, self.f_Add_Entity_Action, id = self.m_menuEntity.GetId() )
        self.Bind( wx.EVT_MENU, self.OnCreditRatings_Manage, id = self.m_menuItem17.GetId() )

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

    def f_add_bond_static_data( self, event ):
        event.Skip()

    def f_load_bond_from_file( self, event ):
        event.Skip()

    def fMenuBondSchedule( self, event ):
        event.Skip()

    def f_Add_Entity_Action( self, event ):
        event.Skip()

    def OnCreditRatings_Manage( self, event ):
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

        m_choice5Choices = [ u"Alexey", u"Olga" ]
        self.m_choice5 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice5Choices, 0 )
        self.m_choice5.SetSelection( 0 )
        fgSizer2.Add( self.m_choice5, 0, wx.ALL|wx.EXPAND, 5 )

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
        self.m_ISIN_input.Bind( wx.EVT_TEXT, self.PortfolioAddBond_ISINEnter )
        self.m_Cancel_button.Bind( wx.EVT_BUTTON, self.f_Cancel_button_pushed )
        self.m_add_to_portfolio.Bind( wx.EVT_BUTTON, self.f_add_to_portfolio )

    def __del__( self ):
        pass


    # Virtual event handlers, override them in your derived class
    def PortfolioAddBond_ISINEnter( self, event ):
        event.Skip()

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


###########################################################################
## Class Add_Instrument
###########################################################################

class Add_Instrument ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Bond form", pos = wx.DefaultPosition, size = wx.Size( 698,561 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        gSizer2 = wx.GridSizer( 0, 2, 0, 0 )

        self.m_staticText12 = wx.StaticText( self, wx.ID_ANY, u"Instrument ISIN", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText12.Wrap( -1 )

        gSizer2.Add( self.m_staticText12, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl11 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_textCtrl11, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText35 = wx.StaticText( self, wx.ID_ANY, u"Instrument type", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText35.Wrap( -1 )

        gSizer2.Add( self.m_staticText35, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice10Choices = [ u"bond", u"equity", u"etf", u"cash" ]
        self.m_choice10 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice10Choices, 0 )
        self.m_choice10.SetSelection( 0 )
        gSizer2.Add( self.m_choice10, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText36 = wx.StaticText( self, wx.ID_ANY, u"Trading place", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText36.Wrap( -1 )

        gSizer2.Add( self.m_staticText36, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice11Choices = [ u"MOEX" ]
        self.m_choice11 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice11Choices, 0 )
        self.m_choice11.SetSelection( 0 )
        gSizer2.Add( self.m_choice11, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText13 = wx.StaticText( self, wx.ID_ANY, u"Tiker", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText13.Wrap( -1 )

        gSizer2.Add( self.m_staticText13, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl12 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_textCtrl12, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText37 = wx.StaticText( self, wx.ID_ANY, u"Instrument currency", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText37.Wrap( -1 )

        gSizer2.Add( self.m_staticText37, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice12Choices = [ u"RUB", u"USD" ]
        self.m_choice12 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice12Choices, 0 )
        self.m_choice12.SetSelection( 0 )
        gSizer2.Add( self.m_choice12, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText14 = wx.StaticText( self, wx.ID_ANY, u"Bond coupon type", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText14.Wrap( -1 )

        gSizer2.Add( self.m_staticText14, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice1Choices = [ u"fixed", u"float", u"linker" ]
        self.m_choice1 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice1Choices, 0 )
        self.m_choice1.SetSelection( 0 )
        gSizer2.Add( self.m_choice1, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText23 = wx.StaticText( self, wx.ID_ANY, u"Float Bond percent base", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText23.Wrap( -1 )

        gSizer2.Add( self.m_staticText23, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice4Choices = [ u"none", u"RUONIA", u"КС" ]
        self.m_choice4 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice4Choices, 0 )
        self.m_choice4.SetSelection( 0 )
        gSizer2.Add( self.m_choice4, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText24 = wx.StaticText( self, wx.ID_ANY, u"Float bond addon to base", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText24.Wrap( -1 )

        gSizer2.Add( self.m_staticText24, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl18 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_textCtrl18, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText15 = wx.StaticText( self, wx.ID_ANY, u"Bond issue date", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText15.Wrap( -1 )

        gSizer2.Add( self.m_staticText15, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl13 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_textCtrl13, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText16 = wx.StaticText( self, wx.ID_ANY, u"Bond maturity date", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText16.Wrap( -1 )

        gSizer2.Add( self.m_staticText16, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl14 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_textCtrl14, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText17 = wx.StaticText( self, wx.ID_ANY, u"Bond call option dates", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText17.Wrap( -1 )

        gSizer2.Add( self.m_staticText17, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl15 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_textCtrl15, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText18 = wx.StaticText( self, wx.ID_ANY, u"Bond put option dates", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText18.Wrap( -1 )

        gSizer2.Add( self.m_staticText18, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl16 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_textCtrl16, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText19 = wx.StaticText( self, wx.ID_ANY, u"Bond credit rating", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText19.Wrap( -1 )

        gSizer2.Add( self.m_staticText19, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice2Choices = [ u"Gov", u"AAA", u"AAA-", u"AA+", u"AA", u"AA-", u"A+", u"A", u"A-", u"BBB+", u"BBB", u"BBB-", u"BB+", u"BB", u"BB-", u"B+", u"B", u"B-", u"CCC", u"CC", u"C", u"D" ]
        self.m_choice2 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice2Choices, 0 )
        self.m_choice2.SetSelection( 0 )
        gSizer2.Add( self.m_choice2, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText20 = wx.StaticText( self, wx.ID_ANY, u"Instrument issuer", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText20.Wrap( -1 )

        gSizer2.Add( self.m_staticText20, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice3Choices = []
        self.m_choice3 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice3Choices, 0 )
        self.m_choice3.SetSelection( 0 )
        gSizer2.Add( self.m_choice3, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_button5 = wx.Button( self, wx.ID_ANY, u"Add instrument", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_button5, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

        self.m_button6 = wx.Button( self, wx.ID_ANY, u"Cancel", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_button6, 0, wx.ALIGN_CENTER|wx.ALL, 5 )


        self.SetSizer( gSizer2 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.m_button5.Bind( wx.EVT_BUTTON, self.fAdd_instrument )
        self.m_button6.Bind( wx.EVT_BUTTON, self.Add_bond_cancel )

    def __del__( self ):
        pass


    # Virtual event handlers, override them in your derived class
    def fAdd_instrument( self, event ):
        event.Skip()

    def Add_bond_cancel( self, event ):
        event.Skip()


###########################################################################
## Class Entity
###########################################################################

class Entity ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 511,327 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        bSizer1 = wx.BoxSizer( wx.VERTICAL )

        self.m_staticText23 = wx.StaticText( self, wx.ID_ANY, u"Manage Entity", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
        self.m_staticText23.Wrap( -1 )

        bSizer1.Add( self.m_staticText23, 0, wx.ALL|wx.EXPAND, 5 )

        gSizer3 = wx.GridSizer( 6, 2, 0, 0 )

        self.m_staticText24 = wx.StaticText( self, wx.ID_ANY, u"UTI (INN)", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText24.Wrap( -1 )

        gSizer3.Add( self.m_staticText24, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl17 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer3.Add( self.m_textCtrl17, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL|wx.EXPAND, 5 )

        self.m_staticText25 = wx.StaticText( self, wx.ID_ANY, u"Short name", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText25.Wrap( -1 )

        gSizer3.Add( self.m_staticText25, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl18 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer3.Add( self.m_textCtrl18, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL|wx.EXPAND, 5 )

        self.m_staticText26 = wx.StaticText( self, wx.ID_ANY, u"Full name", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText26.Wrap( -1 )

        gSizer3.Add( self.m_staticText26, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl19 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer3.Add( self.m_textCtrl19, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL|wx.EXPAND, 5 )

        self.m_staticText27 = wx.StaticText( self, wx.ID_ANY, u"Comment", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText27.Wrap( -1 )

        gSizer3.Add( self.m_staticText27, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl20 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer3.Add( self.m_textCtrl20, 0, wx.ALIGN_CENTER_VERTICAL|wx.ALL|wx.EXPAND, 5 )

        self.m_button7 = wx.Button( self, wx.ID_ANY, u"Add", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer3.Add( self.m_button7, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 5 )

        self.m_button8 = wx.Button( self, wx.ID_ANY, u"Cancel", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer3.Add( self.m_button8, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 5 )


        bSizer1.Add( gSizer3, 1, wx.EXPAND, 5 )


        self.SetSizer( bSizer1 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.m_button7.Bind( wx.EVT_BUTTON, self.f_add_entity )
        self.m_button8.Bind( wx.EVT_BUTTON, self.f_cancel_entity )

    def __del__( self ):
        pass


    # Virtual event handlers, override them in your derived class
    def f_add_entity( self, event ):
        event.Skip()

    def f_cancel_entity( self, event ):
        event.Skip()


###########################################################################
## Class CreditRatings
###########################################################################

class CreditRatings ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Credit ratings", pos = wx.DefaultPosition, size = wx.Size( 750,356 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        fgSizer6 = wx.FlexGridSizer( 0, 1, 0, 0 )
        fgSizer6.AddGrowableCol( 0 )
        fgSizer6.AddGrowableRow( 1 )
        fgSizer6.SetFlexibleDirection( wx.BOTH )
        fgSizer6.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_ALL )

        fgSizer7 = wx.FlexGridSizer( 0, 4, 0, 0 )
        fgSizer7.AddGrowableCol( 0 )
        fgSizer7.AddGrowableCol( 1 )
        fgSizer7.AddGrowableCol( 2 )
        fgSizer7.AddGrowableCol( 3 )
        fgSizer7.AddGrowableRow( 0 )
        fgSizer7.SetFlexibleDirection( wx.BOTH )
        fgSizer7.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )

        self.m_staticText42 = wx.StaticText( self, wx.ID_ANY, u"Select Entity", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
        self.m_staticText42.Wrap( -1 )

        fgSizer7.Add( self.m_staticText42, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice13Choices = []
        self.m_choice13 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice13Choices, 0 )
        self.m_choice13.SetSelection( 0 )
        fgSizer7.Add( self.m_choice13, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText43 = wx.StaticText( self, wx.ID_ANY, u"Enter ISIN", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_CENTER_HORIZONTAL )
        self.m_staticText43.Wrap( -1 )

        fgSizer7.Add( self.m_staticText43, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl29 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer7.Add( self.m_textCtrl29, 0, wx.ALL|wx.EXPAND, 5 )


        fgSizer6.Add( fgSizer7, 1, wx.ALL|wx.EXPAND, 5 )

        fgSizer8 = wx.FlexGridSizer( 5, 2, 0, 0 )
        fgSizer8.AddGrowableCol( 0 )
        fgSizer8.AddGrowableCol( 1 )
        fgSizer8.AddGrowableRow( 0 )
        fgSizer8.SetFlexibleDirection( wx.VERTICAL )
        fgSizer8.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )

        m_listBox2Choices = []
        self.m_listBox2 = wx.ListBox( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_listBox2Choices, 0 )
        fgSizer8.Add( self.m_listBox2, 0, wx.ALL|wx.EXPAND, 5 )

        fgSizer9 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer9.AddGrowableCol( 0 )
        fgSizer9.AddGrowableCol( 1 )
        fgSizer9.SetFlexibleDirection( wx.BOTH )
        fgSizer9.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )

        self.m_staticText34 = wx.StaticText( self, wx.ID_ANY, u"Action", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText34.Wrap( -1 )

        fgSizer9.Add( self.m_staticText34, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice8Choices = []
        self.m_choice8 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice8Choices, 0 )
        self.m_choice8.SetSelection( 0 )
        fgSizer9.Add( self.m_choice8, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText48 = wx.StaticText( self, wx.ID_ANY, u"Date", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText48.Wrap( -1 )

        fgSizer9.Add( self.m_staticText48, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl37 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer9.Add( self.m_textCtrl37, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText49 = wx.StaticText( self, wx.ID_ANY, u"Rating", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText49.Wrap( -1 )

        fgSizer9.Add( self.m_staticText49, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice9Choices = []
        self.m_choice9 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice9Choices, 0 )
        self.m_choice9.SetSelection( 0 )
        fgSizer9.Add( self.m_choice9, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText50 = wx.StaticText( self, wx.ID_ANY, u"Rating agency", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText50.Wrap( -1 )

        fgSizer9.Add( self.m_staticText50, 0, wx.ALL|wx.EXPAND, 5 )

        m_choice14Choices = []
        self.m_choice14 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice14Choices, 0 )
        self.m_choice14.SetSelection( 0 )
        fgSizer9.Add( self.m_choice14, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_staticText51 = wx.StaticText( self, wx.ID_ANY, u"Forecast", wx.DefaultPosition, wx.DefaultSize, wx.ALIGN_RIGHT )
        self.m_staticText51.Wrap( -1 )

        fgSizer9.Add( self.m_staticText51, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_textCtrl40 = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0 )
        fgSizer9.Add( self.m_textCtrl40, 0, wx.ALL|wx.EXPAND, 5 )


        fgSizer8.Add( fgSizer9, 1, wx.ALL|wx.EXPAND, 5 )

        bSizer12 = wx.BoxSizer( wx.VERTICAL )

        self.m_button11 = wx.Button( self, wx.ID_ANY, u"Cancel", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer12.Add( self.m_button11, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 5 )


        fgSizer8.Add( bSizer12, 1, wx.EXPAND, 5 )

        bSizer13 = wx.BoxSizer( wx.VERTICAL )

        self.m_button13 = wx.Button( self, wx.ID_ANY, u"Create", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer13.Add( self.m_button13, 0, wx.ALIGN_CENTER_HORIZONTAL|wx.ALL, 5 )


        fgSizer8.Add( bSizer13, 1, wx.EXPAND, 5 )


        fgSizer6.Add( fgSizer8, 1, wx.ALL|wx.EXPAND, 5 )


        self.SetSizer( fgSizer6 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.m_choice13.Bind( wx.EVT_CHOICE, self.CreditRating_OnEntity )
        self.m_textCtrl29.Bind( wx.EVT_TEXT, self.CreditRating_OnISIN )
        self.m_choice8.Bind( wx.EVT_CHOICE, self.onAction_Selected )
        self.m_button11.Bind( wx.EVT_BUTTON, self.CreditRatings_onCancel )
        self.m_button13.Bind( wx.EVT_BUTTON, self.CreditRating_onAction )

    def __del__( self ):
        pass


    # Virtual event handlers, override them in your derived class
    def CreditRating_OnEntity( self, event ):
        event.Skip()

    def CreditRating_OnISIN( self, event ):
        event.Skip()

    def onAction_Selected( self, event ):
        event.Skip()

    def CreditRatings_onCancel( self, event ):
        event.Skip()

    def CreditRating_onAction( self, event ):
        event.Skip()


###########################################################################
## Class Bond_schedule
###########################################################################

class Bond_schedule ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Bond schedule", pos = wx.DefaultPosition, size = wx.Size( 838,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        bSizer4 = wx.BoxSizer( wx.VERTICAL )

        m_choice13Choices = []
        self.m_choice13 = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_choice13Choices, 0 )
        self.m_choice13.SetSelection( 0 )
        bSizer4.Add( self.m_choice13, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_grid1 = wx.grid.Grid( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, 0 )

        # Grid
        self.m_grid1.CreateGrid( 5, 5 )
        self.m_grid1.EnableEditing( True )
        self.m_grid1.EnableGridLines( True )
        self.m_grid1.EnableDragGridSize( False )
        self.m_grid1.SetMargins( 0, 0 )

        # Columns
        self.m_grid1.EnableDragColMove( False )
        self.m_grid1.EnableDragColSize( True )
        self.m_grid1.SetColLabelAlignment( wx.ALIGN_CENTER, wx.ALIGN_CENTER )

        # Rows
        self.m_grid1.EnableDragRowSize( True )
        self.m_grid1.SetRowLabelAlignment( wx.ALIGN_CENTER, wx.ALIGN_CENTER )

        # Label Appearance

        # Cell Defaults
        self.m_grid1.SetDefaultCellAlignment( wx.ALIGN_LEFT, wx.ALIGN_TOP )
        bSizer4.Add( self.m_grid1, 0, wx.ALL|wx.EXPAND, 5 )


        self.SetSizer( bSizer4 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.m_choice13.Bind( wx.EVT_CHOICE, self.fOnChoice )
        self.m_grid1.Bind( wx.grid.EVT_GRID_CELL_CHANGED, self.fOnGridCellChange )
        self.m_grid1.Bind( wx.grid.EVT_GRID_SELECT_CELL, self.fOnGridSelectCell )
        self.m_grid1.Bind( wx.EVT_SIZE, self.fOnSize )

    def __del__( self ):
        pass


    # Virtual event handlers, override them in your derived class
    def fOnChoice( self, event ):
        event.Skip()

    def fOnGridCellChange( self, event ):
        event.Skip()

    def fOnGridSelectCell( self, event ):
        event.Skip()

    def fOnSize( self, event ):
        event.Skip()


