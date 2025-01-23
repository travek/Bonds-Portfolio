import sqlite3
import bonds_functions_db

connection = sqlite3.connect('portfolio_database.db')   
res=bonds_functions_db.calc_bond_discounted_margine(connection.cursor(), "RU000A1087G4")