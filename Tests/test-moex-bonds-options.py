import sys
sys.path.append('d:\\Alexey\\Python Projects\\Bonds Portfolio')

import bonds_functions_db

moex_data=bonds_functions_db.get_bond_info_moex("RU000A1084B2")
print(f'isin: {moex_data["isin"]}, put_option_date: {moex_data["put_option_date"]}')

moex_data=bonds_functions_db.get_bond_info_moex("RU000A107UU5")
print(f'isin: {moex_data["isin"]}, put_option_date: {moex_data["put_option_date"]}')

moex_data=bonds_functions_db.get_bond_info_moex("RU000A108TU5")
print(f'isin: {moex_data["isin"]}, put_option_date: {moex_data["put_option_date"]}')