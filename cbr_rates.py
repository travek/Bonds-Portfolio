import requests  
import xml.etree.ElementTree as ET  
import datetime
from datetime import timedelta
from sqlalchemy import create_engine, Table, Column, Integer, String, MetaData, select, and_


def get_curs_on_date(date):  
    url = 'https://www.cbr.ru/DailyInfoWebServ/DailyInfo.asmx'  
    headers = {  
        'Content-Type': 'text/xml; charset=utf-8',  
        'SOAPAction': 'http://web.cbr.ru/GetCursOnDate'  
    }  
    body = f'''<?xml version="1.0" encoding="utf-8"?>  
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">  
      <soap:Body>  
        <GetCursOnDate xmlns="http://web.cbr.ru/">  
          <On_date>{date}</On_date>  
        </GetCursOnDate>  
      </soap:Body>  
    </soap:Envelope>'''  
  
    #print(body)
    response = requests.post(url, data=body, headers=headers)  
    response.raise_for_status()  # Поднимаем исключение, если запрос завершился с ошибкой  
    return response.content 

def parse_curs_on_date_response(response_content):  
    root = ET.fromstring(response_content)  
    namespaces = {  
        'soap': 'http://schemas.xmlsoap.org/soap/envelope/',  
        'diffgr': 'urn:schemas-microsoft-com:xml-diffgram-v1',  
        'msdata': 'urn:schemas-microsoft-com:xml-msdata',  
        'xs': 'http://www.w3.org/2001/XMLSchema',  
        'msprop': 'urn:schemas-microsoft-com:xml-msprop'  
    }  
  
    # Найти элемент с датой  
    schema_elem = root.find('.//xs:schema', namespaces)  
    on_date = schema_elem.find('.//xs:element', namespaces).get('{urn:schemas-microsoft-com:xml-msprop}OnDate')  
  
    # Найти все записи валют  
    curs_elements = root.findall('.//diffgr:diffgram/ValuteData/ValuteCursOnDate', namespaces)  
    rates = {}  
  
    for elem in curs_elements:  
        char_code = elem.find('VchCode').text  
        rate = float(elem.find('Vcurs').text)  
        rates[char_code] = rate  
  
    return on_date, rates             

def write2DB_fxrates_on_date():
    #Get FC rates and write to DB
    engine = create_engine('sqlite:///portfolio_database.db')
    
    metadata = MetaData()
    market_data = Table('market_data', metadata, autoload_with=engine)
    
    today = datetime.datetime.today()+datetime.timedelta(days=1)
    today_date = today.date()
    date = today_date.strftime('%Y-%m-%dT%H:%M:%S')   
    
    try:
        #print(date)
        response_content = get_curs_on_date(date)
        on_date, rates = parse_curs_on_date_response(response_content)
        #print(response_content)
        #print(f"Курсы валют на дату: {on_date}")

    
        # Список валют для обработки
        currencies = [
            ('USD', 'USDRUB', 'fx pair', 'USD/RUB'),
            ('EUR', 'EURRUB', 'fx pair', 'EUR/RUB'),
            ('CNY', 'CNYRUB', 'fx pair', 'CNY/RUB'),
        ]
        
        #on_date = datetime.datetime.strptime(on_date, '%Y%m%d').replace(hour=0, minute=0, second=0, microsecond=0)
        
        for currency_code, instrument_id, id_type, instrument_short_name in currencies:
            rate_value = rates.get(currency_code, 'Не найден')
            #print(currency_code, instrument_id)
            if rate_value != "Не найден":                    
                #print(f'Rate added! {instrument_short_name} - {on_date} - {rate_value}')
                print(f'Курс {currency_code}: {rate_value}, дата: {on_date}')
                
                with engine.begin() as conn:
                    # Ищем запись
                    result = conn.execute(
                        select(market_data).where(
                            and_(
                                market_data.c.ID == currency_code,
                                market_data.c.date == on_date
                            )
                        )
                    ).first()
                    
                    if result:
                        # Обновляем
                        conn.execute(
                            market_data.update().where(
                                and_(
                                    market_data.c.ID == currency_code,
                                    market_data.c.date == on_date
                                )
                            ).values(
                                price=rate_value
                            )
                        )
                        print(f"Обновлена существующая запись: {currency_code}, дата {on_date}, курс {rate_value}")
                    else:
                        # Вставляем новую
                        conn.execute(
                            market_data.insert().values(
                                ID=currency_code,
                                ID_type='fx pair',
                                price=rate_value,
                                date=on_date,
                                price_nominal='RUB'
                            )
                        )
                        print(f"Добавлена новая запись: {currency_code}, дата {on_date}, курс {rate_value}")
                        
            else:
                print(f'Курс {currency_code} не найден.')
                
    except requests.RequestException as e:
        print(f'Ошибка при получении курсов валют: {e}')
    except Exception as e:
        print(f'Ошибка в скрипте: {e}')    
        
    return 0

#if __name__ == "__main__":
    #write2DB_fxrates_on_date()
    #print("fin")
    
    
    #if __name__ == "__main__":
        #today = datetime.datetime.today()+datetime.timedelta(days=1)
        #today_date = today.date()
        #date = today_date.strftime('%Y-%m-%dT%H:%M:%S')   
        
        #try:
            ##print(date)
            #response_content = get_curs_on_date(date)
            #on_date, rates = parse_curs_on_date_response(response_content)
            ##print(response_content)
            ##print(f"Курсы валют на дату: {on_date}")
    
        
            ## Список валют для обработки
            #currencies = [
                #('USD', 'USDRUB', 'fx pair', 'USD/RUB'),
                #('EUR', 'EURRUB', 'fx pair', 'EUR/RUB'),
                #('CNY', 'CNYRUB', 'fx pair', 'CNY/RUB'),
            #]
            
            ##on_date = datetime.datetime.strptime(on_date, '%Y%m%d').replace(hour=0, minute=0, second=0, microsecond=0)
            
            #for currency_code, instrument_id, id_type, instrument_short_name in currencies:
                #rate_value = rates.get(currency_code, 'Не найден')
                ##print(currency_code, instrument_id)
                #if rate_value != "Не найден":                    
                    #print(f'Rate added! {instrument_short_name} - {on_date} - {rate_value}')
                    #print(f'Курс {currency_code}: {rate_value}')
                #else:
                    #print(f'Курс {currency_code} не найден.')
                    
        #except requests.RequestException as e:
            #print(f'Ошибка при получении курсов валют: {e}')
        #except Exception as e:
            #print(f'Ошибка в скрипте: {e}')   
            