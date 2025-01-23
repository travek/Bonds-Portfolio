
import datetime

def string_is_date(check_str):
    # Date should be in a format YYYYMMDD
    if not isinstance(check_str, str):
        return False
    
    str_len=len(check_str)
    if str_len!=8:
        return False
    
    try:
        d = datetime.datetime.strptime(check_str, '%Y%m%d')
    except:
        return False
        
    if isinstance(d, datetime.date):
        return True
    
    return False
    
    
    
    