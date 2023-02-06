import xlwings as xw


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet["A1"].value == "Hello xlwings!":
        sheet["A1"].value = "Bye xlwings!"
    else:
        sheet["A1"].value = "Hello xlwings!"


@xw.func
def hello(name):
    return f"Hello {name}!"




### Detect valid email addresses ###
import re
valid_email_re = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

@xw.func
def valid_email(email):
    if(re.fullmatch(valid_email_re, email)):
        return True 
    else:
        return False
####


### Look up timezone of phone number
import phonenumbers
from phonenumbers import timezone

@xw.func
def phone_timezone(phone_number, country_code):
    number_parse = phonenumbers.parse(phone_number, country_code)
    return timezone.time_zones_for_number(number_parse)

###

### Descriptive stats

import pandas as pd
@xw.func
@xw.arg("df", pd.DataFrame, index=False, header=True)
def describe(df):
    return df.describe()
 
###


if __name__ == "__main__":
    xw.Book("first_udf.xlsm").set_mock_caller()
    main()
