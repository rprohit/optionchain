import requests
import lxml.html as lh
import pandas as pd
from bs4 import BeautifulSoup
from datetime import datetime

# Get all get possible expiry date details for the given script
def get_expiry_from_option_chain (symbol):

    # Base url page for the symbole with default expiry date
    Base_url = "https://www.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbol=" + symbol + "&date=-"

    # Load the page and sent to HTML parse
    page = requests.get(Base_url)
    soup = BeautifulSoup(page.content, 'html.parser')

    # Locate where expiry date details are available
    locate_expiry_point = soup.find(id="date")
    # Convert as rows based on tag option
    expiry_rows = locate_expiry_point.find_all('option')

    index = 0
    expiry_list = []
    for each_row in expiry_rows:
        # skip first row as it does not have value
        if index <= 0:
            index = index + 1
            continue
        index = index + 1
        # Remove HTML tag and save to list
        expiry_list.append(BeautifulSoup(str(each_row), 'html.parser').get_text())

    # print(expiry_list)
    return expiry_list # return list

def get_strike_price_from_option_chain(symbol, expdate):

    Base_url = "https://www.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbol=" + symbol + "&date=" + expdate

    page = requests.get(Base_url)
    soup = BeautifulSoup(page.content, 'html.parser')

    table_cls_2 = soup.find(id="octable")
    req_row = table_cls_2.find_all('tr')

    strike_price_list = []

    for row_number, tr_nos in enumerate(req_row):

        # This ensures that we use only the rows with values
        if row_number <= 1 or row_number == len(req_row) - 1:
            continue

        td_columns = tr_nos.find_all('td')
        strike_price = int(float(BeautifulSoup(str(td_columns[11]), 'html.parser').get_text()))
        strike_price_list.append(strike_price)

    # print (strike_price_list)
    return strike_price_list

def get_all_support_resistance():
    url='https://www.eqsis.com/nse-derivative-markets-option-chain'

#Create a handle, page, to handle the contents of the website
    page = requests.get(url)

#Store the contents of the website under doc
    doc = lh.fromstring(page.content)

#Parse data that are stored between <tr>..</tr> of HTML
    tr_elements = doc.xpath('//tr')

#Check the length of the first 12 rows
    [len(T) for T in tr_elements[:12]]

    tr_elements = doc.xpath('//tr')

#Create empty list
    col=[]
    i=0

#For each row, store each first element (header) and an empty list
    for t in tr_elements[0]:
        i+=1
        name=t.text_content()

        #print '%d:"%s"'%(i,name)
        col.append((name,[]))

#Since out first row is the header, data is stored on the second row onwards
    for j in range(1,len(tr_elements)):
        #T is our j'th row
        T=tr_elements[j]
        print('==================')
        print(T)
        print(len(T))		
    #If row is not of size 18, the //tr data is not from our table 
        if len(T)!=18:
            break
    
        #i is the index of our column
        i=0
    
    #Iterate through each element of the row
        for t in T.iterchildren():
            data=t.text_content() 
            #Check if row is empty
            if i>0:
            #Convert any numerical value to integers
                try:
                    data=int(data)
                except:
                    pass
            #Append the data to the empty list of the i'th column
            col[i][1].append(data)
            #Increment i for the next column
            i+=1

    [len(C) for (title,C) in col]

    Dict={title:column for (title,column) in col}
    df=pd.DataFrame(Dict)

    
    #df = pd.read_html(table)
    #df = df[0]
    name = f'E:/Trading/Code/NseOptionsChainScrapper-master/data/master_{datetime.now():%Y-%m-%d}.xlsx'
    df.to_excel(name)

if __name__ == '__main__':
    #print(get_strike_price_from_option_chain('NIFTY',get_expiry_from_option_chain('NIFTY')[0]))
    get_all_support_resistance()
