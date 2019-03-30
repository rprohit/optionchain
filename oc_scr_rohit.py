from constants import symbols
import xlsxwriter
import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime
from string import Template
import lxml.html as lh

start_url_nifty = Template('https://www.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?segmentLink=17&instrument=OPTIDX&symbol=$SYMBOL')
start_url_stocks = Template('https://www.nseindia.com/marketinfo/sym_map/symbolMapping.jsp?symbol=$SYMBOL&instrument=OPTSTK&date=-&segmentLink=17')

def get_expirys(url):
    try:
        pg = requests.get(url)
        bsf = BeautifulSoup(pg.content, 'html5lib')
        exps = bsf.find('select', attrs={'name':'date'}).findAll('option')
        expirys = []
        for e in exps[1:]:
            expirys.append(e.contents[0].strip())
        return expirys
    except Exception as e:
        print(e)

def get_chain(url, date, symbol):
    try:
        url = f'{url}&date={date}'
        pg = requests.get(url)
        bsf = BeautifulSoup(pg.content, 'html5lib')
        table = bsf.find("table", attrs={'id':'octable'})
        table.find('thead')('tr')[0].extract()
        print(table)
        df = pd.read_html(table.prettify())
        #print(df)
        df = df[0]
        df = df.replace('-', 0)
        name = f'E:/Trading/Code/NseOptionsChainScrapper-master/data/{symbol}_{date}_{datetime.now():%Y-%m-%d}.xlsx'
        # name = f'{symbol}_{date}_{datetime.now():%Y-%m-%d}.xlsx'
        # out_path = "E:/Trading/Code/NseOptionsChainScrapper-master/data"
        # writer = pd.ExcelWriter(out_path, engine='xlsxwriter')
        # df.to_excel(writer, name)
		
        df.drop(df.tail(1).index,inplace=True) # drop last n rows , here n=1
        df.to_excel(name)
        print(f'Saved {name} ...')
    except Exception as e:
        print(e)

def get_options_chain(symbol):
    url = start_url_nifty.substitute(SYMBOL=symbol)
    print(url)#The url for option
    exps = get_expirys(url)
    print(exps)# exps is the list of all expirys available.
    for d in exps:
        get_chain(url, d, symbol)
        break# breaking after the first expiry 

def get_options_chain_stocks(symbol):
    url = start_url_stocks.substitute(SYMBOL=symbol)
    print(url)#The url for option
    exps = get_expirys(url)
    print(exps)# exps is the list of all expirys available.
    for d in exps:
        get_chain(url, d, symbol)
        break# breaking after the first expiry 
def get_nifty():
    get_options_chain('NIFTY')

def get_bank_nifty():
    get_options_chain('BANKNIFTY')

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
    df.drop(df.tail(1).index,inplace=True) # drop last n rows , here n=1
    name = f'E:/Trading/Code/NseOptionsChainScrapper-master/data/master_{datetime.now():%Y-%m-%d}.xlsx'
    df.to_excel(name)
	
if __name__ == '__main__':
    get_all_support_resistance()
    get_bank_nifty()
    get_nifty()
    #print(symbols)
	#Below code is to get expiry for All other symbols
    #for symbol in symbols:
        #get_options_chain_stocks(symbol)