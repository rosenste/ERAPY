
import xlsxwriter
from bs4 import BeautifulSoup
import openpyxl
from email.parser import BytesParser
from email import message_from_file, policy
import pandas as pd
import re
import glob
import os
import talon
from talon import quotations
import traceback
import sys 

Directory=os.getcwd()+"\\Staging msg\\"
# returns the list of files inside the directory
file_list = glob.glob(Directory+'*.eml')
# Loop through all the files in the folder
for index, file in enumerate(file_list):
#Read the file
    f = r'{}'.format(file_list[index])
#get the data from the file
    with open(f, 'rb') as fp:
        body = BytesParser(policy=policy.default).parse(fp)
        Sender=body['from']
        print(Sender)
    #Use talon library to extract the reply from the email message, will remove the original message (works in 90% of the cases)
    talon.init()
    try:
        

        htmbody = quotations.extract_from(body.get_body(preferencelist=('html','plain')).get_content(),'text/html')
    # Parse HTML with BeautifulSoup ( library taht is made to scrape information, useful for HTML content)
        soup = BeautifulSoup(htmbody, "html.parser")
        #will check if there are lables in the text - meaning we have HTML tags and we can extract the text as HTML
        label = soup.find("table")
        filename=file[:-4] + '.xlsx'
# Read html and plain body, parse using talon quotations
        if label: # HTML tables
            print("ok")
            try:
            #looking for all tags inside the text -  will aggregate them and create the excel file
                #htmtext = list(dict.fromkeys([tag.text for tag in soup.find('body').find_all(True) if (tag.find_parent('table') is None) & (tag.name not in ['div','table'])]))
                tables = pd.read_html(htmbody, header=None, index_col=None, na_values="")
                #Checks if Flex is the sender
                if "Flex Notification" in Sender:
                    tables[2]=tables[2].replace({'Payment amount:': None})
                df = pd.concat(tables, axis=0).fillna('') # works regardless of number of
                print(df)
                writer = pd.ExcelWriter(filename, engine='xlsxwriter')
                df.to_excel(writer, index=False, header=False)
                writer.save()
            #Checks if there are numbers in the headers(1,2,3 etc as columns names) and if so - removes them
                try:
                    for index,i in enumerate(tables):
                        headers=list(i.columns)
                        print(headers)
                        if headers[0]==0:
                            i=i.rename(columns=i.iloc[0]).drop(i.index[0])
                        i=i.T.drop_duplicates().T
                       # i=i.T.reset_index(drop=True).T
                        tables[index]=i
                        print("Table Done")
                    df = pd.concat(tables, axis=0,ignore_index=False).fillna('').reset_index(drop=True) # works regardless of number of
                    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
                    df.to_excel(writer, index=False, header=True)
                    writer.save()
                    print("Done")
                except:
                    print("Can't find numbers in headers"+sys.exc_info())

            except AttributeError:
                # Append HTML text to tables
            #df = pd.concat([pd.DataFrame(htmtext),df]).fillna('')

            
                print("error")
                #tables = pd.read_html(htmbody, header=None, index_col=None)



        else:
            print("This is txtbody")
            txtbody = quotations.extract_from(body.get_body(preferencelist=('plain')).get_content(),'text/plain')
        # Alternative approach
            writer = pd.ExcelWriter(file[:-4]  + '.xlsx', engine='xlsxwriter')
        # I found these values in the body can throw off the table parsing / cause other noise
            noise = ['=\n', '=20']
            txtbody = re.sub('|'.join(noise), '', txtbody)
            print(txtbody)

        # Split lines (rows) by \n, separated columns by 2 spaces (compared
        # to prior default split [1 space], better for headers/etc), write to Excel as dataframe
            pd.DataFrame([[l for l in ln.strip().split() if l] for ln in txtbody.split('\n') if ln]).fillna('').to_excel(writer, index=False, header=False)
            writer.save()

    except AttributeError:
        print("This is Attribute Error 2")
        txtbody = quotations.extract_from(body.get_body(preferencelist=('plain')).get_content(),'text/plain')
        # Alternative approach
        writer = pd.ExcelWriter(file[:-4]  + '.xlsx', engine='xlsxwriter')
        # I found these values in the body can throw off the table parsing / cause other noise
        noise = ['=\n', '=20']
        txtbody = re.sub('|'.join(noise), '', txtbody)
        # Split lines (rows) by \n, separated columns by 2 spaces (compared
        # to prior default split [1 space], better for headers/etc), write to Excel as dataframe
        pd.DataFrame([[l for l in ln.strip().split()  if l] for ln in txtbody.split('\n') if ln]).fillna('\n').to_excel(writer, index=False, header=False)
        writer.save()
    except :
        f = open('Pythonlog.txt', 'w')
        e = sys.exc_info()[0]
        f.write('An exceptional thing happed - %s' % e)
        f.close()
