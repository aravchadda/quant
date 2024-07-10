#!/usr/bin/env python
# coding: utf-8

# In[1]:


#!/usr/bin/env python
# coding: utf-8

import os
import pickle
import pandas as pd
import numpy as np
import win32timezone
import datetime as dt
from datetime import datetime
print('-'*300)
print("Welcome to the quant Email Scraper.\nDeveloped by quants seniormost/juniormost Python developer.\nA summary of all company reasearch reports from CITI, Jefferies and Macquarie will appear in the same folder as this file titled as ReportSummary <Date_time>.\nTo continue please enter the following details\n\n ")
year=int(input('Enter the year you want the reports from yy eg 24'))
month = int(input('Enter a month as integer 01-12 : '))
day = int(input('Enter a day as integer 01-31 dd: '))
hours =int(float(input('Enter the time in format hhmm eg 1750 for 5:50 p.m. : ')))
minutes=hours%100
hours=int(hours/100)
print(hours,minutes)
try:
	todays_date= datetime(year+2000, month, day, hours, minutes, 0)#+timedelta(hours=5.5)
except:
    print("error in format")
    buffer=input("press any key to exit")
    quit()

print(f"Now showing from :{todays_date} to present")


# In[8]:


#### -*- coding: utf-8 -*-|
import re
alphabets= "([A-Za-z])"
prefixes = "(Mr|St|Mrs|Ms|Dr|www|est|Est|op|Op|approx|mgmt|mgt|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sept|Oct|Nov|Dec|vs|VS|Vs)[.]"
suffixes = "(Inc|Ltd|Jr|Sr|Co|vs|est)"
starters = "(Mr|Mrs|Ms|Dr|Prof|Capt|Cpt|Lt|He\s|She\s|It\s|They\s|Their\s|Our\s|We\s|But\s|However\s|That\s|This\s|Wherever)"
acronyms = "([A-Z][.][A-Z][.](?:[A-Z][.])?)"
websites = "[.](com|net|org|io|gov|edu|me|in|BO|NS)"
digits = "([0-9])"
multiple_dots = r'\.{2,}'

#function defined for splitting entire email body contents into separate sentences
def split_into_sentences(text: str) -> list[str]:
    """
    Split the text into sentences.
    
    If the text contains substrings "<prd>" or "<stop>", they would lead 
    to incorrect splitting because they are used as markers for splitting.
    :param text: text to be split into sentences
    :type text: str
    :return: list of sentences
    :rtype: list[str]
    """
    text = " " + text + "  "
    text = text.replace("\n",".")
    text = re.sub(prefixes,"\\1<prd>",text)
    text = re.sub(websites,"<prd>\\1",text)
    text = re.sub(digits + "[.]" + digits,"\\1<prd>\\2",text)
    text = re.sub(multiple_dots, lambda match: "<prd>" * len(match.group(0)) + "<stop>", text)
    if "Ph.D" in text: text = text.replace("Ph.D.","Ph<prd>D<prd>")
    text = re.sub("\s" + alphabets + "[.] "," \\1<prd> ",text)
    text = re.sub(acronyms+" "+starters,"\\1<stop> \\2",text)
    text = re.sub(alphabets + "[.]" + alphabets + "[.]" + alphabets + "[.]","\\1<prd>\\2<prd>\\3<prd>",text)
    text = re.sub(alphabets + "[.]" + alphabets + "[.]","\\1<prd>\\2<prd>",text)
    text = re.sub(" "+suffixes+"[.] "+starters," \\1<stop> \\2",text)
    text = re.sub(" "+suffixes+"[.]"," \\1<prd>",text)
    text = re.sub(" " + alphabets + "[.]"," \\1<prd>",text)
    if "”" in text: text = text.replace(".”","”.")
    if "\"" in text: text = text.replace(".\"","\".")
    if "!" in text: text = text.replace("!\"","\"!")
    if "?" in text: text = text.replace("?\"","\"?")
    text = text.replace(".",".<stop>")
    text = text.replace("?","?<stop>")
    text = text.replace("!","!<stop>")
    text = text.replace("<prd>",".")
    sentences = text.split("<stop>")
    sentences = [s.strip() for s in sentences]
    if sentences and not sentences[-1]: sentences = sentences[:-1]
    return sentences


# In[92]:


import transformers
from transformers import PegasusTokenizer, PegasusForConditionalGeneration

# Let's load the model and the tokenizer 
def summa(body_content,sub):

    model_name = "human-centered-summarization/financial-summarization-pegasus"
    tokenizer = PegasusTokenizer.from_pretrained(model_name)
    model = PegasusForConditionalGeneration.from_pretrained(model_name) # If you want to use the Tensorflow model 
    try:
        body_content=body_content[:re.search(r'\+?\(?[1-9][0-9 .\-\(\)]{8,}[0-9]', body_content).start()]
    except:
         pass
    st=''
    if body_content.count(' ')>1500:
        body_content=body_content[:10000]
    
    for i in range(int(len(body_content)/500)):
        input_ids = tokenizer(body_content[500*i:500*i+500], return_tensors="pt").input_ids
        # Generate the output (Here, we use beam search but you can also use any other strategy you like)
        output = model.generate(
            input_ids, 
            max_length=32, 
            num_beams=5, 
            early_stopping=True
        )
        st+=str(tokenizer.decode(output[0], skip_special_tokens=True))
    print(st+'-'*200)
    other_row["summary"]=st
    other_row["Subject"]=sub


# In[93]:


#function for jefferies company wise summarizing
def jefferies(body_content,sub):
    body_content=body_content.replace('\n','.')
    row['Analysts Opinion']=f'\'{sub}\''
    list_of_sentences=split_into_sentences(body_content)
    for n,j in enumerate(list_of_sentences):
        if 'Subject' in j:
            continue
        if j.count('|')==3:
            jparts=j.split('|')
            row["Name"]=jparts[0]+'\nJEF' if jparts[0][:4]!='Ltd.' else ''
            row["Target"]=jparts[2].split(':')[1] if jparts[2].count(':')==1 else jparts[2]
            row["% to TP"]=jparts[3].split(':')[1] if jparts[3].count(':')==1 else jparts[3]
        if '<https://javatar.' in j:
            j=j.replace('<https://javatar.','.')
        list_of_sentences[n]=j.replace('\n',".").replace('\r',"");
        ebitda_words=[' ebitda ',' ebit ',' opm ',' ebitdam ',' vnbm ',' nim ',' nims ',' operating profit ',' margin',' vnbm ']
        revenue_words=[' revenue ',' revpar ',' arr ' ,' sales ',' nii ',' operating income ',' fee income ',' ape ',' vnb ',' volume',' pricing',' realization',' realisation',' orders ','bill','acv','tcv','book to bill','book-to-bill','order book','order backlog']
        pat_words=[' pat ',' apat ',' profit after tax ',' net income ',' roev ',' ev ',' earning',' bv ',' bvps ','book vaule']
        res = any(ele in j.lower() for ele in ebitda_words)
        if 'Exhibit ' in j:
            continue
        if 'we ' in j.lower() or ' our' in j.lower():
            row["Analysts Opinion"]+="\n"+j
        elif res:
            row["EBITDA"]+="\n"+j;
        elif any(ele in j.lower() for ele in revenue_words):
            row["Revenue"]+="\n"+j;
        elif any(ele in j.lower() for ele in pat_words):
            row["PAT"]+="\n"+j;
    if row["Name"]=='':
        summa(body_content,'JEF'+sub)


# In[94]:


#function for citi company wise summarizing
def citi(sub,body_content):
    list_of_sentences=split_into_sentences(body_content)
    body_content=body_content.replace("\u2014", "-")
    flag=0
    if sub[-1]==')':
            flag=1
    else:
        row['Analysts Opinion']=f'\'{sub}\''
        for n,j in enumerate(list_of_sentences):
            end=sub.find('(') if sub.find('(') >0 else 999 
            row['Name']=sub[0:min(22,end)]+'\nCiti'
            ebitda_words=[' ebitda ',' ebit ',' opm ',' ebitdam ',' vnbm ',' nim ',' nims ',' operating profit ',' margin',' vnbm ']
            revenue_words=[' revenue ',' revpar ',' arr ' ,' sales ',' nii ',' operating income ',' fee income ',' ape ',' vnb ',' volume',' pricing',' realization',' realisation',' orders ','bill','acv','tcv','book to bill','book-to-bill','order book','order backlog']
            pat_words=[' pat ',' apat ',' profit after tax ',' net income ',' roev ',' ev ',' earning',' bv ',' bvps ','book vaule']
            if '<tel:+91' in j:
                break
            if 'target ' in j.lower() or ' tp ' in j.lower():
                row["Target"]=j
            link_index=j.find('<https:')
            if link_index !=-1:#checking if there is a link in this sentence
                j=j[0:link_index]+j[j.find('>',link_index):]#in order to remove only the part within <> as that is ussualy a hyperlink
            elif 'we ' in j.lower() or ' our ' in j.lower():
                if j not in row["Analysts Opinion"]:
                    row["Analysts Opinion"]+="\n"+j
            elif any(ele in j.lower() for ele in ebitda_words):
                if j not in row["EBITDA"]:
                    row["EBITDA"]+="\n"+j
            elif any(ele in j.lower() for ele in revenue_words):
                if j not in row["Revenue"]:
                    row["Revenue"]+="\n"+j
            elif any(ele in j.lower() for ele in pat_words):
                if j not in row["PAT"]:
                    row["PAT"]+="\n"+j
    if row["Name"]=='' or flag ==1:
        summa(body_content,'Citi\n'+sub)


# In[95]:


#function for macquarie company wise summarization
def macquire(sub,body_content):
    if sub.count('(') <2:
        if 'desk stratergy' in body_content.lower() or True:
            summa(body_content,"Mac\n"+sub)
        return
    if "flashnote" in sub.lower():
        row['Name']=sub[sub.find('flashnote')+9:sub.find('(') ]+'\nMacquarie'
    else:
        row['Name']=sub[0:sub.find('(')]+'\nMacquarie'
    row['Analysts Opinion']=f'\'{sub}\''
    flag=0
    list_of_sentences=split_into_sentences(body_content)
    for n,j in enumerate(list_of_sentences):
        #print(j+"\n")
        #list_of_sentences[n]=j.replace('\n',".").replace('\r',"");
        ebitda_words=[' ebitda ',' ebit ',' opm ',' ebitdam ',' vnbm ',' nim ',' nims ',' operating profit ',' margin',' vnbm ']
        revenue_words=[' revenue ',' revpar ',' arr ' ,' sales ',' nii ',' operating income ',' fee income ',' ape ',' vnb ',' volume',' pricing',' realization',' realisation',' orders ','bill','acv','tcv','book to bill','book-to-bill','order book','order backlog']
        pat_words=[' pat ',' apat ',' profit after tax ',' net income ',' roev ',' ev ',' earning',' bv ',' bvps ','book vaule']
        if '<tel:+91' in j:
            break
            
        link_index=j.find('<http:')
        if link_index !=-1:#checking if there is a link in this sentence
            j=j[0:link_index]+j[j.find('>',link_index):]
        link_index=j.find('<https:')
        if link_index !=-1:#checking if there is a link in this sentence
            j=j[0:link_index]+j[j.find('>',link_index):]
            
        if ('target' in j.lower() or ' tp ' in j.lower() or 'pt ' in j.lower()) and flag==0 :
            r=n+1;flag=1
            while  any(chr.isdigit() for chr in list_of_sentences[r]) and list_of_sentences[r]!='.':
                r=r+1
            row["Target"]=j+'\n'+list_of_sentences[r]
            print(list_of_sentences[r])
        if 'we ' in j.lower() or 'our' in j.lower():
            if j not in row["Analysts Opinion"]:
                row["Analysts Opinion"]+="\n"+j
        elif any(ele in j.lower() for ele in ebitda_words):
            if j not in row["EBITDA"]:
                row["EBITDA"]+="\n"+j
        elif any(ele in j.lower() for ele in revenue_words):
            if j not in row["Revenue"]:
                row["Revenue"]+="\n"+j
        elif any(ele in j.lower() for ele in pat_words):
            if j not in row["PAT"]:
                row["PAT"]+="\n"+j


# In[96]:


def Investec(sub,body_content):
    body_content=body_content.replace('\n','.')
    list_of_sentences=split_into_sentences(body_content)
    row['Analysts Opinion']=f'\'{sub}\''
    for n,j in enumerate(list_of_sentences):
        if 'Subject' in j:
            continue
        link_index=j.find('<http:')
        if link_index !=-1:#checking if there is a link in this sentence
            j=j[0:link_index]+j[j.find('>',link_index):]
        link_index=j.find('<https:')
        if link_index !=-1:#checking if there is a link in this sentence
            j=j[0:link_index]+j[j.find('>',link_index):]
            
        if j.count('|')==2:
            jparts=j.split('|')
            row["Name"]=list_of_sentences[n-3]+list_of_sentences[n-2]+'\nInvestec'
            row["Target"]=jparts[1].split(':')[1] 
        list_of_sentences[n]=j.replace('\n',".").replace('\r',"");
        ebitda_words=[' ebitda ',' ebit ',' opm ',' ebitdam ',' vnbm ',' nim ',' nims ',' operating profit ',' margin',' vnbm ']
        revenue_words=[' revenue ',' revpar ',' arr ' ,' sales ',' nii ',' operating income ',' fee income ',' ape ',' vnb ',' volume',' pricing',' realization',' realisation',' orders ','bill','acv','tcv','book to bill','book-to-bill','order book','order backlog']
        pat_words=[' pat ',' apat ',' profit after tax ',' net income ',' roev ',' ev ',' earning',' bv ',' bvps ','book vaule']
        res = any(ele in j.lower() for ele in ebitda_words)
        if 'IMPORTANT NOTICE' in j:
            break
        if 'we ' in j.lower() or ' our' in j.lower():
            row["Analysts Opinion"]+="\n"+j
        elif res:
            row["EBITDA"]+="\n"+j;
        elif any(ele in j.lower() for ele in revenue_words):
            row["Revenue"]+="\n"+j;
        elif any(ele in j.lower() for ele in pat_words):
            row["PAT"]+="\n"+j;
    if row["Name"]=='':
        summa(body_content,'Investec\n'+sub)


# In[97]:


def CLSA(sub,body_content):
    body_content=body_content.replace('\n','.')
    list_of_sentences=split_into_sentences(body_content)
    row['Analysts Opinion']=f'\'{sub}\''
    for n,j in enumerate(list_of_sentences):
        
        link_index=j.find('<http:')
        if link_index !=-1:#checking if there is a link in this sentence
            j=j[0:link_index]+j[j.find('>',link_index):]
        link_index=j.find('<https:')
        if link_index !=-1:#checking if there is a link in this sentence
            j=j[0:link_index]+j[j.find('>',link_index):]
            
        if 'India >'in j:
            row['Name']=j[7:-5]
        if 'Subject' in j:
            continue                                                    
        list_of_sentences[n]=j.replace('\n',".").replace('\r',"");
        print(j+'\n')
        ebitda_words=[' ebitda ',' ebit ',' opm ',' ebitdam ',' vnbm ',' nim ',' nims ',' operating profit ',' margin',' vnbm ']
        revenue_words=[' revenue ',' revpar ',' arr ' ,' sales ',' nii ',' operating income ',' fee income ',' ape ',' vnb ',' volume',' pricing',' realization',' realisation',' orders ','bill','acv','tcv','book to bill','book-to-bill','order book','order backlog']
        pat_words=[' pat ',' apat ',' profit after tax ',' net income ',' roev ',' ev ',' earning',' bv ',' bvps ','book vaule']
        res = any(ele in j.lower() for ele in ebitda_words)   
        if 'Click to rate this research' in j:
            break
        if 'target ' in j.lower() or ' tp ' in j.lower():
                row["Target"]=j
        if 'we ' in j.lower() or 'our ' in j.lower():
            row["Analysts Opinion"]+="\n"+j
        elif res:
            row["EBITDA"]+="\n"+j;
        elif any(ele in j.lower() for ele in revenue_words):
            row["Revenue"]+="\n"+j;
        elif any(ele in j.lower() for ele in pat_words):
            row["PAT"]+="\n"+j;
    if row["Name"]=='':
        summa(body_content,'CLSA\n'+sub)


# In[98]:


import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
emailno=0
rows_list=[]
other_list=[]
messages = inbox.Items.Restrict(f"[SentOn] > '"+todays_date.strftime('%d/%m/%Y %H:%M %p') +"'")
for found in messages:
    #creating a dictionary with all the columns
    row={key: None for key in ['Name','Date', 'Target', '% to TP', 'EBITDA', 'PAT', 'Revenue','Analysts Opinion'] }
    other_row={key:"" for key in ['Subject','Date','summary']}
    body_content=found.body #assigning the body parameter of the email 
    row={key:"" for key in row}
    # for each email matched, read it (output plain/text to console & save HTML and attachments)
    try:
        row["Date"]=str(found.ReceivedTime)[:16]
    except:
        continue
    other_row["Date"]=row["Date"]
    emailno+=1
    sub=found.subject
    #choosing which firms function to move into
    if 'Investec' in body_content:
        Investec(sub,body_content)
    elif "CITI'S TAKE" in body_content:
        citi(sub,body_content)
    elif "jefferies" in body_content: 
        jefferies(body_content,sub) 
    elif "macquarie" in body_content.lower():
        macquire(sub,body_content)
    elif "clsa" in body_content.lower():
        CLSA(sub,body_content)
    else:
        continue;
    rows_list.append(row)
    other_list.append(other_row)
    
print('No. of emails checked:',emailno)
df=pd.DataFrame(rows_list)
other_df=pd.DataFrame(other_list)
#df = df[['Name','Date', 'Target', '% to TP', 'EBITDA', 'PAT', 'Revenue','Opinion']]
try:    
    df = df[df.Name != ""]
    other_df = other_df[other_df.summary != ""]
except:
    pass


# In[99]:


pd.set_option("max_colwidth", None)
cwd = os.getcwd()
path = cwd + '//'+f"Earnings Reports Summary {str(day)}-{str(month)}-{str(year)}"+".xlsx"
print(cwd)
df=df.drop_duplicates(subset=["Name"])
writer = pd.ExcelWriter(path) 
# Get the xlsxwriter workbook 
workbook  = writer.book
# Add a format.
text_format = workbook.add_format({'text_wrap' : True})
df.to_excel(writer, sheet_name='Earnings Report', index=False,freeze_panes=(1, 1) )
writer.sheets['Earnings Report'].set_column(0,0 ,20,text_format)
writer.sheets['Earnings Report'].set_column(1,1 , 9.7,text_format)
writer.sheets['Earnings Report'].set_column(2,2 ,8,text_format)
writer.sheets['Earnings Report'].set_column(3,3 ,3.5,text_format)
writer.sheets['Earnings Report'].set_column(4,4 ,55,text_format)
writer.sheets['Earnings Report'].set_column(5,5 ,50,text_format)
writer.sheets['Earnings Report'].set_column(6,6 ,55,text_format)
writer.sheets['Earnings Report'].set_column(7,7,60,text_format)
writer.close()


# In[100]:


cwd = os.getcwd()
path = cwd + '//'+f"other {str(day)}-{str(month)}-{str(year)}"+".xlsx"
print(cwd)
other_df=other_df.drop_duplicates()
other_df.to_excel(path,sheet_name='Other Reports',index=False,freeze_panes=(1, 1))


# In[ ]:




