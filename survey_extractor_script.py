import docx
from os import listdir
import re
import pandas as pd

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

def survey_fill():
    #select docx files in survey directory
    files_to_scan=[]
    for i in listdir('surveys'):
        if i.endswith('.docx')==True and i.startswith('~')==False:
            files_to_scan.append(i)
    
    first_n_list=[]
    last_n_list=[]
    birth_list=[]
    sex_list=[]
    address_list=[]
    zip_list=[]
    telephone_list=[]
    mobile_list=[]
    occupation_list=[]
    email_list=[]
    ethnic_list=[]
    kin_list=[]
    kin_relation_list=[]
    kin_tele_list=[]
    mob_list=[]
    other_list=[]
    for i in files_to_scan:
        survey=getText('surveys/'+i).encode('utf-8').strip().decode(errors='ignore').replace('.','')
        #first name
        first_name_regex = r"^First\s*name\:\s*(\w*)\s*Surname"
        first_n_list.append(re.findall(first_name_regex,survey,re.MULTILINE)[0].strip())
        #last name
        last_name_regex = r"Surname\/s\s*(\w*)\s*\n*Date"
        last_n_list.append(re.findall(last_name_regex,survey,re.MULTILINE)[0].strip())
        #birth
        birth_regex = "Date\s*of\s*birth\:\s*(.*)\s*Sex"
        birth_list.append(re.findall(birth_regex,survey,re.MULTILINE)[0].strip())
        #Sex: will have to implement boolean condition on determing if x happens before or after female
        sex_regex = "Sex\:\s*(.*)"
        sex_raw=re.findall(sex_regex,survey,re.MULTILINE)[0].strip()
        sex_raw = 'Male' if 'x' in sex_raw.split('female')[0] else 'Female'
        sex_list.append(sex_raw)
        #address
        adress_regex = '^Address\:\s*(.*)'
        address_list.append(re.findall(adress_regex,survey,re.MULTILINE)[0].strip())
        #zip code
        zip_regex='Zip\scode\:\s*(.*)'
        zip_list.append(re.findall(zip_regex,survey,re.MULTILINE)[0].strip())
        #telephone 
        telphone_regex='Telephone\snumber\:\s*(.*)Mobile'
        telephone_list.append(re.findall(telphone_regex,survey,re.MULTILINE)[0].strip())
        #mobile
        mobile_regex='Mobile\sNumber\:\s*(.*)'
        mobile_list.append(re.findall(mobile_regex,survey,re.MULTILINE)[0].strip())
        #occupation
        occupation_regex='Occupation\:\s*(.*)\s*Email'
        occupation_list.append(re.findall(occupation_regex,survey,re.MULTILINE)[0].strip())
        #email
        email_regex='Email\s*(.*)'
        email_list.append(re.findall(email_regex,survey,re.MULTILINE)[0].strip())
        #ethnicity
        ethnic_regex='Ethnic\sorigin\:\s*(.*)\('
        ethnic_list.append(re.findall(ethnic_regex,survey,re.MULTILINE)[0].strip())
        #next of kin
        kin_regex='Next.*Name\):\s*(.*)R'
        kin_list.append(re.findall(kin_regex,survey,re.MULTILINE)[0].strip())
        #kin_relation
        kin_relation_regex='Relationship\:\s*(.*)'
        kin_relation_list.append(re.findall(kin_relation_regex,survey,re.MULTILINE)[0].strip())
        #kin_phone
        kin_tele_regex='kin\stelephone\snumber\:\s*(.*)'
        kin_tele_list.append(re.findall(kin_tele_regex,survey,re.MULTILINE)[0].strip())
        #mobility
        mob_regex='apply\).*'
        mob_raw=re.findall(mob_regex,survey,re.MULTILINE)[0].strip()
        mob_raw= 'Yes' if 'x' in mob_raw.split('Yes')[0] else 'No'
        mob_list.append(mob_raw)
        #other illness
        other_regex='possible\)((.|\n)*)'
        other_list.append(re.findall(other_regex,survey,re.MULTILINE)[0][0].strip())
    df=pd.DataFrame()
    labels=['first_name','last_name','birthdate','gender','address','zip_code','home_phone_number',\
            'cell_phone_number','occupation','email','ethnicity','next_of_kin_name','next_of_kin_relationship',\
            'next_of_kin_telephone','mobility_hearing_speaking','other_illness_commentary']
    
    lists_to_column=[first_n_list,last_n_list,birth_list,sex_list,address_list,zip_list,
                    telephone_list,mobile_list,occupation_list,email_list,ethnic_list,kin_list,
                    kin_relation_list,kin_tele_list,mob_list,other_list]
    
    for i in labels:
        df[i]=lists_to_column[labels.index(i)]
    print df
    return df.to_csv('survey_responses.csv',index=False)
survey_fill()    