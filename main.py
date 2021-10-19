import pandas as pd
import csv
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt


filename="chat.txt"

def _convert_to_csv(filename):
    df=pd.read_csv(filename,header=None,error_bad_lines=False,encoding='utf8')
    df= df.drop(0)
    df.columns=['Date', 'Chat']
    Message= df["Chat"].str.split("-", n = 1, expand = True) 
    df['Date']=df['Date'].str.replace(",","") 
    df['Time']=Message[0]
    df['Text']=Message[1]
    Message1= df["Text"].str.split(":", n = 1, expand = True) 
    df['Text']=Message1[1]
    df['Name']=Message1[0]
    df=df.drop(columns=['Chat'])
    df['Text']=df['Text'].str.lower()
    df['Text'] = df['Text'].str.replace('<media omitted>','MediaShared')
    df['Text'] = df['Text'].str.replace('this message was deleted','DeletedMsg')    
    df.to_csv("chat.csv",index=False)

#_convert_to_csv(filename)

d = Document()
d.add_heading('Chat con Potato', 0)

p = d.add_paragraph()
p.add_run('Giorgia').bold = True
p.add_run('Mike').bold = True


file = csv.reader(open('chat.csv', 'r'))

# for line in file:
#     if line[3] == " Mike":
#         print('Mike:' + line[2])
#     elif line[3] == ' GiorgiaChiarucci':
#         print('Giorgia:' + line[2])


d.save('chat.docx')