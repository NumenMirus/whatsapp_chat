import pandas as pd
import csv
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from tqdm import tqdm
import os


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
    df['Text'] = df['Text'].str.replace('<media omessi>','**Foto o Video**')
    df['Text'] = df['Text'].str.replace('hai eliminato questo messaggio','Messaggio Eliminato')
    df['Text'] = df['Text'].str.replace('Questo messaggio è stato eliminato','Messaggio Eliminato')    
    df.to_csv("chat.csv",index=False)

def _create_document():
    d = Document()

    #creating the style
    styles = d.styles
    style = styles.add_style('Arial', WD_STYLE_TYPE.PARAGRAPH) #Tahoma is the name I set because that's the font I'm gonna use
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    #composing the document
    d.add_heading('Chat con Potato', 0)
    p = d.add_paragraph()
    p.style = d.styles['Arial']
    p.add_run('Giorgia').bold = True
    p.add_run('\t\t\t\t\t\t\t\t\t\t')
    p.add_run('Mike').bold = True

    file = csv.reader(open('chat.csv', 'r'))

    i = 0
    for line in tqdm(file, desc="Parsing messages..."):
        if line[3] == " Mike":
            p = d.add_paragraph()
            p.paragraph_format.line_spacing = 0.5
            p.style = d.styles['Arial']
            p.alignment = 2
            p.add_run(line[2])
        elif line[3] == ' GiorgiaChiarucci':
            p = d.add_paragraph()
            p.paragraph_format.line_spacing = 0.5
            p.style = d.styles['Arial']
            p.alignment = 0
            p.add_run(line[2])
        i = i+1
        if i == 1000:
            break

    d.save('chat.docx')

if not os.path.isfile('./chat.csv'):
    if os.path.isfile('./chat.txt'):
        print('CSV file not found, converting chat.txt in chat.csv..')
        _convert_to_csv(filename)
    else:
        print('Cannot find chat.txt nor chat.csv in folder, please add the chat source')
else:
    print('\nAll ok, creating document...\n')
    _create_document()