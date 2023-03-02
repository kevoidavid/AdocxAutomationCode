# AdocxAutomationCode
this code iterates through docx files with shapes and textboxes  in a folder, finds a text and replaces it withouth changing the format
# Import the necessary libraries
from docx.shared import Inches
from pathlib import Path  # core python module
from pathlib import Path  # core python module
import win32com.client  # pip install pywin32
import os
import random
YR = 2022

# get_inputs
name = input('enter Place.. ')
Main = name.upper()
DOB = input('enter D.O.B.. ')
SL = input('enter sl... ')
county = input('enter  name... ')

# some data
KISII = ['KISII','NYANZA', 'KSI',  'Z.B MACHOME','016',' S.K MELI10 ']
KWALE = ['KWALE', 'COAST','KWL', 'D.MATHEI','018',' ']
MIGORI = ['MIGORI','NYANZA', 'MGR','J.M.OLOL.TULET','078','3']
NAKURU = ['NAKURU','RIFT VALLEY','NKU','S.K MELI','027','DAVE MUTAI']
NAIROBI = ['NAIROBI','NAIROBI','NRB','JOYCE CHEPKEITANY','026','E.A ATITO']
MOMBASA = ['MOMBASA','COAST','MSA','L.MASHARL','025','J ONYISI']
KISUMU = ['KISUMU','NYANZA','KSM','A.O OMONDI','042','W.S ISAGI']
NYERI = ['NYERI','CENTRAL','NYR','J.N NDATHO','031','S.M KABURU']
GUCHA = ['GUCHA','NYANZA','GCA','P.K OTIENO','055','H.D KOLA11']
NYAMIRA = ['NYAMIRA','NYANZA','NYAM','J.M GORI','076','1']
BARINGO = ['BARINGO', 'RIFT VALLEY','BGO','M.A MUSHIYI','042']
BOMET = ['BOMET','RIFT VALLEY', 'BMT','R.K KORIR','081']
BUNGOMA = ['BUNGOMA','WESTERN','BGM','H.ORENG','002']
BUSIA = ['BUSIA','WESTERN','BSA','J.M.OLOL.TULET','003']
HOMABAY = ['HOMA BAY','NYANZA','HB','V A ONYANGO','033']
KAKAMEGA = ['KAKAMEGA','WESTERN','KKG','TITUS.M.MAINA','011','R.K.WESONGA']
KALOLENI = ['KALOLENI','COAST','KLN','M.N MUTAVI','091']
KIRINYAGA = ['KIRINYAGA','CENTRAL','KRG','E.M MUGAMBI','015']
KITUI = ['KITUI','EASTERN','KTI','H.T MWANIKI','017']
KWANZA = ['KWANZA','COAST','KWZ','ERICK.K.SHIVANDA','142']
LAMU = ['LAMU','COAST','LMU','C.S BUYA','020']
MURANGA = ['MURANGA EAST','CENTRAL', 'MRG',  'J.N MWANGI','007',' S.K MELI10 ']
KILIFI = ['KILIFI','COAST', 'KLF',  'K.S MOHAMMED','014',' S.K MELI10 ']


# my fomula
llist2 = [KILIFI,MURANGA,KISII, KWALE, MIGORI,MOMBASA, NAKURU, LAMU, KWANZA, KITUI,KIRINYAGA,KALOLENI, KAKAMEGA, HOMABAY, BUSIA, BUNGOMA, BOMET, BARINGO, NYAMIRA, GUCHA, NYERI, KISUMU,NAIROBI]
llist1 = ['KILIFI', "MURANGA", 'KISII', 'KWALE','MIGORI','MOMBASA', "NAKURU","LAMU", 'KWANZA', "KITUI", "KIRINYAGA", 'KALOLENI', 'KAKAMEGA', "HOMABAY", "BUSIA", "BUNGOMA","BOMET", "BARINGO", "NYAMIRA", "GUCHA", 'NYERI', "KISUMU", "NAIROBI"]
x = (llist1.index(Main))
T = (llist2[x])


mt = (9, 10, 11)
AUTH = random.randint(1055, 1307)
CA = random.randint(1344, 1900)
dt = random.randint(3, 31)
MT = random.choice(mt)
ET = random.randint(5700, 7985)
ET1 = random.randint(59000, 80099)
ET2 = (ET, ET1)
ET3 = random.choice(ET2)
ENTRY = (f'L.{T[4]}22{ET1}')
ENTRY2 = (F'L.{T[4]}0{ET}/22')
ENTRY3 = (ENTRY, ENTRY2)
ENTRY4 = random.choice(ENTRY3)

DT = (f'{dt}/{MT}/{YR}')

if MT == 9:
    is_true = True
else:
    is_true = False

if is_true:
    z = 'September'
elif (MT == 10) :
    z = 'October'
else :
    z = 'November'

RO = (f'{dt} {z}, {YR}')


print(' ')
print(T[0] + ' ' + T[1] +' ' + T[2] + ' ' + T[3])
print(' ')
print('Registered on ', RO, 'ENTRY NO ', ENTRY4)
print(' ')
bb =(f'AUTH NO.{AUTH}/{T[2]} CA OF {DT}')
cc = (f'{T[2]}/CA NO.{CA} OF {DT}')



# Path settings
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_dir = current_dir / "input"
output_dir = current_dir / "output"
output_dir.mkdir(parents=True, exist_ok=True)


# Find & replace
find_strA = "B"
replace_withA = T[1]

find_str = "A"
replace_with = T[0]

find_str1 = "1"
replace_with1 = ENTRY4

find_str2 = "2"
replace_with2 = f"{SL}"

find_str3 = "3"
replace_with3 = f'{county}'

find_str4 = "4"
replace_with4 = f"{DOB}"

find_str5 = "5"
replace_with5 = "FEMALE"

find_str6 = "6"
replace_with6 = "FATHER"

find_str7 = "7"
replace_with7 = "MOTHER"

find_str8 = "8"
replace_with8 = "Sgd; Self"

find_str9 = "9"
replace_with9 = f"{T[3]}"

find_str10 = "0"
replace_with10 = f"{T[3]}"

find_str11 = "2022"
replace_with11 = f"{T[0]}"

find_str12 = "2027"
replace_with12 = f"{RO}"

find_str13 = "2023"
replace_with13 = f"{bb}\n{cc}"

find_strB = "2024"
replace_withB = f"{dt}"

find_strC = "2025"
replace_withC = f"{MT}"

find_strD = "2026"
replace_withD = f"{YR}"


wd_replace = 2  # 2=replace all occurences, 1=replace one occurence, 0=replace no occurences
wd_find_wrap = 1  # 2=ask to continue, 1=continue search, 0=end if search range is reached


# Open Word
word_app = win32com.client.DispatchEx("Word.Application")
word_app.Visible = True
word_app.DisplayAlerts = False


for doc_file in Path(input_dir).rglob("*.doc*"):
    # Open each document and replace string
    word_app.Documents.Open(str(doc_file))
    # API documentation: https://learn.microsoft.com/en-us/office/vba/api/word.find.execute
    word_app.Selection.Find.Execute(
        FindText=find_str1,
        ReplaceWith=replace_with1,
        Replace=wd_replace,
        Forward=True,
        MatchCase=True,
        MatchWholeWord=False,
        MatchWildcards=True,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Wrap=wd_find_wrap,
        Format=True,
    )

    word_app.Selection.Find.Execute(
        FindText=find_str2,
        ReplaceWith=replace_with2,
        Replace=wd_replace,
        Forward=True,
        MatchCase=True,
        MatchWholeWord=False,
        MatchWildcards=True,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Wrap=wd_find_wrap,
        Format=True,
    )

    word_app.Selection.Find.Execute(
        FindText=find_str3,
        ReplaceWith=replace_with3,
        Replace=wd_replace,
        Forward=True,
        MatchCase=True,
        MatchWholeWord=False,
        MatchWildcards=True,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Wrap=wd_find_wrap,
        Format=True,
    )

    word_app.Selection.Find.Execute(
        FindText=find_str4,
        ReplaceWith=replace_with4,
        Replace=wd_replace,
        Forward=True,
        MatchCase=True,
        MatchWholeWord=False,
        MatchWildcards=True,
        MatchSoundsLike=False,
        MatchAllWordForms=False,
        Wrap=wd_find_wrap,
        Format=True,
    )
    # -- Replace str in shapes
    # VBA SO reference: https://stackoverflow.com/a/26266598
    # Loop through all the shapes
    for i in range(word_app.ActiveDocument.Shapes.Count):
        if word_app.ActiveDocument.Shapes(i + 1).TextFrame.HasText:
            words = word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words
            # Loop through each word. This method preserves formatting.
            for j in range(words.Count):
                # If a word exists, replace the text of it, but keep the formatting.
                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str1:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with1

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str2:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with2

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str3:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with3

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str4:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with4

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str5:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with5

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str6:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with6

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str7:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with7

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str8:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with8

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str9:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with9

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str10:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with10

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str11:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with11

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str12:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with12

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str13:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with13

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_str:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_with

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_strA:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_withA

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_strB:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_withB

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_strC:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_withC

                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text == find_strD:
                    word_app.ActiveDocument.Shapes(i + 1).TextFrame.TextRange.Words.Item(j + 1).Text = replace_withD


    # Save the new file
    output_path = output_dir / f"{doc_file.stem}_replaced{doc_file.suffix}"
    word_app.ActiveDocument.SaveAs(str(output_path))
    word_app.ActiveDocument.Close(SaveChanges=False)
word_app.Application.Quit()



