# coding: utf-8

from PyPDF2 import PdfFileReader
import re
import xlwt
import os
import datetime
from calendar import monthrange

from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfpage import PDFPage
import io

date_regex_format   = '\d{1,2}\.\ (\S+)\ (\d{4})'
date_regex_format_2 = '\d{1,2}\ (\S+)\ (\d{4})' #alternative format

def find_start_date(arg_full_text):
    danish_month_list = ['januar', 'februar', 'marts', 'april', 'maj', 'juni',
                         'juli', 'august', 'september', 'oktober', 'november', 'december']

    date_matches = []
    # Find list of all entries in the text that matches the general date format (two groups: mounth and year)
    date_matches = re.findall(date_regex_format, arg_full_text , flags=(re.IGNORECASE | re.MULTILINE))
    # There should be a check that the found matches are the same
    
    if not date_matches: #no match - try alternative format
        date_matches = re.findall(date_regex_format_2, arg_full_text , flags=(re.IGNORECASE | re.MULTILINE))
        
    if not date_matches: #no match
        print("Error: No start date found")
        exit(-1)

    first_date_match = date_matches[0] # (mounth, year)
    month_match_index = 0
    #print(date_matches)
    for month_index, danish_month in enumerate(danish_month_list):
        if re.match(danish_month, first_date_match[0], flags=(re.IGNORECASE | re.MULTILINE)):
            month_match_index = month_index

    start_date = datetime.date(int(first_date_match[1]), month_match_index+1, 1)

    print("Start date: " + start_date.strftime('%A d. %d/%m-%y'))

    return start_date



file_list = [os.path.join("input_pdfs", "Wod_April_2019.pdf"),
             os.path.join("input_pdfs", "Stort_Hold_April_2019.pdf"),
             os.path.join("input_pdfs", "Øvet_April_2019.pdf")
             ]


wod_seperator = "--------------------------------"
excel_base_filename = "WODs"
excel_style = xlwt.XFStyle()
excel_style.alignment.wrap = 1
excel_column_width = 10000

book = xlwt.Workbook()
wod_sheet = book.add_sheet("WODs")
start_date = 0

for file_index, file_name in enumerate(file_list):

    print("Processing file number %d, named: %s" %(file_index, file_name))

    full_text = ""  # Reset the collected text
    with open(file_name, 'rb') as f:
        pdf = PdfFileReader(f)
        number_of_pages = pdf.getNumPages()

        # Print pages:
        page_range = range(number_of_pages)
        for page_num in page_range:
            page = pdf.getPage(page_num)
            # print("Page info: ", page)
            # print('Page type: {}'.format(str(type(page))))
            page_text = page.extractText()

            #Remove unwanted text
            page_text = re.sub('Uklarheder.+Fang mig p.+$',          '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('.*@crossfitcopenhagen\.dk.*',        '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('.*@crossfys\.dk.*',                  '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('.*kalender.*',                       '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('.*calendar.*',                       '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('^Dato$',                             '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('^Workout Of the Day$',               '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('^StortHold.+',                       '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('^.vetHold.+',                        '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('^.vet WOD.+',                        '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('^Program',                           '', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('^\ *\d+ af \d+',                     '', page_text, flags=(re.IGNORECASE | re.MULTILINE)) #f.eks. 3 af 7
           
            
            #Correct danish charecters
            page_text = re.sub('¾',                                  'æ', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('¿',                                  'ø', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            page_text = re.sub('„',                                  'å', page_text, flags=(re.IGNORECASE | re.MULTILINE))
            

            page_text = str.strip(page_text)

            full_text = full_text + page_text + '\n'

    start_date = find_start_date(full_text)

    #Split into WOD list
    text_with_seperaters = re.sub('.*' + date_regex_format,     wod_seperator, full_text, flags=(re.IGNORECASE | re.MULTILINE))
    text_with_seperaters = re.sub('.*' + date_regex_format_2,   wod_seperator, text_with_seperaters, flags=(re.IGNORECASE | re.MULTILINE))
    text_with_seperaters = re.sub('Forkortelser',               wod_seperator, text_with_seperaters, flags=(re.IGNORECASE | re.MULTILINE))
    text_with_seperaters = re.sub('Forkortelser|Abbreviations', wod_seperator, text_with_seperaters, flags=(re.IGNORECASE | re.MULTILINE))
    
    wod_array = re.split(wod_seperator, text_with_seperaters)
    del wod_array[0]  # Because the PDF starts with a date, the first entry is empty

    column_index = file_index + 1  # First column is the dates
    wod_sheet.col(column_index).width = excel_column_width
    for day_index, wod_entry in enumerate(wod_array):
        wod_entry = re.sub('(?P<orig>^-+.*(rest|pause).*-\ *$)', '\n\g<orig>\n', wod_entry, flags=(re.IGNORECASE | re.MULTILINE)) #Add newlines before ---- rest ---
        wod_entry = re.sub('(?P<orig>(?<!\n\n)^[ABCDE]\d*\.)', '\n\g<orig>', wod_entry, flags=(re.MULTILINE)) #Add newlines for ie. A. B. C. - but not if there already is a double newline
        
        wod_entry = str.strip(wod_entry)
        wod_sheet.write(day_index, column_index, wod_entry, excel_style)
        print("Writing the WOD to row %d column %d" %(day_index, column_index))

# Write time column
wod_sheet.col(0).width = int(excel_column_width/2)
number_of_days_in_months = monthrange(start_date.year,start_date.month)[1]

day_range = range(number_of_days_in_months)
for day_number in day_range:
    date = start_date + datetime.timedelta(days=day_number)
    wod_sheet.write(day_number, 0, date.strftime('%A d. %d/%m'), excel_style)

book.save(excel_base_filename + '_' + start_date.strftime('%B_%Y') + ".xls")
        
print("Finished")
