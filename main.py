from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.listitems.listitem import ListItem
from bs4 import BeautifulSoup
import pandas as pd
import xlsxwriter
import matplotlib.pyplot as plt
import openpyxl
from credentials import USERNAME,PASSWORD,SITE,SOURCE_DIR
from pathlib import Path
import xlwings as xw
import os
import smtplib , ssl
from email.message import EmailMessage
import time


username = USERNAME
password = PASSWORD

site_url = SITE

def get_user_input():
    
    query_dictionary = dict()
    status_input = "yes"
    status = True
    while status:
        
        if (status_input == "No" or status_input == "no"):
            status = False
        
        else:
            user_input = input("Please Enter the term against which you want to query the tickets : ")
            query_dictionary[user_input]=[]
            status_input = input("Do you want to continue ? Yes or No : ")
    
    return (query_dictionary)


def query_sharepoint():
    
    list_data = []
    ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))
    target_list = ctx.web.lists.get_by_title("IT Tickets")
    paged_items = target_list.items.paged(50).get().execute_query()
    
    for index,item in enumerate(paged_items):
        get_body = item.properties["ProblemStatement"]
        
        if get_body is not None:
            
            soup = BeautifulSoup(get_body)
            text_out  = soup.get_text()
            list_data.append(text_out)
        
        
    return list_data


def organize_data(query_dictionary,list_data):
    
    for key,value in query_dictionary.items():
        for data in list_data:
            if key == None:
                pass
            if key in data:
                value.append(data)
            
    
    return query_dictionary

def generate_array(query_dictionary):
    
    issue_array = ["Issues"]
    Issue_frequeny = ["No of issues"]
    for key,value in query_dictionary.items():
        issue_array.append(key)
        Issue_frequeny.append(len(query_dictionary[key]))
    
    return (issue_array,Issue_frequeny)
    

def generate_excel(query_dictionary,issue_array,Issue_frequeny):
    
    workbook = xlsxwriter.Workbook("Tickets.xlsx")
    worksheet = workbook.add_worksheet()
    
    array = [issue_array,
             Issue_frequeny]
    
    row  = 0
    
    for column, data in enumerate(array):
        worksheet.write_column(row,column,data)
    
    workbook.close()
    
    x_axis = issue_array[1:]
    
    description_array = []
    for key,values in query_dictionary.items():
        description_array.append(values)
        
    df = pd.DataFrame(description_array).T
    df.to_excel("Description.xlsx")
    y_axis = Issue_frequeny[1:]
    
    plt.figure(1)
    plt.subplot(211)
    plt.bar(x_axis, y_axis, color ='maroon',width = 0.4) 
    plt.xlabel("Reported Issues")
    plt.ylabel("No of issues ")
    plt.title("IT Tickets")
    plt.savefig('bar_plot.png',transparent = True, bbox_inches = 'tight', pad_inches = 0)
    plt.show()
    
    plt.figure(2)
    plt.subplot(212)
    plt.pie(y_axis,labels=x_axis)
    plt.savefig("pie.png",transparent = True, bbox_inches = 'tight', pad_inches = 0)
    plt.show()


    wb = openpyxl.load_workbook("Tickets.xlsx")
    ws = wb.active
    
    img = openpyxl.drawing.image.Image("bar_plot.png")
    img.anchor = "A8"
    
    img2 = openpyxl.drawing.image.Image("pie.png")
    img2.anchor = "H8"
    
    ws.add_image(img)
    ws.add_image(img2)
    wb.save("Tickets.xlsx")
    
def merge_excel():
    
    excel_files = Path(SOURCE_DIR).glob('*.xlsx')
    with xw.App(visible=False) as app:
        combined_wb = app.books.add()
        for excel_file in excel_files:
            wb = app.books.open(excel_file)
            for sheet in wb.sheets:
                sheet.copy(after=combined_wb.sheets[0])
            wb.close()
            
        combined_wb.sheets[0].delete()
        combined_wb.save(f"Issues_Report.xlsx")
        combined_wb.close()

def send_mail_with_excel(recipient_email, subject,content, excel_file):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = USERNAME
    msg['To'] = recipient_email
    msg.set_content(content)

    with open(excel_file, 'rb') as f:
        file_data = f.read()
    msg.add_attachment(file_data, maintype="application", subtype="xlsx", filename=excel_file)

    with smtplib.SMTP('smtp.office365.com', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(USERNAME, PASSWORD)
        smtp.send_message(msg)

def main():
        
    dictionary = get_user_input()
    
    list_dat = query_sharepoint()
    
    test_dat = organize_data(dictionary,list_dat)
    
    x,y = generate_array(dictionary)
    
    generate_excel(dictionary,x,y)
    
    merge_excel()
    
    subject = "IT Tickets Report"
    excel_file = "Issues_Report.xlsx"
    content = "PFA the attached IT Tickets Report"
    send_mail_with_excel("recepient_email", subject,content, excel_file)
    time.sleep(10)
    os.remove("Issues_Report.xlsx")

main()
    
