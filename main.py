from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.listitems.listitem import ListItem
from bs4 import BeautifulSoup


username = "email"
password = "password"

site_url = "Site"

def get_user_input():
    
    query_dictionary = dict()
    query_array = []
    status_input = "yes"
    status = True
    while status:
        
        if (status_input == "No" or status_input == "no"):
            status = False
        
        else:
            user_input = input("Please Enter the term against which you want to query the tickets : ")
            query_dictionary[user_input]=query_array
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
    
    for key,value in enumerate(query_dictionary.items()):
        for data in list_data:
            if key == None:
                pass
            if key in data:
                print(key)
                value.append(data)
            
    
    return query_dictionary

dictionary = get_user_input()

list_dat = query_sharepoint()

test_dat = organize_data(dictionary,list_dat)
