import requests, json, datetime

def authenticate():
    url = "https://accounts.rydoo.com/connect/token"
    payload = "grant_type=client_credentials&scope=company%20structure%3Aread%20structure%3Awrite%20expenses%3Aread%20expenses%3Awrite%20fields%3Aread%20fields%3Awrite%20users%3Aread%20users%3Awrite&client_id=eoBdEYIBzfzgJTnWstfl2yM1rEKjPWM&client_secret=ew6f7vtzyaBEYpIe3N2wuGxbfb"
    headers = {
        "Accept": "application/json",
        "Content-Type": "application/x-www-form-urlencoded"
    }
    response = requests.post(url, data=payload, headers=headers)
    print(response.text)
    json_obj = json.loads(response.text)
    token_string = json_obj["access_token"].encode("UTF-8", "ignore")
    return token_string.decode('UTF-8')

def get_doc_count():
    token = authenticate()
    url = "https://api.rydoo.com/v2/expenses/exported"
    headers = {
        "Accept": "application/json",
        "Authorization": "Bearer %s" % token
    }
    response = requests.get(url, headers=headers)
    json_obj = json.loads(response.text)
    return json_obj["totalCount"]

def fetch_all_exp(paging_size,limit):
    token = authenticate()
    filename2 = r'C:\Users\mahoo8\OneDrive - Inter IKEA Group\Rydoo_PBI_Workfolder\rydoo_expenses_' + datetime.datetime.now().strftime(
        "%Y%m%d-%H%M%S") + '.txt'
    with open(filename2, 'w', encoding="utf-8") as f:
        offset_range = int(limit/paging_size) + int(1)
        for ix in range(offset_range):
            url_limit = str(paging_size)
            url_offset = str((ix + 1) * paging_size)
            url = "https://api.rydoo.com/v2/expenses/exported?limit=" + url_limit + "&offset=" + url_offset
            print("calling: " +url)
            headers = {
                "Accept": "application/json",
                "Authorization": "Bearer %s" % token
            }
            response = requests.get(url, headers=headers)
            json_obj = json.loads(response.text)

            inner_obj = json_obj["data"]
            print("parsing into file.. ")
            for i in enumerate(inner_obj):
                project_dict = i[1]
                collection = ("|".join([str(e) for e in project_dict.values()]))
                f.write(str(collection))
                f.write('\n')

if __name__ == '__main__':
    doc_count = get_doc_count()
    print(doc_count)
    fetch_all_exp(paging_size=100,limit=doc_count)




