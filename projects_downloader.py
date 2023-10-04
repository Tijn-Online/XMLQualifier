import requests, json, datetime, re

def authenticate():
    global token
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
    token = token_string.decode('UTF-8')
    return token_string.decode('UTF-8')


def fetch_all_id(paging_size):
    global token
    reconnect_counter = 0
    url = "https://api.rydoo.com/v2/projects/count?isActive=true"
    headers = {
        "Accept": "application/json",
        "Authorization": "Bearer %s" %token
    }
    response = requests.get(url, headers=headers)
    json_obj = json.loads(response.text)
    inner_obj = json_obj["data"]
    limit = inner_obj['count']
    offset_range = int(limit/paging_size) + int(1)
    for ix in range(offset_range):

        url_limit = str(paging_size)
        url_offset = str((ix + 1) * paging_size)
        url = "https://api.rydoo.com/v2/projects?isActive=true&limit=" + url_limit + "&offset=" + url_offset
        headers = {
            "Accept": "application/json",
            "Authorization": "Bearer %s" % token
        }
        response = requests.get(url, headers=headers)
        json_obj = json.loads(response.text)
        inner_obj = json_obj["data"]
        for i in enumerate(inner_obj):
            project_dict = i[1]

            print("processing row/total: " + str(ix*paging_size + i[0]) + "/" + str(limit) + " updated: " + str(reconnect_counter) )
            isactive = project_dict['isActive']
            unique_id = project_dict['id']
            unique_name = str(project_dict['name'].encode('utf-8'))
            unique_name = re.sub('[^0-9a-zA-Z]+', ' ', unique_name)
            if isactive:
                disable_project(unique_id,token,unique_name)
                reconnect_counter += 1
                if not reconnect_counter % 200:
                    token = authenticate()



def disable_project(guid,token, name):
    url = "https://api.rydoo.com/v2/projects/" + guid
    current_datetime = datetime.datetime.now()
    payload = "{\"name\":\"Disabled %s - %s\",\"isActive\":false}" % (current_datetime.strftime("%Y-%m-%d %H:%M:%S"),name)
    headers = {
        "accept": "application/json",
        "Authorization": "Bearer %s" % token,
        "content-type": "application/*+json"
    }
    response = requests.put(url, data=payload, headers=headers)
    print(response.text)



if __name__ == '__main__':
    token = authenticate()
    fetch_all_id(paging_size=1)
    #disable_project("bd733df6-adcc-404c-8301-7d54f0b86dd3",token, "Business transformation (PJRB2304718563)")




