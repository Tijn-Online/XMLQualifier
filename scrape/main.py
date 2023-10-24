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

def fetch_all_id(token,paging):
    return_l = []
    for i in range(1000000):
        limit = paging
        offset = i*paging
        url = "https://api.rydoo.com/v2/users/?search=&email=&firstName=&lastName=&limit=%s&offset=%s" % (limit, offset)
        headers = {
            "Accept": "application/json",
            "Authorization": "Bearer %s" % token
        }
        response = requests.get(url, headers=headers)
        json_obj = json.loads(response.text)
        inner_obj = json_obj["data"]
        if bool(inner_obj):
            for j in enumerate(inner_obj):
                user_dict = j[1]
                return_l.append(user_dict['id'])
            print("fetched " + str(len(return_l)) + " users to process")
        else:
            break
    return return_l

def download_user(id_list,token):
    with open(filename1,'w') as f:
        for user_id in enumerate(id_list):
            print("download user for " + str(user_id))
            if(not user_id[0] % 150): token = authenticate()
            url_user = "https://api.rydoo.com/v2/users/%s" %user_id[1]
            url_groupassignment = "https://api.rydoo.com/v2/users/%s/groups" %user_id[1]

            headers = {
                "Accept": "application/json",
                "Authorization": "Bearer %s" %token
            }
            response_user = requests.get(url_user, headers=headers)
            response_groupid = requests.get(url_groupassignment, headers=headers)
            try:
                # process basic data
                json_obj = json.loads(response_user.text)
                inner_obj = json_obj["data"]
                write_id = inner_obj['id']
                write_enabled = inner_obj['enabled']
                write_email = inner_obj['email']
                write_businessUnitId = inner_obj['businessUnitId']
                write_refId = inner_obj['refId']
                enter_BU = inner_obj['customFields'][0]
                write_BU = enter_BU['valueCode']
                enter_CF = inner_obj['customFields'][1]
                write_CF = enter_CF['valueCode']

                # process company and CC
                json_obj = json.loads(response_groupid.text)
                cclist = []
                s = "#"
                for i,groups in enumerate(json_obj["data"]):
                    outer_company = groups['group']
                    write_cc = outer_company['refId']
                    startdate = groups['startDate']
                    if startdate is None: startdate = "2022-01-01"
                    enddate = groups['endDate']
                    if enddate is None: enddate = "2099-01-01"
                    startdate_dateobject = datetime.datetime.strptime(startdate, '%Y-%m-%d').date()
                    enddate_dateobject = datetime.datetime.strptime(enddate, '%Y-%m-%d').date()

                    if startdate_dateobject <= datetime.datetime.now().date() <= enddate_dateobject:
                        cclist.append(write_cc)
                        write_company = outer_company['branch']['refId']
                    else:
                        pass
                write_cc = s.join(cclist)
                collection = write_id, write_enabled, write_email, write_businessUnitId, write_refId, write_BU, write_CF, write_cc, write_company
                f.write(str(collection))
                f.write('\n')
            except:
                print("Fetch user details exception occurred at " + str(user_id))

def download_approvers(id_list,token):
    with open(filename2, 'w') as f:
        for user_id in enumerate(id_list):
            print("download approver for " + str(user_id))
            if (not user_id[0] % 300): token = authenticate()
            url_approver = "https://api.rydoo.com/v2/users/%s/expenseapprovers"  %user_id[1]
            headers = {
                "Accept": "application/json",
                "Authorization": "Bearer %s" % token
            }
            response_approver = requests.get(url_approver, headers=headers)
            try:
                json_obj = json.loads(response_approver.text)
                inner_obj = json_obj["data"]
                inner_obj = inner_obj["approvers"]
                for key in enumerate(inner_obj):
                    approverlist = key[1]
                    collection = (user_id[1],approverlist["id"])
                    f.write(str(collection))
                    f.write('\n')
            except:
                collection = (user_id[1], "No approver found")
                f.write(str(collection))
                f.write('\n')

def quick_test(userid):
    token = authenticate()
    url = "https://api.rydoo.com/v2/users/%s" % userid
    headers = {"Accept": "application/json","Authorization": "Bearer %s" % token }
    response = requests.get(url, headers=headers)
    json_obj = json.loads(response.text)
    print(json_obj["data"])
    url = "https://api.rydoo.com/v2/users/%s/groups" % userid
    response = requests.get(url, headers=headers)
    json_obj = json.loads(response.text)
    cclist = []
    s = "#"
    for i,groups in enumerate(json_obj["data"]):
        outer_company = groups['group']
        write_cc = outer_company['refId']
        startdate = groups['startDate']
        if startdate is None: startdate = "2022-01-01"
        enddate = groups['endDate']
        if enddate is None: enddate = "2099-01-01"
        startdate_dateobject = datetime.datetime.strptime(startdate, '%Y-%m-%d').date()
        enddate_dateobject = datetime.datetime.strptime(enddate, '%Y-%m-%d').date()
        if startdate_dateobject <= datetime.datetime.now().date() <= enddate_dateobject:
            cclist.append(write_cc)
            write_company = outer_company['branch']['refId']
        else:
            pass
    print(s.join(cclist),write_company)


    url_approver = "https://api.rydoo.com/v2/users/%s/expenseapprovers" % userid
    headers = {
        "Accept": "application/json",
        "Authorization": "Bearer %s" % token
    }
    response_approver = requests.get(url_approver, headers=headers)
    print(response_approver.text)

if __name__ == '__main__':
    filename1 = r'C:\Users\mahoo8\OneDrive - Inter IKEA Group\Rydoo_PBI_Workfolder\rydoo_users_' + datetime.datetime.now().strftime(
        "%Y%m%d-%H%M%S") + '.txt'
    filename2 = r'C:\Users\mahoo8\OneDrive - Inter IKEA Group\Rydoo_PBI_Workfolder\rydoo_approvers_' + datetime.datetime.now().strftime(
        "%Y%m%d-%H%M%S") + '.txt'
    token = authenticate()
    id_list = fetch_all_id(token,100)
    download_user(id_list,token)
    download_approvers(id_list,token)
    #quick_test('59c53729-677a-4ab5-b850-1e677ab7cad6')
