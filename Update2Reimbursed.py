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

def update(expense):
    url = "https://api.rydoo.com/v2/expenses/reimburse"
    payload = "[{\"expenseId\":\"3901831931\",\"paymentReference\":\"reference\"}]"
    headers = {
        "accept": "application/json",
        "content-type": "application/*+json"
    }
    response = requests.put(url, data=payload, headers=headers)
    print(response.text)