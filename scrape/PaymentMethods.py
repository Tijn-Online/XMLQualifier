import requests
import json

class RydooAPI:
    def __init__(self, client_id, client_secret):
        self.client_id = client_id
        self.client_secret = client_secret
        self.token = self.authenticate()

    def authenticate(self):
        url = "https://accounts.rydoo.com/connect/token"
        payload = f"grant_type=client_credentials&scope=company%20structure%3Aread%20structure%3Awrite%20expenses%3Aread%20expenses%3Awrite%20fields%3Aread%20fields%3Awrite%20users%3Aread%20users%3Awrite&client_id={self.client_id}&client_secret={self.client_secret}"
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/x-www-form-urlencoded"
        }
        response = requests.post(url, data=payload, headers=headers)
        json_obj = json.loads(response.text)
        token_string = json_obj["access_token"].encode("UTF-8", "ignore")
        return token_string.decode('UTF-8')

    def count_cards(self):
        url = "https://api.rydoo.com/v2/paymentmethods/all/count"
        headers = {
            "Accept": "application/json",
            "Authorization": f"Bearer {self.token}"
        }
        params = {
            "isActive": True,
            "validationStatus": "Validated",
            "accountType": "LodgeCard"
        }
        response = requests.get(url, params=params, headers=headers)
        response_data = response.json()
        count = response_data['data']['count']
        return count

    def payment_methods(self):
        url = "https://api.rydoo.com/v2/paymentmethods/all"
        limit = 25  # Batch size
        count = self.count_cards()
        payment_info_list = []  # List to store dictionaries with payment info

        headers = {
            "Accept": "application/json",
            "Authorization": f"Bearer {self.token}"
        }

        allowed_card_numbers = ["122089xxxx64266", "122089xxxx07908", "122089xxxx92937", "122089xxxx64266",
                                "122089xxxx69758"]

        for offset in range(0, count, limit):
            params = {
                "limit": limit,
                "offset": offset,
                "isActive": True,
                "validationStatus": "Validated",
                "accountType": "LodgeCard"
            }
            response = requests.get(url, params=params, headers=headers)

            if response.status_code == 200:
                response_data = response.json()

                payment_methods = response_data["data"]

                for method in payment_methods:
                    card_number = method["cardNumber"]
                    if card_number in allowed_card_numbers:
                        payment_info = {
                            "id": method["id"],
                            "cardName": method["name"],
                            "userId": method["user"]["id"]
                        }
                        payment_info_list.append(payment_info)
            else:
                print("Error:", response.status_code)

        return payment_info_list

    def update_payment_method_status(self, payment_method_id, payment_method_cardname, payment_method_UserID, is_active):
        url = f"https://api.rydoo.com/v2/paymentmethods/{payment_method_id}"
        headers = {
            "Accept": "application/json",
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }

        payload = {
            "user": {
                "id": payment_method_UserID
            },
            "paymentMethodType": "LodgeCard",
            "accountOwnership": "Company",
            "accountNumber": "0",
            "travellerId": "0",
            "cardNumber": "0",
            "isDefault": "true",
            "isAutoCreateExpenseEnabled": "false",
            "cardName": payment_method_cardname,
            "isActive": is_active
        }

        response = requests.put(url, json=payload, headers=headers)
        print(response.text)
        if response.status_code == 200:
            print("Payment method status updated successfully.")
        else:
            print(f"Error updating payment method status: {response.status_code}")


if __name__ == '__main__':
    client_id = "eoBdEYIBzfzgJTnWstfl2yM1rEKjPWM"
    client_secret = "ew6f7vtzyaBEYpIe3N2wuGxbfb"

    rydoo_api = RydooAPI(client_id, client_secret)
    ids = rydoo_api.payment_methods()

    for payment_info in ids:
        print(payment_info)
        payment_method_id = payment_info["id"]
        payment_method_cardname = payment_info["cardName"]
        payment_method_UserID = payment_info["userId"]
        rydoo_api.update_payment_method_status(payment_method_id, payment_method_cardname, payment_method_UserID, is_active=False)

