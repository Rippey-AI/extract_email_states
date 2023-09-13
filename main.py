import requests

import json
import pandas as pd
from bs4 import BeautifulSoup
import re

def read_request():
    try:
        with open("emails_request.json", "r") as request:
            content_json = json.load(request)
    except FileNotFoundError:
        print("File not found.")
    except Exception as e:
        print("An error occurred:", e)
    return content_json
request_json = read_request()


url = request_json["url"]
queries = request_json["queries"]
headers = request_json["headers"]

def get_html_content(url):
    response = requests.get(url)
    if response.status_code == 200:
        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')
        # Extract and print the text from the HTML
        extracted_text = soup.get_text()
        refine_text = re.sub(r'\n+', '\n', extracted_text)
        return refine_text
    else:
        print(f"Failed to fetch HTML content. Status code: {response.status_code}")


def save_result(filename, output):
    with open(filename, "w") as file:
        json.dump(output, file , indent=4)


emails_details = []
except_emails = ["susan@rippey.ai","bhuwan@rippey.ai"]
def send_request(url, headers, query):
    response = requests.get(url=url, headers=headers, params=query).json()
    save_result("response.json", response)
    for emails in response["data"]["results"]:
        is_error = emails["error_info"]["error_info"]["value"]
        if is_error == "ERROR":
            if emails["sender_email_address"] not in except_emails:
                instance_id = emails["instance_id"]
                results = get_emails_state(instance_id)
                quote_type = results[2]["value"]
                model_extraction = results[3]["value"]
                mapping = results[4]["value"]
                generate_quotes = results[5]["value"]
                email_extraction = results[0]["value"]["email_file_url"]
                email_content = get_html_content(email_extraction)

                emails_details.append(
                    {
                        "Date": emails["email_date"],
                        "Subject": emails["email_subject"],
                        "Sender Name": emails["sender_user_name"],
                        "quote_type": json.dumps(quote_type,indent= 4),
                        "model_extraction": json.dumps(model_extraction, indent=4),
                        "mapping": json.dumps(mapping, indent=4),
                        "generate_quotes": json.dumps(generate_quotes, indent=4),

                        "email_content": email_content,
                    }
                )



def get_emails_state(instance_id):
    url = f"https://appserver-preprod.rpalabsinternal.com/analytics/instances/{instance_id}/states"
    state_response = requests.get(url=url,headers=headers).json()
    results = state_response["data"]["results"]
    return results


def to_excel(data):
    dataframe = pd.DataFrame(data)
    dataframe.columns= ["Date","Subject","Sender Name","quote_type","model_extraction","mapping","generate_quotes","email_content"]
    output_file = "result.xlsx"
    dataframe.to_excel(output_file, index=False)
    print("Done.................")

send_request(url=url,headers=headers,query=queries)
to_excel(emails_details)