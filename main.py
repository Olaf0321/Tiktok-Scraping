import datetime
import requests
import json
import pandas as pd
import pandas as pd
import os
import pytz
import time
import random

# Define the Excel file name
excel_filename = "processed_data.xlsx"
sheet_name = "Sheet1"

def convert_to_japanese_time(unix_timestamp):
    japan_tz = pytz.timezone('Asia/Tokyo')
    utc_time = datetime.datetime.fromtimestamp(unix_timestamp, tz=pytz.utc)
    japan_time = utc_time.astimezone(japan_tz)
    japan_time_str = japan_time.strftime('%Y-%m-%d %H:%M:%S %Z%z')
    return japan_time_str

def append_to_excel(data, filename, sheet_name):
    df_new = pd.DataFrame(data)
    if os.path.exists(filename):
        df_existing = pd.read_excel(filename, sheet_name=sheet_name)
        last_row = df_existing.tail(1)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        df_combined = df_new
    df_combined.to_excel(filename, sheet_name=sheet_name, index=False)

def get_last_row_from_excel(filename, sheet_name):
    if os.path.exists(filename):
        df_existing = pd.read_excel(filename, sheet_name=sheet_name)
        if not df_existing.empty:
            return df_existing.tail(1).to_dict(orient='records')[0]
    return None

def fetch_comments(user_input):
    url = 'https://igcomment.com/tiktok-display-comments/'
    
    data = {
        'user_input': user_input,
        'cursor': '8'  # Adjust the cursor value if needed
    }

    try:
        response = requests.post(url, data=data)
        response.raise_for_status()  # Raise an error for bad status codes
        
        try:
            json_data = response.json()
            # Extract comments array
            comments = json_data.get('comments', [])
            # Return the array of comment texts
            return [comment.get('text', '') for comment in comments]
        except ValueError:
            # Return an empty list if JSON parsing fails
            return []
    except requests.exceptions.RequestException as e:
        print(f"Request failed: {e}")
        return []

def getPostInfo(username, cursor, max_retries=5, delay=2):
    url = "https://freetik.co/api/tiktok"
    headers = {
        "Content-Type": "application/json",
        "Accept": "*/*",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "en-US,en;q=0.9,en-IN;q=0.8",
        "Priority": "u=1, i"
    }
    json_payload = {"max_cursor": cursor, "username": username}

    for attempt in range(max_retries):
        try:
            response = requests.post(url, headers=headers, json=json_payload, timeout=15)
            response.raise_for_status()  # Raise an error if the request fails
            data = response.json()
            
            # Check if 'status_msg' indicates no more videos
            if data.get('status_msg') == "No more videos":
                return 0

            # Retrieve 'max_cursor', return 0 if missing
            max_cursor = data.get("max_cursor")
            if max_cursor is None:
                print("Warning: 'max_cursor' not found in API response.")
                return 0  # Or handle differently

            return [data, max_cursor]

        except requests.exceptions.RequestException as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(delay * (2 ** attempt))  # Exponential backoff
            else:
                return 0  # Final failure returns 0

file_path = 'split_1.csv'
reg_url = 'https://api.tokbackup.com/api/auth/register'
basic_url = 'https://api.tokbackup.com/api/basicdata'
data = pd.read_csv(file_path)
usernames = data['username']
all_processed_data = []
sheetId = '1sUd1Pi0q3FVYbAQisRMdatZcLUs-kg8C6M1rNMbFROM'

all_processed_data = []
val = 0
for username in usernames:
    if not isinstance(username, str):
        print(f"Non-string value encountered: {username}. Breaking the loop.")
        break

    print(username)

    random_number = random.randint(1000000, 9999999)

    reg_payload = {
        "email": f"user{random_number}@gmail.com",
        "name": f"user{random_number}",
        "password": "Qwer4321",
        "phone": "+810900909090",
        "username": username
    }
    try:
        reg_response = requests.post(reg_url, json=reg_payload)
        reg_response.raise_for_status()
        reg_data = reg_response.json()
        user_id = reg_data["user"]["id"]
        print(user_id)
    except requests.exceptions.RequestException as e:
        print(f"Registration request failed: {e}")
        print("Response Code:", reg_response.status_code)
        print("Response Body:", reg_response.text)
        continue
    except (json.JSONDecodeError, KeyError) as e:
        print(f"Error parsing registration response: {e}")
        continue

    basic_payload = {
        "username": username,
        "platform": "tiktok",
        "user": user_id
    }

    try:
        basic_response = requests.post(basic_url, json=basic_payload)
        basic_response.raise_for_status()
        basic_data = basic_response.json()
        print(basic_data)
    except requests.exceptions.RequestException as e:
        print(f"Basic info request failed: {e}")
        continue
    except json.JSONDecodeError as e:
        print(f"Error parsing basic info response: {e}")
        continue

    if basic_data.get('error'):
        continue

    user_processed_data = {
        'ID': basic_data.get("id", ""),
        'ユーザー': basic_data.get("username", ""),
        'ユーザー名': basic_data.get("name", ""),
        '自己紹介': basic_data.get("bio", ""),
        'フォロワー数': basic_data.get("followerCount", 0),
        '投稿動画数': basic_data.get("videoCount", 0),
        'ハート数': basic_data.get("heartCount", 0),
        'フォロー数': basic_data.get("followingCount", 0),
        'アバターURL': basic_data.get("avatar", ""),
    }

    status = ''
    cursor = 9007199254740991
    update_count = 0

    while status != "No more videos":
        result = getPostInfo(username, cursor)
        if result == 0:
            break
        cursor = result[1]
        info = result[0]["aweme_list"]

        for subData in info:
            if isinstance(subData, str):
                try:
                    subData = json.loads(subData)
                except json.JSONDecodeError:
                    print(f"Error decoding JSON for post data: {subData}")
                    continue

            post_url = f'https://www.tiktok.com/@{username}/video/' + subData['statistics']['aweme_id']
            print(post_url)
            comments = fetch_comments(post_url)

            processedData = {
                f'投稿タイトル-{update_count}': subData['desc'],
                f'投稿日-{update_count}': convert_to_japanese_time(subData['create_time']),
                f'投稿URL-{update_count}': post_url,
                f'再生数-{update_count}': subData['statistics']['play_count'],
                f'いいね数-{update_count}': subData['statistics']['digg_count'],
                f'コメント数-{update_count}': subData['statistics']['comment_count'],
                f'保存数-{update_count}': subData['statistics']['collect_count'],
                f'共有数-{update_count}': subData['statistics']['share_count'],
            }

            for comment_index, comment in enumerate(comments):
                processedData[f'コメント-{update_count}-{comment_index}'] = comment

            user_processed_data.update(processedData)
            update_count += 1

            if update_count >= 10:
                break

        if update_count >= 10:
            break

    print(user_processed_data)
    all_processed_data.append(user_processed_data)

    if len(all_processed_data) >= 1:
        append_to_excel(all_processed_data, excel_filename, sheet_name)
        all_processed_data = []

if len(all_processed_data) > 0:
    append_to_excel(all_processed_data, excel_filename, sheet_name)

print("Successfully Done!")
