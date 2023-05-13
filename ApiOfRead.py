# -*- coding: UTF-8 -*-
import os
import requests as req
import json, time, random
from typing import TypedDict, List

app_num = os.getenv("APP_NUM")
if not app_num or app_num == "":
    app_num = "1"
access_token_list = ["wangziyingwen"] * int(app_num)

###########################
# config options
# 0: off, 1: on
# api_rand: whether to randomly sort api (turn on to randomly extract 12, turn off to default to the initial version of 10). Default is 1 to turn on.
# rounds: number of rounds, i.e. how many rounds to run each time.
# rounds_delay: whether to enable random delay between rounds, the latter two parameters represent the interval of delay. Default is 0 to turn off.
# api_delay: whether to enable delay between api, default is 0 to turn off.
# app_delay: whether to enable delay between accounts, default is 0 to turn off.
########################################

class ConfigType(TypedDict):
    api_rand: int
    rounds: int
    rounds_delay: List[int]
    api_delay: List[int]
    app_delay: List[int]

config: ConfigType = {
    "api_rand": 1,
    "rounds": 3,
    "rounds_delay": [0, 60, 120],
    "api_delay": [0, 2, 6],
    "app_delay": [0, 30, 60],
}

api_list = [
    r"https://graph.microsoft.com/v1.0/me/",
    r"https://graph.microsoft.com/v1.0/users",
    r"https://graph.microsoft.com/v1.0/me/people",
    r"https://graph.microsoft.com/v1.0/groups",
    r"https://graph.microsoft.com/v1.0/me/contacts",
    r"https://graph.microsoft.com/v1.0/me/drive/root",
    r"https://graph.microsoft.com/v1.0/me/drive/root/children",
    r"https://graph.microsoft.com/v1.0/drive/root",
    r"https://graph.microsoft.com/v1.0/me/drive",
    r"https://graph.microsoft.com/v1.0/me/drive/recent",
    r"https://graph.microsoft.com/v1.0/me/drive/sharedWithMe",
    r"https://graph.microsoft.com/v1.0/me/calendars",
    r"https://graph.microsoft.com/v1.0/me/events",
    r"https://graph.microsoft.com/v1.0/sites/root",
    r"https://graph.microsoft.com/v1.0/sites/root/sites",
    r"https://graph.microsoft.com/v1.0/sites/root/drives",
    r"https://graph.microsoft.com/v1.0/sites/root/columns",
    r"https://graph.microsoft.com/v1.0/me/onenote/notebooks",
    r"https://graph.microsoft.com/v1.0/me/onenote/sections",
    r"https://graph.microsoft.com/v1.0/me/onenote/pages",
    r"https://graph.microsoft.com/v1.0/me/messages",
    r"https://graph.microsoft.com/v1.0/me/mailFolders",
    r"https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
    r"https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta",
    r"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules",
    r"https://graph.microsoft.com/v1.0/me/messages?$filter=importance eq 'high'",
    r'https://graph.microsoft.com/v1.0/me/messages?$search="hello world"',
    r"https://graph.microsoft.com/beta/me/messages?$select=internetMessageHeaders&$top",
]


# Microsoft refresh_token acquisition
def getmstoken(ms_token, appnum):
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "grant_type": "refresh_token",
        "refresh_token": ms_token,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": "http://localhost:53682/",
    }
    html = req.post("https://login.microsoftonline.com/common/oauth2/v2.0/token", data=data, headers=headers)
    jsontxt = json.loads(html.text)
    if "refresh_token" in jsontxt:
        print("Microsoft key obtained successfully for account/application " + str(appnum))
    else:
        print(
            "Failed to obtain Microsoft key for account/application "
            + str(appnum)
            + "\n"
            + "Please check if the format and content of CLIENT_ID, CLIENT_SECRET, and MS_TOKEN in the secret are correct, and then reset them"
        )
    refresh_token = jsontxt["refresh_token"]
    access_token = jsontxt["access_token"]
    return access_token


# Call the function
def runapi(apilist, a):
    localtime = time.asctime(time.localtime(time.time()))
    access_token = access_token_list[a - 1]
    headers = {"Authorization": "bearer " + access_token, "Content-Type": "application/json"}
    for b in range(len(apilist)):
        if req.get(api_list[apilist[b]], headers=headers).status_code == 200:
            print(f"[{localtime}] API " + str(apilist[b]) + " called successfully")
        else:
            print("pass")
        if config["api_delay"][0] == 1:
            time.sleep(random.randint(config["api_delay"][1], config["api_delay"][2]))


# Get access_token in bulk to reduce the acquisition rate
for a in range(1, int(app_num) + 1):
    client_id = os.getenv("CLIENT_ID_" + str(a))
    client_secret = os.getenv("CLIENT_SECRET_" + str(a))
    ms_token = os.getenv("MS_TOKEN_" + str(a))
    access_token_list[a - 1] = getmstoken(ms_token, a)

# Randomize API sequence
fixed_api = [0, 1, 5, 6, 20, 21]
# Ensure that APIs for outlook and onedrive are included
ex_api = [2, 3, 4, 7, 8, 9, 10, 22, 23, 24, 25, 26, 27, 13, 14, 15, 16, 17, 18, 19, 11, 12]
# Additional APIs to fill in
fixed_api.extend(random.sample(ex_api, 6))
random.shuffle(fixed_api)
final_list = fixed_api

# Actual execution
if int(app_num) > 1:
    print("In multi-account/application mode, there may be a bunch of *** in the log report, which is normal")
print("If the number of APIs is less than the specified value, it is because the API authorization has not been done properly, or OneDrive has not been initialized successfully. Please re-authorize and obtain Microsoft keys to replace them, or wait a few days for OneDrive to initialize.")
print("Total " + str(app_num) + r" accounts/applications, " + r"each account/application " + str(config["rounds"]) + " rounds")
for r in range(1, config["rounds"] + 1):
    if config["rounds_delay"][0] == 1:
        time.sleep(random.randint(config["rounds_delay"][1], config["rounds_delay"][2]))
    for a in range(1, int(app_num) + 1):
        if config["app_delay"][0] == 1:
            time.sleep(random.randint(config["app_delay"][1], config["app_delay"][2]))
        client_id = os.getenv("CLIENT_ID_" + str(a))
        client_secret = os.getenv("CLIENT_SECRET_" + str(a))
        print("\n" + "Round " + str(r) + " for application/account " + str(a) + " " + time.asctime(time.localtime(time.time())) + "\n")
        if config["api_rand"] == 1:
            print("Random order is enabled, there are twelve APIs in total, count them yourself")
            apilist = final_list
        else:
            print("Original order, there are ten APIs in total, count them yourself")
            apilist = [5, 9, 8, 1, 20, 24, 23, 6, 21, 22]
        runapi(apilist, a)
