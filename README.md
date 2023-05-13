# AutoApi v6.1 (2021-2-10) ———— E5 Automatic Renewal
AutoApi Series: ~~AutoApi(v1.0)~~, ~~AutoApiSecret(v2.0)~~, ~~AutoApiSR(v3.0)~~, ~~AutoApiS(v4.0)~~, ~~AutoApiP(v5.0)~~

## Description ##
* E5 automatic renewal program, but **does not guarantee renewal**
* Set to **not start** automatic calling on Saturdays and Sundays (UTC time), and automatically start every 6 hours from Monday to Friday (see tutorial for modification)
* Call API to keep alive:
     * Query APIs: onedrive, outlook, notebook, site, etc.
     * Create APIs: automatically send emails, upload files, modify Excel, etc.
     
### Related ###
* AutoApi: https://github.com/wangziyingwen/AutoApi
* **Errors and Solutions/Renewal-related Knowledge/Update Log**: https://github.com/wangziyingwen/Autoapi-test
   * Most error descriptions have been updated into the program, please see the action log report for details after running
* Video tutorial:
   * Bilibili: https://www.bilibili.com/video/BV185411n7Mq/

## Steps ##
* Preparation:
   * E5 developer account (**not personal/private account**)
       * Administrator account ———— Required
       * Sub-account ———— Optional (not sure if Microsoft will count the activity of sub-accounts, choose to run selectively if you want to)
   * rclone software, [download address rclone.org](https://downloads.rclone.org/v1.53.3/rclone-v1.53.3-windows-amd64.zip), (windows 64)
   * If the tutorial pictures cannot be seen, please use a VPN
   
* Steps outline:
   * Microsoft's preparation work (get application ID, password, and key)
   * GIHTHUB's preparation work (get Github key, set secret)
   * Trial run
   

#### Microsoft Preparation Work ####

* **Step 1, Register the application and get the application ID and secret**

   * 1) Click to open the [dashboard](https://aad.portal.azure.com/), click **All Services** on the left, find **App Registration**, and click +**New Registration**
   
    ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/creatapp.png)
   
   * 2) Fill in the name, select any of the first three supported account types, fill in http://localhost:53682/ for redirect, and click **Register**
   
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/creatapp2.png)
   
   * 3) Copy the application (client) ID to Notepad for backup (**Got the application ID!**), click **Certificates & secrets** on the left, click +**New client secret**, click Add, and save the **value** of the new client secret (**Got the application password!**)
   
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/creatapp3.png)
   
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/creatapp4.png)
   
   * 4) Click **API permissions** on the left, click +**Add a permission**, click **Microsoft Graph** in the common Microsoft API (the blue crystal), click **Delegated permissions**, and then select the following required permissions in the following clauses, and finally click **Add permissions** at the bottom
   
   **When assigning API permissions, select the following 13**
  
            Calendars.ReadWrite, Contacts.ReadWrite, Directory.ReadWrite.All,
            
            Files.ReadWrite.All, Group.ReadWrite.All, MailboxSettings.ReadWrite,
            
            Mail.ReadWrite, Mail.Send, Notes.ReadWrite.All,
            
            People.Read.All, Sites.ReadWrite.All, Tasks.ReadWrite,
            
            User.ReadWrite.All
                        
   
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/creatapp5.png)
   
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/creatapp6.png)
    
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/creatapp8.png)
   
   * 5) After adding, it will automatically jump back to the permission homepage, click **Grant admin consent**.
       
       If it is a **sub-account** operation, please log in to the [dashboard](https://aad.portal.azure.com/) with the administrator account to find the **application registered by the sub-account**, and click "Representative administrator authorization". 
   
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/creatapp7.png)
   
* **Step 2, Get the refresh_token (Microsoft key)**

   * 1) In the rclone.exe folder, shift+right-click to open powershell here, enter the modified content below, and press Enter to jump out of the browser, log in to the e5 account, click Accept, and return to the powershell window to see a string of things.
         
            ./rclone authorize "onedrive" "Application (client) ID" "Application password"
            
   * 2) Find "refresh_token": in that string of things, select from the double quotes to ","expiry":2021 (the string in the double quotes after refresh_token, without double quotes), as shown in the figure below, right-click to copy and save (**Got the Microsoft key!**)
   
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApi/token地方.png)
   
 ____________________________________________________
 
 #### Preparations for GITHUB ####

 * **Step 1: Fork this project**
 
     Log in or create a GitHub account, return to the project page, and click the fork button in the top-right corner to fork the project code to your own account. Then, you will see an identical project under your account, and all subsequent operations will be performed in this project.
     
     ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApi/fork.png)
     
 * **Step 2: Create a new GitHub secret**
 
    * 1) Go to your personal settings page (click on the avatar in the top-right corner and choose Settings, not the repository settings), select Developer settings -> Personal access tokens -> Generate new token.

    ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApi/Settings.png)
    
    ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApi/token.png)
    
    * 2) Set the name to **GH_TOKEN**, then check repo, click Generate token, and finally **copy and save** the generated GitHub secret (**you've got the GitHub secret**; once you leave the page, you won't see it next time!).
   
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/repo.png)
  
 * **Step 3: Create new secrets**
 
    * 1) Click on the Settings tab in the top-right corner of the page, then on the left sidebar click Secrets -> New repository secret in the top-right corner, and create 6 new secrets: **GH_TOKEN, MS_TOKEN, CLIENT_ID, CLIENT_SECRET, CITY, EMAIL**.
   
    ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/setting.png)
    
    ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApiP/secret2.png)
    
     **(Make sure not to have any spaces or blank lines before or after the following content)**
 
     GH_TOKEN
     ```shell
     GitHub secret (obtained in Step 2), for example, if the obtained secret is abc...xyz, paste it directly into the secret page without any modification, just make sure there are no spaces or blank lines before or after
     ```
     MS_TOKEN
     ```shell
     Microsoft secret (refresh_token obtained in Step 2)
     ```
     CLIENT_ID
     ```shell
     Application ID (obtained in Step 1)
     ```
     CLIENT_SECRET
     ```shell
     Application password (obtained in Step 1)
     ```
     CITY
     ```shell
     City (e.g., Beijing, used for sending weather emails automatically)
     ```
     EMAIL
     ```shell
     Recipient's email address (used for sending weather emails automatically)
     ```

________________________________________________

#### Trial Run ####

   * 1) Click on the Actions tab in the middle of the top bar to enter the run log page. There should be a green button in the middle (I understand my workflow...), click it.
   
      After automatic refresh, there will be three processes on the left, one Run api.Read, one Run api.Write, and one Update Token.
      
       Workflow Description
          Run api.Write: Create system api, run once a day
          Run api.Read: Query system api, run every 6 hours
          Update Token: Microsoft key update, run every 2 days
          
      These three process names should all have yellow exclamation marks in front of them.
   
      Click on each one, and then you will see a yellow bar (this schedule was disabled...), click the enable workflow button, **you need to do this for all three processes!**
   
      (Not sure if all of them need to be done, I found some when I was making a video tutorial. If you don't have it, just ignore it and continue, as long as it can run normally)
   
   * 2) Click the star button (next to the fork button) twice to start the action,
   
      Then click on the Actions tab again, select the Run api.Read or api.Write process -> build -> run api to see the running log each time.

      (You must go into the build and run api.XXX to see if the api has been called properly, if the operation was successful, and if there were any errors)

      ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApi/日志.png)
    
   * 3) Click the star button twice again to see if it can run successfully again.
   
      Then click on the update token process in the Actions tab -> build -> update token, and the log will show "Microsoft key uploaded successfully".
      
      At the same time, click on the Settings tab on the page -> left column Secrets (the third step of Github preparation), and you should see MS_TOKEN displayed as just updated.
      
      (This step is to ensure that the token re-uploaded to the secret is correct)
   
      
#### End of Tutorial ####

   The program will start automatically according to the plan, so don't worry about it.
   
      However, Github has updated the rules to prevent cheating. If there is no change in the repository for 60 days, the Action will be suspended, but an email notification will be sent. So please pay attention to your email. If you receive an email, please come up and manually start the action.
      (I have never received this email, but it is said that there will be a start link in the email, or just click the star button twice)
   
   **P version (AutoApiP) users, please pay attention to whether this suspension rule will be triggered. Due to the new plan adopted by the P version, can it skip the Github detection of activity? If the P version receives a suspension email, it is best to leave a message in this post in the issues [Trigger Suspension Statistics](https://github.com/wangziyingwen/AutoApiP/issues/7)**
   
### Tutorial Complete ###

__________________________________________________________________________

## Additional Settings (if you don't understand, please ignore) ##
   * **Modify scheduled start time**

   * **Multi-account/application support**
    
   * **Advanced parameter settings**

#### Modify Scheduled Start Time ####
   
   I have set it to run automatically every 6 hours (not on Saturdays and Sundays), with 3 rounds of calls each time (clicking on the star in the top-right corner can also invoke it immediately). You can adjust the settings as needed (I'm not sure how many times and how often to maintain activity):

  * To modify the scheduled start time, go to the .github/workflow/autoapi.yml file (only modify this one), and look up the cron scheduling format on your own. The shortest interval is once every 5 minutes.
   
   ![image](https://github.com/wangziyingwen/ImageHosting/blob/master/AutoApi/定时.png)
    
#### Multi-account/Application Support ####

   If you want to enter a second account or application, please follow the steps above to obtain the **ID, password, and Microsoft secret for the second application**:
 
   Then follow these steps:
 
   1) Add secret
 
   Click Settings in the top-right corner of the page, then on the left sidebar click Secrets -> New repository secret, and add new secrets: APP_NUM, MS_TOKEN_2, CLIENT_ID_2, CLIENT_SECRET_2
 
   APP_NUM
   ````shell
   Number of accounts/applications (for example, if there are two accounts/applications now, then it's 2; if there are three accounts, then fill in 3; if you want to increase the number in the future, modify APP_NUM)
   ```
   MS_TOKEN_2
   ````shell
   Microsoft secret for the second account (refresh_token from Step 2), (for the third account/application, it would be MS_TOKEN_3, and so on)
   ```
   CLIENT_ID_2
   ````shell
   Application ID for the second account (obtained in Step 1), (for the third account/application, it would be CLIENT_ID_3, and so on)
   ```
   CLIENT_SECRET_2
   ````shell
   Application password for the second account (obtained in Step 1), (for the third account/application, it would be CLIENT_SECRET_3, and so on)
   ```
   
   2) Modify the two yml files in .github/workflows/ (**if you have more than 5 accounts, you need to change this; if you have 5 or fewer accounts, you don't need to modify the files for now, so ignore this step**)
    
   The yml files are already annotated, so you can just follow the instructions. I've already added templates for 5 accounts, so copying and pasting should be very simple (I haven't found a better automatic solution yet).
  
#### Advanced Parameter Settings ####
 
   In the ApiOfRead.py and ApiOfWrite.py files, there is a config around line 11, and the specific parameter settings are explained in the files.
   
   This includes random delays for account API calls, API random sorting, and the number of rounds per time, etc.
     
   
### Conclusion ###

If you have any issues, please create an issue.
    




