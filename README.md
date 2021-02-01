# CreateUsersAD
Starts the logging process
Generates a random password
Picks up all files stored on under the Drop folder
Adds headers to all the csv files and moves the files to the Create folder
Sets the parameters based on the csv
Interrogates the global catalogue for the account owner domain so that it can place the requested account under the same domain
Creates the user based on the parameters, and sets the extension attributes (extensionAttribute1 and extensionAttribute2) and the accountType
In case of failure due to duplicate accounts it sends an email notification to the requestor and to IT, while the csv file is moved to the Error folder
In case of failure due to other reasons, IT is informed, and the csv file is moved to the Error folder
Once the user is created successfully, the script sends an email to the owner with the username and password in an HTML format
The processed files are moved to the Archive
Logs are sent to IT

Please see full documentation link here: https://www.accesa.eu/insights/automating-account-creation-in-on-prem-ad-with-power-automate.
