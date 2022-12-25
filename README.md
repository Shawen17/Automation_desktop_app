# Automation_desktop_app

#### About

![alt text](https://github.com/Shawen17/Automation_desktop_app/blob/main/desktop-app.jpg)

This is a desktop application (GUI) that is designed to automate Report preparation and also send out prepared reports.
It automates the activity of an Excel sheet and sends out the desired report by a bulk sms platform.
The report is for both 2G and 3G base-station activities. The desktop application declares sites up, record the downtime duration and sends base stations currently down to field
engineers. It has a progress bar used to monitor the task completion state.
The Tkinter python module is used to develop this desktop application, pandas is used in manipulating the Excel sheet and Selenium is used to Scrap the bulk sms web platform and 
automate sending of the report to the regions involved.

#### Modules
For this application to function, the following python module must be installed;
* Tkinter
* Selenium
* Pandas
* bs4
