# Introduction
Report Phishing Addin for Outlook client (2010/2013/2016)
The addin helps users who want to report a suspicious looking email to their internal Information Security team. An outlook add-in is added to the home ribbon. The user highlights the emails in question and clicks on the button to submit the email to a common inbox. The email is sent with the user details retrieved from the Address book along with the logged in AD user and the hostname of the computer. Following this the email is deleted from the inbox. 

# Getting Started

1.	Dependencies
    Microsoft Visual studio 2017 Installer Projects
    NET Framework 4.5

2.  Running the code
    ReportPhishing - The main code for the add-in
    Connect.h (ReportPhishing) - contains the main code used for implementing the add-in
    ReportPhishingPS - The code for the setup file (creates the required registry keys,etc)


