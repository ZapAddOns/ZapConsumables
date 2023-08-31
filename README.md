# ZapConsumables
PowerShell script to send a list of treated patients from last month by mail

### Intro
This PowerShell script creates two lists of all patients treated in the last month with the Zap. The first list contains the UUID of the plan, the second one contains patients first and last name. Both are send as one mail via a SMTP server to one or more recipents.

### How it works
The script should run in the Zap network (10.0.0.xxx). The best would be on the DATABASEPC. If you start it with the Task Scheduler, it will run each month without any interaction. It connects to the SQLServer used by the Broker and collects the data. This data is then formated and send via SMTP to the recipents.

### How to setup the script
There are a few things to provide to the script. All are collected in the first part of the script.

1. In line 17 you have to set the username, which you need for the access to the SMTP server
2. In line 18 you have to set the password for user username, which you need for the access to the SMTP server. This is created by the instructions below. It had to be in one line, no line breaks allowed.
3. In line 19 you set all the mail addresses, that should used as receivers of this mail (To addresses). It could contain one or more entries, separated by a comma.
4. In line 20 you set all the mail addresses, that should used as carbon copys of this mail (Cc addresses). It could contain one or more entries, separated by a comma.
5. In line 21 you set the mail addresse of the sender of this mail.
6. In line 22 you set the clear name of the sender.
7. In line 23 you set the IP address of the SMTP server you use.
8. In line 24 you set the port of the SMTP server you use.

### How to create passwords in the right format
A short description is also at the begining of the script.

1. Open a PowerShell
2. Enter the following line and replace "Your password" with your real password
      "Your password" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
3. Copy the given string with numbers in the right field below
4. Save the PowerShell script

### How to setup the Microsoft Task Scheduler
If you setup a task with the Microsoft Task Scheduler, the mail is sent each month automatically at all mail addresses you selected. So you never forget to collect the data.

1. Start Task Scheduler
2. Go to \<Action\>\<Create New Task...\> in the menu
3. Add a "Name" and a "Description" to help you to find the task
4. Add a user with right security to run this script
5. Check "Run whether user is logged in or not"
6. Change page to "Triggers"
7. Add a "New..." one
8. Select "On a schedule" for "Begin the task"
9. Set the rest as you like. Normally "Monthly" and the first in each month would be the best
10. Change page to "Actions"
11. Add a "New..." one
12. Set "Action" to "Start a program"
13. Set "Program/script" to "PowerShell.exe". Perhaps you have to add the path to PowerShell.
14. Set "Add arguments (optional)" to the file name of this script. Add the path too.
15. Check the other things, if they make sense to you
16. Create the task by pressing "OK"