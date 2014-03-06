outofoffice
===========

**********************************************************************************************
* Out of office tool by Kellan Damm
*
* originating powershell scripts were used from: 
* http://gsexdev.blogspot.com/2011/11/creating-out-of-office-board-using.html
* This currently works with Exchange 2010.
* Uses Exchange Web Services API
**********************************************************************************************

SYNOPSIS:

Included are 3 PowerShell Scripts 
1. Create dynamic html to be used for groups or entire companies.  
2. Create html file for Front Desk employees with Menu
3. Send email with Out of Office info for a specific group

Each PowerShell Script runs through in increments of 100.
Included are .css files that are used for dymnamic html sites. CSS is embedded in email script
Create a scheduled task for each script based on how often you want to update

**********************************************************************************************

REQUIREMENTS:

1. Windows based machine with Exchange Management Console installed
2. Exchange user with Management Role for Exchange Web Services
	RUN THIS in Exchange shell: 
	New-ManagementRoleAssignment -Name:"OOF EWS" -Role:ApplicationImpersonation –User:"OOF USER"
3. Creating scheduled Task for each group you want to get Out of Office info for.
4. Scheduled task must run as user with Management Role permissions



**********************************************************************************************

FILE DESCRIPTIONS AND REQUIRED CHANGES:

Below I give a brief description of each file and lines that you will want to change. 

**********************************************************************************************

Location.ps1

Creates a dynamic site for email distros. Included is a top bar if you wanted
to link to different group. Hovering over blocks will give info.

LINE 20 - change alias
LINE 32 - 33 change to the times you are open are script will run though each user and grab OOF info.
LINE 38 - location header
LINES 200 -210 change links for top header
LINE 215 - change location to save html file
File requires menu.css

You can change color throughout script to your liking.

**********************************************************************************************

Frontdesk.ps1

Creates a site for Front Desk employees that man a switchboard. Included is a menu page that will 
pop out each link so employee could have multiple groups open on a screen. The group pages created 
by script are black, grey and white only.

LINE 20 - change alias
LINE 32 - 33 change to the times you are open are script will run though each user and grab OOF info.
LINE 38 - location header
LINE 192 - change location to save html file
Script requires menu2.css
FrontdeskMENU requires menu1.css

You can change color throughout script to your liking

***********************************************************************************************

Email.ps1

Creates an email that sends the out of office info for an email distro in html format CSS is included.
LINE 20 - change alias
LINE 77 - change to Group name.
LINES 223-226 - change email info

You can change color throughout script to your liking

************************************************************************************************

