####################
# Calendar Script  #
# Author: DBemrose #
# Build: POC-Dev   #
# Build V: 2.3.0   #
####################
$Script:Version = "2.3.0"
$Script:Build = "DEV"

####################
# Script Settings  #
####################




####################
# Logging Function #
####################
Function EnableLogging{
    $VerbosePreference = "Continue"
    $Script:OwnPath = Get-Location
    $Script:LogPath = "$OwnPath\logs"
    $Script:LogPathName = Join-Path -Path $LogPath -ChildPath "$("Log")-$(Get-Date -Format 'dd-MM-yyyy').log"
    Start-Transcript $Script:LogPathName -Append
    #Stop script logging on exit
    Register-EngineEvent PowerShell.Exiting -Action {Stop-Transcript}
}


#Login MFA function is the universal way to login as it accepts MFA and non MFA clients. This does not handle delegate accounts and will need you to login to the client's admin portal
Function LoginMFA{
    #This imports the ExchangeOnlineManagement Module. This is assumed that the module is already installed but can be easily obtained or an auto install function can be created if needed
    Import-Module -Name ExchangeOnlineManagement
    #initiates a connection with no details so the technician can login as whoever they need
    Connect-ExchangeOnline -ShowBanner:$False
    #calls the MainMenu function
    MainMenu
}

#Gets all the mailboxes within the organization. This is used for the dynamic menu creation for selecting the initial user that owns the calendar
Function Get-Users{
    #clears all the text that is on screen
    clear-host
    IF($Script:users.count -lt 1){
    #Writes to the screen what will be happening next
    Write-host "Getting all users..."
    #Creates a variable that can be used by only the script that will get all the names for the mailboxes in the organization
    $Script:users = Get-Mailbox | Select -ExpandProperty Name
    #Creates an empty hashtable that is accessed only by the script and will be used to store the index number and the username. This is used to create a dynamic menu system.
    $Script:UsersTable = @{}
    #Creates a variable that is used to add 1 after each username has been passed.
    $iteration = 0
    #For each loop which will loop through every username and add the iteration(Index) and username to the UsersTable hashtable. This is used as the base for the dynamic menu system.
    foreach($Script:Username in $Script:Users){
        #Adds the variables per username to the hashtable
        $Script:UsersTable.add($iteration, $Script:Username)
        #Adds 1 to the iteration variable so we can create the menu properly.
        $iteration = $iteration + 1
    }
    #Writes to the screen how many users it had found based on the unique entries within the users variable
    Write-Host ('Found {0} Users' -f $Script:users.count)
    #Calls the SelectUserMenu function
    SelectUserMenu
}   Else{
        SelectUserMenu
    }
}

#This function does the exact same thing as the one above but for the user you would like to edit for. 
#This is the barbaric way of doing it until I can implement a universal function which would work regardless of the type you need it for.
Function Get-UserAdjust{
    clear-host
    Write-host "Getting all users..."
    $Script:usersa = Get-Mailbox | Select -ExpandProperty Name
    $Script:UsersTablea = @{}
    $iterationa = 0
    foreach($Script:Usernamea in $Script:Usersa){
        $Script:UsersTablea.add($iterationa, $Script:Usernamea)
        $iterationa = $iterationa + 1
    }
    Write-Host ('Found {0} Users' -f $Script:usersa.count)
    SelectUserMenua
}

#This function is to get the calendars for the selected user. This is done by getting all the folder statistics then filtering the name value to search for anything containing calendar
#then selecting the identity value to get the result for all the calendars. This will then add it to the UsersCalendars hashtable just like in the previous 2 functions
Function Get-CalendarsUser{
    #Gets the folder statistics for the mailbox for the user selected. Then filters the search for anything containing Calendar Then selects the identity variable
    $Script:UserCalendars = Get-MailboxFolderStatistics -Identity $Script:UserSelected | Where Name -like "*Calendar*" | Select -ExpandProperty Identity
    #Creates the userscalendarstable hashtable to be used for storing the calendars for the dynamic menu.
    $Script:UsersCalendarsTable = @{}
    #Creates the iteration variable used for the indexing for the dynamic menu
    $iteration = 0
    #will loop through each user calendar and will add it to the UCT hashtable to be used as the base for the dynamic menu.
    foreach($Script:UserCalendar in $Script:UserCalendars){
        #Adds the values to the hashtable
        $Script:UsersCalendarsTable.add($iteration, $Script:UserCalendar)
        #Increases the index by 1
        $iteration = $iteration + 1
    }
    #Outputs the number of calendars found for the user.
    Write-Host ('Found {0} Calendars for User {1}' -f $Script:UserCalendars.Count, $Script:UserSelected)
    #Ends the function and goes back to the point just after the function was called.
    Return
}

#Gets all the user permissions on the selected calendar and outputs it as a table.
#This is in the planning stage to build a hashtable to make it a lot more versatile as currently it has minimal manipulation that can be done to it.
Function Get-Calendar-Permissions{
    #creates a variable which will store the calendar permissions for the selected calendar in a table format containing the username and their access rights
    $Script:CalendarPermissions = Get-MailboxFolderPermission -Identity $Script:CalendarSelectedR | ft User, AccessRights 
    #Outputs the variable to the console to display to the user
    $Script:CalendarPermissions
    #Barbaric way of waiting for user input. This will be updated with a cleaner function so it looks nicer but there is nothing inherently wrong with this approach.
    Read-Host 'Press Any key to conitnue...'
    #Ends the function and goes back to the point just after the function was called.
    Return
}

#Builds a menu system asking the user which permissions they would like to EDIT and will then edit them with the settings supplied
#A quick options at the top is used for the 2 options that are the main ones used when actioning requests.
#This will provide an error if the user to edit is NOT already added to the mailbox. It will not break anything it just wont edit it and you will need to add it instead.
Function Set-Calendar-Permission{
    #Prints the menu system to the user
    Write-Host "Permission to apply:"
    Write-Host "Quick permissions"
    Write-Host "-------------------------"
    Write-Host "1. Reviewer"
    Write-Host "2. Editor"
    Write-Host "-------------------------"
    Write-Host "All Permissions"
    Write-Host "-------------------------"
    Write-Host "3. Publishing Editor"
    Write-Host "4. Publishing Author"
    Write-Host "5. Author"
    Write-Host "6. Nonediting Author"
    Write-Host "7. Contributor"
    #Waits for user input for the menu system.
    $Script:PermR = Read-Host "Select Permission"
    #Uses the user input provided to set the permission variable. If a value is provided that is not within the menu system it will recall
    Switch($Script:PermR){
        1{$Script:Perm = "Reviewer"}
        2{$Script:Perm = "Editor"}
        3{$Script:Perm = "Publishing Editor"}
        4{$Script:Perm = "Publishing Author"}
        5{$Script:Perm = "Author"}
        6{$Script:Perm = "Nonediting Author"}
        7{$Script:Perm = "Contributor"}
        default{Set-Calendar-Permission}
    }
    #clears the text on screen to make it cleaner for the user as the menu is no longer required
    Clear-Host
    #Prints to the console the settings that have been selected and will then attempt to edit the calendar permissions.
    #Currently there is no error handling so the "permissions have been set!" message is dumb like a dumb switch. There are plans to have error handling
    #and base the output on the results
    Write-Host "Attempting to edit the calendar "$Script:CalendarSelectedR "With the following settings:"
    Write-Host "User to edit:" $Script:UserSelecteda
    Write-Host "Permissions:" $Script:Perm
    Write-Host "Editing the calendar..."
    #Sets the folder permission on the mailbox. This most of the time will output some feedback to the console but I have noticed it does not always do it. 
    #It usually does not do it if an error had occured within the last 3 but I have not determined the exact cause. Plans are to make it so it doesn't output the MS result
    #to the console and to build a smart output to give proper feedback
    Set-MailboxFolderPermission -Identity $Script:CalendarSelectedR -user $Script:UserSelecteda -AccessRights $Script:Perm
    #dumb message stating the previous command has finished
    Write-Host "Permissions have been set!"
    #Ends the function and goes back to the point just after the function was called.
    Return
}

#Function to be implemented. This function will provide the ability to remove calendar permissions from a calendar.
Function Remove-Calendar-Permission{
    $Script:RemoveConfirm = Read-Host "Are you sure you want to remove the permissions to $Script:CalendarSelectedR for the user $Script:UserSelecteda (Y/N)"
    Switch($Script:RemoveConfirm){
        Y{Write-Host "Attempting to remove the permissions..."; Remove-MailboxFolderPermission -Identity $Script:CalendarSelectedR -user $Script:UserSelecteda; Return}
        N{Return}
        Default{Return}
    }
}

#This is the exact same as the edit calendar permission function but for adding the user permissions when there is not already one applied to the calendar.
Function Add-Calendar-Permission{
    Write-Host "Permission to apply:"
    Write-Host "Quick permissions"
    Write-Host "-------------------------"
    Write-Host "1. Reviewer"
    Write-Host "2. Editor"
    Write-Host "-------------------------"
    Write-Host "All Permissions"
    Write-Host "-------------------------"
    Write-Host "3. Publishing Editor"
    Write-Host "4. Publishing Author"
    Write-Host "5. Author"
    Write-Host "6. Nonediting Author"
    Write-Host "7. Contributor"
    $Script:PermR = Read-Host "Select Permission"
    Switch($Script:PermR){
        1{$Script:Perm = "Reviewer"}
        2{$Script:Perm = "Editor"}
        3{$Script:Perm = "Publishing Editor"}
        4{$Script:Perm = "Publishing Author"}
        5{$Script:Perm = "Author"}
        6{$Script:Perm = "Nonediting Author"}
        7{$Script:Perm = "Contributor"}
        default{Add-Calendar-Permission}
    }
    Clear-Host
    Write-Host "Attempting to add the permission to the calendar "$Script:CalendarSelectedR "With the following settings:"
    Write-Host "User to add:" $Script:UserSelecteda
    Write-Host "Permissions:" $Script:Perm
    Write-Host "Adding the Permission..."
    Add-MailboxFolderPermission -Identity $Script:CalendarSelectedR -user $Script:UserSelecteda -AccessRights $Script:Perm
    Write-Host "Permissions have been set!"
    Return
}

#This function is no longer used but is currently kept in the code in case it is needed. This was originally used as it was 2 different menu's but I decided to
#combine the 2 menus into 1 to provide a sleeker user experience. 
Function SettingMenu{
    Clear-Host
    Write-Host "Permissions Menu"
    If ([string]::IsNullOrEmpty($Script:CalendarSelectedR)){Write-Host "1. Select Calendar"}Else{Write-Host "Selected Calendar" -NoNewline; Write-Host " (" -f gray -NoNewline; Write-Host $Script:CalendarSelectedR -f green -NoNewline; Write-Host ")" -f gray}
    If ([string]::IsNullOrEmpty($Script:UserSelecteda)){Write-Host "2. Select User for editing"}Else{Write-Host "Selected User for editing" -NoNewline; Write-Host " (" -f gray -NoNewline; Write-Host $Script:UserSelecteda -f green -NoNewline; Write-Host ")" -f gray}
    Write-Host "3. Edit Permission"
    Write-Host "4. Add Permission"
    Write-host "5. Main Menu"
    Write-Host "Q. Quit"
    $Script:SettingR = Read-Host "Option"
    Switch($Script:SettingR){
        1{If ([string]::IsNullOrEmpty($Script:CalendarSelectedR)){Get-CalendarsUser}Else{SettingMenu}}
        2{Get-UserAdjust}
        3{Set-Calendar-Permission}
        4{Add-Calendar-Permission}
        5{MainMenu}
        default{SettingMenu}
        "Q"{Exit}
    }
}

#Builds the main menu system and waits for user input
Function MainMenu{
    Clear-Host
    #Puts the version and build within the output
    Write-Host "Calendar Powershell v $Script:Version($Script:Build)"
    Write-Host "Main Menu:"
    #Login type is not currently used but is planned to be replaced with the email that was logged in with
    Write-Host "1. Login To Exchange" -NoNewline; Write-Host "($Script:LoginType)" -ForegroundColor Red
    Write-Host "2. Calendar Options"
    Write-Host "3. Miscellaneous" -NoNewline; Write-Host " (" -NoNewline -f Gray; Write-Host "Coming Soon" -NoNewline -f Red; Write-Host ")" -f Gray
    Write-Host "Q. Quit"
    #Waits for the user to provide a menu option
    $Script:MMResult = Read-Host "Option"
    #Based on the users output it will go to a specific menu
    Switch($Script:MMResult){
        1{LoginMenu}
        #Some basic logic to make sure the user has logged into the exchange account and loaded the needed commands. If it is not they will be taken to the login menu instead.
        2{If($Script:Login = $null){LoginMenu}Else{SelectionMenu}}
        4{MainMenu}
        default{MainMenu}
        "Q"{Exit}
    }
}

#Basic menu system for providing access to Exchange Online.
#Basic login has been removed as the MFA option can handle both
Function LoginMenu{
    Clear-Host
    Write-Host "Login to Exchange"
    Write-Host "1. Login"
    Write-Host "2. Delegate Login"
    Write-Host "3. Back to main menu"
    Write-Host "Q. Quit"
    $LMResult = Read-Host "Option"
    Switch($LMResult){
        1{LoginMFA}
        2{LoginDelegate}
        3{MainMenu}
        default{LoginMenu}
        "Q"{Exit}
    }
}

#Creates a basic menu system that has checks on some of the menu outputs to add the value of the option if it has already been selected.
#For example the user that has been selected will show on the menu system so you know who was chosen.
Function SelectionMenu {
    #Clears the screen
    clear-host
    #states the menu the tech is currently in
    Write-Host "Selection Menu"
    #when there has been no user selected it will just display the basic "1. Select User" message but when a user has been selected it will display the user in brackets next to it.
    #Example: 1. Select User (DBemrose)
    If ([string]::IsNullOrEmpty($Script:UserSelected)){Write-Host "1. Select User"}Else{Write-Host "1. Select User" -NoNewline; Write-Host " (" -f gray -NoNewline; Write-Host $Script:UserSelected -f green -NoNewline; Write-Host ")" -f gray}
    #When no calendar has been selected it will display the basic "2. Select Calendar" message but when a calendar has been selected it will display the calendar in brackets.
    #Example: 2. Select Calendar (Dbemrose:\Calendar)
    If ([string]::IsNullOrEmpty($Script:CalendarSelected)){Write-Host "2. Select Calendar"}Else{Write-Host "2. Select Calendar" -NoNewline; Write-Host " (" -f gray -NoNewline; Write-Host $Script:CalendarSelectedR -f green -NoNewline; Write-Host ")" -f gray}
    Write-Host "3. List Calendar Permissions"
    #When no user to apply the permissions to has been selected it will display the basic "4. Select User for editing" message but when a user has been selected it will display the
    #user in brackets. Example: 4. Selected User for editing (DBemrose)
    If ([string]::IsNullOrEmpty($Script:UserSelecteda)){Write-Host "4. Select User for editing"}Else{Write-Host "4. Selected User for editing" -NoNewline; Write-Host " (" -f gray -NoNewline; Write-Host $Script:UserSelecteda -f green -NoNewline; Write-Host ")" -f gray}
    Write-Host "5. Edit Permission"
    Write-Host "6. Add Permission"
    Write-Host "7. Remove Calendar Permission (TESTING)"
    Write-Host "8. Back to main menu"
    Write-Host "R1. Refresh User Select"
    Write-Host "Q. Quit"
    $SMResult = Read-Host "Option"
    Switch($SMResult){
        #Runs the Get-Users function and will then run the SelectUserMenu function once the previous one has completed.
        1{Get-Users; SelectUserMenu}
        #Runs the Get-CalendarsUser function and will then run the SelectCalendarMenu function once the previous one has been completed.
        2{Get-CalendarsUser; SelectCalendarMenu}
        #Runs the Get-Calendar-Permissions function and will then run it's own function again to update the menu.
        3{Get-Calendar-Permissions; SelectionMenu}
        #Runs the GetUserAdjust function and will then wait for a user input. Then it will update the menu
        4{Get-UserAdjust; Read-Host 'Press Any key to conitnue...'; SelectionMenu}
        #Runs the Set-Calendar-Permission and will then wait for user input. Then it will update the menu
        5{Set-Calendar-Permission; Read-Host 'Press Any key to conitnue...'; SelectionMenu}
        #Runs the Add-Calendar-Permission and will then wait for user input. Then it will update the menu
        6{Add-Calendar-Permission; Read-Host 'Press Any key to conitnue...'; SelectionMenu}
        #This option is currently being implemented
        7{Remove-Calendar-Permission; Read-Host 'Press Any key to conitnue...'; SelectionMenu}
        #Takes you back to the main menu
        8{MainMenu}
        "R1"{$Script:Users = $null; $Script:UserSelected; SelectionMenu}
        #Will reload the selection menu if an option was proivided that was not on the list
        default{SelectionMenu}
        "Q"{Exit}
    }
}

#Builds the dynamic menu system for selecting the user who owns the calendar
Function SelectUserMenu{
    #This will print out the index for the user and then the users name.
    #Example: 1. DBemrose
    foreach($script:user in $script:users){
        '{0} - {1}' -f ($script:users.indexof($script:user) + 1), $script:user
    }

    #Creates an empty variable which is used within a while loop that checks if it is empty.
    #This is used to wait for the user input as when the user provides input it will cause the variable to no longer be empty thus breaking the while loop
    $script:UserSelectedIndex = $null
    while ([string]::IsNullOrEmpty($script:UserSelectedIndex)){
        #Creates a blank line for better readabilty
        Write-Host
        #Waits for the user input
        $script:UserSelectedIndex = Read-Host 'Select a user'
        #If the value provided is not within the dynamic menu it will set the check variable back to empty and will state the range of valid choices
        if ($script:UserSelectedIndex -notin 1..$script:users.Count){
            Write-Warning ('    Your selection [ {0} ] is not valid' -f $script:UserSelectedIndex)
            Write-Warning ('    The valid choices are 1 through {0}.' -f $script:users.count)
            Write-Warning  '    Please try again'
            Pause
            $script:UserSelectedIndex = $null
        }
    }
    #Creates the selected user variable based on the users input. The reason why it is minusing 1 is due to the fact that index's start at 0 but the menu system starts at 1
    #This is for user readability
    $Script:UserSelected = $Script:UsersTable[($script:UserSelectedIndex -1)]
    #Prints out the value the user selected and will then run the selection menu function
    Write-Host ('You selected {0}' -f $Script:UserSelected)
    SelectionMenu
}

#Exact same as the above script but with changed variable names so it can be used for the user to adjust the permissions for.
Function SelectUserMenua{
    foreach($script:usera in $script:usersa){
        '{0} - {1}' -f ($script:usersa.indexof($script:usera) + 1), $script:usera
    }

    $script:UserSelectedIndexa = $null
    while ([string]::IsNullOrEmpty($script:UserSelectedIndexa)){
        Write-Host
        $script:UserSelectedIndexa = Read-Host 'Select a user'
        if ($script:UserSelectedIndexa -notin 1..$script:usersa.Count){
            Write-Warning ('    Your selection [ {0} ] is not valid' -f $script:UserSelectedIndexa)
            Write-Warning ('    The valid choices are 1 through {0}.' -f $script:usersa.count)
            Write-Warning  '    Please try again'
            Pause
            $script:UserSelectedIndexa = $null
        }
    }
    $Script:UserSelecteda = $Script:UsersTablea[($script:UserSelectedIndexa -1)]
    Write-Host ('You selected {0}' -f $Script:UserSelecteda)
    Return
}

#Creates a dynamic menu for selecting the calendar
Function SelectCalendarMenu{
    foreach($Script:UserCalendar in $Script:UserCalendars){
        '{0} - {1}' -f ($Script:UserCalendars.IndexOf($Script:UserCalendar) + 1),$Script:UserCalendar
    }

    $Script:CalendarSelectedIndex = $null
    While ([string]::IsNullOrEmpty($Script:CalendarSelectedIndex)){
        Write-Host
        $Script:CalendarSelectedIndex = Read-Host 'Select a calendar'
        if ($Script:CalendarSelectedIndex -notin 1..$Script:UserCalendars.count){
            Write-Warning ('    Your selection [ {0} ] is not valid' -f $Script:CalendarSelectedIndex)
            Write-Warning ('    The valid choices are 1 through {0}.' -f $Script:UserCalendars.count)
            Write-Warning  '    Please try again'
            Pause
            $Script:CalendarSelectedIndex = $null
        }
    }
    $Script:CalendarSelected = $Script:UsersCalendarsTable[($Script:CalendarSelectedIndex -1)]
    $Script:CalendarSelectedR = $Script:CalendarSelected.Insert($Script:CalendarSelected.IndexOf('\'),':')
    Write-Host ('You selected {0}' -f $Script:CalendarSelected)
    Write-Host ('This is interpreted as {0}' -f $Script:CalendarSelectedR)
    SelectionMenu
}



#Enable Script Logging
EnableLogging
#Initial call to the main menu
MainMenu