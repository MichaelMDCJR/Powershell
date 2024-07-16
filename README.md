I have been working on this script for around a week in tandem with Power Automate script to take members of an AD OU, put them into Excel files, and then send emails to the correct people to verify the users. 
Here is how to modify the script to work with any AD 

This Powershell file was created to automate the process of going through Active Directory and recording a group's members into an excel file. Currently it is set up to go through the OU found at MCN\MCN_Groups, going into each of the sub OUs and getting the members from the groups located there. It also grabs from another specific OU that can be found at AT\Cloud_Services\MCN\MCN_Managed, along with the OU inside of it, ResourceCalendars. The main block of code that covers everything but the specific OUs and groups, is from line 1 to 152. The specific groups are covered in two for loops below it. 

However, here is how to modify the main file to work with any OU:

To start we have this line at the beginning of the file:
$listOUs = (Get-ADOrganizationalUnit -LDAPFilter "(name=*)" -searchbase "OU=MCN_Groups, OU=MCN, DC=ad, DC=ilstu, DC=edu" -SearchScope OneLevel -Properties *).Name
This gets a list of all the OUs we want to iterate through. In this case, it is all the OUs under MCN_Groups. To select a different OU, lay out the file path like so: OU=childfolder ,OU=parentfolder

We then have this line that gets the groups from each of these OUs previously grabbed
$groups = (Get-ADGroup -Filter * -searchbase "OU=$childOU, OU=MCN_Groups, OU=MCN, DC=ad, DC=ilstu, DC=edu" -Properties *).Name
Again, change the OU path to match the one given at the top of the file. 
You will also have to change the path around line 107 to match, to start it should look something like this:
$groupName = "CN=$group, OU=$childOU, OU=MCN_Groups, OU=MCN, DC=ad, DC=ilstu, DC=edu"

There is also a variable to store the path (around line 25) where we want to store our generated excel files. I recommend storing all of the files in a single folder, which can then be leveraged with a Power Automate script I created to automatically send emails to the correct person for review.
$Path = "C:\Users\ULID\OneDrive - IL State University\Documents\AD Group Members\$group.csv"
Here, I have the files going into a folder called 'AD Group Members' in my Onedrive, but you will have to change this to fit your needs. You will also have to change this line at the other location it appears in the main block of code, around lines 113.


By changing all these values, you can search through any given OU and send the results to any given file. However from line 154 onward, we have two for loops that cover specific scenarios. Should you want to get a select few groups from a specific OU path, I recommend copying one of the for loops and pasting it in the file to modify. 

The first for loop gets 8 specific groups found at the OU path AT\Cloud_Services\MCN\MCN_Managed. To alter the list of groups collected, add or remove group names from '$specificGroupArray.'

If you are getting groups from a different OU path, you will also have to change this line
$groupName = "CN=$group, OU=MCN_Managed, OU=MCN, OU=Cloud_Services, OU=AT, DC=ad, DC=ilstu, DC=edu"

Again, change the path to place files where you want, around line 165.


The second for loop will mostly likely be the most helpful as it grabs all of the groups in a single OU. You will have to change the OU path to match where you want to go with the variables '$groupsRC' and 'groupName'

Again, change the path, and that should be it!


