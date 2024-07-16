# The main OU that contains all of our sub OUs that have our groups
$listOUs = (Get-ADOrganizationalUnit -LDAPFilter "(name=*)" -searchbase "OU=MCN_Groups, OU=MCN, DC=ad, DC=ilstu, DC=edu" -SearchScope OneLevel -Properties *).Name

# Loops through the childOUs and creates csvs of the groups that each contains
foreach ($childOU in $listOUs) {
    # Groups in given child OU
    $groups = (Get-ADGroup -Filter * -searchbase "OU=$childOU, OU=MCN_Groups, OU=MCN, DC=ad, DC=ilstu, DC=edu" -Properties *).Name

    # Loop through the groups and put all the members into a csv
    foreach($group in $groups) {
        <#
            Because we want to combine groups with _modify and _read to a single csv, we have to do some logic
            If there is a group that ends with _read, there will always be a group that ends with _modify. If this changes in the future, this will have to be updated
            However not every group that ends with _modify will have a counterpart with _read
        #>

        # Here we grab groups that end with modify and also get their would be _read counterpart
        if($group -like "*_modify") {
            $readName = $group -replace "_modify", "_read"
            $modGroupMembers = Get-ADGroupMember -Identity $group -Recursive
      
            # Define the output path for the CSV file
            #$Path = "C:\Users\isu_mdcarl2\OneDrive - IL State University\Documents\AD Group Members\$group.csv"
            # A path must be chosen, I recommend putting all the created excel files in a folder, here that folder is 'AD Group Members'
            $Path = "C:\Users\ULID\OneDrive - IL State University\Documents\AD Group Members\$group.csv"

            # Initialize an empty array to store custom objects
            $users = @()

            # Iterate through each modify group member
            foreach ($member in $modGroupMembers) {
                # Check if the member is a user
                if ($member.objectClass -eq "user") {
                    # Retrieve user name
                    $user = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName

                    # Create a custom object with ID and DisplayName properties
                    $userObject = [PSCustomObject]@{
                        ULID = $member.Name
                        DisplayName = $user.DisplayName
                        Access = "modify"
                    }

                    # Add the custom object to the array
                    $users += $userObject
                }
            }

            # Ensure that the read file does exist as there are some files with only _modify
            try
            {
                $readGroupMembers = Get-ADGroupMember -Identity $readName -Recursive
                If($readGroupMembers)
                {
                    Write-Host "Group exists"

                    # Iterate through each _read group members
                    foreach ($member in $readGroupMembers) {
                        # Check if the member is a user
                        if ($member.objectClass -eq "user") {
                            # Retrieve user name
                            $user = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName

                            # Create a custom object with ID and DisplayName properties
                            $userObject = [PSCustomObject]@{
                                ULID = $member.Name
                                DisplayName = $user.DisplayName
                                Access = "read"
                            }

                            # Add the custom object to the array
                            $users += $userObject
                        }
                    }

                }
            }

            # If there is no _read group, print "Group does not exist" and do nothing
            catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
            {
                Write-Host "Group does not exist" 
            }
       
            # Export the array of objects to a CSV file
            
            # Export the array of objects to a CSV file
            if ($users[0].ULID -eq $null) {
                # Create a custom object with ULID and DisplayName properties
                $userObject = [PSCustomObject]@{
                    ULID = "N/A"
                    DisplayName = "No Users"
                }

                # Add the custom object to the array
                $users += $userObject
            }
       
            $users | Export-Csv -Path $Path -NoTypeInformation
        
  
        }
        # If the group does not end with _modify, check if it ends with _read, and if so, skip it
        else {
            if ($group -notlike "*_read") {
                #Get the group name and list of members
                $groupName = "CN=$group, OU=$childOU, OU=MCN_Groups, OU=MCN, DC=ad, DC=ilstu, DC=edu"
                $groupMembers = Get-ADGroupMember -Identity $groupName -Recursive

                # Define the output path for the CSV file
                #$Path = "C:\Users\isu_mdcarl2\OneDrive - IL State University\Documents\AD Group Members\$group.csv"
                # A path must be chosen, I recommend putting all the created excel files in a folder, here that folder is 'AD Group Members'
                $Path = "C:\Users\ULID\OneDrive - IL State University\Documents\AD Group Members\$group.csv"

                # Initialize an empty array to store custom objects
                $users = @()

                # Iterate through each group member
                foreach ($member in $groupMembers) {
                    # Check if the member is a user
                    if ($member.objectClass -eq "user") {
                        # Retrieve user name
                        $user = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName

                        # Create a custom object with ID and DisplayName properties
                        $userObject = [PSCustomObject]@{
                            ULID = $member.Name
                            DisplayName = $user.DisplayName
                        }

                        # Add the custom object to the array
                        $users += $userObject
                    }
                }

                # Export the array of objects to a CSV file
                if ($users[0].ULID -eq $null) {
                    # Create a custom object with ULID and DisplayName properties
                    $userObject = [PSCustomObject]@{
                        ULID = "N/A"
                        DisplayName = "No Users"
                    }
                    # Add the custom object to the array
                    $users += $userObject
                }
       
                $users | Export-Csv -Path $Path -NoTypeInformation
        
            }
        }
    }
}

# An array of all the specific groups we need to loop through
$specificGroupArray = @("MCNg_APSP_FullAccess", "MCNg_MCNAcademics", "MCNg_Twilio_FullAccess", "MCNg_PstLicenseClinic", "MCNg_MCNPostlicensureHealth", "MCNg_MCNPrelicensureHealth", "MCNg_MCNPrelicensureHealth_SendAs", "MCNg_TheFlame")

foreach($group in $specificGroupArray) {
        # Get the group name and list of members
        $groupName = "CN=$group, OU=MCN_Managed, OU=MCN, OU=Cloud_Services, OU=AT, DC=ad, DC=ilstu, DC=edu"
        $groupMembers = Get-ADGroupMember -Identity $groupName -Recursive
        
        # Define the output path for the CSV file
        #$Path = "C:\Users\isu_mdcarl2\OneDrive - IL State University\Documents\AD Group Members\$group.csv"
        # A path must be chosen, I recommend putting all the created excel files in a folder, here that folder is 'AD Group Members'
        $Path = "C:\Users\ULID\OneDrive - IL State University\Documents\AD Group Members\$group.csv"

        # Initialize an empty array to store custom objects
        $users = @()

        # Iterate through each group member
        foreach ($member in $groupMembers) {
            # Check if the member is a user
            if ($member.objectClass -eq "user") {
                # Retrieve user name
                $user = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName

                # Create a custom object with ULID and DisplayName properties
                $userObject = [PSCustomObject]@{
                    ULID = $member.Name
                    DisplayName = $user.DisplayName
                }

                # Add the custom object to the array
                $users += $userObject
            }
        }

        # Export the array of objects to a CSV file
        if ($users[0].ULID -eq $null) {
            # Create a custom object with ULID and DisplayName properties
                $userObject = [PSCustomObject]@{
                    ULID = "N/A"
                    DisplayName = "No Users"
                }

                # Add the custom object to the array
                $users += $userObject
        }
       
        $users | Export-Csv -Path $Path -NoTypeInformation
        

}

# For looping through a specific AD File, ResourceCalendars
$groupsRC = (Get-ADGroup -Filter * -searchbase "OU=ResourceCalendars, OU=MCN_Managed, OU=MCN, OU=Cloud_Services, OU=AT, DC=ad, DC=ilstu, DC=edu" -Properties *).Name

foreach($group in $groupsRC) {
        # Get the group name and list of members
        $groupName = "CN=$group, OU=ResourceCalendars, OU=MCN_Managed, OU=MCN, OU=Cloud_Services, OU=AT, DC=ad, DC=ilstu, DC=edu"
        $groupMembers = Get-ADGroupMember -Identity $groupName -Recursive
  
        # Define the output path for the CSV file
        #$Path = "C:\Users\isu_mdcarl2\OneDrive - IL State University\Documents\AD Group Members\$group.csv"
        # A path must be chosen, I recommend putting all the created excel files in a folder, here that folder is 'AD Group Members'
        $Path = "C:\Users\ULID\OneDrive - IL State University\Documents\AD Group Members\$group.csv"

        # Initialize an empty array to store custom objects
        $users = @()

        # Iterate through each group member
        foreach ($member in $groupMembers) {
            # Check if the member is a user
            if ($member.objectClass -eq "user") {
                # Retrieve user name
                $user = Get-ADUser -Identity $member.DistinguishedName -Properties DisplayName

                # Create a custom object with ULID and DisplayName properties
                $userObject = [PSCustomObject]@{
                    ULID = $member.Name
                    DisplayName = $user.DisplayName
                }

                # Add the custom object to the array
                $users += $userObject
            }
        }

        # Export the array of objects to a CSV file
        if ($users[0].ULID -eq $null) {
            # Create a custom object with ULID and DisplayName properties
                $userObject = [PSCustomObject]@{
                    ULID = "N/A"
                    DisplayName = "No Users"
                }

                # Add the custom object to the array
                $users += $userObject
        }
       
        $users | Export-Csv -Path $Path -NoTypeInformation
        

}
