# windows-migrate-ad-accounts
This program will take an Excel document with a list of Active Directory users, and their AD properties, and write New-ADUser and Set-ADUser commands for each user to the console.

The excel spreadsheet can be created from the output of the Get-ADUser listing the users of the domain to be migrated.  In addition to creating the user account, the Set-ADUser sets the new account with the old domain email address as a second ProxyAddress.

The program considers that some users do not have all the properties set, so for each New-ADUser command it only includes the properties present for the user.

The utility is useful for migrating users between Active Directory organizations that are not connected with each other in any way.
