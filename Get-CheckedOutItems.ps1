<#
.SYNOPSIS
    This script goes through site collections on a SharePoint farm and searches through document libraries to find checked out files and files with no check in version.
.DESCRIPTION
    This script will iterate through site collections on a farm to find all checked out files.
    When it finds a file that is checked out, it stores the file information (value) with the email of the user who has it checked out (key).
    When it is finshed searching the sites, it iterates through the hashtable's keys and sends an email of all users and all checked out files to a specified user.
    It also has the option to send each user a personal list of the files that they have checked out along with a link to edit the properties of the file and check it in.
    Currently it writes all progress output to the console but by uncommenting certain lines, it can also write to an output txt file.
.NOTES
    File Name:   Get-CheckedOutItems.ps1
       Author:   Owen Brown - Co-op Student
         Term:   Summer 2019
#>

# start a timer to keep track of how long it takes the script to run
$timer = [system.diagnostics.stopwatch]::StartNew()

# variables for the number of days since check out to alert users
$fullListAlerts = 90    # full list of all checked out items
$userListAlerts = 45    # user list of just their checked out items

# add the Snapin required to work with SharePoint objects
Add-PSSnapin Microsoft.SharePoint.PowerShell
$webApp = Get-SPWebApplication https://my.SharePointSite.com

# user login credentials - used for specific searches (DOMAIN\username)
#$userlogin = "DOMAIN\username"

# output file for writing results
#$Output = "C:\Desktop\output.txt"
#Clear-Content $Output

# print the time of the search at the beginning
$now = Get-Date 
Write-Host ("Date: {0} `r`n `r`n" -f $now)
#Add-Content $Output ("Date: {0} `r`n `r`n" -f $now)

# create an array to store the file objects
$checkedOutItems = New-Object Collections.Generic.List[PSObject]

# iterate through each site collection on the server farm (web applications)
foreach ($site in $webApp.Sites)
{

  <# skip specific sites ***(USED FOR FASTER TESTING)***
   if($site.Url -like '*/it' -or $site.Url -like '*/hr' -or $site.Url -like '*/finance')
   { 
      continue 
   }
#>

   # display current site being scaned
   Write-Host "Scanning Site: $($site.Url)" -ForegroundColor Cyan
#   Add-Content $Output ("Scanning Site: {0}" -f $site.URL)

   # iterate through each site in the site collection
   foreach ($web in $site.AllWebs)
   {
       # count of checked lists for progress bar
       $checkedLists = 0;

       # get total number of lists/libraries on the site
       $allLists = $web.Lists | ? { $_.BaseType -eq [Microsoft.SharePoint.SPBaseType]::DocumentLibrary }
       $listCount = $allLists.Count

       # iterate through each list in the site
       foreach ($list in $allLists)
       {
           # display current list being scanned
           Write-Host "    Scanning List: $($list.RootFolder.ServerRelativeUrl)" -ForegroundColor Green
#           Add-Content $Output ("    Scanning List: {0}" -f $list.RootFolder.ServerRelativeUrl)

           # add to counter after list is scanned
           $checkedLists++;

           # progress bar
           Write-Progress -Activity ("Search in progress. {0}" -f $site.Url) `
                -Status ("{0:0}% Complete" -f ($checkedLists/$listCount*100)) `
                -PercentComplete ($checkedLists/$listCount*100)




           # iterate through items with no check in version in the current list
           foreach ($item in $list.CheckedOutFiles)
           {  
               Write-Host "--------------------------------------------------------------------"
               Write-Host "             File: " $item.Url -ForegroundColor Yellow
               Write-Host "   Checked Out To: " $item.CheckedOutBy -ForegroundColor Yellow
               Write-Host "            Email: " $item.CheckedOutByEmail -ForegroundColor Yellow
               Write-Host "--------------------------------------------------------------------"
<#             
               # write to output file
               Add-Content $Output "--------------------------------------------------------------------"
               Add-Content $Output ("              File: {0}" -f $item.Url)
               Add-Content $Output ("    Checked Out To: {0}" -f $item.CheckedOutBy)
               Add-Content $Output ("             Email: {0}" -f $item.CheckedOutByEmail)
               Add-Content $Output "--------------------------------------------------------------------"
#>             
               
               # create an object to hold the file's information
               $checkedOutItem = New-Object PSObject
               
               # check if the user does not have an email 
               if($item.CheckedOutByEmail -eq "")
               {
                   # no email so add the user's name
                   $checkedOutItem | Add-Member -Type NoteProperty -Name Email -Value $item.CheckedOutByName
               } 
               else
               {
                   # email found so add the user's email
                   $checkedOutItem | Add-Member -Type NoteProperty -Name Email -Value $item.CheckedOutByEmail
               }
               
               # store the file info as properties of the checkedOutItem data type
               # the variable is the data that the user will see in the email and the variableLink is the web link for that data
               $siteUrl = $site.Url -split "/"
               $s = $siteUrl[4]
               $checkedOutItem | Add-Member -Type NoteProperty -Name Site -Value ("{0}" -f $s)
               $checkedOutItem | Add-Member -Type NoteProperty -Name SiteLink -Value $site.Url
               
               $lib = $item.Url -split "/"
               $checkedOutItem | Add-Member -Type NoteProperty -Name Library -Value $lib[2]
               $checkedOutItem | Add-Member -Type NoteProperty -Name LibraryLink -Value ("{0}/{1}" -f $site.Url, $lib[2])
               
               # shorten then file name if it is longer than 20 chars
               if($item.LeafName.length -gt 20)
               {
                   $checkedOutItem | Add-Member -Type NoteProperty -Name File -Value ("{0}..." -f $item.LeafName.SubString(0,20))
               }
               else
               {
                   $checkedOutItem | Add-Member -Type NoteProperty -Name File -Value ($item.LeafName)
               }
               $checkedOutItem | Add-Member -Type NoteProperty -Name FileFull -Value ($item.LeafName)# needed a file name that is full for the individual user list because we dont need to shorten it
               $checkedOutItem | Add-Member -Type NoteProperty -Name FileLink -Value ("{0}//{1}/{2}" -f $siteUrl[0], $site.HostName, $item.Url)
               
               $checkedOutItem | Add-Member -Type NoteProperty -Name Version -Value "None"
		       $checkedOutItem | Add-Member -Type NoteProperty -Name DispUrl -Value ("{0}//{1}/{2}/Forms/DispForm.aspx?ID={3}" -f $siteUrl[0], $site.HostName, $item.DirName, $item.ListItemID)
               
               # number of days the file has been checked out for
               $today = Get-Date
               $date = $item.TimeLastModified
               $daysSince = (New-TimeSpan -Start $date -End $today).Days
               $checkedOutItem | Add-Member -Type NoteProperty -Name daysSinceCheckOut -Value $daysSince

               # format the date the file was checked out for the email
               $dispDate = $date.ToString("yyyy/MM/dd")
               $checkedOutItem | Add-Member -Type NoteProperty -Name checkedOutSince -Value $dispDate
               
               
               # add file object to the array
               $checkedOutItems.Add($checkedOutItem)
           }




           # iterate through all items in the current list that do have checked in/published versions
           foreach ($item in $list.Items)
           {
               # check if file is not checked out by anyone or checked out by a specific person and skip it
               if(!$item.File.CheckedOutByUser -or $item.File.CheckedOutByUser -like "SHAREPOINT\system") # in this case, all custom content types and folders are considered as files checked out by "SHAREPOINT\system" so skip those
               {
                   continue
               }

               # check if file is a record and skip it
               if([Microsoft.Office.RecordsManagement.RecordsRepository.Records]::IsRecord($item))
               {
                   continue
               }

               # check if the item is checked out        & filter by a specific username (optional)
               if ($item.File.CheckOutStatus -ne "None" )# -and $item.File.CheckedOutByUser -like $userlogin)
               {
                   # write to console
                   Write-Host "--------------------------------------------------------------------"
                   Write-Host "             File: " $item.Url -ForegroundColor Yellow
                   Write-Host "   Checked Out To: " $item.File.CheckedOutByUser -ForegroundColor Yellow
                   Write-Host "            Email: " $item.File.CheckedOutByUser.Email -ForegroundColor Yellow
                   Write-Host "--------------------------------------------------------------------"
<#
                   # write to output file
                   Add-Content $Output "--------------------------------------------------------------------"
                   Add-Content $Output ("              File: {0}" -f $item.File.Url)
                   Add-Content $Output ("    Checked Out To: {0}" -f $item.File.CheckedOutByUser)
                   Add-Content $Output ("             Email: {0}" -f $item.File.CheckedOutByUser.Email)
                   Add-Content $Output "--------------------------------------------------------------------"
#>

                   # create an object to hold the file's information
                   $checkedOutItem = New-Object PSObject

                   # check if the user does not have an email 
                   if($item.File.CheckedOutByUser.Email -eq "")
                   {
                       # no email so add the user's name
                       $checkedOutItem | Add-Member -Type NoteProperty -Name Email -Value $item.File.CheckedOutByUser.Name
                   } 
                   else
                   {
                       # email found so add the user's email
                       $checkedOutItem | Add-Member -Type NoteProperty -Name Email -Value $item.File.CheckedOutByUser.Email
                   }

                   # store the file info as properties of the checkedOutItem data type
                   # the plain variable is the data that the user will see in the email and the variableLink will be the link for that data
                   $siteUrl = $site.Url -split "/"
                   $checkedOutItem | Add-Member -Type NoteProperty -Name Site -Value $siteUrl[4]
                   $checkedOutItem | Add-Member -Type NoteProperty -Name SiteLink -Value $site.Url

                   $lib = $item.Url -split "/"
                   $checkedOutItem | Add-Member -Type NoteProperty -Name Library -Value $lib[0]
                   $checkedOutItem | Add-Member -Type NoteProperty -Name LibraryLink -Value ('{0}/{1}' -f $site.Url, $lib[0])

                   # shorten then file name if it is longer than 20 chars
                   if($item.Name.length -gt 20)
                   {
                       $checkedOutItem | Add-Member -Type NoteProperty -Name File -Value ("{0}..." -f $item.Name.SubString(0,20))
                   }
                   else
                   {
                       $checkedOutItem | Add-Member -Type NoteProperty -Name File -Value ($item.Name)
                   }
                   $checkedOutItem | Add-Member -Type NoteProperty -Name FileFull -Value ($item.Name)
                   $checkedOutItem | Add-Member -Type NoteProperty -Name FileLink -Value ("{0}/{1}" -f $web.Url, $item.File.Url)
                   
                   $checkedOutItem | Add-Member -Type NoteProperty -Name Version -Value $item.File.UIVersionLabel
                   $checkedOutItem | Add-Member -Type NoteProperty -Name DispUrl -Value ("{0}//{1}{2}?ID={3}" -f $siteUrl[0], $site.HostName, $item.ParentList.DefaultDisplayFormUrl, $item.ID)

                   # number of days the file has been checked out for
                   $today = Get-Date
                   $date = $item.File.CheckedOutDate
                   $daysSince = (New-TimeSpan -Start $date -End $today).Days
                   $checkedOutItem | Add-Member -Type NoteProperty -Name daysSinceCheckOut -Value $daysSince
                   
                   # format the date the file was checked out for the email
                   $dispDate = $date.ToString("yyyy/MM/dd")
                   $checkedOutItem | Add-Member -Type NoteProperty -Name checkedOutSince -Value $dispDate

                   # add file object to the array
                   $checkedOutItems.Add($checkedOutItem)
               }
           }
       }
   }
}

# create the hash table and set it empty
$hTable = @{}

# loop through the object array
foreach ($item in $checkedOutItems)
{
<#
    # write to output file
    Add-Content $Output (" File: {0}" -f $file)
    Add-Content $Output ("Email: {0}" -f $item.Email)
    Add-Content $Output ("-----------------------------------------")
#>

    # set the user's email/name to the key of the hashtable
    $key = $item.Email

    # add the file object to the value of the appropriate key in the hashtable
    $hTable.$key += @($item)
}

# clear the message variables before writing the email
$msgBody = "";
$fullBody = "";



# add notes to the messages
$fullBody += "<p style='font-size: 18px'>
                This email contains all checked out files from " + $webApp.Url + ". <br>
                A yellow box means that the file has no checked in version and cannot be seen by anyone. <br>
                A red box means that the file has been checked out for over " + $fullListAlerts + " days.
              </p>"



# loop through the keys in the hastable to sort the data
foreach($userEmail in $hTable.keys)
{
    # fullBody is for the full list of all users
    $fullBody += "-----------------------------------------------------------------------"
    $fullBody += "<h3>" + $userEmail + "</h3>
        <table style='border: 1px solid black; border-collapse: collapse'>
            <tr>
                <th style='width: 200px; border: 1px solid black; padding-left: 5px; background-color: #dae0eb; text-align: left'>File</th>
                <th style='width:  70px; border: 1px solid black; padding-left: 5px; background-color: #dae0eb; text-align: left'>Version</th>
		        <th style='width: 110px; border: 1px solid black; padding-left: 5px; background-color: #dae0eb; text-align: left'>Site</th>
		        <th style='width: 170px; border: 1px solid black; padding-left: 5px; background-color: #dae0eb; text-align: left'>Library</th>
		        <th style='width: 120px; border: 1px solid black; padding-left: 5px; background-color: #dae0eb; text-align: left'>Edit/Check In</th>
                <th style='width: 120px; border: 1px solid black; padding-left: 5px; background-color: #dae0eb; text-align: left'>Last Modified</th>
            </tr>"

    # msgBody is for the list of individual users
    $msgBody += "<h3>" + $userEmail + "</h3>"

    $msgBody += "<p style='font-size: 18px'> 
                    This email contains your checked out files on SharePoint.<br>
                    A yellow box means that the file has no checked in version and cannot be seen by anyone but you.<br>
                    A red box means that the file has been checked out for over " + $userListAlerts + " days.<br>
                    Please save any changes you have made to the file and check it in so others can have access to it.<br>
                    If there are missing required property fields, please make sure to fill them in.
                 </p>"

    # add each file checked out to each user to the list
    foreach($value in $hTable.$userEmail)
    {
        $fullBody += "<tr>
                        <td style='width: 200px; padding-left: 5px; border: 1px solid black; text-align: left'><a href='" + $value.FileLink + "'>" + $value.File + "</a></td>"

        # if a file has no checked in version, colour the background yellow to draw attention
        if($value.Version -like "None")
        {
            $fullBody += "<td style='width: 70px; padding-left: 5px; border: 1px solid black; text-align: left; background-color: #f0da60'>" + $value.Version + "</td>"
        }
        else
        {
            $fullBody += "<td style='width: 70px; padding-left: 5px; border: 1px solid black; text-align: left'>" + $value.Version + "</td>"
        }                
        
        $fullBody +=   "<td style='width: 110px; padding-left: 5px; border: 1px solid black; text-align: left'><a href='" + $value.SiteLink + "'>" + $value.Site + "</a></td>
                        <td style='width: 170px; padding-left: 5px; border: 1px solid black; text-align: left'><a href='" + $value.LibraryLink + "'>" + $value.Library + "</a></td>
                        <td style='width: 120px; padding-left: 5px; border: 1px solid black; text-align: left'><a href='" + $value.DispUrl + "'>Edit/Check In</a></td>"

        # if a file has been checked out for over $fullListAlerts days, colour the background red to draw attention (for the full list)
        if($value.daysSinceCheckOut -gt $fullListAlerts)
        {
            $fullBody += "<td style='width: 120px; padding-left: 5px; border: 1px solid black; text-align: left; background-color: #ff7873'>" + $value.checkedOutSince + "</td>"
        }
        else
        {
            $fullBody += "<td style='width: 120px; padding-left: 5px; border: 1px solid black; text-align: left'>" + $value.checkedOutSince + "</td>"
        }        
        $fullBody += "</tr>"

        
        $msgBody += "<table style='border: 1px solid black; border-collapse: collapse'>
                        <tr>
                           <th style='width: 120px; border: 1px solid black; padding-right: 5px; background-color: #dae0eb; text-align: right'>File:</th>
                           <td style='width: 500px; padding-left: 5px; border: 1px solid black'><a href='" + $value.FileLink + "'>" + $value.FileFull + "</a></td>
                        </tr>
                        <tr>
                           <th style='width: 120px; border: 1px solid black; padding-right: 5px; background-color: #dae0eb; text-align: right'>Site:</th>
                           <td style='width: 500px; padding-left: 5px; padding-right: 3px; border: 1px solid black'><a href='" + $value.SiteLink + "'>" + $value.Site + "</a></td>
                        </tr>
                        <tr>
                           <th style='width: 120px; border: 1px solid black; padding-right: 5px; background-color: #dae0eb; text-align: right'>Library:</th>
                           <td style='width: 500px; padding-left: 5px; padding-right: 3px; border: 1px solid black'><a href='" + $value.LibraryLink + "'>" + $value.Library + "</a></td>
                        </tr>
                        <tr> 
                           <th style='width: 120px; border: 1px solid black; padding-right: 5px; background-color: #dae0eb; text-align: right'>Version:</th>"

        # if a file has no checked in version, colour the background yellow to draw attention
        if($value.Version -like "None")
        {
            $msgBody +=   "<td style='width: 500px; padding-left: 5px; border: 1px solid black; background-color: #f0da60'>No Checked In Version</td>"
        }
        else
        {
            $msgBody +=   "<td style='width: 500px; padding-left: 5px; border: 1px solid black'>" + $value.Version + "</td>"
        }
                        
        $msgBody +=    "</tr>
                        <tr>
                           <th style='width: 120px; border: 1px solid black; padding-right: 5px; background-color: #dae0eb; text-align: right'>Last Modified:</th>"
        
        # if a file has been checked out for over $userListAlerts days, colour the background red to draw attention (for the user list)
        if($value.daysSinceCheckOut -gt $userListAlerts)
        {
            $msgBody +=   "<td style='width: 500px; padding-left: 5px; border: 1px solid black; background-color: #ff7873'>" + $value.checkedOutSince + "</td>"
        }
        else
        {                   
            $msgBody +=   "<td style='width: 500px; padding-left: 5px; border: 1px solid black'>" + $value.checkedOutSince + "</td>"
        }
        
        $msgBody +=    "</tr>
                        <tr>
                           <th style='width: 120px; border: 1px solid black; padding-right: 5px; background-color: #dae0eb; text-align: right'>Edit/Check In:</th>
                           <td style='width: 500px; padding-left: 5px; border: 1px solid black'><a href='" + $value.DispUrl + "'>Edit Properties/Check In</a></td>
                        </tr>                        
                     </table>
                     <br/><br/>"
    }
    $fullBody += "</table>"

    # check if the user has a company email
    if($userEmail -like "*mycompany.com") 
    {
        # send the individual user an email with all of their checked out files listed
#############        $send = [Microsoft.SharePoint.Utilities.SPUtility]::SendEmail($web,0,0,$userEmail,"Your Checked Out Files", $msgBody.ToString()) ############# uncomment when ready to send to all users
    } # else: skip the user because they either dont have a simcoe email or it's an adm account
    
    # reset the msgBody variable for the next user
    $msgBody = ""
}

# send the full list
$send = [Microsoft.SharePoint.Utilities.SPUtility]::SendEmail($web,0,0,"myemail@mycompany.ca","All Checked Out Files", $fullBody.ToString())

# stop the timer and display how long the script took to execute
$timer.Stop()
Write-Host ("Finished in {0:00}:{1:00}" -f $timer.Elapsed.Minutes, $timer.Elapsed.Seconds)



