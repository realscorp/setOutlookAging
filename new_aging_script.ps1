#############################################################
##        Set aging recursively on Outlook folder          ##
##                     +++++++++++++++                     ##
##               Sergey Krasnov, 27.09.2017                ##
#############################################################

cls
Write-Output "=================================================================="

# Add Outlook .Net class
try
{
    Add-Type -assembly "Microsoft.Office.Interop.Outlook"
}
catch
{
    Write-Output "ERROR: Cannot add Microsoft.Office.Interop.Outlook assembly. Exit now."
    Exit
}


# Create COM-Object connection
try
{
    $app = New-Object -ComObject Outlook.Application
}
catch
{
    Write-Output "ERROR: Cannot Create COM-Object connection Outlook.Application. Exit now."
    Exit
}

# Add namespace
try
{
    $namespace = $app.GetNamespace("MAPI")

}
catch
{
    Write-Output "ERROR: Cannot Create MAPI connection. Exit now."
    Exit
}


# Set MAPI attributes we want to change in Outlook folders
# If you are planning GPO script deployment you only need $PR_AGING_DEFAULT
$PR_AGING_AGE_FOLDER = "http://schemas.microsoft.com/mapi/proptag/0x6857000B"
$PR_AGING_PERIOD = "http://schemas.microsoft.com/mapi/proptag/0x36EC0003"
$PR_AGING_GRANULARITY = "http://schemas.microsoft.com/mapi/proptag/0x36EE0003"
$PR_AGING_DELETE_ITEMS = "http://schemas.microsoft.com/mapi/proptag/0x6855000B"
$PR_AGING_FILE_NAME_AFTER9 = "http://schemas.microsoft.com/mapi/proptag/0x6859001E"
$PR_AGING_DEFAULT = "http://schemas.microsoft.com/mapi/proptag/0x685E0003"

# Accesibility MAPI attribute which is needed cause if we'll try to change public folder setting we'll get an exception
$PR_ACCESS = "http://schemas.microsoft.com/mapi/proptag/0x0FF40003"

# Set aging parameters
# $Granularity - sets what units is used for aging time
# 0 - months, 1 - weeks, 2 - days
#
# $Period - set aging time period. Elements that old is treated as expired.
# you can set 0-999
#
# $AgeFolder - allow folder aging
#
# $DeleteItems - allow expired items deletion
#
# $Default - should we use default settings
# 0 - None default setting is used
# 1 - Archive file default location is used only. Flags are ticked: 
# "Archive this folder using this settings" and "Move old items to default archive folder" 
# 3 - Every setting is set to default and flag ticked:
# "Archive items in this folder using the default settings" 
$Granularity=2
$Period=6
$AgeFolder=$True
$Default=3
$DeleteItems=$false

# Check if parameters is invalid
if ($Granularity -lt 0 -or $Granularity -gt 2) 
{
    Write-Output "ERROR: Invalid Granularity value: $Granularity. Exit now."
    Exit
}

if ($Period -lt 1 -or $Period -gt 999)
{
    Write-Output "ERROR: Invalid Period value: $Period. Exit now."
    Exit
}

if ($Default -lt 0 -or $Default -eq 2 -or $Default -gt 3)
{
    Write-Output "ERROR: Invalid Default value: $Default. Exit now."
    Exit
}


# Aging function that would be called for every folder
Function SetAgingOnFolder ($objFolder)
{
    # Check if input paramaters is invalid
	if ($objFolder -eq $null)
	{
		return $False
        Write-Output ("ERROR: Invalid Folder input object passed to (SetAgingOnFolder), folder name:" + $objFolder.Name.ToString())
	}
	
	Try
	{
        # Getting access to folder hidden item thas stores aging settings
		$oStorage = $objFolder.GetStorage("IPC.MS.Outlook.AgingProperties", 2)
        # Getting access to item's properties interface
		$oPA = $oStorage.PropertyAccessor
        
		If ($AgeFolder -eq $false)
		{
			oPA.SetProperty PR_AGING_AGE_FOLDER, False
		}
		Else
		{
			# Set aging properties. You should uncomment what you want to use. In my case it is $PR_AGING_DEFAULT only
			#$oPA.SetProperty($PR_AGING_AGE_FOLDER, $True)
			#$oPA.SetProperty($PR_AGING_GRANULARITY, $Granularity)
			#$oPA.SetProperty($PR_AGING_DELETE_ITEMS, $DeleteItems)
			#$oPA.SetProperty($PR_AGING_PERIOD, $Period)
			$oPA.SetProperty($PR_AGING_DEFAULT, $Default)
            
	    		# Set Filename property if it's not Null
			if ($FileName -ne $null)
			{
				 #$oPA.SetProperty($PR_AGING_FILE_NAME_AFTER9, $FileName)
			}
		}

		# Save hidden item's changes and return function success
		$oStorage.Save()
        Write-Output ("SUCCESS: Set aging properties on folder >>> " + $objFolder.Name.ToString())
	}
	
    # Catch errors
	Catch
	{
        Write-Output ("ERROR: Something went wrong in (SetAgingOnFolder) function on folder: " + $objFolder.Name.ToString())
		Write-Output $_
	}
}

# Function that will go through all the folders recursively and call for SetAgingOnFolder function
Function RecursiveRoutine ($objParent)
{
#    SetAgingOnFolder ($objFolder)
    $objChildren = $objParent
    foreach ($objChild in $objChildren)
    {
        SetAgingOnFolder ($objChild)
        RecursiveRoutine ($objChild.folders)
    }
}

# Main
###############
Write-Output "Searching for set up Outlook profile"

# Check if we have Outlook profile set up
try
{
    if ($app.DefaultProfileName -eq $null)
    {
        Write-Output "ERROR: We don't have any Outlook profile. Exit now."
        Exit
    }
}
catch
{
    Write-Output "ERROR: Something went wrong when we searched for Outlook profile. Exit now."
    Exit
}


# Check if we have set up Email account in profile
try
{
    $accCount = ($app.Explorers.Session.Accounts | Measure-Object).count
    if ($accCount -eq 0)
    {
        Write-Output "ERROR: No Email account configured. Exit now."
        Exit
    }
}
catch
{
    Write-Output "ERROR: Something went wrong when we searched for Email accounts. Exit now."
    Exit
}

# Get folder objects list
try
{
    $objAllfolders = $namespace.Session.Folders
}
catch
{
    Write-Output "ERROR: Can't get folder list. Exit now."
    Exit
}


# For every root folder except of Public ones we run recursive routine
ForEach ($objFolder in $objAllfolders) 
{
    Write-Output " "
    Write-Output "====================================================="
    Write-Output ("Enter folder root: " + $objFolder.Name)

    try
    {
        $accessLevel = $objFolder.PropertyAccessor.GetProperty($PR_ACCESS)
    }
    catch
    {
        Write-Output "ERROR: Can't get PR_ACCESS attribute, skip this folder"
        Continue
    }

    if ($accessLevel -eq 63)
    {
        RecursiveRoutine $objFolder.Folders
    }
    else
    {
        Write-Output "System folder, skipping..."
        Continue
    }
}
