#Requires -Version 5.1
# +--------------------------------------------------------------------------------------------------------------------------------------
# | File     :                                    
# | Version  : 2.1                                 
# | Purpose  : Update-MSEdgeManagedProfile
# | Synopsis : 
# | Usage    : ???
# +--------------------------------------------------------------------------------------------------------------------------------------
# | Maintenance History                                            
# | -------------------                                            
# | Name	           Date		      Version		C/R		Description        
# | -------------------------------------------------------------------------------------------------------------------------------------
# | Craig Cram		   2024-08-12	  1.8                  Inital Version
# | Craig Cram		   2024-08-14	  1.9                  Added Check for Sync Account Profile
# | Craig Cram		   2024-08-14	  1.10                 Updated logging to include found Edge Profiles
# | Craig Cram		   2025-06-13	  2.1                  Updated cloud only machines
# +--------------------------------------------------------------------------------------------------------------------------------------
# ***************************************************************************************************************************************

#
#
# Config Section
#
#

$sitelistFilename = "$($env:temp)\ManagedProfileList.csv"
$ManagedProfileName = "ITWCAP"
$sitelistURL = "https://raw.githubusercontent.com/ITWCAP-ITOps/ManagedEdgeProfiles/refs/heads/main/ManagedProfileList.csv"

$resultcode = Invoke-WebRequest -Uri $sitelistURL -OutFile $sitelistFilename 

## Script Logging Enabled
$logFolder = "$env:TEMP\MSEdgeManagedProfile"
if ($false -eq (Test-Path $logFolder)) { mkdir $logFolder -ErrorAction SilentlyContinue | out-null}
get-childitem -Path "$logFolder\" -Filter "MSEdgeManagedProfile*.log" | Sort-Object name | Select-Object -skiplast 9 | Remove-Item -filter *.log
$randomchar = -join ((65..90) + (97..122) | Get-Random -Count 5 | % {[char]$_})
$filenamedatetime = "$(get-date -uFormat '%Y-%m-%d-%H-%M-%S')"
$script:LogFile="$logFolder\MSEdgeManagedProfile-$randomchar-$filenamedatetime.log" 
##-----------------------------------------------------------------------
## Function: Out-ToLine
## Purpose: Used to write output to the screen
##-----------------------------------------------------------------------
function Out-ToLine {
	param ([string]$strOut, [string]$strType)
    
    #$strDateTime = get-date -uFormat "%d-%m-%Y %H:%M"
    #$strOut = $strDateTime + ": " + $strOut
    
	if ($strOut -match "ERR:" -or $strType -eq "Error") {
		    write-host $strOut -ForegroundColor RED 
	} elseif ($strOut -match "WARN:"  -or $strType -eq "WARNING") {
		    write-host $strOut -ForegroundColor Blue 
	} elseif ($strOut -match "WARNING:"  -or $strType -eq "WARNING") {
		    write-host $strOut -ForegroundColor Blue 
	} elseif ($strOut -match "INFO:"  -or $strType -eq "INFO") {
		    write-host $strOut -ForegroundColor CYAN 
	} elseif ($strOut -match "DEBUG:"  -or $strType -eq "DEBUG") {
		    write-host $strOut -ForegroundColor DARKGREEN -BackgroundColor WHITE
	} elseif ($strOut -match "HIGHLIGHT:"  -or $strType -eq "HIGHLIGHT") {
		    write-host $strOut -ForegroundColor Yellow
	} else {
		    write-host $strOut
	}
	Out-ToFile $strOut
}
##-----------------------------------------------------------------------
## Function: Out-ToFile
## Purpose: Used to write output to a log file
##-----------------------------------------------------------------------
function Out-ToFile {
	param ([string]$strOut)
	if ($script:LogFile) {"$strOut" | out-file -filepath $script:LogFile -append}
}
#
# Main Section
#
# Get profile lists from localstate file
$BrowserSettingsList = get-childitem "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\Local State" -ErrorAction SilentlyContinue| Select-Object -ExpandProperty fullname
$BrowserUserProfileList = [System.Collections.Generic.List[object]]::new()
foreach ($BrowserSettings in $BrowserSettingsList){
    if ($BrowserSettings -match '\\Users\\([^\\]+)\\') {
        $username = $matches[1]
    } else {
        $username = "Unknown"
    }

    $BrowserLocalstate = get-Content $BrowserSettings | convertfrom-json
    foreach ($BrowserUserProfile in $BrowserLocalstate.profile.info_cache.psobject.properties.name) {
        if ([Regex]::Match($BrowserSettings, "\\Local\\([^\\]+)\\").Groups[1].Value -eq "Microsoft") {
            $BrowserType = "Edge"
            $ProfileFullPath = "$($env:SystemDrive)\Users\$UserName\AppData\Local\Microsoft\Edge\User Data\$browserUserProfile\Preferences"
        } elseif ([Regex]::Match($BrowserSettings, "\\Local\\([^\\]+)\\").Groups[1].Value -eq "Google") {
            $BrowserType = "Google"
            $ProfileFullPath = "$($env:SystemDrive)\Users\$UserName\AppData\Local\Google\Chrome\User Data\$browserUserProfile\Preferences"
        }
        $BrowserUserProfileList.add([pscustomobject][ordered]@{
            Username        = $username
            Browser         = $BrowserType
            Profile         = $BrowserLocalstate.profile.info_cache.$BrowserUserProfile.shortcut_name
            ProfileFolder   = $BrowserUserProfile
            ProfileFullPath = $ProfileFullPath
            activeTime      = (Get-Date 01.01.1970).AddSeconds($BrowserLocalstate.profile.info_cache.$BrowserUserProfile.active_time).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") 
            SyncUsername    = $BrowserLocalstate.profile.info_cache.$BrowserUserProfile.user_name
            }
        )
    }

}

Out-ToLine "Starting Proflie Extract:"
out-toline ($BrowserUserProfileList | select * | ft -AutoSize| Out-String)
#
# Check for the Managed Prfolie exists - Create if missing
#
if(@($BrowserUserProfileList | Where-Object {$_.Browser -eq "Edge"} ).count -ge 1) {  #check to makesure the default profile already exisits - if not leave Edge well alone until next user logon!
    Out-ToLine "Starting Process:"
    Start-Sleep (Get-Random -Maximum 10 -Minimum 0)
    if ($null -eq ($BrowserUserProfileList | Where-Object {$_.Browser -eq "Edge" -and $_.profile -eq $ManagedProfileName}) ) {
        
        if (Get-Process msedge -ErrorAction SilentlyContinue) {
            Out-ToLine "..Stopping all edge tasks" 
            get-process msedge -ErrorAction SilentlyContinue | stop-process -Force
        }
        Out-ToLine "..Creating Profile with name: $ManagedProfileName"
        $profilepath = $ManagedProfileName -ireplace " ",""
        $result = Start-Process -FilePath "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" -ArgumentList "--profile-directory=$profilePath --no-startup-window --no-first-run --no-default-browser-check " -PassThru
        Out-ToLine "..waiting for MSEdge to create profile"
        start-sleep 15
        get-process msedge -ErrorAction SilentlyContinue | stop-process -Force
        Out-ToLine "..loading MSEdge Config"
        $BrowserLocalstate = get-Content "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\Local State" | convertfrom-json
    
        Out-ToLine "..update Profile Naming"
        $browserLocalstate.profile.info_cache.$ManagedProfileName.name = $ManagedProfileName
        $browserLocalstate.profile.info_cache.$ManagedProfileName.shortcut_name = $ManagedProfileName
        $browserLocalstate | ConvertTo-Json -Compress -Depth 100 | out-file "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\Local State" -Encoding UTF8
        #update registry to match new name
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Edge\Profiles\$ManagedProfileName" -Name "ShortcutName" -Value $ManagedProfileName -Force
        #update profile prefswith the correct nme
        $ManagedProfileNamePref = get-Content -raw "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\$ManagedProfileName\Preferences" | ConvertFrom-Json
        $ManagedProfileNamePref.profile.name = $ManagedProfileName
    
        Write-Output "..Switching: show_hub_apps_tower to false"
        if ($null -eq $ManagedProfileNamePref.browser.show_hub_apps_tower) {
            $ManagedProfileNamePref.browser | add-member -Name "show_hub_apps_tower" -value $false -MemberType NoteProperty
        } else {
            $ManagedProfileNamePref.browser.show_hub_apps_tower = $false
        }
        Write-Output "..Switching: show_hub_apps_tower_pinned to false"
        if ($null -eq $ManagedProfileNamePref.browser.show_hub_apps_tower_pinned) {
            $ManagedProfileNamePref.browser | add-member -Name "show_hub_apps_tower_pinned" -value $false -MemberType NoteProperty
        }else {
            $ManagedProfileNamePref.browser.show_hub_apps_tower_pinned = $false
        }
        Write-Output "..Switching: show_toolbar_learning_toolkit_button to false"
        if ($null -eq $ManagedProfileNamePref.browser.show_toolbar_learning_toolkit_button) {
            $ManagedProfileNamePref.browser | add-member -Name "show_toolbar_learning_toolkit_button" -value $false -MemberType NoteProperty
        } else {
            $ManagedProfileNamePref.browser.show_toolbar_learning_toolkit_button = $false
        }
        Write-Output "..Switching: Start Page features off"
        if ($null -eq $ManagedProfileNamePref.ntp) {
            $blockvalue = '{"background_image_type":"imageAndVideo","hide_default_top_sites":false,"layout_mode":3,"news_feed_display":"off","num_personal_suggestions":1,"prerender_contents_height":823,"prerender_contents_width":1185,"quick_links_options":0}'
            $ManagedProfileNamePref | add-member -Name "ntp" -value (Convertfrom-Json $blockvalue) -MemberType NoteProperty
        } else {
            $blockvalue = '{"background_image_type":"imageAndVideo","hide_default_top_sites":false,"layout_mode":3,"news_feed_display":"off","num_personal_suggestions":1,"prerender_contents_height":823,"prerender_contents_width":1185,"quick_links_options":0}'
            $ManagedProfileNamePref | add-member -Name "ntp" -value (Convertfrom-Json $blockvalue) -MemberType NoteProperty -Force
        }
        $ManagedProfileNamePref | ConvertTo-Json -Compress -Depth 100 | out-file "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\$ManagedProfileName\Preferences" -Encoding UTF8
    
        
    } else {
        Out-ToLine "..Skipping Profile already created."
    }
    
        
    Out-ToLine "Starting Profile Switch Site Check"
    $sitelist = import-csv $sitelistFilename -header "site","profile"
    
    Out-ToLine "..Prepairing Site List:"
    if ($null -eq ($BrowserUserProfileList | Where-Object {$_.profile -eq "Default"})) {
        $NewDefaultProfile = ($BrowserUserProfileList | where {$_.profile -ne $ManagedProfileName } |Sort-Object Profile -Descending | select -First 1).ProfileFolder # yes to should be ProfileFolder not profilename
        out-toline "..Flipped Default Ref in CSV to $NewDefaultProfile"
    }
    $newlist = @()
    foreach ($site in $sitelist ) {
        if ($($site.profile) -eq $ManagedProfileName -or $($site.profile) -eq "Default") {
            if ($site.profile -eq "Default") {
                $site.profile = $NewDefaultProfile
            }
            $newlist = $newlist + ([pscustomobject][ordered]@{
                action = 1
                profile = $site.profile
                type = 0
                url = $site.site
                }
            )
            Out-ToLine "..$($site.profile) will be used for: $($site.site)"
        } else {
            $message = "Bad Profile named in sitelist: $($site.profile) should be: $ManagedProfileName"
            Out-ToLine $message
        }
    }
    
    $BrowserLocalstate = get-Content "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\Local State" | convertfrom-json 
    $LocalStateChanged = $false
    if ((($newlist | sort-object url) | ConvertTo-Json) -ne (($BrowserLocalstate.profiles.edge.guided_switch_pref | Sort-Object url) | ConvertTo-Json)) {
        Out-ToLine "..Updating Site Profile List"
        get-process msedge -ErrorAction SilentlyContinue | stop-process -Force
        $BrowserLocalstate = get-Content "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\Local State" | convertfrom-json
        $BrowserLocalstate.profiles.edge | add-member -Name "guided_switch_pref" -value $null -MemberType NoteProperty -ErrorAction SilentlyContinue
        $BrowserLocalstate.profiles.edge.guided_switch_pref = $newlist | Sort-Object profile
        $LocalStateChanged = $true
    } else {
        Out-ToLine "..Skipping Site Profile List"
    }
    
 #   if ($($BrowserLocalstate.profiles.edge.ext_link_open_behavior) -ne 1 -and $($BrowserLocalstate.profiles.edge.ext_link_open_profile_path -ne "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\Default" ) ) {
 #       Out-ToLine "Updating Default Profile for new Links"
 #       $BrowserLocalstate.profiles.edge | add-member -Name "ext_link_open_profile_path" -value $null -MemberType NoteProperty -ErrorAction SilentlyContinue
 #       $BrowserLocalstate.profiles.edge | add-member -Name "ext_link_open_behavior" -value $null -MemberType NoteProperty -ErrorAction SilentlyContinue
 #       $BrowserLocalstate.profiles.edge.ext_link_open_profile_path = "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\Default"
 #       $BrowserLocalstate.profiles.edge.ext_link_open_behavior     = 1
 #       $LocalStateChanged = $true
 #   } else {
 #       Out-ToLine "..Skipping Default Profile for New Links"
 #   }
    
    
    if($LocalStateChanged -eq $true) {
        Out-ToLine "..Saving Changes Localstate to MS Edge"
        Out-ToLine "....Killing Live MSEDGE.EXE"
        get-process msedge -ErrorAction SilentlyContinue | stop-process -Force
        $browserLocalstate | ConvertTo-Json -Compress -Depth 100 | out-file "$($env:LOCALAPPDATA)\Microsoft\Edge\User Data\Local State" -Encoding UTF8
        Out-ToLine "....Save Localstate Configuration Done!"
    }
    
} else {
    Out-ToLine "Default Profile Missing...   Exiting Script"
}
