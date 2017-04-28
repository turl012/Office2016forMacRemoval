#!/bin/

## This script removes Microsoft Office 2016 for Mac, including the licence
## This script was written by Matthew Turley for use at the Technology Service Desk at Pitt

## Check if user is sudo
if [[ $EUID -ne 0 ]]; then
    echo "Please Run As: \"sudo sh office16_removal.sh\" "
    exit 1
else
    ## Remove the Applications
    echo "Removing Microsoft Office 2016 Applications ..."
    rm -rf /Applications/Microsoft\ Word.app
    rm -rf /Applications/Microsoft\ Excel.app
    rm -rf /Applications/Microsoft\ PowerPoint.app
    rm -rf /Applications/Microsoft\ Outlook.app
    rm -rf /Applications/Microsoft\ OneNote.app

    ## Remove the items from /Library
    echo "Removing System Wide Support Files ..."
    rm -rf /Library/Application\ Support/Microsoft
    rm /Library/LaunchDaemons/com.microsoft.office.*
    rm -rf /Library/PrivilegedHelperTools/com.microsoft.office.*
    rm -rf /Library/Preferences/ByHost/com.microsoft*
    rm -rf /Library/Receipts/Office2016_*
    rm -rf /Library/Fonts/Microsoft

    ## Remove the items from ~/Library
    echo "Removing User Support Files ..."
    rm -rf ~/Library/Application\ Support/Microsoft
    rm -rf ~/Library/Application\ Scripts/com.microsoft.*
    rm -rf ~/Library/Caches/Microsoft
    rm -rf ~/Library/Caches/com.microsoft.*

    ## Removing files from ~/Library/Containers (Files Microsoft Says to Delete)
    ## https://support.office.com/en-us/article/Uninstall-Office-2016-for-Mac-eefa1199-5b58-43af-8a3d-b73dc1a8cae3
    rm -rf ~/Library/Containers/com.microsoft.errorreporting
    rm -rf ~/Library/Containers/com.microsoft.Excel
    rm -rf ~/Library/Containers/com.microsoft.netlib.shipassertprocess
    rm -rf ~/Library/Containers/com.microsoft.Office365ServiceV2
    rm -rf ~/Library/Containers/com.microsoft.Outlook
    rm -rf ~/Library/Containers/com.microsoft.Powerpoint
    rm -rf ~/Library/Containers/com.microsoft.RMS-XPCService
    rm -rf ~/Library/Containers/com.microsoft.Word
    rm -rf ~/Library/Containers/com.microsoft.onenote.mac

    ## Removing Files from ~/Library/Group Containers
    rm -rf ~/Library/Group\ Containers/UBF8T346G9.*

    ## Forget Office 2016 Compleatly
    echo "Forget Office 2016 Install ..."
    pkgutil --forget com.microsoft.package.Fonts
    pkgutil --forget com.microsoft.package.Microsoft_AutoUpdate.app
    pkgutil --forget com.microsoft.package.Microsoft_Excel.app
    pkgutil --forget com.microsoft.package.Microsoft_OneNote.app
    pkgutil --forget com.microsoft.package.Microsoft_Outlook.app
    pkgutil --forget com.microsoft.package.Microsoft_PowerPoint.app
    pkgutil --forget com.microsoft.package.Microsoft_Word.app
    pkgutil --forget com.microsoft.package.Proofing_Tools
    pkgutil --forget com.microsoft.package.licensing

    echo "Please Restart Your Computer
          Some Microsoft Applications may need reinstalled"

fi
