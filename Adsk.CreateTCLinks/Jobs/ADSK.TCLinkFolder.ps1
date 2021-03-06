#=========================================================#
# PowerShell script sample for coolOrange powerJobs                                                         #
# Create a Thinclient link and save it as property  						                                      #
#                                                                                                                                    #
# Copyright (c) autodesk/coolOrange - All rights reserved.                                                   #
#                                                                                                                                    #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER      #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES   #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.    #
#=========================================================#

#Open-VaultConnection -Server '192.168.85.128' -User 'Administrator' -Vault 'VLT-MKDE'
#$folder = '$/Konstruktion/SR-0009'
#Import-Module C:\ProgramData\coolOrange\powerJobs\Modules\ADSK.CreateTCLinks.psm1

Add-Log "Starting job 'Create-Add TC Link' for folder '$($folder.FullName)' ..."

    $TCLink = Adsk.CreateTcFolderLink $folder.FullName
    if (-not $TCLink)
    {
        throw("Function 'CreateTcFolderLink' did not return a link for:  $($folder.FullName)!")
    }
    else 
    {
        $mFldPropUpdated = mUpdateFldrProperties $folder.Id 'ThinClient Link' $TCLink
        if(-not $mFldPropUpdated)
        {
            throw("Failed to update folder property 'ThinClient Link'. Check that this property is available for folders. ")
        }
    }

Add-Log "...Finishing job 'Create-Add TC Link' for folder '$($folder.FullName)'."