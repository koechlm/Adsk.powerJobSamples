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
#$file = Get-VaultFile -File '$/Konstruktion/01-0387.ipt'
#Import-Module C:\ProgramData\coolOrange\powerJobs\Modules\ADSK.CreateTCLinks.psm1

Add-Log "Starting job 'Create-Add TC Link' for file '$($file._Name)' ..."

    $TCLink = Adsk.CreateTcFileLink $file._FullPath
    if (-not $TCLink)
    {
        throw("Function 'CreateTcFileLink' did not return a link for:  $($file._Name)!")
    }
    else 
    {
        Update-VaultFile -File $file._FullPath -Properties @{'ThinClient Link' = $TCLink}
    }

Add-Log "...Finishing job 'Create-Add TC Link' for file '$($file._Name)'."