#===============================================================================#
# PowerShell script sample for coolOrange powerJobs								#
# Create a Thinclient link and save it as property  							#
#																				#
# Copyright (c) autodesk/coolOrange - All rights reserved.						#
#																				#
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER     #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES   #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.    #
#===============================================================================#

#Open-VaultConnection -Server '192.168.85.128' -User 'Administrator' -Vault 'VLT-MKDE'
#$file = Get-VaultFile -File '$/Konstruktion/01-0387.ipt'
#Import-Module C:\ProgramData\coolOrange\powerJobs\Modules\ADSK.CreateTCLinks.psm1

Add-Log "Starting job 'Create-Add TC Link' for Item '$($Item._Name)' ..."

    $TCLink = Adsk.CreateTcItemLink $Item.Id
    if (-not $TCLink)
    {
        throw("Function 'CreateTcFileLink' did not return a link for:  $($Item._Name)!")
    }
    else 
    {
	    $mVaultItems = $vault.ItemService.GetItemsByIds(@($Item.Id))
	    $mVaultItem = $mVaultItems[0]

	    #check that the item is editable for the current user, if not, we shouldn't add the files, before we try to attach
	    try{
		    $vault.ItemService.EditItems(@($mVaultItem.RevId))
		    Add-Log "Item is accessible for JobProcessor user"
		    $_ItemIsEditable = $true
			$vault.ItemService.UndoEditItems(@($mVaultItem.RevId))
			$vault.ItemService.DeleteUncommittedItems($true)
			#Add-Log "Item Lock Removed to continue"
	    }
	    catch {
		    $_ItemIsEditable = $false
            throw ("Item is NOT accessible for JobProcessor user.")
	    }
	    If($_ItemIsEditable)
	    {
		    $mItemProps = @{}
		    $mItemProps.Add("ThinClient Link", $TCLink)
            Try
            {
                Update-VaultItem -Number $Item._Number -Properties $mItemProps 
            }
            catch { 
                throw("Failed to update Item property 'ThinClient Link'. Check that this property is available for items. ")
            }
	    }			
    }

Add-Log "...Finishing job 'Create-Add TC Link' for Item '$($file._Name)'."