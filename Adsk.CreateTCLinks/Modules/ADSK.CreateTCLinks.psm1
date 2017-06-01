#=============================================================================#
# PowerShell script sample for Vault Data Standard / coolOrange powerJobs                           #
# Create Thin Client Web Link for file from Autodesk Vault                    #
#                                                                             #
# Copyright (c) Autodesk - All rights reserved.                               #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

function Adsk.CreateTcFileLink([string]$FileFullVaultPath )
{
    $FullPaths = @($FileFullVaultPath)

    $Files = $vault.DocumentService.FindLatestFilesByPaths($FullPaths)

    $IDs = @($Files[0].Id)
     
    $PersIDs = $vault.KnowledgeVaultService.GetPersistentIds("FILE", $IDs, [Autodesk.Connectivity.WebServices.EntPersistOpt]::Latest)
    $PersID = $PersIDs[0].TrimEnd("=")

    $serverUri = [System.Uri]$Vault.InformationService.Url

    $vaultName = $VaultConnection.Vault
    $Server = $VaultConnection.Server

    $TCLink = $serverUri.Scheme + "://" + $Server + "/AutodeskTC/" + $Server + "/" + $vaultName + "/#/Entity/Details?id=m" + "$PersID" + "&itemtype=File"
    
    return $TCLink
}

function Adsk.CreateTcFolderLink([string]$FolderFullVaultPath)
{
    $Folder = $vault.DocumentService.GetFolderByPath($FolderFullVaultPath)

    $IDs = @($Folder.Id)
     
    $PersIDs = $vault.KnowledgeVaultService.GetPersistentIds("FLDR", $IDs, [Autodesk.Connectivity.WebServices.EntPersistOpt]::Latest)
    $PersID = $PersIDs[0].TrimEnd("=")

    $serverUri = [System.Uri]$Vault.InformationService.Url

    $vaultName = $VaultConnection.Vault
    $Server = $VaultConnection.Server

    $TCLink = $serverUri.Scheme + "://" + $Server + "/AutodeskTC/" + $Server + "/" + $vaultName + "/#/Entity/Entities?folder=m" + "$PersID" + "&start=0"
    return $TCLink
}

function mGetFolderPropertyDefId ([STRING] $mDispName) {
	$PropDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")
	$propDefIds = @()
	$PropDefs | ForEach-Object {
		$propDefIds += $_.Id
	} 
	$mPropDef = $propDefs | Where-Object { $_.DispName -eq $mDispName}
	Return $mPropDef.Id
}

function mUpdateFldrProperties([Long] $FldId, [String] $mDispName, [Object] $mVal)
{
	$ent_idsArray = @()
	$ent_idsArray += $FldId
	$propInstParam = New-Object Autodesk.Connectivity.WebServices.PropInstParam
	$propInstParamArray = New-Object Autodesk.Connectivity.WebServices.PropInstParamArray
	$mPropDefId = mGetFolderPropertyDefId $mDispName
 	$propInstParam.PropDefId = $mPropDefId
	$propInstParam.Val = $mVal
	$propInstParamArray.Items += $propInstParam
	$propInstParamArrayArray += $propInstParamArray
	Try{
        $vault.DocumentServiceExtensions.UpdateFolderProperties($ent_idsArray, $propInstParamArrayArray)
	    return $true
    }
    catch { return $false}
}

function Adsk.CreateTCItemLink ([Long]$ItemId)
{
	$IDs = @($ItemId)
    $PersIDs = $vault.KnowledgeVaultService.GetPersistentIds("ITEM", $IDs, [Autodesk.Connectivity.WebServices.EntPersistOpt]::Latest)
    $PersID = $PersIDs[0].TrimEnd("=")

    $serverUri = [System.Uri]$Vault.InformationService.Url

    $vaultName = $VaultConnection.Vault
    $Server = $VaultConnection.Server

    $TCLink = $serverUri.Scheme + "://" + $Server + "/AutodeskTC/" + $Server + "/" + $vaultName + "/#/Entity/Details?id=m" + "$PersID" + "&itemtype=Item"
    return $TCLink
}

