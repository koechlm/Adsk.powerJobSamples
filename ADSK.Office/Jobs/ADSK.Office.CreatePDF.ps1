#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a PDF file and add it to Autodesk Vault							  #
#                                                                             #
# Copyright (c) coolOrange s.r.l./Autodesk GmbH Germany - All rights reserved.#
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

Add-Log "Starting job 'Create PDF as attachment' for file '$($file._Name)' ..."
$workingDirectory = "C:\Temp\powerJob"
$localPDFFullFileName = "$workingDirectory\$($file._Name).pdf"
$vaultPDFfileLocation = $file._EntityPath +"/"+ (Split-Path -Leaf $localPDFFullFileName)
if( @("docx","xlsx", "pptx") -notcontains $file._Extension ) {
    Add-Log "Files with extension: '$($file._Extension)' are not supported. Job execution finished"
    return
}

$mLocalFile = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory
Switch($mLocalFile._Extension)
{
	"docx"
	{
		$word = New-Object -ComObject Word.Application
		Sleep -Seconds 5
		$process = Get-Process winword -ErrorAction SilentlyContinue
		$word.Visible = $false
		#Opens the wordfile to save as pdf-file
		$openResult = $word.documents.open($mLocalFile.LocalPath)
		#Creates the pdf-file
		IF ($openResult)
		{
			Try 
			{
				$openResult.ExportAsFixedFormat($localPDFFullFileName, [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF)
				$mExpFile = $true
				$openResult.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
				$mCloseResult = $true
			}
			Catch
			{
				throw("Word opened the file, but could not save as PDF")
			}
			
		}
		$word.Quit()
	}

	"xlsx"
	{
		$excel = New-Object -ComObject Excel.Application
		Sleep -Seconds 5
		$process = Get-Process excel -ErrorAction SilentlyContinue
		$excel.Visible = $false
		$openResult = $excel.Workbooks.Open($mLocalFile.LocalPath)
		IF ($openResult)
		{
			Try 
			{
				#Creates the pdf-file
				$xlFixedFormat = "Microsoft.Office.Interop.Excel.xlFixedFormatType" -as [type] 
				$openResult.ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $localPDFFullFileName)
				$mExpFile = $true
				$openResult.Close($false) #close and skipe save as question
				$mCloseResult = $true
			}
			catch 
			{
				throw("Excel opened the file, but could not save as PDF")
			}
			$excel.Quit()
		}
	}

	"pptx"
	{
		$powerpoint = New-Object -ComObject Powerpoint.Application
		Sleep -Seconds 5
		$process = Get-Process POWERPNT #-ErrorAction SilentlyContinue
		#$powerpoint.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
		$openResult = $powerpoint.presentations.Open($mLocalFile.LocalPath)
		IF ($openResult)
		{
			Try 
			{
				#Creates the pdf-file
				$xlFixedFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
				$openResult.SaveAs($localPDFFullFileName, $xlFixedFormat::ppSaveAsPDF)
				$mExpFile = $true
				$openResult.Close()
				$mCloseResult = $true
			}
			catch 
			{
				throw("powerPoint opened the file, but could not save as PDF")
			}
			$powerpoint.Quit()
			$powerpoint = $null
			Stop-Process $process
		}
 }
    } # end switch
	
	if($mExpFile) 
	{
        $PDFfile = Add-VaultFile -From $localPDFFullFileName -To $vaultPDFfileLocation 
        $mUpdatedfile = Update-VaultFile -File $mLocalFile._FullPath -AddAttachments @($PDFfile._FullPath) -Comment "PDF file attached"   
	}

Clean-Up -folder $workingDirectory

if(-not $openResult) {
    throw("Failed to open document $($mLocalFile.LocalPath)! Reason: $($openResult.ErrorMessage)")
}
if(-not $mExpFile) {
    throw("Failed to export document $($mLocalFile.LocalPath) to $localPDFFullFileName! Reason: $($mExpError)")
}
if(-not $mCloseResult){
	throw("Failed to close document $($mLocalFile.LocalPath) ")
}
if(-not $mUpdatedfile){
	throw("Failed to attach PDF document to $($vaultPDFfileLocation)")
}

Add-Log "Completed job 'Create PDF as attachment'"