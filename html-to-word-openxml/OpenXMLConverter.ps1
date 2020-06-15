<#
  .SYNOPSIS
  Convert Html file to Word document.

  .DESCRIPTION
  The OpenXMLConverter.ps1 script copies an empty existing word and include on it the html.

  .INPUTS
  None. You cannot pipe objects to OpenXMLConverter.ps1.

  .OUTPUTS
  None. OpenXMLConverter.ps1 does not generate any output.

  .EXAMPLE
  C:\PS> .\OpenXMLConverter.ps1
#>


Add-Type -Path $PSScriptRoot\DocumentFormat.OpenXml.2.11.0\lib\net35\DocumentFormat.OpenXml.dll
function Open-Using-Object
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowEmptyCollection()]
        [AllowNull()]
        [Object]
        $InputObject,

        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock
    )

    try
    {
        . $ScriptBlock
    }
    finally
    {
        if ($null -ne $InputObject -and $InputObject -is [System.IDisposable])
        {
            $InputObject.Dispose()
        }
    }
}

function Convert-Word{
    (
        [string]$htmlPath,
        [string]$wordPath
    )    
    
    $templateWordPath = "$PSScriptRoot\Template.docx"

    Copy-Item $templateWordPath -Destination $wordPath -Force

    Open-Using-Object ($myDoc = [DocumentFormat.OpenXml.Packaging.WordprocessingDocument]::Open($wordPath,$true)) {
	
        $altChunkId = "AltChunkId1";
    
        $mainPart = $myDoc.MainDocumentPart;
    
        $chunk = $mainPart.AddAlternativeFormatImportPart([DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPartType]::Xhtml, $altChunkId);
        
        Open-Using-Object ($fileStream = [System.IO.File]::Open($htmlPath,[System.IO.FileMode]::Open)){
            $chunk.FeedData($fileStream);
        }
       
        $altChunk = New-Object DocumentFormat.OpenXml.Wordprocessing.AltChunk
    
        $altChunk.Id = $altChunkId;
    
        $lastElement = $mainPart.Document.Body.Elements()
    
        $last = [Linq.Enumerable]::Last($lastElement)
    
        $mainPart.Document.Body.InsertBefore($altChunk,$last)
    
        $mainPart.Document.Save()
    }
}

$htmlPath = "$PSScriptRoot\Test.html"
$wordPath = "$PSScriptRoot\Test.docx"

Convert-Word $htmlPath $wordPath
