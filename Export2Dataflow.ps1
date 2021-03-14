<# 
.SYNOPSIS
	This script exports the Power Query to a Dataflow JSON file
.DESCRIPTION
	This Powershell script can be called as an External Tool from Power BI Desktop, which then extracts the Power Queries via the 
	TOM (Tabular Object Model) and stores them in the Dataflow JSON format.
.NOTES
    File Name	: Export2Dataflow.ps1
	Date		: 03/14/2021
	Version		: 1.0.0 
    Author		: Marcus Wegener

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
	EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
	MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
	IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
	OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
	ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
	OTHER DEALINGS IN THE SOFTWARE.
.LINK 
	https://github.com/MarcusWegener/Export2Dataflow
.Parameter Server
	The server name and portnumber of the local instance of Analysis Services Tabular for imported/DirectQuery data models
.Parameter Database
	The database name of the model hosted in the local instance of Analysis Services Tabular for imported/DirectQuery data models
#>


# Below section defines the server and database based on the input captured from the External tools integration 
# This is defined as arguments \"%server%\" and \"%database%\" in the external tools json 
Param(
	[string]$server,
	[string]$database
)

# Write Intro, Server and Database information to screen
Write-Host -ForegroundColor White  '========================================================================================================================'
Write-Host -ForegroundColor Yellow '                ______                           __  ___    ____          __          ____ __                           '
Write-Host -ForegroundColor Yellow '               / ____/_  __ ____   ____   _____ / /_|__ \  / __ \ ____ _ / /_ ____ _ / __// /____  _      __            '
Write-Host -ForegroundColor Yellow '              / __/  | |/_// __ \ / __ \ / ___// __/__/ / / / / // __ `// __// __ `// /_ / // __ \| | /| / /            '
Write-Host -ForegroundColor Yellow '             / /___ _>  < / /_/ // /_/ // /   / /_ / __/ / /_/ // /_/ // /_ / /_/ // __// // /_/ /| |/ |/ /             '
Write-Host -ForegroundColor Yellow '            /_____//_/|_|/ .___/ \____//_/    \__//____//_____/ \__,_/ \__/ \__,_//_/  /_/ \____/ |__/|__/              '
Write-Host -ForegroundColor Yellow '                        /_/                                                                                             '
Write-Host -ForegroundColor Yellow '                                                                          By Marcus Wegener  v1.0.0.0   03/14/2021      '
Write-Host -ForegroundColor White  '========================================================================================================================'
Write-Host -ForegroundColor White  "Your Power BI Model currently runs with the following connection details:"
Write-Host -ForegroundColor White  "Server: " $server 
Write-Host -ForegroundColor White  "Database: " $database 

# Install latest package (if not already installed) of Microsoft.AnalysisServices.retail.amd64 for current user
Write-Host -ForegroundColor White  '========================================================================================================================'
Write-Host -ForegroundColor Gray  "Install latest package (if not already installed) of Microsoft.AnalysisServices.retail.amd64 for current user..."
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Package Microsoft.AnalysisServices.retail.amd64 -Scope CurrentUser -Source "https://www.nuget.org/api/v2"
$installedPackages = Find-Package -Name Microsoft.AnalysisServices.retail.amd64 -Source "$env:USERPROFILE\AppData\Local\PackageManagement\NuGet\Packages\" -AllVersions
$maxVersion = $installedPackages[0].Version

# Load Assembly Files from insalled package Microsoft.AnalysisServices.retail.amd64
Write-Host -ForegroundColor Gray "Load Assembly Files from insalled package Microsoft.AnalysisServices.retail.amd64..."
$assemblyPathTabular = "$env:USERPROFILE\AppData\Local\PackageManagement\NuGet\Packages\Microsoft.AnalysisServices.retail.amd64.$maxVersion\lib\net45\Microsoft.AnalysisServices.Tabular.dll"

Write-Host -ForegroundColor White  '========================================================================================================================'

# Connect to Power BI Model
Write-Host -ForegroundColor Gray "Connect to Power BI Model..."

Add-Type -Path $assemblyPathTabular

$as = New-Object Microsoft.AnalysisServices.Tabular.Server
$as.Connect($server)
$db = $as.Databases[$database]
$dbModel = $db.Model

$modelCulture = $dbModel.Culture
$modelModifiedTime = Get-Date($dbModel.ModifiedTime) -format s

# Create query groups
Write-Host -ForegroundColor Gray "`Create query groups..."

$queryGroups = @{}
foreach($queryGroup in $dbModel.QueryGroups) {
		$groupStructur = $queryGroup.Folder.Split("\")
		$groupName = $groupStructur[-1]
		$groupParentName = ""

		if($groupStructur.Length -gt 1) {
			$groupParentName = $groupStructur[0..($groupStructur.Length -2)] -join "\"
		}
		
        $queryGroups[$queryGroup.Name] = [ordered]@{
			"id" = New-Guid
			"name" = $groupName
			"description" = $dbModel.QueryGroups[0].Annotations.description
			"parentId" = $queryGroups[$groupParentName].id
			"order" = $groupStructur.Length
		}
}

$pbiQueryGroups = ""

if($queryGroups.Count -le 1) {
    $pbiQueryGroups = "["
}

$pbiQueryGroups += $queryGroups.values | ConvertTo-Json -Compress 

if($queryGroups.Count -le 1) {
    $pbiQueryGroups += "]"
}


# Create queries
Write-Host -ForegroundColor Gray "`nCreate queries..."

$queriesMetadata = @{}
$entities = @()
$document = "section Section1;`n"


foreach($e in $dbModel.Expressions) {
    $queriesMetadata[$e.Name] =  [ordered]@{ 
		queryId = New-Guid
		queryName = $e.Name
    }
    if($e.QueryGroup) {
        $queriesMetadata[$e.Name].queryGroupId = $queryGroups[$e.QueryGroup.Folder].id
    }
    
    $tableName = $e.Name
    if($tableName -notmatch '^[a-z0-9]+$') {
        $tableName = "#"""+$tableName+""""
    }

    $document += "shared " + $tableName + " = " + $e.Expression + ";`n"
}

foreach($t in $dbModel.Tables) {
    if($t.Partitions.SourceType -eq "M") {
        $queriesMetadata[$t.Name] = [ordered]@{ 
			queryId = New-Guid
			queryName = $t.Name
			"loadEnabled" = $true
		}
        if($t.Partitions.QueryGroup) {
			$queriesMetadata[$t.Name].queryGroupId = $queryGroups[$t.Partitions.QueryGroup.Folder].id
        }

        $tableName = $t.Name
        if($tableName -notmatch '^[a-z0-9]+$') {
          $tableName = "#"""+$tableName+""""
        }

        $document += "shared " + $tableName + " = " + $t.Partitions.Source.Expression + ";`n"

        $attributes = @()

        foreach($c in $t.Columns) {
			if ($c.Type -eq "Data") {
               $attributes += [ordered]@{ 
					name = $c.Name 
					dataType = $c.DataType.ToString().ToLower()
				}
			}
        }

        $entities += [ordered]@{ 
			'$type' = "LocalEntity"
			"name" = $t.Name
			"description" = $t.Description
			"pbi:refreshPolicy"= [ordered]@{
				'$type' = "FullRefreshPolicy"
				"location" = [uri]::EscapeDataString($t.Name) + ".csv"
			}
			"attributes" = $attributes
        }
    }
}

$document = $document.Replace("`n","`r`n")

Write-Host -ForegroundColor Gray "`nDisconnect from Power BI Model..."

$as.Disconnect();

# Save file dialog
Write-Host -ForegroundColor Gray "Export Dataflow model to file..."

Add-Type -AssemblyName System.Windows.Forms

$fileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter = 'JSON Files (*.json)|*.json'
    Title = 'Export Dataflow model to file:'
}

$null = $fileBrowser.ShowDialog()
$fileName = $fileBrowser.FileName

if(!$fileName){
    Write-Host -ForegroundColor Yellow "No file chosen. Terminating..."
    Start-Sleep -Seconds 1.5
    exit
}

$fileNameWithoutPath = Split-path $fileBrowser.FileName -leaf


Write-Host -ForegroundColor Gray "$fileName chosen..."

Write-Host -ForegroundColor White  '========================================================================================================================'

$dataflowName = [regex]::Match($fileNameWithoutPath, ".+?(?=\.json)").Value

$json = [ordered]@{ 
	"name" = $dataflowName
	"description" = ""
	"version" = "1.0"
	"culture" = $modelCulture
	"modifiedTime" = $modelModifiedTime
	"pbi:mashup" = [ordered]@{
		"fastCombine" = $false
		"allowNativeQueries" = $false
		"skipAutomaticHeaderAndTypeDetection" = $false
		"queriesMetadata" = $queriesMetadata
		"document" = $document
	}
	"annotations" = @(
		[ordered]@{
			"name" = "pbi:QueryGroups"
			"value" = $pbiQueryGroups
		}
	)
	"entities" = $entities
 }

Write-Host -ForegroundColor Gray "`nWrite into JSON file..."

$json | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding UTF8 "$fileName"

Start-Process -FilePath C:\Windows\explorer.exe -ArgumentList "/select, ""$fileName"""

Write-Host -ForegroundColor Gray "Dataflow model file ($fileName) created..."
Write-Host -ForegroundColor White  '========================================================================================================================'
Start-Sleep -Seconds 1.5