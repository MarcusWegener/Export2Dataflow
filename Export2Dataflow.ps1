<# 
.SYNOPSIS
	This script exports the Power Query to a Dataflow JSON file or publishes it to the Power BI service.
.DESCRIPTION
	This Powershell script can be called as an External Tool from Power BI Desktop, which then extracts the Power Queries via the 
	TOM (Tabular Object Model) and stores them in the Dataflow JSON format or publishes it to the Power BI service.
.NOTES
    File Name	: Export2Dataflow.ps1
	Date		: 03/28/2021
	Version		: 1.1.0 
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
.Parameter Deployment
	Deployment options : 
	0 - saves to Dataflow JSON file
	1 - published in Power BI service
#>

# Below section defines the server, database and deployment based on the input captured from the External tools integration 
# This is defined as arguments \"%server%\" and \"%database%\" in the external tools json 
Param(
	[string] $server   = $null,
	[string] $database = $null,
	[int] $deployment  = 0
)

Set-StrictMode -Version Latest

# =================================================================================================================================================
# DEFINE FUNCTIONS
# =================================================================================================================================================

# Function of Marc Lelijveld
# https://github.com/marclelijveld/Power-BI-Automation/blob/master/PowerBI_MoveDataflows.ps1
function _postDataflowDefinition([string] $GroupID, [string]$DataflowDefinition, [string]$NameConflict) {

    $UserAccessToken = Get-PowerBIAccessToken
    $bearer = $UserAccessToken.Authorization.ToString()
    
    $url = [string]::Format("https://api.powerbi.com/v1.0/myorg/groups/{0}/imports?datasetDisplayName=model.json&nameConflict={1}", $GroupID, $NameConflict)

    $boundary = [System.Guid]::NewGuid().ToString("N")
    $LF = [System.Environment]::NewLine
		
    $body = (
        "--$boundary",
        "Content-Disposition: form-data; name=`"`"; filename=`"model.json`"",
        "Content-Type: application/json$LF",
        $DataflowDefinition,
        "--$boundary--$LF"
    ) -join $LF

    $headers = @{
        'Authorization' = "$bearer"
        'Content-Type'  = "multipart/form-data; boundary=--$boundary"
    }
   
    $postFlow = Invoke-RestMethod -Uri $url -ContentType 'multipart/form-data' -Method POST -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($body))

    return $postFlow
}

# Form
function _showSelectionForm($mode, $workspace, $selectionList){

    #create GUI
    $watermarkText = ""
    $labelText = ""
    $textBoxWidth = 560
    if ($mode -eq "workspace") {
        $watermarkText = "Search"
        $labelText = "Select a destination:"
    } elseif ($mode -eq "dataflow") {
        $watermarkText = "name for a new dataflow"
        $labelText = "Workspace: " + $workspace + "`n" + "Create a new dataflow or select an existing dataflow:"
        $textBoxWidth = 480
    }
    #create the window settings
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Publish to Power BI'
    $form.Size = New-Object System.Drawing.Size(620,400)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedSingle'
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.BackColor = "White"
    $form.ShowIcon = $false

    #create the label settings
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20,10)
    $label.Size = New-Object System.Drawing.Size(560,40)
    $label.Font = [System.Drawing.Font]::new($label.Font.FontFamily.Name, 10)
    $label.Text = $labelText
    $form.Controls.Add($label)

    #create Input settings
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Size(20,50)
    $textBox.Size = New-Object System.Drawing.Size($textBoxWidth,20)
    $textBox.Font = [System.Drawing.Font]::new($textBox.Font.FontFamily.Name, 10)
    $textBox.MaxLength = 250
    $textBox.ForeColor = 'LightGray'
    $textBox.Text = $watermarkText
    $form.Controls.Add($textBox)

    if ($mode -eq "dataflow") {
        #create the create button settings
        $createButton = New-Object System.Windows.Forms.Button
        $createButton.Location = New-Object System.Drawing.Point(500,49)
        $createButton.Size = New-Object System.Drawing.Size(80,25)
        $createButton.Font = [System.Drawing.Font]::new($createButton.Font.FontFamily.Name, 10)
        $createButton.Text = 'Create'
        $createButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
        $createButton.BackColor = "#f2c811"
        $form.AcceptButton = $createButton
        $form.Controls.Add($createButton)
    }

    #create list Box settings for the selection List
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(20,77)
    $listBox.Font = [System.Drawing.Font]::new($listBox.Font.FontFamily.Name, 10)
    $listBox.Height = 243
    $listBox.Width = 560
    $listBox.Sorted = $true
    $listBox.HorizontalScrollbar = $true

    #add selectionList to the listBox
    $listBox.Items.Clear()
    [void] $listBox.Items.AddRange($selectionList)

    $form.Controls.Add($listBox)

    #create the select button settings
    $selectButton = New-Object System.Windows.Forms.Button
    $selectButton.Location = New-Object System.Drawing.Point(400,318)
    $selectButton.Size = New-Object System.Drawing.Size(80,27)
    $selectButton.Font = [System.Drawing.Font]::new($selectButton.Font.FontFamily.Name, 10)
    $selectButton.Text = 'Select'
    $selectButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $selectButton.Enabled = $false
    $selectButton.BackColor = "lightgray"
    $form.AcceptButton = $selectButton
    $form.Controls.Add($selectButton)

    #create the cancel button settings
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(500,318)
    $cancelButton.Size = New-Object System.Drawing.Size(80,27)
    $cancelButton.Font = [System.Drawing.Font]::new($cancelButton.Font.FontFamily.Name, 10)
    $cancelButton.Text = 'Cancel'
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButton
    $form.Controls.Add($cancelButton)

    $form.Topmost = $true

    #behavior if you enter the input textbox
    $textBox.Add_Enter({
        if($textBox.Text -eq $watermarkText -and $textBox.ForeColor -eq 'LightGray')
        {
            #Clear the text
            $textBox.Text = ""
            $textBox.ForeColor = 'WindowText' 
            $listBox.ClearSelected()
        }
        })
    #behavior if you leave the input textbox
    $textBox.Add_Leave({
        if($textBox.Text -eq "")
        {
            #Display the watermark
            $textBox.Text = $watermarkText
            $textBox.ForeColor = 'LightGray'
            if ($mode -eq "workspace") {
                $listBox.Items.Clear()
                [void] $listBox.Items.AddRange($selectionList)
            }
        }
    })

    if ($mode -eq "workspace") {
        #behavior if you type inside the input textbox
        $textBox.Add_TextChanged({
            $textBoxText = [regex]::Escape($textBox.Text)
            if ($textBoxText) {
                $listBox.Items.Clear()
                forEach ($selection in $selectionList) {
                    if($selection -match $textBoxText){
                        [void] $listBox.Items.Add($selection)
                    }
                }
            }
        })
    }

    $listBox.Add_SelectedIndexChanged({
        if($listBox.SelectedIndex -ge 0){
            $selectButton.Enabled = $true
            $selectButton.BackColor = "#f2c811"
        } else {
            $selectButton.Enabled = $false
            $selectButton.BackColor = "lightgray"
        }
    })

    $result = $form.ShowDialog()

    if($result -eq [System.Windows.Forms.DialogResult]::OK){
        return $listBox.SelectedItem
    } elseif ($result -eq [System.Windows.Forms.DialogResult]::Yes){
        if ($textBox.TextLength -lt 1 -or $textBox.Text -eq $watermarkText) {
            return $null
        } else {
            return $textBox.Text
        }
    } else {
        Write-Host -ForegroundColor Yellow "No " + $mode + " was selected.. Terminating.."
        Start-Sleep -Seconds 1.5
        exit
    }
}

# =================================================================================================================================================
# RUN APPLICATION
# =================================================================================================================================================

# Write Intro, Server and Database information to screen
Write-Host -ForegroundColor White  '========================================================================================================================'
if ($deployment  -eq 0) {
	Write-Host -ForegroundColor Yellow '                ______                           __  ___    ____          __          ____ __                           '
	Write-Host -ForegroundColor Yellow '               / ____/_  __ ____   ____   _____ / /_|__ \  / __ \ ____ _ / /_ ____ _ / __// /____  _      __            '
	Write-Host -ForegroundColor Yellow '              / __/  | |/_// __ \ / __ \ / ___// __/__/ / / / / // __ `// __// __ `// /_ / // __ \| | /| / /            '
	Write-Host -ForegroundColor Yellow '             / /___ _>  < / /_/ // /_/ // /   / /_ / __/ / /_/ // /_/ // /_ / /_/ // __// // /_/ /| |/ |/ /             '
	Write-Host -ForegroundColor Yellow '            /_____//_/|_|/ .___/ \____//_/    \__//____//_____/ \__,_/ \__/ \__,_//_/  /_/ \____/ |__/|__/              '
	Write-Host -ForegroundColor Yellow '                        /_/                                                                                             '
} else {
	Write-Host -ForegroundColor Yellow '               ____          __     __ _        __    ___    ____          __          ____ __                          '
	Write-Host -ForegroundColor Yellow '              / __ \ __  __ / /_   / /(_)_____ / /_  |__ \  / __ \ ____ _ / /_ ____ _ / __// /____  _      __           '
	Write-Host -ForegroundColor Yellow '             / /_/ // / / // __ \ / // // ___// __ \ __/ / / / / // __ `// __// __ `// /_ / // __ \| | /| / /           '
	Write-Host -ForegroundColor Yellow '            / ____// /_/ // /_/ // // /(__  )/ / / // __/ / /_/ // /_/ // /_ / /_/ // __// // /_/ /| |/ |/ /            '
	Write-Host -ForegroundColor Yellow '           /_/     \__,_//_.___//_//_//____//_/ /_//____//_____/ \__,_/ \__/ \__,_//_/  /_/ \____/ |__/|__/             '
	Write-Host -ForegroundColor Yellow '                                                                                                                        '
}
Write-Host -ForegroundColor Yellow '                                                                          By Marcus Wegener  v1.1.0     03/28/2021      '
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
        $groupParentId = $null

		if($groupStructur.Length -gt 1) {
			$groupParentName = $groupStructur[0..($groupStructur.Length -2)] -join "\"
            $groupParentId = $queryGroups[$groupParentName].id
		}

        $queryGroups[$queryGroup.Name] = [ordered]@{
			"id" = New-Guid
			"name" = $groupName
			"description" = $queryGroup.Description
			"parentId" = $groupParentId
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

$as.Disconnect()

$dataflowName = ""
$conflict = "Ignore"

if ($deployment  -eq 0) {
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

    $dataflowName = [regex]::Match($fileNameWithoutPath, ".+?(?=\.json)").Value
} else {

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName PresentationFramework

    [System.Windows.Forms.Application]::EnableVisualStyles()

    #Get Module "MicrosoftPowerBIMgmt"
    $moduleName = Get-Module -ListAvailable -Verbose:$false | Where-Object { $_.Name -eq "MicrosoftPowerBIMgmt" } | Select-Object -ExpandProperty Name
    if ([string]::IsNullOrEmpty($moduleName)) {
        Write-Host -ForegroundColor White "`n========================================================================================================================"
        Write-Host -ForegroundColor Gray  "Install module MicrosoftPowerBIMgmt..."
        Install-Module MicrosoftPowerBIMgmt -SkipPublisherCheck -AllowClobber -Force -Scope CurrentUser
        Write-Host -ForegroundColor White "========================================================================================================================"
    }

    Write-Host -ForegroundColor Gray "`nConnect to PowerBI service"
    Connect-PowerBIServiceAccount

    #Get all Workspaces of connected user
    $workspaces = Get-PowerBIWorkspace -All

    $workspaceName = _showSelectionForm -mode "workspace" -workspace $null -selectionList $workspaces.Name

    $selectedWorkspace = $workspaces | Where Name -eq $workspaceName

    $dataflowsInWorkspace = Get-PowerBIDataflow -Workspace $selectedWorkspace

    do {
        $dataflowName = _showSelectionForm -mode "dataflow" -workspace $workspaceName -selectionList $dataflowsInWorkspace.Name

        if ($dataflowName -eq $null) {
            [void] [System.Windows.MessageBox]::Show("You need to write at least 1 character"
                                                    ,""
                                                    ,[System.Windows.MessageBoxButton]::OK
                                                    ,[System.Windows.MessageBoxImage]::Exclamation)
        } elseif ($dataflowsInWorkspace | Where Name -eq $dataflowName){
            $confirmationResult = [System.Windows.MessageBox]::Show("Are you sure you want to overwrite the dataflow '$dataflowName'?"
                                                                   ,"Confirm Overwriting dataflow"
                                                                   ,[System.Windows.MessageBoxButton]::YesNoCancel
                                                                   ,[System.Windows.MessageBoxImage]::Exclamation)
            
            switch ($confirmationResult) {
                "Yes"{
                    Write-Host -ForegroundColor Yellow "Overwriting Dataflow '$dataflowName'"
                    $conflict = "Overwrite"
                }"No"{
                    Write-Host "Not overwriting Dataflow '$dataflowName'.. Navigating back to Dataflow-Selection"
                    $dataflowName = $null
                }"Cancel"{
                    Write-Host -ForegroundColor Yellow "Cancel-Button clicked.. Terminating.."
                    Start-Sleep -Seconds 1
                    exit
                }
            }
        }
    } while ($dataflowName -eq $null)

    Write-Host -ForegroundColor Gray $dataflowName

}

Write-Host -ForegroundColor White  '========================================================================================================================'

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


if ($deployment  -eq 0) {
Write-Host -ForegroundColor Gray "`nWrite into JSON file..."

$json | ConvertTo-Json -Depth 10 -Compress | Out-File -Encoding UTF8 "$fileName"

Start-Process -FilePath C:\Windows\explorer.exe -ArgumentList "/select, ""$fileName"""

Write-Host -ForegroundColor Gray "Dataflow model file ($fileName) created..."
} else {
    Write-Host -ForegroundColor Gray "`nGenerate JSON..."

    $dataflowJSON = $json | ConvertTo-Json -Depth 10 -Compress 

    $newDataFlow = _postDataflowDefinition -GroupID $selectedWorkspace.Id -DataflowDefinition $dataflowJSON -NameConflict $conflict
    
    Write-Host -ForegroundColor White ( [string]::Format("New dataflow with id '{0}' created in workspace '{1}'", $newDataFlow.id, $selectedWorkspace.Name ))
}
Write-Host -ForegroundColor White  '========================================================================================================================'
Start-Sleep -Seconds 1.5