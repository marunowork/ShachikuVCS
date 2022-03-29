# @name         VCS MENU
# @command      powershell.exe -ExecutionPolicy Bypass -file "%EXTENSION_PATH%" -SessionUrl "!E" -RepoFilePath "!/" -RepoFileName "/!" -pause
# @description  Start SHACHIKU life
# @flag         RemoteFiles 
# @version      1
# @shortcut     Shift+Ctrl+Alt+K
# @homepage     https://github.com/marunowork/ShachikuVCS
# @require      WinSCP 5.13.4
# @option       Pause -config checkbox "&Pause at the end" -pause -pause

using namespace System.Xml

Param (
    $SessionUrl,
    [Parameter(Mandatory = $True)]
    [string]$RepoFilePath,
    [Parameter(Mandatory = $True)]
    [string]$RepoFileName,
    [Switch]$pause
)
$VS_DEBUG_MODE = $false
if ($VS_DEBUG_MODE) {
    $WSCP_PATH = 'C:\Program Files (x86)\WinSCP\'
}

$USR_NAME = if ([Net.Dns]::GetHostName()) { [Net.Dns]::GetHostName() } else { $env:COMPUTERNAME }

$WSCP_PATH = if ($env:WINSCP_PATH) { $env:WINSCP_PATH } else { $PSScriptRoot }

$CONF_FILE = "${HOME}\ShckConfig.xml"
$MAP_SHCKCONF = @(
    @{ "Name"  = "pgname"
        "Type" = "String" 
    },
    @{ "Name"  = "version"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_lblHistory"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_txtFindComment"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnClear"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnFind"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_ttlDT"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_ttlUsrName"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_ttlRevisionID"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_ttlFileName"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_ttlRemoteDirectory"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_ttlLocalPath"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_ttlComment"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_ttlRevertFlag"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_lblRevID"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_lblRmtPath"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_lblComment"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_lblCmtComment"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnCheckRcnt"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnCheckSel"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnRevert"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnExport"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnComment"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnUpdate"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnCommit"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnCan"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_btnCcl"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_msgCfm"
        "Type" = "String" 
    },
    @{ "Name"  = "capt_msgReboot"
        "Type" = "String" 
    }
)
$MAP_REPOHISTORY = @(
    @{ "Name"  = "rev_id"
        "Type" = "String" 
    },
    @{ "Name"  = "file_name"
        "Type" = "String" 
    },
    @{ "Name"  = "remote_dir"
        "Type" = "String" 
    },
    @{ "Name"  = "local_dir"
        "Type" = "String" 
    },
    @{ "Name"  = "comment"
        "Type" = "String" 
    },
    @{ "Name"  = "usr_name"
        "Type" = "String" 
    },
    @{ "Name"  = "revert_flg"
        "Type" = "String" 
    }
)

<#
Setup ConfigFile
#>
function Set-ConfDef {

    $WKDIR_DEF = "${HOME}\.shchk\"
    $HSTRY_DEF = "${HOME}\.shchk\ShckHistory.xml"

    try {

        if (!(Test-Path -Path $WKDIR_DEF)) {
            New-Item $WKDIR_DEF -ItemType Directory
        }

        if (!(Test-Path -Path $CONF_FILE)) {

            $arrDef = @{}
            $atrKey = ""
            $txtVal = ""
            $objCli = New-Object System.Net.WebClient
            $objStrm = $objCli.OpenRead('https://raw.githubusercontent.com/marunowork/ShachikuVCS/main/lang/ShckConfig.ja_JP.xml')
            $objXr = [System.Xml.XmlReader]::Create($objStrm)
            while ( $objXr.Read() ) {
                switch ($objXr.NodeType) {
                    ([System.Xml.XmlNodeType]::Element).ToString() {
                        if ( $objXr.HasAttributes -and $objXr.AttributeCount -gt 1) {
                            $atrKey = $objXr.GetAttribute(0)
                        }
                    }
                    ([System.Xml.XmlNodeType]::Text).ToString() { $txtVal = $objXr.Value }
                    ([System.Xml.XmlNodeType]::EndElement).ToString() {
                        $arrDef.Add($atrKey, $txtVal) 
                        $atrKey = ""
                        $txtVal = ""            
                    }
                }
            }
            
            [xml]$doc = [XmlDocument]::new()
            $dec = $doc.CreateXmlDeclaration("1.0", "UTF-8", $null)
            $doc.AppendChild($dec) | Out-Null
            
            $settings = $doc.CreateNode("element", "settings", "")
            $settings.SetAttribute("version", "1.0")

            foreach ($aryShckConfMap in $MAP_SHCKCONF) {
                foreach ($aryConfMap in $aryShckConfMap) {

                    $defaults = $doc.CreateNode("element", "defaults", "")
                    $defaults.SetAttribute("Name", $aryConfMap.Name)
                    $defaults.SetAttribute("Type", $aryConfMap.Type)        
                    $defaults.set_InnerText($arrDef[$aryConfMap.Name])
                    $settings.AppendChild($defaults) | Out-Null    
                }
            }

            $defaults = $doc.CreateNode("element", "defaults", "")
            $defaults.SetAttribute("Name", "history")
            $defaults.SetAttribute("Type", "String")
            $defaults.set_InnerText($HSTRY_DEF)
            $settings.AppendChild($defaults) | Out-Null

            $defaults = $doc.CreateNode("element", "defaults", "")
            $defaults.SetAttribute("Name", "work_dir")
            $defaults.SetAttribute("Type", "String")
            $defaults.set_InnerText($WKDIR_DEF)
            $settings.AppendChild($defaults) | Out-Null    
            
            $doc.AppendChild($settings) | Out-Null
            $doc.Save($CONF_FILE) | Out-Null
            return 0
        }
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)"
        return 1
    }
    return 0
}

<#
Setup HistoryFile
#>
function Set-HistoryDef {

    try {

        $xmlHistoryFile = Get-ShckConfig "history"

        while (!(Test-Path $xmlHistoryFile)) {

            [xml]$doc = [XmlDocument]::new()
            $dec = $doc.CreateXmlDeclaration("1.0", "UTF-8", $null)
            $doc.AppendChild($dec) | Out-Null

            $histories = $doc.CreateNode("element", "histories", "")
            $histories.SetAttribute("version", "1.0")
            $checklists = $doc.CreateNode("element", "checklists", "")
            $histories.AppendChild($checklists) | Out-Null                

            foreach ($aryShckHistoryMap in $MAP_REPOHISTORY) {
                foreach ($aryHistoryMap in $aryShckHistoryMap) {

                    $checkpoint = $doc.CreateNode("element", "checkpoint", "")
                    $checkpoint.SetAttribute("Name", $aryHistoryMap.Name)
                    $checkpoint.SetAttribute("Type", $aryHistoryMap.Type)        
                    $checkpoint.set_InnerText("")
                    $checklists.AppendChild($checkpoint) | Out-Null    
                }
            }

            $doc.AppendChild($histories) | Out-Null    
            $doc.Save($xmlHistoryFile) | Out-Null
            Write-Host "Load Default Files..."

            Start-Sleep -Seconds 5
        }
        return New-Object System.Data.Datatable
    }
    catch {
        Write-Host "Set-HistoryDef Error: $($_.Exception.Message)"
        return 1
    }
    return 0
}

<#
Show Massage Box
#>
function Show-Msgbox(
    [string]$Text, `
        [string]$Caption = "", `
        [System.Windows.Forms.MessageBoxButtons]$MessageBoxButtons = [System.Windows.Forms.MessageBoxButtons]::OK, `
        [System.Windows.Forms.MessageBoxIcon]$MessageBoxIcon = [System.Windows.Forms.MessageBoxIcon]::Information, `
        [System.Windows.Forms.MessageBoxDefaultButton]$MessageBoxDefaultButton = [System.Windows.Forms.MessageBoxDefaultButton]::Button1
) {
    $Caption = Get-ShckConfig "pgname"
    [System.Windows.Forms.MessageBox]::Show($Text, $Caption, $MessageBoxButtons, $MessageBoxIcon, $MessageBoxDefaultButton)
}

<#
Show Confirm Box
#>
function Show-MsgboxCfm(
    [string]$Text, `
        [string]$Caption = "", `
        [System.Windows.Forms.MessageBoxButtons]$MessageBoxButtons = [System.Windows.Forms.MessageBoxButtons]::YesNo, `
        [System.Windows.Forms.MessageBoxIcon]$MessageBoxIcon = [System.Windows.Forms.MessageBoxIcon]::Question, `
        [System.Windows.Forms.MessageBoxDefaultButton]$MessageBoxDefaultButton = [System.Windows.Forms.MessageBoxDefaultButton]::Button2
) {
    $Caption = Get-ShckConfig "pgname"
    $msgBoxInput = [System.Windows.Forms.MessageBox]::Show($Text, $Caption, $MessageBoxButtons, $MessageBoxIcon, $MessageBoxDefaultButton)
    if ($msgBoxInput -eq 'Yes') { return $True }
    return $False
}

<#
Get Config Parameter
#>
function Get-ShckConfig {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$key_name
    )

    while (!(Test-Path -Path $CONF_FILE)) {

        if (Set-ConfDef eq -1) { return $null }
        Write-Host "Load Default Settings..."

        Start-Sleep -Seconds 5
    }

    $xmlConf = [xml](Get-Content -Encoding utf8 $CONF_FILE)
    $confDef = $xmlConf.settings.defaults | Where-Object { $_.Name.Contains($key_name) }
    return $confDef.InnerText
}

<#
Get WinSCP Session
#>
function Get-WSCPSession {

    Add-Type -Path (Join-Path $WSCP_PATH "WinSCPnet.dll")
    Add-Type -Assembly PresentationCore

    $sessionOptions = New-Object WinSCP.SessionOptions
    $sessionOptions.ParseUrl($SessionUrl)
    $getSession = New-Object WinSCP.Session

    try {
        $getSession.Open($sessionOptions) 
        Write-Host "Get-WSCPSession: open session"
        return $getSession
    }
    catch {
        Write-Host "Get-WSCPSession Error: $($_.Exception.Message)"
        return 1
    }
}

<#
Get RepoFile Format
#>
function Get-RepoFileFmt {
    Param (
        $rev_id,
        $remoteFilePath
    )
    try {

        $rmtFileID = $remoteFilePath.Replace("/", "_")
        $repoFileFmt = '{0}-{1}_{2}' -f $rmtFileID, $rev_id, $USR_NAME
        return $repoFileFmt
    }
    catch {
        Write-Host "Get-RepoFileFmt Error: $($_.Exception.Message)"
        return 1
    }
}

<#
Revert RepoFile
#>
function Redo-RepoFile {
    Param (
        [Parameter(Mandatory = $True)]
        $localPath,
        [Parameter(Mandatory = $True)]
        $rmtFilePath
    )
    try {
        
        $getSession = Get-WSCPSession
        $transferResult = $getSession.PutFiles($localPath, $rmtFilePath, $True)
        Write-Host "Redo-RepoFile ${rmtFilePath} -> ${localPath}"
        $transferResult.Check()

        foreach ($transfer in $transferResult.Transfers) {
            Push-RepoHistory $RemotefilePath $RemotefileName $tbxComment.Text
            return 0
        }
    }
    catch {
        Write-Host "Redo-RepoFile Error: $($_.Exception.Message)"
        return 1
    }
}

<#
Import RepoFile
#>
function Import-RepoFile {
    Param (
        [Parameter(Mandatory = $True)]
        $localPath,
        $rev_id,
        $remotePath,
        $remoteFile
    )

    $remoteFilePath = '{0}{1}' -f $remotePath, $remoteFile 

    if ($rev_id -ne "") {
        $localDirName = $localPath
        $repoFileName = Get-RepoFileFmt $rev_id $remoteFilePath
        $backupFilePath = '{0}{1}' -f $localDirName, $repoFileName 
    }
    
    else {
        $backupFilePath = '{0}{1}' -f $localPath, $remoteFile 
    }

    try {

        $getSession = Get-WSCPSession
        $getSession.GetFiles($remoteFilePath, $backupFilePath).Check()
        Write-Host "Import-RepoFile ${remoteFilePath} -> ${backupFilePath}"

        return 0
    }
    catch {
        Write-Host "Import-RepoFile Error: $($_.Exception.Message)"
        return 1
    }
}

<#
Copy RepoFile
#>
function Copy-ExportFile {
    Param (
        [Parameter(Mandatory = $True)]
        $backUpDir,
        $rev_id,
        [Parameter(Mandatory = $True)]
        $remotePath,
        $remoteFile
    )
    try {

        $remoteFilePath = '{0}{1}' -f $remotePath, $remoteFile 

        Write-Host "Copy-ExportFile: ${remoteFilePath}"

        $repoDir = Get-ShckConfig "work_dir"
        if ([string]::IsNullOrEmpty($repoDir)) { return 1 }
    
        $repoFile = Get-RepoFileFmt $rev_id $remoteFilePath
        $repoFilePath = '{0}{1}' -f $repoDir, $repoFile

        $backupFilePath = '{0}\{1}' -f $backUpDir, $remoteFile 
        
        Write-Host "Copy-ExportFile: ${repoFilePath} -> ${backupFilePath} "
        Copy-Item -Path $repoFilePath -Destination $backupFilePath
        return 0
    }
    catch {
        Write-Host "Copy-ExportFile Error: $($_.Exception.Message)"
        return 1
    }
}

<#
Add History Data
#>
function Add-RepoHistory {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0)]
        [string]$_rev_id,
        [Parameter(Position = 1)]
        [string]$_file_name,
        [Parameter(Position = 2)]
        [string]$_remote_dir,
        [Parameter(Position = 3)]
        [string]$_local_dir,
        [Parameter(Position = 4)]
        [string]$_comment,
        [Parameter(Position = 5)]
        [string]$_usr_name,
        [Parameter(Position = 6)]
        [string]$_revert_flg
    )
    try {

        $xmlHistoryFile = Get-ShckConfig "history"
        if ([string]::IsNullOrEmpty($xmlHistoryFile)) { return 1 }

        if (!(Test-Path $xmlHistoryFile)) {
            return Set-HistoryDef 
        }
        else {
            $xmlHistory = [xml](Get-Content -Encoding utf8 $xmlHistoryFile)
            
            [xml]$doc = [XmlDocument]::new()
            $dec = $doc.CreateXmlDeclaration("1.0", "UTF-8", $null)
            $doc.AppendChild($dec) | Out-Null
            
            $histories = $doc.CreateNode("element", "histories", "")
            $histories.SetAttribute("version", "1.0")

            $key_name = ""
            foreach ($_checklists in $xmlHistory.histories.checklists) {
                foreach ($_checkpoint in $_checklists.checkpoint) {
                    $key_name = $_checkpoint.Name

                    if ($key_name -eq "rev_id") {
                        $checklists = $doc.CreateNode("element", "checklists", "")
                        $histories.AppendChild($checklists) | Out-Null                
                    }
                
                    $checkpoint = $doc.CreateNode("element", "checkpoint", "")
                    $checkpoint.SetAttribute("Name", $_checkpoint.Name)
                    $checkpoint.SetAttribute("Type", $_checkpoint.Type)
                    $checkpoint.set_InnerText($_checkpoint.InnerText)
                    $checklists.AppendChild($checkpoint) | Out-Null
                }
            }
            $doc.AppendChild($histories) | Out-Null
            
            $checklists_add = $doc.CreateNode("element", "checklists", "")
            $histories.AppendChild($checklists_add) | Out-Null

            foreach ($aryHistoryMap in $MAP_REPOHISTORY) {
                foreach ($aryRepoMap in $aryHistoryMap) {

                    $getValName = $aryRepoMap.Name
                    $getValPrm = Get-variable -ValueOnly "_$getValName"

                    $checkpoint_add = $doc.CreateNode("element", "checkpoint", "")
                    $checkpoint_add.SetAttribute("Name", $aryRepoMap.Name)
                    $checkpoint_add.SetAttribute("Type", $aryRepoMap.Type)
                    $checkpoint_add.set_InnerText($getValPrm)
                    $checklists_add.AppendChild($checkpoint_add) | Out-Null        
                }
            }
            
            $doc.AppendChild($histories) | Out-Null
            $doc.Save($xmlHistoryFile) | Out-Null

            Write-Host "Add-RepoHistory:Success"
            return 0
        }
    }
    catch {
        Write-Host "Add-RepoHistory Error: $($_.Exception.Message)"
        return 1
    }

}

<#
Edit History
#>
function Edit-RepoHistory {
    [CmdletBinding()]
    Param(
        [Parameter(Position = 0)]
        [string]$key_rev_id,
        [Parameter(Position = 1)]
        [string]$_file_name,
        [Parameter(Position = 2)]
        [string]$_remote_dir,
        [Parameter(Position = 3)]
        [string]$_local_dir,
        [Parameter(Position = 4)]
        [string]$_comment,
        [Parameter(Position = 5)]
        [string]$_usr_name,
        [Parameter(Position = 6)]
        [string]$_revert_flg
    )
    try {

        $xmlHistoryFile = Get-ShckConfig "history" 
        if ([string]::IsNullOrEmpty($xmlHistoryFile)) { return 1 }

        if (!(Test-Path $xmlHistoryFile)) {
            return Set-HistoryDef 
        }
        else {
            $xmlHistory = [xml](Get-Content -Encoding utf8 $xmlHistoryFile)
            
            [xml]$doc = [XmlDocument]::new()
            $dec = $doc.CreateXmlDeclaration("1.0", "UTF-8", $null)
            $doc.AppendChild($dec) | Out-Null
            
            $histories = $doc.CreateNode("element", "histories", "")
            $histories.SetAttribute("version", "1.0")

            $key_name = ""
            $isTargetData = $False
            foreach ($_checklists in $xmlHistory.histories.checklists) {
                foreach ($_checkpoint in $_checklists.checkpoint) {

                    $key_name = $_checkpoint.Name
                    $setInnerText = $_checkpoint.InnerText

                    if ($key_name -eq "rev_id") {
                        $isTargetData = ($key_rev_id -eq $_checkpoint.InnerText)
                        $checklists = $doc.CreateNode("element", "checklists", "")
                        $histories.AppendChild($checklists) | Out-Null 
                    }
                    else {
                        if ($isTargetData) {
                            $setInnerText = Get-variable -ValueOnly "_$key_name"
                        }
                    }

                    $checkpoint = $doc.CreateNode("element", "checkpoint", "")
                    $checkpoint.SetAttribute("Name", $_checkpoint.Name)
                    $checkpoint.SetAttribute("Type", $_checkpoint.Type)
                    $checkpoint.set_InnerText($setInnerText)
                    $checklists.AppendChild($checkpoint) | Out-Null
                }
            }
            $doc.AppendChild($histories) | Out-Null    
            $doc.Save($xmlHistoryFile) | Out-Null

            Write-Host "Edit-RepoHistory:Success"
            $result = 0
        }
    }
    catch {
        Write-Host "Edit-RepoHistory Error: $($_.Exception.Message)"
        $result = 1
    }
    return $result
}

<#
Commit History
#>
function Push-RepoHistory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True)]
        [string]$remotePath,
        [string]$remoteFile,
        [string]$txtComment
    )
    try {
        
        $xmlHistoryFile = Get-ShckConfig "history"
        if ([string]::IsNullOrEmpty($xmlHistoryFile)) { return 1 }

        if (!(Test-Path $xmlHistoryFile)) {
            return Set-HistoryDef 
        }
        else {

            $local_dir = Get-ShckConfig "work_dir"
            if ([string]::IsNullOrEmpty($local_dir)) { return 1 }
            
            $rev_id = Get-Date -UFormat "%Y%m%d%H%M%S"
            $revert_flg = "0"

            Add-RepoHistory $rev_id $remoteFile $remotePath $local_dir $txtComment $USR_NAME $revert_flg
            Write-Host "Add-RepoHistory:Success"

            if (!(Test-Path -Path $local_dir)) {
                New-Item $local_dir -ItemType Directory
            }

            Import-RepoFile $local_dir $rev_id $remotePath $remoteFile
            Write-Host "Import-RepoFile:Success"

            $result = 0
        }
    }
    catch {
        Write-Host "Push-RepoHistory Error: $($_.Exception.Message)"
        $result = 1
    }
    return $result
}

<#
Search History Data With Linq
#>
function Search-RepoHistory {

    Param(
        [String]$file_name,
        [String]$keyword = ""
    )
    try {

        [void][Reflection.Assembly]::LoadWithPartialName("System.Data.DataSetExtensions")
        [void][Reflection.Assembly]::LoadWithPartialName("System.Data.Linq")
    
        $xmlHistoryFile = Get-ShckConfig "history"
        if ([string]::IsNullOrEmpty($xmlHistoryFile)) { return 1 }

        if (!(Test-Path $xmlHistoryFile)) {
            return Set-HistoryDef 
        }
        else {

            $xmlHistory = [xml](Get-Content -Encoding utf8 $xmlHistoryFile)
            $dt = New-Object System.Data.Datatable "CheckLists"
            [void]$dt.Columns.Add("rev_id")
            [void]$dt.Columns.Add("usr_name")
            [void]$dt.Columns.Add("file_name")
            [void]$dt.Columns.Add("remote_dir")
            [void]$dt.Columns.Add("local_dir")
            [void]$dt.Columns.Add("comment")
            [void]$dt.Columns.Add("revert_flg")

            foreach ($_checklists in $xmlHistory.histories.checklists) {

                $recXml = @{}
                foreach ($_checkpoint in $_checklists.checkpoint) {
                    $recXml.Add($_checkpoint.Name, $_checkpoint.InnerText)
                }

                [void]$dt.Rows.Add(
                    $recXml["rev_id"],
                    $recXml["usr_name"],
                    $recXml["file_name"],
                    $recXml["remote_dir"],
                    $recXml["local_dir"],
                    $recXml["comment"],
                    $recXml["revert_flg"]
                )
            }

            $list = [System.Data.DataTableExtensions]::AsEnumerable($dt)

            if ($keyword -ne "") {
                $wherequery = [System.Func[System.Object, bool]] { 
                    Param($row) (($row.'file_name' -eq $file_name) `
                            -and ($row.'revert_flg' -eq "0") `
                            -and ($row.'comment'.Contains($keyword)))
                }
            }
            else {
                $wherequery = [System.Func[System.Object, bool]] { 
                    Param($row) (($row.'file_name' -eq $file_name) `
                            -and ($row.'revert_flg' -eq "0"))
                }    
            }

            $orderbyQuery = [System.Func[System.Data.DataRow, string]] { param($row) $row.'rev_id' }
            $orderby = [System.Linq.Enumerable]::OrderByDescending($list, $orderbyQuery)
            
            $outQuery = [System.Linq.Enumerable]::Where($orderby, $wherequery)
            $outQueryArry = [System.Linq.Enumerable]::ToArray($outQuery);
            return $outQueryArry
        }
    }
    catch {
        Write-Host "Search-RepoHistory Error: $($_.Exception.Message)"
        return 1
    }
}

<#
Conpare Remote File
#>
function Show-DiffRemote { 
    Param (
        [String]$localPath,
        [String]$remoteFilePath
    )
    try {
        
        $getSession = Get-WSCPSession
        $tmpFileName = $remoteFilePath.Replace("/", "_")
        $cacheFilePath = $getSession.xmlLogPath.Replace(".tmp", $tmpFileName)
        Write-Host "L:${cacheFilePath} R:${remoteFilePath}"

        $getSession.GetFiles($remoteFilePath, $cacheFilePath).Check()
        Write-Host "L:${localPath} R:${cacheFilePath}"

        Invoke-WSCPExtCompFiles $localPath $cacheFilePath
        $result = 0
    }
    catch {
        Write-Host "Show-DiffRemote Error: $($_.Exception.Message)"
        $result = 1
    }
    return $result 
}

<#
Compare Last RepoFile
#>
function Show-DiffRepoLast { 
    Param (
        [String]$curRevId,
        [String]$remoteFilePath,
        [System.Windows.Forms.DataGridView]$dgView
    )
    try {

        $repoDir = Get-ShckConfig "work_dir"
        if ([string]::IsNullOrEmpty($repoDir)) { return 1 }

        $lastRevId = ""
        for ($i = 0; $i -lt $dgView.RowCount; $i++) { 
            if (($dgView.Rows[$i].Cells[0].Value -eq $curRevId)`
                    -and (($i + 1) -lt $dgView.RowCount)) {
                $lastRevId = $dgView.Rows[$i + 1].Cells[0].Value 
                break
            }
        }
        if ($lastRevId -eq "") { $lastRevId = $curRevId }

        $curFile = Get-RepoFileFmt $curRevId $remoteFilePath
        $repoCurFilePath = '{0}{1}' -f $repoDir, $curFile

        $lastFile = Get-RepoFileFmt $lastRevId $remoteFilePath
        $repoLastFilePath = '{0}{1}' -f $repoDir, $lastFile

        Invoke-WSCPExtCompFiles $repoLastFilePath $repoCurFilePath 
        $result = 0
    }
    catch {
        Write-Host "Show-DiffRepoLast Error: $($_.Exception.Message)"
        $result = 1
    }
    return $result 
}

<#
Call WinSCP Extensions
(CompareFiles.WinSCPextension.ps1)
#>
function Invoke-WSCPExtCompFiles { 
    Param (
        [String]$filePathLeft,
        [String]$filePathRight
    )
    try {

        $scpExtRoot = '{0}\Extensions' -f $WSCP_PATH 

        Set-Location -path $scpExtRoot
        Write-Host "L:${filePathLeft} R:${filePathRight}"        
    
        $params = "-localPath `"$filePathLeft`" -remotePath `"$filePathRight`" -tool `"WinMerge`"" 

        $ScriptFile = ".\CompareFiles.WinSCPextension.ps1"
        $Argument = "-File `"$ScriptFile`" $params"

        Start-Process -FilePath powershell.exe -NoNewWindow -Wait -ArgumentList $Argument
        $result = 0
    }
    catch {
        Write-Host "Invoke-WSCPExtCompFiles Error: $($_.Exception.Message)"
        $result = 1
    }
    return $result 
}

<#
Show Main Menu
#>
function Start-ShckMain {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$RemotefilePath,
        [Parameter(Mandatory)]
        [string]$RemotefileName
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $remoteRepoFilePath = '{0}{1}' -f $RemotefilePath, $RemotefileName 
    $wSelComment = ""
    $isChkDisable = $true

    <#
    Form definition
    #>
    $appTitle = Get-ShckConfig "version"
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text            = $appTitle
        Size            = New-Object Drawing.Size(800, 600)
        MaximizeBox     = $false
        FormBorderStyle = 'FixedDialog'
        Font            = New-Object Drawing.Font('Meiryo UI', 8.5)
    }

    $tbxHidden = New-Object System.Windows.Forms.TextBox -Property @{
        Location = New-Object Drawing.Point(50, 10)
        Width    = 300
        Text     = $SessionUrl
        Visible  = $VS_DEBUG_MODE
        ReadOnly = $True
    }
    
    $lblHistory = New-Object System.Windows.Forms.Label -Property @{
        Location = New-Object System.Drawing.Point(10, 15)
        Size     = New-Object System.Drawing.Size(80, 20)
        Text     = Get-ShckConfig "capt_lblHistory"
    }

    $txtFindComment = New-Object Windows.Forms.TextBox -Property @{
        Location = New-Object Drawing.Point(110, 15) 
        Size     = New-Object Drawing.Size(230, 30)    
    }

    $btnClear = New-Object System.Windows.Forms.Button -Property @{
        Location  = New-Object Drawing.Point(350, 10) 
        Size      = New-Object Drawing.Size(50, 30)
        Text      = Get-ShckConfig "capt_btnClear"
        FlatStyle = "popup"
    }

    $btnFind = New-Object System.Windows.Forms.Button -Property @{
        Location  = New-Object Drawing.Point(410, 10) 
        Size      = New-Object Drawing.Size(90, 30)
        Text      = Get-ShckConfig "capt_btnFind"
        FlatStyle = "popup"
    }

    $Grid = New-Object System.Windows.Forms.DataGridView -Property @{
        Location            = New-Object Drawing.Point(10, 50)
        Size                = New-Object Drawing.Size(495, 240)
        AutoSizeColumnsMode = "Fill"
        AutoSizeRowsMode    = "None"
        ReadOnly            = $True
        AllowUserToAddRows  = $false
        ColumnCount         = 8
        RowHeadersVisible   = $false
        MultiSelect         = $false
        SelectionMode       = 'FullRowSelect'
    }

    $Grid.Columns[0].Name = Get-ShckConfig "capt_ttlRevisionID"
    $Grid.Columns[1].Name = Get-ShckConfig "capt_ttlDT"
    $Grid.Columns[2].Name = Get-ShckConfig "capt_ttlUsrName"
    $Grid.Columns[3].Name = Get-ShckConfig "capt_ttlFileName"
    $Grid.Columns[4].Name = Get-ShckConfig "capt_ttlRemoteDirectory"
    $Grid.Columns[5].Name = Get-ShckConfig "capt_ttlLocalPath"
    $Grid.Columns[6].Name = Get-ShckConfig "capt_ttlComment"
    $Grid.Columns[7].Name = Get-ShckConfig "capt_ttlRevertFlag"

    $Grid.Columns[0].Visible = $false
    $Grid.Columns[2].Visible = $false
    $Grid.Columns[3].Visible = $false
    $Grid.Columns[4].Visible = $false
    $Grid.Columns[5].Visible = $false
    $Grid.Columns[7].Visible = $false

    foreach ($row in Search-RepoHistory $RemotefileName ) {

        $wDspComment = $row["comment"]
        [void]$Grid.Rows.Add(
            $row["rev_id"],
            [DateTime]::ParseExact($row["rev_id"], "yyyyMMddHHmmss", $null),
            $row["usr_name"],
            $row["file_name"],
            $row["remote_dir"],
            $row["local_dir"],
            $wDspComment.Replace("`n", "`r`n"),
            $row["revert_flg"]
        )
    }

    if ($Grid.Rows.Count -gt 0) {
        $Grid.Rows[0].Selected = $true
        $wSelComment = $Grid.Rows[0].Cells[6].Value
    }
    else {
        $isChkDisable = $false
    }

    $lblRevID = New-Object System.Windows.Forms.Label  -Property @{
        Location = New-Object System.Drawing.Point(520, 30)
        Size     = New-Object System.Drawing.Size(250, 20)
        Text     = Get-ShckConfig "capt_ttlFileName"    
    }

    $tbxRevID = New-Object System.Windows.Forms.TextBox -Property @{
        Location = New-Object Drawing.Point(520, 50)
        Width    = 250
        Text     = $RemotefileName
        ReadOnly = $True
    }

    $lblRmtPath = New-Object System.Windows.Forms.Label -Property @{
        Location = New-Object System.Drawing.Point(520, 80)
        Size     = New-Object System.Drawing.Size(250, 20)
        Text     = Get-ShckConfig "capt_lblRmtPath"
    }

    $tbxRmtPath = New-Object System.Windows.Forms.TextBox -Property @{
        Location = New-Object Drawing.Point(520, 100)
        Width    = 250
        Text     = $RemotefilePath
        ReadOnly = $True
    }

    $lblComment = New-Object System.Windows.Forms.Label -Property @{
        Location = New-Object System.Drawing.Point(520, 130)
        Size     = New-Object System.Drawing.Size(250, 20)
        Text     = Get-ShckConfig "capt_lblComment"    
    }

    $tbxComment = New-Object System.Windows.Forms.TextBox -Property @{
        Location      = New-Object Drawing.Point(520, 150)
        Multiline     = $True
        AcceptsReturn = $True
        AcceptsTab    = $True
        WordWrap      = $True
        Width         = 250
        Height        = 180
        ScrollBars    = [System.Windows.Forms.ScrollBars]::Vertical
        Anchor        = (([System.Windows.Forms.AnchorStyles]::Left) `
                -bor ([System.Windows.Forms.AnchorStyles]::Top) `
                -bor ([System.Windows.Forms.AnchorStyles]::Right) `
                -bor ([System.Windows.Forms.AnchorStyles]::Bottom))
        Text          = $wSelComment
        ReadOnly      = $True
    }
    
    $lblCmtComment = New-Object System.Windows.Forms.Label -Property @{
        Location = New-Object System.Drawing.Point(10, 400)
        Size     = New-Object System.Drawing.Size(250, 20)
        Text     = Get-ShckConfig "capt_lblCmtComment"    
    }

    $txtCmtComment = New-Object Windows.Forms.TextBox -Property @{
        Location      = New-Object Drawing.Point(10, 420) 
        Multiline     = $True
        AcceptsReturn = $True
        AcceptsTab    = $True
        WordWrap      = $True
        Width         = 495
        Height        = 130
        ScrollBars    = [System.Windows.Forms.ScrollBars]::Vertical
        Anchor        = (([System.Windows.Forms.AnchorStyles]::Left) `
                -bor ([System.Windows.Forms.AnchorStyles]::Top) `
                -bor ([System.Windows.Forms.AnchorStyles]::Right) `
                -bor ([System.Windows.Forms.AnchorStyles]::Bottom))        
    }

    $btnCheckRcnt = New-Object System.Windows.Forms.Button -Property @{
        Location  = New-Object Drawing.Point(10, 300)
        Size      = New-Object Drawing.Size(240, 30)
        Text      = Get-ShckConfig "capt_btnCheckRcnt"
        FlatStyle = "popup"
        Enabled   = $isChkDisable
    }

    $btnCheckSel = New-Object System.Windows.Forms.Button -Property @{
        Location  = New-Object Drawing.Point(265, 300)
        Size      = New-Object Drawing.Size(240, 30)
        Text      = Get-ShckConfig "capt_btnCheckSel"
        FlatStyle = "popup"
        Enabled   = $isChkDisable
    }

    $btnRevert = New-Object System.Windows.Forms.Button -Property @{
        Location  = New-Object Drawing.Point(10, 340)
        Size      = New-Object Drawing.Size(240, 30)
        Text      = Get-ShckConfig "capt_btnRevert"
        FlatStyle = "popup"
        Enabled   = $isChkDisable
    }

    $fbdSelect = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        ShowNewFolderButton = $false
    }
    $btnExport = New-Object System.Windows.Forms.Button -Property @{
        Location  = New-Object Drawing.Point(265, 340)
        Size      = New-Object Drawing.Size(240, 30)
        Text      = Get-ShckConfig "capt_btnExport"
        FlatStyle = "popup"
        Enabled   = $isChkDisable
    }

    $btnComment = New-Object System.Windows.Forms.Button -Property @{
        Location  = New-Object Drawing.Point(520, 340) 
        Size      = New-Object Drawing.Size(250, 30)
        Text      = Get-ShckConfig "capt_btnComment"
        FlatStyle = "popup"
        Enabled   = $isChkDisable
    }

    $btnCommit = New-Object System.Windows.Forms.Button -Property @{
        Location  = New-Object Drawing.Point(520, 430) 
        Size      = New-Object Drawing.Size(250, 60)
        Text      = Get-ShckConfig "capt_btnCommit"
        FlatStyle = "popup"
    }

    $btnCan = New-Object System.Windows.Forms.Button -Property @{
        Location     = New-Object Drawing.Point(520, 500)
        Size         = New-Object Drawing.Size(250, 40)
        Text         = Get-ShckConfig "capt_btnCan"
        FlatStyle    = "popup"
        DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    }

    <#
    Lambda Function definition
    #>
    $scrbFindGrid = {
        $Grid.Rows.Clear()
        foreach ($row in Search-RepoHistory $RemotefileName $txtFindComment.Text) { 
            $wDspComment = $row["comment"]

            [void]$Grid.Rows.Add(
                $row["rev_id"],
                [DateTime]::ParseExact($row["rev_id"], "yyyyMMddHHmmss", $null),
                $row["usr_name"],
                $row["file_name"],
                $row["remote_dir"],
                $row["local_dir"],
                $wDspComment.Replace("`n", "`r`n"),
                $row["revert_flg"]
            )
        }

        if ($Grid.Rows.Count -gt 0) {
            $Grid.Rows[0].Selected = $true
            $tbxComment.Text = $Grid.Rows[0].Cells[6].Value
            $isChkDisable = $true
        } 
        else {
            $isChkDisable = $false
        }

        $btnCheckRcnt.Enabled = $isChkDisable
        $btnCheckSel.Enabled = $isChkDisable
        $btnRevert.Enabled = $isChkDisable
        $btnExport.Enabled = $isChkDisable
        $btnComment.Enabled = $isChkDisable
    }

    $actBtnComment = {
        if ($tbxComment.ReadOnly) {
            $btnComment.Enabled = $true
            $txtCmtComment.ReadOnly = $true
            $tbxComment.ReadOnly = $false
            $btnComment.Text = Get-ShckConfig "capt_btnCcl" 
            $btnCommit.Text = Get-ShckConfig "capt_btnUpdate"
        }
        else {
            $tbxComment.Text = $Grid.SelectedRows.Cells[6].Value
            $txtCmtComment.ReadOnly = $false
            $tbxComment.ReadOnly = $true
            $btnComment.Text = Get-ShckConfig "capt_btnComment" 
            $btnCommit.Text = Get-ShckConfig "capt_btnCommit"
        }
    }

    <#
    Action definition
    #>
    $Grid.Add_CellMouseClick({
            $tbxComment.Text = $Grid.SelectedRows.Cells[6].Value
        }
    )

    $btnClear.add_click({
            $txtFindComment.Text = ""
            . $scrbFindGrid
        }
    )
    
    $btnFind.add_click({
            . $scrbFindGrid
        }
    )

    $btnCheckRcnt.add_click({

            $btnCheckRcnt.Enabled = $false
    
            $repoFileName = Get-RepoFileFmt $Grid.SelectedRows.Cells[0].Value $remoteRepoFilePath
            $localRepoFile = Join-Path $Grid.SelectedRows.Cells[5].Value $repoFileName

            if (Show-DiffRemote $localRepoFile $remoteRepoFilePath -eq 0) {
                $btnCheckRcnt.Enabled = $true
            }
        }
    )

    $btnCheckSel.add_click({
    
            $btnCheckSel.Enabled = $false

            if (Show-DiffRepoLast $Grid.SelectedRows.Cells[0].Value $remoteRepoFilePath $Grid -eq 0) {
                $btnCheckSel.Enabled = $true
            }
        }
    )

    $btnRevert.add_click({
            $txtMsgCfm = Get-ShckConfig "capt_msgCfm"
            if (Show-MsgboxCfm -text $txtMsgCfm) {

                $btnRevert.Enabled = $false

                $repoFileName = Get-RepoFileFmt $Grid.SelectedRows.Cells[0].Value $remoteRepoFilePath
                $localRepoFile = Join-Path $Grid.SelectedRows.Cells[5].Value $repoFileName

                $result = Redo-RepoFile $localRepoFile $remoteRepoFilePath

                $result = Edit-RepoHistory `
                    $Grid.SelectedRows.Cells[0].Value $Grid.SelectedRows.Cells[3].Value `
                    $Grid.SelectedRows.Cells[4].Value $Grid.SelectedRows.Cells[5].Value `
                    $Grid.SelectedRows.Cells[6].Value $USR_NAME 1
                if ($result -eq 0) { $btnRevert.Enabled = $true }
            }
            . $scrbFindGrid
        }
    )

    $btnExport.add_click({
            if ($fbdSelect.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {

                $btnExport.Enabled = $false
                $select_dir = $fbdSelect.SelectedPath

                $result = Copy-ExportFile $select_dir $Grid.SelectedRows.Cells[0].Value $RemotefilePath $RemotefileName
                Write-Host "export success"

                if ($result -eq 0) { $btnExport.Enabled = $true }
            }
        }
    )

    $btnComment.add_click({
            . $actBtnComment
        }
    )

    $btnCommit.add_click({
            $btnCommit.Enabled = $false

            if ($tbxComment.ReadOnly) {

                if ($txtCmtComment.Text -eq "") {
                    $getMsg = Get-ShckConfig "capt_lblCmtComment"
                    Show-Msgbox -text $getMsg
                    $btnCommit.Enabled = $true
                    return 1
                }

                $result = Push-RepoHistory $RemotefilePath $RemotefileName $txtCmtComment.Text
                . $scrbFindGrid
                Write-Host "commit success"
                if ($result -eq 0) {
                    $txtCmtComment.Text = ""
                    $btnCommit.Enabled = $true
                }
            }
            else {

                if ($tbxComment.Text -eq "") {
                    $getMsg = Get-ShckConfig "capt_lblCmtComment"
                    Show-Msgbox -text $getMsg
                    $btnCommit.Enabled = $true
                    return 1
                }

                $result = Edit-RepoHistory `
                    $Grid.SelectedRows.Cells[0].Value $Grid.SelectedRows.Cells[3].Value `
                    $Grid.SelectedRows.Cells[4].Value $Grid.SelectedRows.Cells[5].Value `
                    $tbxComment.Text $USR_NAME $Grid.SelectedRows.Cells[7].Value
                
                . $actBtnComment
                
                . $scrbFindGrid
                Write-Host "comment update"

                if ($result -eq 0) { $btnCommit.Enabled = $true }
            }
        }
    )

    $form.Controls.AddRange(@(
            $tbxHidden,
            $lblHistory,
            $txtFindComment, 
            $btnClear,
            $btnFind,
            $Grid, 
            $lblRevID,
            $tbxRevID,
            $lblRmtPath,
            $tbxRmtPath,
            $tbxComment,
            $lblComment,
            $lblCmtComment, 
            $txtCmtComment, 
            $btnComment,
            $btnCommit, 
            $btnCheckRcnt, 
            $btnCheckSel, 
            $btnRevert, 
            $btnExport,
            $btnCan))
    $form.Showdialog()
}

try {

    $tempFileName = Split-Path $RepoFileName -Leaf

    Start-ShckMain $RepoFilePath $tempFileName
    $result = 0
}
catch {

    Write-Host "Main Error: $($_.Exception.Message)"
    $result = 1
}

if ($pause) {
    Write-Host "Press any key to exit..."
    [System.Console]::ReadKey() | Out-Null
}
exit $result
