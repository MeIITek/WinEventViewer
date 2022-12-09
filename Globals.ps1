#######################################
#           VARIABLES
#######################################
#region

$Global:Me = [hashtable]::Synchronized(@{
		Script = @{
			Contact = 'mitch.ermey@altairglobal.com'
			Name    = $null
			Path    = $null
			Root    = $null
			Version = $null
		}
		StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
	})

$DefaultLogs = @(
    'Application'
    'Security'
    'System'
    'Windows Powershell'
)

$global:Credentials = $null
$global:Sessions = @{
}

$Properties = @(
    'RecordNumber',
    'EventCode',
    'TimeGenerated',
    'Type',
    'Message',
    'User',
    'SourceName'
)

$Title = 'Win Event Viewer'

If ($HostInvocation)
{
    $Me.Script.Path = $hostinvocation.MyCommand.path
}
Else
{
    $Me.Script.Path = $script:MyInvocation.MyCommand.Path
}

$Data = [System.IO.FileInfo]$Me.Script.Path

If ($Data.VersionInfo.FileVersion)
{
    $Me.Script.Version = $Data.VersionInfo.FileVersion
}
Else
{
    $Me.Script.Version = $Data.LastWriteTimeUtc.ToString('yyyy.MMdd.HHmmss')
}

$Me.Script.Name = ([System.IO.Path]::GetFileNameWithoutExtension($Me.Script.Path)).Replace('.Run', '')
$Me.Script.Root = [System.IO.Directory]::GetParent($Me.Script.Path)
#endregion
#######################################
#           FUNCTIONS
#######################################
#region
Function Update-ComboBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ComboBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ComboBox control.
	
	.PARAMETER ComboBox
		The ComboBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ComboBox's Items collection.
	
	.PARAMETER DisplayMember
		Indicates the property to display for the items in this control.
		
	.PARAMETER ValueMember
		Indicates the property to use for the value of the control.
	
	.PARAMETER Append
		Adds the item(s) to the ComboBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red" -Append
		Update-ComboBox $combobox1 "White" -Append
		Update-ComboBox $combobox1 "Blue" -Append
	
	.EXAMPLE
		Update-ComboBox $combobox1 (Get-Process) "ProcessName"
	
	.NOTES
		Additional information about the function.
#>
    
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        [System.Windows.Forms.ComboBox]$ComboBox,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNull()]
        $Items,
        
        [Parameter(Mandatory = $false)]
        [string]$DisplayMember,
        
        [Parameter(Mandatory = $false)]
        [string]$ValueMember,
        
        [switch]$Append
    )
    
    If (-not $Append)
    {
        $ComboBox.Items.Clear()
    }
    
    If ($Items -is [Object[]])
    {
        $ComboBox.Items.AddRange($Items)
    }
    ElseIf ($Items -is [System.Collections.IEnumerable])
    {
        $ComboBox.BeginUpdate()
        ForEach ($obj In $Items)
        {
            $ComboBox.Items.Add($obj)
        }
        $ComboBox.EndUpdate()
    }
    Else
    {
        $ComboBox.Items.Add($Items)
    }
    
    $ComboBox.DisplayMember = $DisplayMember
    $ComboBox.ValueMember = $ValueMember
}
#endregion
###########################################################
##                                                       ##
##                    END OF GLOBALS                     ##
##                                                       ##
###########################################################