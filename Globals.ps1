#######################################
#           VARIABLES
#######################################
#region
$Script:Me = [hashtable]::Synchronized(@{
        Cache = @{
			Credentials = $null
            EndDate     = [DateTime]::Now
            Entries     = New-Object System.Collections.Generic.List['Object']
            Node        = $ENV:COMPUTERNAME
			StartDate   = [DateTime]::Now.AddDays(-3)
        }
        DefaultLogs = @(
            'Application'
            'Security'
            'System'
            'Windows Powershell'
        )
        Script = @{
			Contact = 'mitch.ermey@altairglobal.com'
			Name    = $null
			Path    = $null
            Root    = $null
            Title   = 'Win Event Viewer'
			Version = $null
		}
        StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
        Validation = @{
            Node  = $true
            type  = $true
            Log   = $true
            Start = $true
            End   = $true
            Count = $true
        }
	})

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
#             CLASSES
#######################################
#region
Class LogEntries
{
    [Int]$Count
    [String]$Node
    [String]$Log
    [Int]$EventID
    [DateTime]$TimeStamp
    [String]$EntryType
    [String]$Message
    [String]$Provider
    
    LogEntries ($InputObject)
    {
        If (-not [String]::IsNullOrEmpty($InputObject.MachineName))
        {
            $this.Node = $InputObject.MachineName.Split('.')[0]
        }
        
        If (-not [String]::IsNullOrEmpty($InputObject.ContainerLog))
        {
            $this.Log = $InputObject.ContainerLog
        }
        
        If ($InputObject.ID -gt 0)
        {
            $this.EventID = $InputObject.ID
        }
        
        If ($InputObject.TimeCreated -ne $null)
        {
            $this.TimeStamp = $InputObject.TimeCreated
        }
        
        If (-not [String]::IsNullOrEmpty($InputObject.LevelDisplayName))
        {
            $this.EntryType = $InputObject.LevelDisplayName
        }
        Else
        {
            $this.EntryType = 'Information'
        }
        
        If (-not [String]::IsNullOrEmpty($InputObject.Message))
        {
            $this.Message = $InputObject.Message
        }
        
        If (-not [String]::IsNullOrEmpty($InputObject.ProviderName))
        {
            $this.Provider = $InputObject.ProviderName
        }
    }
}
#endregion
#######################################
#             STRINGS
#######################################
#region

#endregion
#######################################
#           SCRIPT BLOCK
#######################################
#region

#endregion
#######################################
#           FUNCTIONS
#######################################
#region
function Update-ComboBox
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
	
	param
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
	
	if (-not $Append)
	{
		$ComboBox.Items.Clear()
	}
	
	if ($Items -is [Object[]])
	{
		$ComboBox.Items.AddRange($Items)
	}
	elseif ($Items -is [System.Collections.IEnumerable])
	{
		$ComboBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ComboBox.Items.Add($obj)
		}
		$ComboBox.EndUpdate()
	}
	else
	{
		$ComboBox.Items.Add($Items)
	}
	
	if ($DisplayMember)
	{
		$ComboBox.DisplayMember = $DisplayMember
	}
	
	if ($ValueMember)
	{
		$ComboBox.ValueMember = $ValueMember
	}
}

Function Set-ControlTheme
{
<#
	.SYNOPSIS
		Applies a theme to the control and its children.
	
	.PARAMETER Control
		The control to theme. Usually the form itself.
	
	.PARAMETER Theme
		The color theme:
		Light
		Dark

	.PARAMETER CustomColor
		A hashtable that contains the color values.
		Keys:
		WindowColor
		ContainerColor
		BackColor
		ForeColor
		BorderColor
		SelectionForeColor
		SelectionBackColor
		MenuSelectionColor
	.EXAMPLE
		PS C:\> Set-ControlTheme -Control $form1 -Theme Dark
	
	.EXAMPLE
		PS C:\> Set-ControlTheme -Control $form1 -CustomColor @{ WindowColor = 'White'; ContainerBackColor = 'Gray'; BackColor... }
	.NOTES
		Created by SAPIEN Technologies, Inc.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.ComponentModel.Component]$Control,
		[ValidateSet('Light', 'Dark')]
		[string]$Theme = 'Dark',
		[System.Collections.Hashtable]$CustomColor
	)
	
	$Font = [System.Drawing.Font]::New('Segoe UI', 9)
	
	#Initialize the colors
	if ($Theme -eq 'Dark')
	{
		$formMain.Tag = 'Dark'
		$WindowColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
		$ContainerColor = [System.Drawing.Color]::FromArgb(45, 45, 45)
		$BackColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
		$ForeColor = [System.Drawing.Color]::White
		$BorderColor = [System.Drawing.Color]::DimGray
		$SelectionBackColor = [System.Drawing.SystemColors]::Highlight
		$SelectionForeColor = [System.Drawing.Color]::White
		$MenuSelectionColor = [System.Drawing.Color]::DimGray
	}
	else
	{
		$formMain.Tag = 'Light'
		$WindowColor = [System.Drawing.Color]::White
		$ContainerColor = [System.Drawing.Color]::WhiteSmoke
		$BackColor = [System.Drawing.Color]::Gainsboro
		$ForeColor = [System.Drawing.Color]::Black
		$BorderColor = [System.Drawing.Color]::DimGray
		$SelectionBackColor = [System.Drawing.SystemColors]::Highlight
		$SelectionForeColor = [System.Drawing.Color]::White
		$MenuSelectionColor = [System.Drawing.Color]::LightSteelBlue
	}
	
	if ($CustomColor)
	{
		#Check and Validate the custom colors:
		$Color = $CustomColor.WindowColor -as [System.Drawing.Color]
		if ($Color) { $WindowColor = $Color }
		$Color = $CustomColor.ContainerColor -as [System.Drawing.Color]
		if ($Color) { $ContainerColor = $Color }
		$Color = $CustomColor.BackColor -as [System.Drawing.Color]
		if ($Color) { $BackColor = $Color }
		$Color = $CustomColor.ForeColor -as [System.Drawing.Color]
		if ($Color) { $ForeColor = $Color }
		$Color = $CustomColor.BorderColor -as [System.Drawing.Color]
		if ($Color) { $BorderColor = $Color }
		$Color = $CustomColor.SelectionBackColor -as [System.Drawing.Color]
		if ($Color) { $SelectionBackColor = $Color }
		$Color = $CustomColor.SelectionForeColor -as [System.Drawing.Color]
		if ($Color) { $SelectionForeColor = $Color }
		$Color = $CustomColor.MenuSelectionColor -as [System.Drawing.Color]
		if ($Color) { $MenuSelectionColor = $Color }
	}
	
	#Define the custom renderer for the menus
	#region Add-Type definition
	try
	{
		[SAPIENTypes.SAPIENColorTable] | Out-Null
	}
	catch
	{
		if ($PSVersionTable.PSVersion.Major -ge 7)
		{
			$Assemblies = 'System.Windows.Forms', 'System.Drawing', 'System.Drawing.Primitives'
		}
		else
		{
			$Assemblies = 'System.Windows.Forms', 'System.Drawing'
		}
		Add-Type -ReferencedAssemblies $Assemblies -TypeDefinition "
using System;
using System.Windows.Forms;
using System.Drawing;
namespace SAPIENTypes
{
    public class SAPIENColorTable : ProfessionalColorTable
    {
        Color ContainerBackColor;
        Color BackColor;
        Color BorderColor;
		Color SelectBackColor;

        public SAPIENColorTable(Color containerColor, Color backColor, Color borderColor, Color selectBackColor)
        {
            ContainerBackColor = containerColor;
            BackColor = backColor;
            BorderColor = borderColor;
			SelectBackColor = selectBackColor;
        } 
		public override Color MenuStripGradientBegin { get { return ContainerBackColor; } }
        public override Color MenuStripGradientEnd { get { return ContainerBackColor; } }
        public override Color ToolStripBorder { get { return BorderColor; } }
        public override Color MenuItemBorder { get { return SelectBackColor; } }
        public override Color MenuItemSelected { get { return SelectBackColor; } }
        public override Color SeparatorDark { get { return BorderColor; } }
        public override Color ToolStripDropDownBackground { get { return BackColor; } }
        public override Color MenuBorder { get { return BorderColor; } }
        public override Color MenuItemSelectedGradientBegin { get { return SelectBackColor; } }
        public override Color MenuItemSelectedGradientEnd { get { return SelectBackColor; } }      
        public override Color MenuItemPressedGradientBegin { get { return ContainerBackColor; } }
        public override Color MenuItemPressedGradientEnd { get { return ContainerBackColor; } }
        public override Color MenuItemPressedGradientMiddle { get { return ContainerBackColor; } }
        public override Color ImageMarginGradientBegin { get { return BackColor; } }
        public override Color ImageMarginGradientEnd { get { return BackColor; } }
        public override Color ImageMarginGradientMiddle { get { return BackColor; } }
    }
}"
	}
	#endregion
	
	$colorTable = New-Object SAPIENTypes.SAPIENColorTable -ArgumentList $ContainerColor, $BackColor, $BorderColor, $MenuSelectionColor
	$render = New-Object System.Windows.Forms.ToolStripProfessionalRenderer -ArgumentList $colorTable
	[System.Windows.Forms.ToolStripManager]::Renderer = $render
	
	#Set up our processing queue
	$Queue = New-Object System.Collections.Generic.Queue[System.ComponentModel.Component]
	$Queue.Enqueue($Control)
	
	Add-Type -AssemblyName System.Core
	
	#Only process the controls once.
	$Processed = New-Object System.Collections.Generic.HashSet[System.ComponentModel.Component]
	
	#Apply the colors to the controls
	while ($Queue.Count -gt 0)
	{
		$target = $Queue.Dequeue()
		
		#Skip controls we already processed
		if ($Processed.Contains($target)) { continue }
		$Processed.Add($target)
		
		#Set the text color
		$target.ForeColor = $ForeColor
		
		#region Handle Controls
		if ($target -is [System.Windows.Forms.Form])
		{
			#Set Font
			$target.Font = $Font
			$target.BackColor = $ContainerColor
		}
		elseif ($target -is [System.Windows.Forms.SplitContainer])
		{
			$target.BackColor = $BorderColor
		}
		elseif ($target -is [System.Windows.Forms.PropertyGrid])
		{
			$target.BackColor = $BorderColor
			$target.ViewBackColor = $BackColor
			$target.ViewForeColor = $ForeColor
			$target.ViewBorderColor = $BorderColor
			$target.CategoryForeColor = $ForeColor
			$target.CategorySplitterColor = $ContainerColor
			$target.HelpBackColor = $BackColor
			$target.HelpForeColor = $ForeColor
			$target.HelpBorderColor = $BorderColor
			$target.CommandsBackColor = $BackColor
			$target.CommandsBorderColor = $BorderColor
			$target.CommandsForeColor = $ForeColor
			$target.LineColor = $ContainerColor
		}
		elseif ($target -is [System.Windows.Forms.ContainerControl] -or
			$target -is [System.Windows.Forms.Panel])
		{
			#Set the BackColor for the container
			$target.BackColor = $ContainerColor
			
		}
		elseif ($target -is [System.Windows.Forms.GroupBox])
		{
			$target.FlatStyle = 'Flat'
		}
		elseif ($target -is [System.Windows.Forms.Button])
		{
			$target.FlatStyle = 'Flat'
			$target.FlatAppearance.BorderColor = $BorderColor
			$target.BackColor = $BackColor
		}
		elseif ($target -is [System.Windows.Forms.CheckBox] -or
			$target -is [System.Windows.Forms.RadioButton] -or
			$target -is [System.Windows.Forms.Label])
		{
			#$target.FlatStyle = 'Flat'
		}
		elseif ($target -is [System.Windows.Forms.ComboBox])
		{
			$target.BackColor = $BackColor
			$target.FlatStyle = 'Flat'
		}
		elseif ($target -is [System.Windows.Forms.TextBox])
		{
			$target.BorderStyle = 'FixedSingle'
			$target.BackColor = $BackColor
		}
		elseif ($target -is [System.Windows.Forms.DataGridView])
		{
			$target.GridColor = $BorderColor
			$target.BackgroundColor = $ContainerColor
			$target.DefaultCellStyle.BackColor = $WindowColor
			$target.DefaultCellStyle.SelectionBackColor = $SelectionBackColor
			$target.DefaultCellStyle.SelectionForeColor = $SelectionForeColor
			$target.ColumnHeadersDefaultCellStyle.BackColor = $ContainerColor
			$target.ColumnHeadersDefaultCellStyle.ForeColor = $ForeColor
			$target.EnableHeadersVisualStyles = $false
			$target.ColumnHeadersBorderStyle = 'Single'
			$target.RowHeadersBorderStyle = 'Single'
			$target.RowHeadersDefaultCellStyle.BackColor = $ContainerColor
			$target.RowHeadersDefaultCellStyle.ForeColor = $ForeColor
			
		}
		elseif ($PSVersionTable.PSVersion.Major -le 5 -and $target -is [System.Windows.Forms.DataGrid])
		{
			$target.CaptionBackColor = $WindowColor
			$target.CaptionForeColor = $ForeColor
			$target.BackgroundColor = $ContainerColor
			$target.BackColor = $WindowColor
			$target.ForeColor = $ForeColor
			$target.HeaderBackColor = $ContainerColor
			$target.HeaderForeColor = $ForeColor
			$target.FlatMode = $true
			$target.BorderStyle = 'FixedSingle'
			$target.GridLineColor = $BorderColor
			$target.AlternatingBackColor = $ContainerColor
			$target.SelectionBackColor = $SelectionBackColor
			$target.SelectionForeColor = $SelectionForeColor
		}
		elseif ($target -is [System.Windows.Forms.ToolStrip])
		{
			
			$target.BackColor = $BackColor
			$target.Renderer = $render
			
			foreach ($item in $target.Items)
			{
				$Queue.Enqueue($item)
			}
		}
		elseif ($target -is [System.Windows.Forms.ToolStripMenuItem] -or
			$target -is [System.Windows.Forms.ToolStripDropDown] -or
			$target -is [System.Windows.Forms.ToolStripDropDownItem])
		{
			$target.BackColor = $BackColor
			foreach ($item in $target.DropDownItems)
			{
				$Queue.Enqueue($item)
			}
		}
		elseif ($target -is [System.Windows.Forms.ListBox] -or
			$target -is [System.Windows.Forms.ListView] -or
			$target -is [System.Windows.Forms.TreeView])
		{
			$target.BackColor = $WindowColor
		}
		else
		{
			$target.BackColor = $BackColor
		}
		#endregion
		
		if ($target -is [System.Windows.Forms.Control])
		{
			#Queue all the child controls
			foreach ($child in $target.Controls)
			{
				$Queue.Enqueue($child)
			}
		}
	}
}

Function Set-PanelExecte
{
    [CmdletBinding(ConfirmImpact = 'Low',
                   PositionalBinding = $true)]
    [OutputType([void])]
    Param ()
    
    Process
    {
        $Enabled = $true
        
        ForEach ($Key in $Me.Validation.Keys)
        {
            If ($Me.Validation.$Key -ne $true)
            {
                $Enabled = $false
                break
            }
        }
        
        If ($Enabled -eq $true)
        {
			$panelExecute.Enabled = $true
			
			If ($formMain.Tag -eq 'Dark')
			{
				$panelExecute.ForeColor = 'White'
			}
			Else
			{
				$panelExecute.ForeColor = 'Black'
			}
        }
        Else
        {
			$panelExecute.Enabled = $false
			$panelExecute.ForeColor = 'Gray'
        }
    }
}

Function Get-LogEntries
{
    [CmdletBinding(ConfirmImpact = 'Low',
                   PositionalBinding = $true)]
    [OutputType([LogEntries])]
    Param ()
    
    Process
	{
        $EventParam = @{
            FilterHashtable = $null
            ComputerName    = $textboxHost.Text.Trim()
            MaxEvents       = [int]$numericupdownRecordCount.Text.Trim()
            ErrorAction     = 'SilentlyContinue'
        }
        
        $FilterHashtable = @{
            Logname   = $comboboxLogs.Text.Trim()
            StartTime = $Me.Cache.StartDate
            EndTime   = $Me.Cache.EndDate
        }
        
        If ($Me.Cache.Node -ne '127.0.0.1' -and $Me.Cache.Node -ne $Env:COMPUTERNAME)
		{
			If ($Me.Cache.Credentials -eq $null)
			{
				$Me.Cache.Credentials = Get-Credential
			}
			
            $EventParam.Add('Credential', $Me.Cache.Credentials)
        }
        
        Switch ($comboboxType.Text.Trim())
        {
            'Critilcal'
            {
                $FilterHashtable.Add('Level', 1)
                Break
            }
            
            'Error'
            {
                $FilterHashtable.Add('Level', 2)
                Break
            }
            
            'Warning'
            {
                $FilterHashtable.Add('Level', 3)
                Break
            }
            
            'Information'
            {
                $FilterHashtable.Add('Level', 4)
                Break
            }
            
            'Verbose'
            {
                $FilterHashtable.Add('Level', 5)
                Break
            }
        }
		
		$Cnt = 0
		$EventParam.FilterHashtable = $FilterHashtable
		$Me.Cache.Entries = New-Object System.Collections.Generic.List['Object']
        
        Get-WinEvent @EventParam | Sort-Object TimeCreated -Descending | ForEach-Object `
		{
            $Data = [LogEntries]::New($_)
            $Data.Count = ($Cnt++)
            [Void]$Me.Cache.Entries.Add($Data)
		}
		
		Write-Output $Me.Cache.Entries
	}
}
#endregion
###########################################################
##                                                       ##
##                    END OF GLOBALS                     ##
##                                                       ##
###########################################################