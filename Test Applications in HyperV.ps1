<#

Written by Ryan Ephgrave

This tool will create a snapshot of the VM and then install all the applications selected, 
reverting to the previous snapshot after each install. Results will be logged in the same folder in
ApplicationInstalls.log

Usage:
Computer Name = Name of VM you wish to install Applications on as listed in AD
VM Name = Name of the VM as listed in Hyper-V

This needs to be run on the computer running Hyper-V!

#>

$ScriptName = $MyInvocation.MyCommand.path
$Directory = Split-Path $ScriptName
$Popup = new-object -comobject wscript.shell
$Script:LogFile = "$Directory\ApplicationInstalls.log"
$Script:SaveCMMLogFiles = $false
$Script:CreateCheckPointPerApp = $false
$Script:AppList = @()
$Script:strVMName = ""
$Script:strCompName = ""
$Script:VMObject = $null
$Script:SaveLogsSuccessful = $false
$Script:SaveLogsError = $false
$Script:SaveCheckPointSuccessful = $false
$Script:SaveCheckPointError = $false

If(!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]"Administrator")){
	Start-Process Powershell.exe -ArgumentList "-STA -noprofile -file `"$ScriptName`"" -Verb RunAs
	Exit
}

Try { Import-Module Hyper-V }
Catch { 
    $Popup.Popup("This must be run on the Hyper-V machine with the Hyper-V Cmdlets!",0,"Error",16)
    Exit
}

Function Log {
    Param (
		[Parameter(Mandatory=$false)]
		$Message,
 
		[Parameter(Mandatory=$false)]
		$ErrorMessage,
 
		[Parameter(Mandatory=$false)]
		$Component,
 
		[Parameter(Mandatory=$false)]
		[int]$Type,
		
		[Parameter(Mandatory=$true)]
		$LogFile
	)
<#
Type: 1 = Normal, 2 = Warning (yellow), 3 = Error (red)
#>
	$Time = Get-Date -Format "HH:mm:ss.ffffff"
	$Date = Get-Date -Format "MM-dd-yyyy"
 
	if ($ErrorMessage -ne $null) {$Type = 3}
	if ($Component -eq $null) {$Component = " "}
	if ($Type -eq $null) {$Type = 1}
 
	$LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"
	$LogMessage | Out-File -Append -Encoding UTF8 -FilePath $LogFile
    $Message = "$Message - $ErrorMessage"
    Write-Host $Message
}

Function Translate-EvaluationState {
    Param ($EvaluationState)

    $strEvaluationState = ""
    Switch ($EvaluationState) {
        0 { $strEvaluationState = "No state information is available." }
        1 { $strEvaluationState = "Application is enforced to desired/resolved state." }
        2 { $strEvaluationState = "Application is not required on the client." }
        3 { $strEvaluationState = "Application is available for enforcement (install or uninstall based on resolved state). Content may/may not have been downloaded." }
        4 { $strEvaluationState = "Application last failed to enforce (install/uninstall)." }
        5 { $strEvaluationState = "Application is currently waiting for content download to complete." }
        6 { $strEvaluationState = "Application is currently waiting for content download to complete." }
        7 { $strEvaluationState = "Application is currently waiting for its dependencies to download." }
        8 { $strEvaluationState = "Application is currently waiting for a service (maintenance) window." }
        9 { $strEvaluationState = "Application is currently waiting for a previously pending reboot." }
        10 { $strEvaluationState = "Application is currently waiting for serialized enforcement." }
        11 { $strEvaluationState = "Application is currently enforcing dependencies." }
        12 { $strEvaluationState = "Application is currently enforcing." }
        13 { $strEvaluationState = "Application install/uninstall enforced and soft reboot is pending." }
        14 { $strEvaluationState = "Application installed/uninstalled and hard reboot is pending." }
        15 { $strEvaluationState = "Update is available but pending installation." }
        16 { $strEvaluationState = "Application failed to evaluate." }
        17 { $strEvaluationState = "Application is currently waiting for an active user session to enforce." }
        18 { $strEvaluationState = "Application is currently waiting for all users to logoff." }
        19 { $strEvaluationState = "Application is currently waiting for a user logon." }
        20 { $strEvaluationState = "Application in progress, waiting for retry." }
        21 { $strEvaluationState = "Application is waiting for presentation mode to be switched off." }
        22 { $strEvaluationState = "Application is pre-downloading content (downloading outside of install job)." }
        23 { $strEvaluationState = "Application is pre-downloading dependent content (downloading outside of install job)." }
        24 { $strEvaluationState = "Application download failed (downloading during install job)." }
        25 { $strEvaluationState = "Application pre-downloading failed (downloading outside of install job)." }
        26 { $strEvaluationState = "Download success (downloading during install job)." }
        27 { $strEvaluationState = "Post-enforce evaluation." }
        28 { $strEvaluationState = "Waiting for network connectivity." }
    }
    return $strEvaluationState
}

Function LoadApplications {
    Param ($CompName)
    $ApplicationArray = New-Object System.Collections.ArrayList
	Try {
		Get-WmiObject -Query "select * from CCM_Application" -Namespace root\ccm\clientsdk -ComputerName $CompName | ForEach-Object {
			$Results = Select-Object -InputObject "" Name, Installed, Required, LastEvaluated
			$Results.Name = $_.Name
			if ($Results.Name -ne $null) {$FoundApps = $true}
			$Results.Installed = $_.InstallState
			$ResolvedState = $_.ResolvedState
			If ($ResolvedState -eq "Available") {$Results.Required = "False"}
			else {$Results.Required = "True"}
			$LastEvalTime = $_.LastEvalTime
			if ($LastEvalTime -ne $null) {
				$EvalTime = $_.ConvertToDateTime($_.LastEvalTime)
				$Results.LastEvaluated = $EvalTime.ToShortDateString() + " " + $EvalTime.ToShortTimeString()
			}
			$ApplicationArray += $Results
		}
        Log -Message "Finished loading application list!" -LogFile $LogFile
	}
    Catch {Log -Message "Error loading applications -" -ErrorMessage $_.Exception.Message -LogFile $LogFile}

    return $ApplicationArray
}

Function LoadWPF {
    Param ($XAML)
	Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
	$XMLReader = (New-Object System.Xml.XmlNodeReader $XAML)
	$XAML = $XAML.OuterXML
	$SplitXaml = $XAML.Split("`n")
	$Script:Window = [Windows.Markup.XamlReader]::Load($XMLReader)
    foreach ($Line in $SplitXaml) {
        if ($Line.ToLower().Contains("x:name")) {
    		$SplitLine = $Line.Split("`"")
    		$Count = 0
    		foreach ($instance in $SplitLine) {
    			$Count++
    			if ($instance.ToLower().Contains("x:name")) {
    				$ControlName = $SplitLine[$Count]
                    $strExpression = "`$Script:" + "$ControlName = `$Window.FindName(`"$ControlName`")"
                    Invoke-Expression $strExpression
    			}
    		}
    	}
    }
}

[xml]$xaml = @'

<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Application Tester" ResizeMode="NoResize" WindowStartupLocation="CenterScreen" SizeToContent="WidthAndHeight" >
    <StackPanel Orientation="Vertical" Margin="5">
        <StackPanel Orientation="Horizontal" Margin="5">
            <Label Content="Computer Name:"/>
            <TextBox x:Name="ComputerName" Width="125" Margin="5,0,0,0" VerticalContentAlignment="Center" TextWrapping="NoWrap"/>
            <Label Content="VM Name:" Margin="5,0,0,0"/>
            <TextBox x:Name="VMName" Width="125" Margin="5,0,0,0" VerticalContentAlignment="Center" TextWrapping="NoWrap"/>
            <Button x:Name="LoadApplications" Width="100" Margin="5,0,0,0" Content="Load Apps"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal">
            <StackPanel Orientation="Vertical">
                <Label Content="Applications Advertised to Computer" HorizontalContentAlignment="Center"/>
                <DataGrid x:Name="ApplicationGrid" IsReadOnly="True" Width="450" Height="200">
                    <DataGrid.ContextMenu>
                        <ContextMenu>
                            <MenuItem x:Name="AddToList" Header="Add to list"/>
                        </ContextMenu>
                    </DataGrid.ContextMenu>
                </DataGrid>
            </StackPanel>
            <StackPanel Orientation="Vertical" Margin="5,0,0,0">
                <Label Content="Applications To Test" HorizontalAlignment="Center"/>
                <ListBox x:Name="ApplicationList" Height="200" Width="200">
                    <ListBox.ContextMenu>
                        <ContextMenu>
                            <MenuItem x:Name="RemoveFromList" Header="Remove from list"/>
                        </ContextMenu>
                    </ListBox.ContextMenu>
                </ListBox>
            </StackPanel>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="5">
            <Label Content="Save Logs"/>
            <ComboBox x:Name="ComboSaveLogs" Width="150" Margin="5,0,0,0" SelectedIndex="0">
                <ComboBoxItem Content="Never"/>
                <ComboBoxItem Content="Only when there is an error"/>
                <ComboBoxItem Content="Always"/>
            </ComboBox>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="5">
            <Label Content="Create Checkpoint"/>
            <ComboBox x:Name="ComboCheckPoint" Width="150" Margin="5,0,0,0" SelectedIndex="0">
                <ComboBoxItem Content="Never"/>
                <ComboBoxItem Content="Only when there is an error"/>
                <ComboBoxItem Content="Always"/>
            </ComboBox>
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Center" Margin="5">
            <Button x:Name="StartBtn" Content="Start" Width="75"/>
        </StackPanel>
    </StackPanel>
</Window>

'@

LoadWPF -XAML $XAML

$LoadApplications.Add_Click({
    $ApplicationGrid.ItemsSource = LoadApplications -CompName $ComputerName.Text
})

$RemoveFromList.Add_Click({
    $SelectedItems = $ApplicationList.SelectedItems
    $tempAppArray = $ApplicationList.Items
    $tempNewArray = @()
    Foreach ($instance in $tempAppArray) {
        $AddItem = $true
        foreach ($item in $SelectedItems) {
            if ($item -eq $instance) { $AddItem = $false }
        }
        if ($AddItem) { $tempNewArray += @($instance) }
    }
    $ApplicationList.ItemsSource = $tempNewArray
})

$AddToList.Add_Click({
    $SelectedItems = $ApplicationGrid.SelectedItems
    Foreach ($Item in $SelectedItems) {
        $AddToArray = $true
        Foreach ($instance in $ApplicationList.Items) {
            If ($instance -eq $Item.Name) { $AddToArray = $false }
        }
        If ($AddToArray) { $ApplicationList.Items.Add($Item.Name) }
    }
})

$StartBtn.Add_Click({
    $Message = "Do you want to install these apps on the VM?`nNote, if you click yes the UI will go away and you can track progress in the log file and console window`n"
    Foreach ($instance in $ApplicationList.Items) {
        $Script:AppList += @($instance)
        $Message = $Message + "`n" + $instance
    }
    $Answer = $Popup.Popup($Message,0,"Are you sure?",1)
    if ($Answer -eq 1) {
        $Script:strVMName = $VMName.Text
        Switch ($ComboSaveLogs.SelectedIndex) {
            0 { 
                $Script:SaveLogsSuccessful = $false
                $Script:SaveLogsError = $false
            }
            1 {  
                $Script:SaveLogsSuccessful = $false
                $Script:SaveLogsError = $true
            }
            2 {  
                $Script:SaveLogsSuccessful = $true
                $Script:SaveLogsError = $true
            }
        }
        Switch ($ComboCheckPoint.SelectedIndex) {
            0 { 
                $Script:SaveCheckPointSuccessful = $false
                $Script:SaveCheckPointError = $false
            }
            1 {  
                $Script:SaveCheckPointSuccessful = $false
                $Script:SaveCheckPointError = $true
            }
            2 {  
                $Script:SaveCheckPointSuccessful = $true
                $Script:SaveCheckPointError = $true
            }
        }
        $Script:strCompName = $ComputerName.Text
        $ContinueScript = $true
        try {
            $Script:VMObject = Get-VM -Name $strVMName
            if ($Script:VMObject -eq $null) {
                $ContinueScript = $false
            }
        }
        catch {
            $Popup.Popup("Could not find $strVMName",0,"Error!",16)
        }
        If ($ContinueScript) {
            $Window.Close() | Out-Null
        }
    }
})

$Window.ShowDialog() | Out-Null

Try { Checkpoint-VM -Name $Script:strVMName -SnapshotName "TestAppScript-Original" }
catch { 
    Log -Message "Could not create VM checkpoint" -ErrorMessage $_.Exception.Message -LogFile $LogFile
    Exit
}

Log -Message "Successfully created checkpoint TestAppScript-Original" -LogFile $LogFile
Log -Message "Starting install of applications" -LogFile $LogFile

foreach ($instance in $Script:AppList) {
    
    Log "Installing $instance" -LogFile $LogFile
    $MakeCheckPoint = $false
    $CopyLogFiles = $false
    $StopLoop = $false
    $Count = 0
    Do {
        $Script:VMObject = Get-VM -Name $Script:strVMName
        If ($VMObject.State -eq "Running") {
            Start-Sleep 10
            $StopLoop = $true
        }
        else { 
            $VMObject | Start-VM
            Start-Sleep 10
            $Count++
            If ($Count -gt 10) { 
                Log -Message "Cannot restart VM!" -ErrorMessage "Error" -LogFile $LogFile
                Exit
            }
        }
    } while ($StopLoop -ne $true)
	
    $AppErrorCodes = 8,16,17,18,19,21,24,25, 4
    $AppInProgressCodes = 0,3,5,6,7,10,11,12,15,20,22,23,26,27,28
    $AppSuccessfulCodes = 1,2
    $AppRestartCodes = 13,14,9

    Try {
		$WMIPath = "\\" + $Script:strCompName + "\root\ccm\clientsdk:CCM_Application"
		$WMIClass = [WMIClass] $WMIPath
        $ApplicationID = ""
        $ApplicationRevision = ""
        $IsMachineTarget = ""
		Get-WmiObject -ComputerName $Script:strCompName -Query "select * from CCM_Application" -Namespace root\ccm\ClientSDK | ForEach-Object {
			if ($_.Name -eq $instance) {
				$ApplicationRevision = $_.Revision
				$IsMachineTarget = $_.IsMachineTarget
				$EnforcePreference = $_.EnforcePreference
				$ApplicationID = $_.ID
			}
		}
		$WMIClass.Install($ApplicationID, $ApplicationRevision, $IsMachineTarget, "", "1", $false) | Out-null
        
        $EndLoop = $false
        do {
            $InstallState = "NotInstalled"
		    Get-WmiObject -ComputerName $Script:strCompName -Query "select * from CCM_Application" -Namespace root\ccm\ClientSDK | ForEach-Object {
			    if ($_.Name -eq $instance) {
				    $InstallState = $_.InstallState
                    $EvaluationState = $_.EvaluationState
			    }
		    }
            $TranslatedEvaluationState = Translate-EvaluationState -EvaluationState $EvaluationState
            If ($InstallState -eq "Installed") {
                $CopyLogFiles = $Script:SaveLogsSuccessful
                $MakeCheckPoint = $Script:SaveCheckPointSuccessful
                $EndLoop = $true
                Log "$instance - Successfully installed. Evaluation State: $TranslatedEvaluationState" -LogFile $LogFile
            }
            else {
                
                If ($AppErrorCodes -contains $EvaluationState) {
                    Log "$instance - Error installing application" -ErrorMessage $TranslatedEvaluationState -LogFile $LogFile
                    $MakeCheckPoint = $Script:SaveCheckPointError
                    $CopyLogFiles = $Script:SaveLogsError
                    $EndLoop = $true
                }
                elseif ($AppInProgressCodes -contains $EvaluationState) {
                    Log "$instance - Still in progress. Sleeping 30 seconds. Evaluation State: $TranslatedEvaluationState" -LogFile $LogFile
                    Start-Sleep 30
                }
                elseif ($AppRestartCodes -contains $EvaluationState) {
                    Log "$instance - Requires restart to complete install. Will restart computer and wait 180 seconds now... Evaluation State: $TranslatedEvaluationState" -LogFile $LogFile
                    try {
                        shutdown /r /t 0 /m "\\$Script:strCompName"
                        Start-Sleep 180
                    }
                    catch { Log "$instance - Error restarting computer!" -ErrorMessage $_.Exception.Message -LogFile $LogFile }
                    try {
                        Log "$instance - Starting CCMExec service so the Application install can restart" -LogFile $LogFile
                        (Get-WmiObject -ComputerName $Script:strCompName -Query "Select * From Win32_Service where Name like 'ccmexec'").StartService()
                        Start-Sleep 60
                    }
                    catch { Log "$instance - Error starting ccmexec on remote computer" -ErrorMessage $_.Exception.Message -LogFile $LogFile }
                    try {
                        Log "$instance - Triggering application install again to re-check detection method..." -LogFile $LogFile
                        $WMIPath = "\\" + $Script:strCompName + "\root\ccm\clientsdk:CCM_Application"
		                $WMIClass = [WMIClass] $WMIPath
                        $ApplicationID = ""
                        $ApplicationRevision = ""
                        $IsMachineTarget = ""
		                Get-WmiObject -ComputerName $Script:strCompName -Query "select * from CCM_Application" -Namespace root\ccm\ClientSDK | ForEach-Object {
			                if ($_.Name -eq $instance) {
				                $ApplicationRevision = $_.Revision
				                $IsMachineTarget = $_.IsMachineTarget
				                $EnforcePreference = $_.EnforcePreference
				                $ApplicationID = $_.ID
			                }
		                }
		                $WMIClass.Install($ApplicationID, $ApplicationRevision, $IsMachineTarget, "", "1", $false) | Out-null
                        Start-Sleep 20
                    }
                    catch { 
                        Log "$instance - Error triggering application install" -ErrorMessage $_.Exception.Message -LogFile $LogFile
                        $MakeCheckPoint = $Script:SaveCheckPointError
                        $CopyLogFiles = $Script:SaveLogsError
                        $EndLoop = $true
                    }
                }
                elseif ($AppSuccessfulCodes -contains $EvaluationState) {
                    $MakeCheckPoint = $Script:SaveCheckPointSuccessful
                    $CopyLogFiles = $Script:SaveLogsSuccessful
                    $EndLoop = $true
                    Log "$instance - Successfully installed. Evaluation State: $TranslatedEvaluationState" -LogFile $LogFile
                }
            }
        } while ($EndLoop -ne $true)
	}
	Catch { 
        Log -Message "Error installing $AppName -" -ErrorMessage $_.Exception.Message -LogFile $LogFile
        $MakeCheckPoint = $Script:SaveCheckPointError
        $CopyLogFiles = $Script:SaveLogsError
    }
    
    If ($MakeCheckPoint) {
        Try {
            Log -Message "$instance - Creating checkpoint" -LogFile $LogFile
            $CheckpointName = "App-" + $instance
            Checkpoint-VM -Name $Script:strVMName -SnapshotName $CheckpointName
            Log -Message "$instance - Created checkpoint!" -LogFile $LogFile
        }
        catch { Log -Message "$instance - Error creating checkpoint" -ErrorMessage $_.Exception.Message -LogFile $LogFile }
    }
    
    if ($CopyLogFiles) {
        Try {
            $AppLogDirectory = "$Directory\Logs\$instance"
            Log -Message "Saving $instance log files to $AppLogDirectory" -LogFile $LogFile
            $CopyPath = "\\" + $Script:strCompName + "\c$\windows\ccm\logs"
            Copy-Item $CopyPath $AppLogDirectory -Recurse -Force
        }
        catch { Log -Message "$instance - Error copying log files!" -ErrorMessage $_.Exception.Message -LogFile $LogFile }
    }

    Try {
        Log -Message "Reverting VM to previous checkpoint" -LogFile $LogFile
        Restore-VMSnapshot -VMName $Script:strVMName -Name "TestAppScript-Original" -Confirm:$false
        Start-Sleep 10
    }
    catch {
        Log -Message "Error reverting to previous checkpoint!" -ErrorMessage $_.Exception.Message -LogFile $LogFile
        exit
    }
}

Log -Message "Finished!" -Type 2 -LogFile $LogFile


