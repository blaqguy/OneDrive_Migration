#Allow for multi threading here so GUI doesn't hang. We're going to be splitting up the different blocks into different threads
$Global:syncHash = [hashtable]::Synchronized(@{})
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)


# Load WPF assembly if necessary
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

$psCmd = [PowerShell]::Create().AddScript({
    [xml]$xaml = @"
    <Window x:Class="TakeUno.MainWindow"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:local="clr-namespace:TakeUno"
    mc:Ignorable="d"
    Title="OneDrive Migration Tool" Height="603.018" Width="742.053">
    <Grid>
        <Grid.Background>
            <LinearGradientBrush EndPoint="0.5,1" StartPoint="0.5,0">
                <GradientStop Color="Black" Offset="0"/>
                <GradientStop Color="#FFD1B0F1" Offset="1"/>
            </LinearGradientBrush>
        </Grid.Background>
        <RichTextBox x:Name="VerboseOutput" HorizontalAlignment="Left" Height="403" Margin="162,120,0,0" VerticalAlignment="Top" Width="445">
            <FlowDocument/>
        </RichTextBox>
        <Button x:Name="InitGPUpdate" Content="GPUpdate" HorizontalAlignment="Left" Margin="39,266,0,0" VerticalAlignment="Top" Width="75" Background="#FFF2716C"/>
        <Button x:Name="InitMigration" Content="Migrate Data" HorizontalAlignment="Left" Margin="39,355,0,0" VerticalAlignment="Top" Width="75" Background="#FF9B7DC7"/>
        <Button x:Name="Close" Content="Close" HorizontalAlignment="Left" Margin="649,533,0,0" VerticalAlignment="Top" Width="75"/>
        <ProgressBar x:Name="ProgressBar" IsIndeterminate="True" Visibility="Hidden" HorizontalAlignment="Left" Height="24" Margin="162,91,0,0" VerticalAlignment="Top" Width="445"/>
        <Label x:Name="PBStatus" Visibility="Hidden" Content="Migrating Your Data" HorizontalAlignment="Left" Margin="313,91,0,0" VerticalAlignment="Top" Height="24" Width="122"/>
    </Grid>
</Window>
"@ 

    # Remove XML attributes that break a couple things.
    #   Without this, you must manually remove the attributes
    #   after pasting from Visual Studio. If more attributes
    #   need to be removed automatically, add them below.
    $AttributesToRemove = @(
        'x:Class',
        'mc:Ignorable'
    )

    foreach ($Attrib in $AttributesToRemove) {
        if ( $xaml.Window.GetAttribute($Attrib) ) {
             $xaml.Window.RemoveAttribute($Attrib)
        }
    }
    
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    
    $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )

    [xml]$XAML = $xaml
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
        #Find all of the form types and add them as members to the synchash
        $syncHash.Add($_.Name,$syncHash.Window.FindName($_.Name) )

    }

    $Script:JobCleanup = [hashtable]::Synchronized(@{})
    $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

    #region Background runspace to clean up jobs
    $jobCleanup.Flag = $True
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()        
    $newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup)     
    $newRunspace.SessionStateProxy.SetVariable("jobs",$jobs) 
    $jobCleanup.PowerShell = [PowerShell]::Create().AddScript({
        #Routine to handle completed runspaces
        Do {    
            Foreach($runspace in $jobs) {            
                If ($runspace.Runspace.isCompleted) {
                    [void]$runspace.powershell.EndInvoke($runspace.Runspace)
                    $runspace.powershell.dispose()
                    $runspace.Runspace = $null
                    $runspace.powershell = $null               
                } 
            }
            #Clean out unused runspace jobs
            $temphash = $jobs.clone()
            $temphash | Where {
                $_.runspace -eq $Null
            } | ForEach {
                $jobs.remove($_)
            }        
            Start-Sleep -Seconds 1     
        } while ($jobCleanup.Flag)
    })
    $jobCleanup.PowerShell.Runspace = $newRunspace
    $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  
    #endregion Background runspace to clean up jobs

    $syncHash.InitMigration.Add_Click({
        #Start-Job -Name Sleeping -ScriptBlock {start-sleep 5}
        #while ((Get-Job Sleeping).State -eq 'Running'){
            $x+= "."
        #region Boe's Additions
        $newRunspace =[runspacefactory]::CreateRunspace()
        $newRunspace.ApartmentState = "STA"
        $newRunspace.ThreadOptions = "ReuseThread"          
        $newRunspace.Open()
        $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
        $PowerShell = [PowerShell]::Create().AddScript({

$NewRootPath = "$env:USERPROFILE\OneDrive - Aura"
$regUSF = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
$regKeys = @("My Pictures", "My Video", "My Music", "Personal", "Desktop")
$FolderNames = @("Pictures", "Videos", "Music", "Documents", "Desktop")

Function Move-KnownFolderPath {
    Param (
            [Parameter(Mandatory = $true)]
            [ValidateSet('Pictures', 'Videos', 'Music', 'Documents', 'Desktop')]
            [string]$KnownFolder,
            [Parameter(Mandatory = $true)]
            [string]$Path
    )

    # Known Folder GUIDs
    $KnownFolders = @{
        'Pictures' = @('33E28130-4E1E-4676-835A-98395C3BC3BB','0ddd015d-b06c-45d5-8c4c-f59713854639');
        'Videos' = @('18989B1D-99B5-455B-841C-AB7C74E4DDFC','35286a68-3c57-41a1-bbb1-0eae73d76c95'); 
        'Music' = @('4BD8D571-6D19-48D3-BE97-422220080E43','a0c69a99-21c8-4671-8703-7934162fcf1d');                
        'Documents' = @('FDD39AD0-238F-46AF-ADB4-6C85480369C7','f42ee2d3-909f-4907-8871-4c22fc0bf756');
        'Desktop' = @('B4BFCC3A-DB2C-424C-B029-7FE99A87C641');
    }

    $FolderGUID = ([System.Management.Automation.PSTypeName]'KnownFolders').Type
	#create the Type entry if it does not yet exist / relates to the Registry eventually
    If (-not $FolderGUID) {
        $KnownFolderGUID = @'
[DllImport("shell32.dll")]
public extern static int SHSetKnownFolderPath(ref Guid folderId, uint flags, IntPtr token, [MarshalAs(UnmanagedType.LPWStr)] string path);
'@
        $FolderGUID = Add-Type -MemberDefinition $KnownFolderGUID -Name 'KnownFolders' -Namespace 'SHSetKnownFolderPath' -PassThru
    }

    If(!(Test-Path $Path)){
		#in case folder does yet exist, we create it
		Try {
			New-Item -Path $Path -Type Directory -Force -ErrorAction SilentlyContinue
		} Catch {}
    }
	#set the new folder path
	ForEach ($TmpGUID in $KnownFolders[$KnownFolder]) {
		$tmp = $FolderGUID::SHSetKnownFolderPath([ref]$TmpGUID, 0, 0, $Path)	
	}	
	Attrib +r $Path	#needed to retain the icon 
}            
Function Update-Window {
        Param (
            $Control,
            $Property,
            $Value,
            [switch]$AppendContent
        )

        # This is kind of a hack, there may be a better way to do this
        If ($Property -eq "Close") {
            $syncHash.Window.Dispatcher.invoke([action]{$syncHash.Window.Close()},"Normal")
            Return
        }

        # This updates the control based on the parameters passed to the function
        $syncHash.$Control.Dispatcher.Invoke([action]{
            # This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps
            If ($PSBoundParameters['AppendContent']) {
                $syncHash.$Control.AppendText($Value)
            } Else {
                $syncHash.$Control.$Property = $Value
            }
        }, "Normal")
    }

0..($regKeys.Count-1) | ForEach-Object {
    $regkey,$foldername = $regKeys[$_],$FolderNames[$_]
    Update-Window -Control VerboseOutput -property text -value "Working on: $regkey" -AppendContent
    Write-Output "Working on: $regkey" >> $env:USERPROFILE\migrationlog.txt
    $SourcePath = (Get-ItemProperty -Path $regUSF -Name $regkey).$regkey
    If ($SourcePath.Length -gt 0)
    {
        Write-Output "..Checking source path: $SourcePath" >> $env:USERPROFILE\migrationlog.txt
        If ((Test-Path -Path $SourcePath)){
            Write-Output "....$SourcePath is accessible" >> $env:USERPROFILE\migrationlog.txt
            $CompSource = Get-Item -Path $SourcePath
            $CompareResult = $false
            If ((Test-Path -Path ($NewRootPath + "\" + $CompSource.Name))){
                $CompTarget = Get-Item -Path ($NewRootPath + "\" + $CompSource.Name)
                Write-Output "..Comparing Source Path: $CompSource" >> $env:USERPROFILE\migrationlog.txt
                Write-Output "....with Target Path: $CompTarget" >> $env:USERPROFILE\migrationlog.txt
                If ($CompSource.FullName.ToLower() -eq $CompTarget.FullName.ToLower()){
                    $CompareResult = $true
                } Else {
                    $CompareResult = $false
                }
            } Else {
                Write-Output "..$CompTarget does not exist" >> $env:USERPROFILE\migrationlog.txt
            }
            If ($CompareResult -eq $true){
                Write-Output "..Source and Target path are identical: $CompTarget" >> $env:USERPROFILE\migrationlog.txt
                Write-Output "....Not applying any changes!" >> $env:USERPROFILE\migrationlog.txt
            } Else {
                $SourceFolder = Get-Item -Path $SourcePath
                Write-Output "..Accessing source folder $SourceFolder" >> $env:USERPROFILE\migrationlog.txt
                If (!(Test-Path -Path ($NewRootPath + '\' + $SourceFolder.Name))){
                    Move-KnownFolderPath -KnownFolder ("" + $foldername + "") -Path ($NewRootPath + '\' + $foldername)
                    $TargetPath = Get-Item -Path ("$NewRootPath\" + $foldername)
                    Write-Output "..Created new folder $TargetPath" >> $env:USERPROFILE\migrationlog.txt			
                    Write-Output "..Moving data from: $SourcePath" >> $env:USERPROFILE\migrationlog.txt
                    Write-Output "....to folder:      $TargetPath" >> $env:USERPROFILE\migrationlog.txt
                    Write-Output "....This can take several minutes..." >> $env:USERPROFILE\migrationlog.txt
                    Update-Window -Control Progressbar -property Visibility -value "Visible"
                    Update-Window -Control PBStatus -property Visibility -value "Visible"
                    #Ask Shawn if he wants a report on what was migrated. If so use RoboCopy
                    robocopy "$SourcePath " "$TargetPath " /e /LOG+:$env:USERPROFILE\datamigration.log
                    #Move-Item -Path ("" + $SourcePath + "\*") -Destination $TargetPath -Force -ErrorAction SilentlyContinue
                    Update-Window -Control Progressbar -property Visibility -value "Hidden"
                    Update-Window -Control PBStatus -property Visibility -value "Hidden"              
                    Write-Output "....Done moving data" >> $env:USERPROFILE\migrationlog.txt
                } Else {
                    Write-Output "..Warning - Target Path exists already: $CompTarget" >> $env:USERPROFILE\migrationlog.txt
                    Write-Output "....Files will not be moved, registry still will be adjusted to the target path!" >> $env:USERPROFILE\migrationlog.txt
                    Move-KnownFolderPath -KnownFolder ("" + $foldername + "") -Path ($NewRootPath + '\' + $foldername)
                    Write-Output "....Finished adjusting registry" >> $env:USERPROFILE\migrationlog.txt
                }
            }
        } Else {
            Write-Output "....SourceFolder Path invalid: $SourcePath" >> $env:USERPROFILE\migrationlog.txt
        }
    } 
    Update-Window -Control VerboseOutput -property text -value "`n ..Done processing $foldername `n" -AppendContent
    Write-Output "..Done processing $foldername" >> $env:USERPROFILE\migrationlog.txt
}
#Change the non GUID entries to point to OD - Aura. Sometimes these don't get updated.
Set-Itemproperty -path $regUSF -Name 'Personal' -value '%USERPROFILE%\OneDrive - Aura\Documents'
Set-Itemproperty -path $regUSF -Name 'My Video' -value '%USERPROFILE%\OneDrive - Aura\Videos'
Set-Itemproperty -path $regUSF -Name 'My Pictures' -value '%USERPROFILE%\OneDrive - Aura\Pictures'
Set-Itemproperty -path $regUSF -Name 'My Music' -value '%USERPROFILE%\OneDrive - Aura\Music'
Set-Itemproperty -path $regUSF -Name 'Desktop' -value '%USERPROFILE%\OneDrive - Aura\Desktop'
#Restart explorer process
Stop-Process -Name explorer
#Tell the user migration is complete
Add-Type -AssemblyName PresentationCore,PresentationFramework
[System.Windows.MessageBox]::Show("Migration Complete")    
})
$PowerShell.Runspace = $newRunspace
[void]$Jobs.Add((
    [pscustomobject]@{
        PowerShell = $PowerShell
        Runspace = $PowerShell.BeginInvoke()
    }
))
})

$syncHash.InitGPUpdate.Add_Click({
    #Start-Job -Name Sleeping -ScriptBlock {start-sleep 5}
    #while ((Get-Job Sleeping).State -eq 'Running'){
        $x+= "."
    #region Boe's Additions
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
    $PowerShell = [PowerShell]::Create().AddScript({

Function Update-Window {
        Param (
            $Control,
            $Property,
            $Value,
            [switch]$AppendContent
        )

        # This is kind of a hack, there may be a better way to do this
        If ($Property -eq "Close") {
            $syncHash.Window.Dispatcher.invoke([action]{$syncHash.Window.Close()},"Normal")
            Return
        }

        # This updates the control based on the parameters passed to the function
        $syncHash.$Control.Dispatcher.Invoke([action]{
            # This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps
            If ($PSBoundParameters['AppendContent']) {
                $syncHash.$Control.AppendText($Value)
            } Else {
                $syncHash.$Control.$Property = $Value
            }
        }, "Normal")
    }

Add-Type -AssemblyName PresentationCore,PresentationFramework
Update-Window -Control VerboseOutput -property text -value "Updating Group Policy definitions.... please wait. `n" -AppendContent
[System.Windows.MessageBox]::Show("Updating Group Policy definitions. Your computer will automatically restart upon completion")
$output_gpupdate = gpupdate /force /boot

        })
        $PowerShell.Runspace = $newRunspace
        [void]$Jobs.Add((
            [pscustomobject]@{
                PowerShell = $PowerShell
                Runspace = $PowerShell.BeginInvoke()
            }
        ))
    })

    $syncHash.Close.Add_Click({$syncHash.Window.Close()})
    #region Window Close 
    $syncHash.Window.Add_Closed({
        Write-Verbose 'Halt runspace cleanup job processing'
        $jobCleanup.Flag = $False

        #Stop all runspaces
        $jobCleanup.PowerShell.Dispose()      
    })

    #endregion Window Close 
    #endregion Boe's Additions

    #$x.Host.Runspace.Events.GenerateEvent( "TestClicked", $x.test, $null, "test event")

    #$syncHash.Window.Activate()
    $syncHash.Window.ShowDialog() | Out-Null
    $syncHash.Error = $Error
})
$psCmd.Runspace = $newRunspace
$data = $psCmd.BeginInvoke()
