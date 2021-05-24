Add-Type -AssemblyName PresentationFramework

<# Set up the Runspace #>
$Runspace = [runspacefactory]::CreateRunspace()
$Runspace.ApartmentState = "STA"
$Runspace.ThreadOptions = "ReuseThread"
$Runspace.Name = "Teams Helper"
$Runspace.Open()

$Code = {
[xml]$Xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Name="Window"
    Title="Teams Helper"
    WindowStartupLocation="CenterScreen"
    Width="300"
    Height="300"
    ShowInTaskbar="True">
    <Grid Name="Grid">

        <Label Name="Label_Heading" Content="Re-Profile Microsoft Teams" Margin="0,25,0,0" FontSize="14" HorizontalContentAlignment="Center"/>

        <Button Name="Button_ReProfile" >
            <Button.Resources>
                <Style TargetType="Border">
                    <Setter Property="CornerRadius" Value="35"/>
                </Style>
            </Button.Resources>
        </Button>

    </Grid>
</Window>
"@

    <# Set up the Synchronized Hash Table as this will allow access to the shared data whilst threading #>
    $Script:SyncHash    = [hashtable]::Synchronized(@{})
    $SyncHash.Host      = $host
    $Reader             = (New-Object System.Xml.XmlNodeReader $Xaml)
    $SyncHash.Window    = [Windows.Markup.XamlReader]::Load( $Reader )

    Function Invoke-TeamsReprofile {

        <#
        .SYNOPSIS
        Invoke-TeamsReprofile is a simple function that will re-profile the MS Teams
        user profile for the Teams desktop client on a windows 10 pc.

        .DESCRIPTION
        This Function will test for MS Teams & MS Outlook running then will close them
        both in advance of re-naming the 'Teams' folder in %APPDATA%. Once the re-naming,
        or re-profiling has occurred it will restart MS Teams & MS Outlook if they were
        previously running

        .EXAMPLE
        PS C:\> Invoke-TeamsReprofile

        .NOTES

        Author:  AlanPs1
        Website: http://AlanPs1.io
        Twitter: @AlanO365

        #>

        [OutputType()]
        [CmdletBinding()]
        Param()

        BEGIN {

            # Create a small guid type value to help avoid folder re-naming clashes
            $SyncHash.Guid = (New-Guid).Guid.Split('-')[4]

            # Populate $Date variable to suit either UK or US format
            If ((Get-Culture).LCID -eq "1033") {

                $SyncHash.Date = (Get-Date).tostring("MM_dd_yy")

            }
            Else {

                $SyncHash.Date = (Get-Date).tostring("dd_MM_yy")

            }

            # Build the unique name for the folder re-naming
            $SyncHash.NewFolderName = "Teams.Old_$($SyncHash.Date)_$($SyncHash.Guid)"

            # Capture logged on user's username
            $SyncHash.LoggedOnUser = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object UserName).UserName.Split('\')[1]

            # Construct the logged on user's equivelant to $Home variable
            $SyncHash.LoggedOnUserHome = "C:\Users\$($SyncHash.LoggedOnUser)"

            # Get MS Teams process. Only using 'SilentlyContinue' as we test this below
            $SyncHash.TeamsProcess = Get-Process Teams -ErrorAction SilentlyContinue

            # Get Outlook process. Only using 'SilentlyContinue' as we test this below
            $SyncHash.OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue

            # Get Teams Folder
            $SyncHash.TeamsFolder = Get-Item "$($SyncHash.LoggedOnUserHome)\AppData\Roaming\Microsoft\Teams" -ErrorAction SilentlyContinue

        }

        PROCESS {

            If ($SyncHash.TeamsProcess) {
        
                # If 'Teams' process is running, stop it else do nothing
                $SyncHash.TeamsProcess | Stop-Process -Force
                

            }
            Else {

                # Do something - removed Write-Host

            }

            If ($SyncHash.OutlookProcess) {
        
                # If 'Outlook' process is running, stop it else do nothing
                $SyncHash.OutlookProcess | Stop-Process -Force
                

            }
            Else {

                # Do something - removed Write-Host

            }

            # Give the processes a little time to completely stop to avoid error
            Start-Sleep -Seconds 10

            If ($SyncHash.TeamsFolder) {
        
                # If 'Teams' folder exists in %APPDATA%\Microsoft\Teams, rename it
                Rename-Item "$($SyncHash.LoggedOnUserHome)\AppData\Roaming\Microsoft\Teams" "$($SyncHash.LoggedOnUserHome)\AppData\Roaming\Microsoft\$($SyncHash.NewFolderName)"
            }
            Else {

                # If 'Teams' folder does not exist notify user then break
                # Do something - removed Write-Host
                break
            }

            # Give the folder a couple of seconds to fully rename to avoide error
            Start-Sleep -Seconds 2

            # Restart MS Teams desktop client
            If ($SyncHash.TeamsProcess) { 

                Start-Process -File "$($SyncHash.LoggedOnUserHome)\AppData\Local\Microsoft\Teams\Update.exe" -ArgumentList '--processStart "Teams.exe"'
            }
            Else {

                # Do something - removed Write-Host

            }
        
            # Restart Outlook
            If ($SyncHash.OutlookProcess) {

                $SyncHash.OutlookExe = Get-ChildItem -Path 'C:\Program Files\Microsoft Office\root\Office16' -Filter Outlook.exe -Recurse -ErrorAction SilentlyContinue -Force | Where-Object { $_.Directory -notlike "*Updates*" } | Select-Object Name, Directory

                If (!$SyncHash.OutlookExe) {

                    $SyncHash.OutlookExe = Get-ChildItem -Path 'C:\Program Files (x86)\Microsoft Office\root\Office16' -Filter Outlook.exe -Recurse -ErrorAction SilentlyContinue -Force | Where-Object { $_.Directory -notlike "*Updates*" } | Select-Object Name, Directory
                }
                
                Start-Process -File "$($SyncHash.OutlookExe.Directory)\$($SyncHash.OutlookExe.Name)"

            }
            Else {

                # Do something - removed Write-Host

            }

        }

        END {

            # Check for newly renamed folder
            $SyncHash.NewlyRenamedFolder = Get-Item "$($SyncHash.LoggedOnUserHome)\AppData\Roaming\Microsoft\$($SyncHash.NewFolderName)" -ErrorAction SilentlyContinue

            If ($SyncHash.NewlyRenamedFolder) { 

                # Do something - removed Write-Host

            }
            Else {

                # Do Something - removed Write-Host

            }

        }

    }

    # Use .FindName() to locate the WPF element prior to styling
    $SyncHash.Button_ReProfile = $SyncHash.Window.FindName("Button_ReProfile")

    $SyncHash.Button_ReProfile.Content = "C L I C K   M E"
    $SyncHash.Button_ReProfile.Margin = "40"
    $SyncHash.Button_ReProfile.VerticalAlignment = "Bottom"
    $SyncHash.Button_ReProfile.Width = "140"
    $SyncHash.Button_ReProfile.Height = "140"
    $SyncHash.Button_ReProfile.Background = "#7eabfd"
    $SyncHash.Button_ReProfile.BorderBrush = "White"
    $SyncHash.Button_ReProfile.Foreground = "White"
    $SyncHash.Button_ReProfile.Padding = "8"
    $SyncHash.Button_ReProfile.ToolTip = "Click to Re-profile Teams"

    $SyncHash.Button_ReProfile.Add_MouseEnter( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Hand
        $SyncHash.Button_ReProfile.ForeGround = '#7eabfd'

    })

    $SyncHash.Button_ReProfile.Add_MouseLeave( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $SyncHash.Button_ReProfile.ForeGround = '#ffffff'

    })

    $SyncHash.Button_ReProfile.Add_Click( {
        
        Invoke-TeamsReprofile
        $SyncHash.Window.Close()

    })

    $SyncHash.Window.ShowDialog() | Out-Null

}

$PSinstance            = [powershell]::Create().AddScript($Code)
$PSinstance.Runspace   = $Runspace
$Job                   = $PSinstance.BeginInvoke()