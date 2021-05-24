Add-Type -AssemblyName PresentationFramework

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
$Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
$Window = [Windows.Markup.XamlReader]::Load($Reader)

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
        $Guid = (New-Guid).Guid.Split('-')[4]

        # Populate $Date variable to suit either UK or US format
        If ((Get-Culture).LCID -eq "1033") {

            $Date = (Get-Date).tostring("MM_dd_yy")

        }
        Else {

            $Date = (Get-Date).tostring("dd_MM_yy")

        }

        # Build the unique name for the folder re-naming
        $NewFolderName = "Teams.Old_$($Date)_$($Guid)"

        # Capture logged on user's username
        $LoggedOnUser = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object UserName).UserName.Split('\')[1]

        # Construct the logged on user's equivelant to $Home variable
        $LoggedOnUserHome = "C:\Users\$LoggedOnUser"

        # Get MS Teams process. Only using 'SilentlyContinue' as we test this below
        $TeamsProcess = Get-Process Teams -ErrorAction SilentlyContinue

        # Get Outlook process. Only using 'SilentlyContinue' as we test this below
        $OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue

        # Get Teams Folder
        $TeamsFolder = Get-Item "$LoggedOnUserHome\AppData\Roaming\Microsoft\Teams" -ErrorAction SilentlyContinue

    }

    PROCESS {

        If ($TeamsProcess) {
    
            # If 'Teams' process is running, stop it else do nothing
            $TeamsProcess | Stop-Process -Force
            

        }
        Else {

            # Do something - removed Write-Host

        }

        If ($OutlookProcess) {
    
            # If 'Outlook' process is running, stop it else do nothing
            $OutlookProcess | Stop-Process -Force
            

        }
        Else {

            # Do something - removed Write-Host

        }

        # Give the processes a little time to completely stop to avoid error
        Start-Sleep -Seconds 10

        If ($TeamsFolder) {
    
            # If 'Teams' folder exists in %APPDATA%\Microsoft\Teams, rename it
            Rename-Item "$LoggedOnUserHome\AppData\Roaming\Microsoft\Teams" "$LoggedOnUserHome\AppData\Roaming\Microsoft\$NewFolderName"
        }
        Else {

            # If 'Teams' folder does not exist notify user then break
            # Do something - removed Write-Host
            break
        }

        # Give the folder a couple of seconds to fully rename to avoide error
        Start-Sleep -Seconds 2

        # Restart MS Teams desktop client
        If ($TeamsProcess) { 

            Start-Process -File "$LoggedOnUserHome\AppData\Local\Microsoft\Teams\Update.exe" -ArgumentList '--processStart "Teams.exe"'
        }
        Else {

            # Do something - removed Write-Host

        }
    
        # Restart Outlook
        If ($OutlookProcess) {

            $OutlookExe = Get-ChildItem -Path 'C:\Program Files\Microsoft Office\root\Office16' -Filter Outlook.exe -Recurse -ErrorAction SilentlyContinue -Force | Where-Object { $_.Directory -notlike "*Updates*" } | Select-Object Name, Directory

            If (!$OutlookExe) {

                $OutlookExe = Get-ChildItem -Path 'C:\Program Files (x86)\Microsoft Office\root\Office16' -Filter Outlook.exe -Recurse -ErrorAction SilentlyContinue -Force | Where-Object { $_.Directory -notlike "*Updates*" } | Select-Object Name, Directory
            }
            
            Start-Process -File "$($OutlookExe.Directory)\$($OutlookExe.Name)"

        }
        Else {

            # Do something - removed Write-Host

        }

    }

    END {

        # Check for newly renamed folder
        $NewlyRenamedFolder = Get-Item "$LoggedOnUserHome\AppData\Roaming\Microsoft\$NewFolderName" -ErrorAction SilentlyContinue

        If ($NewlyRenamedFolder) { 

            # Do something - removed Write-Host

        }
        Else {

            # Do Something - removed Write-Host

        }

    }

}

# Use .FindName() to locate the WPF element prior to styling
$Button_ReProfile = $window.FindName("Button_ReProfile")

$Button_ReProfile.Content = "C L I C K   M E"
$Button_ReProfile.Margin = "40"
$Button_ReProfile.VerticalAlignment = "Bottom"
$Button_ReProfile.Width = "140"
$Button_ReProfile.Height = "140"
$Button_ReProfile.Background = "#7eabfd"
$Button_ReProfile.BorderBrush = "White"
$Button_ReProfile.Foreground = "White"
$Button_ReProfile.Padding = "8"
$Button_ReProfile.ToolTip = "Click to Re-profile Teams"

$Button_ReProfile.Add_MouseEnter( {
    
    $Window.Cursor = [System.Windows.Input.Cursors]::Hand
    $Button_ReProfile.ForeGround = '#7eabfd'

})

$Button_ReProfile.Add_MouseLeave( {
    
    $Window.Cursor = [System.Windows.Input.Cursors]::Arrow
    $Button_ReProfile.ForeGround = '#ffffff'

})

$Button_ReProfile.Add_Click( {
    
    Invoke-TeamsReprofile
    $Window.Close()

})

$Window.ShowDialog() | Out-Null