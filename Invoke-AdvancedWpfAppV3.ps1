Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, system.windows.forms

<# Set up the Runspace #>
$Runspace = [runspacefactory]::CreateRunspace()
$Runspace.ApartmentState = "STA"
$Runspace.ThreadOptions = "ReuseThread"
$Runspace.Name = "V3"
$Runspace.Open()
$Runspace.SessionStateProxy.SetVariable("SyncHash", $Script:SyncHash)

$Code = {
[xml]$Xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    Name="Window"
    Title="MS Teams Assistant"
    WindowStartupLocation="CenterScreen"
    Width="360"
    Height="300"
    ShowInTaskbar="True">
    <Grid Name="Grid_Main">

        <Grid.RowDefinitions>
                <RowDefinition/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Name="HomePanel1" Orientation="Horizontal" Grid.Row="0">

            <Grid Name="Grid_Inner1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <Grid Grid.Column="0" Width="180">
                    <Label Name="Label_TeamsProcessConstant"/>
                </Grid>

                <Grid Grid.Column="1">
                    <Label Name="Label_OutlookProcessConstant"/>
                </Grid>
            </Grid>

        </StackPanel>

        <StackPanel Name="HomePanel2" Orientation="Horizontal" Grid.Row="1">

            <Grid Name="Grid_Inner2">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <Grid Grid.Column="0" Width="180">
                    <Button Name="Button_TeamsStart">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="25"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                </Grid>

                <Grid Grid.Column="1">
                    <Button Name="Button_OutlookStart">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="25"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                </Grid>
            </Grid>

        </StackPanel>

        <StackPanel Name="HomePanel3" Orientation="Horizontal" Grid.Row="2">

            <Grid Name="Grid_Inner3">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <Grid Grid.Column="0" Width="180">
                    <Button Name="Button_TeamsKill">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="25"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                </Grid>

                <Grid Grid.Column="1">
                    <Button Name="Button_OutlookKill">
                        <Button.Resources>
                            <Style TargetType="Border">
                                <Setter Property="CornerRadius" Value="25"/>
                            </Style>
                        </Button.Resources>
                    </Button>
                </Grid>
            </Grid>

        </StackPanel>

        <StackPanel Name="HomePanel4" Orientation="Horizontal" Grid.Row="3">

            <Button Name="Button_TeamsReprofile">
                <Button.Resources>
                    <Style TargetType="Border">
                        <Setter Property="CornerRadius" Value="25"/>
                    </Style>
                </Button.Resources>
            </Button>

        </StackPanel>

    </Grid>
</Window>
"@

    <# Set up the Synchronized Hash Table as this will allow access to the shared data whilst threading #>
    $Script:SyncHash    = [hashtable]::Synchronized(@{})
    $SyncHash.Host      = $Host
    $Reader             = (New-Object System.Xml.XmlNodeReader $Xaml)
    $SyncHash.Window    = [Windows.Markup.XamlReader]::Load( $Reader )

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
    $SyncHash.LoggedOnUser = 'Alan' #Test only

    # Construct the logged on user's equivelant to $Home variable
    $SyncHash.LoggedOnUserHome = "$($Home.Split($($SyncHash.LoggedOnUser))[0])$($SyncHash.LoggedOnUser)"

    # Get Teams Folder
    $SyncHash.TeamsFolder = Get-Item "$($SyncHash.LoggedOnUserHome)\AppData\Roaming\Microsoft\Teams" -ErrorAction SilentlyContinue

    Function Invoke-TeamsStatus {

        $SyncHash.TeamsTester = Get-Process Teams -ErrorAction SilentlyContinue

        If ($SyncHash.TeamsTester) { 
            $SyncHash.TeamsProcessConstant = $true
            $SyncHash.Label_TeamsProcessConstant.Foreground = "Green"
        } Else { 
            $SyncHash.TeamsProcessConstant = $false
            $SyncHash.Label_TeamsProcessConstant.Foreground = "Red"
        }

        $SyncHash.Window.Dispatcher.Invoke(
            [action] {
                $SyncHash.Label_TeamsProcessConstant.Content = "Teams Running: $($SyncHash.TeamsProcessConstant)"
            },
            "Normal"
        )

    }

    Function Invoke-OutlookStatus {

        $SyncHash.OutlookTester = Get-Process Outlook -ErrorAction SilentlyContinue

        If ($SyncHash.OutlookTester) { 
            $SyncHash.OutlookProcessConstant = $true 
            $SyncHash.Label_OutlookProcessConstant.Foreground = "Green"
        } Else { 
            $SyncHash.OutlookProcessConstant = $false
            $SyncHash.Label_OutlookProcessConstant.Foreground = "Red"
        }

        $SyncHash.Window.Dispatcher.Invoke(
            [action] {
                $SyncHash.Label_OutlookProcessConstant.Content = "Outlook Running: $($SyncHash.OutlookProcessConstant)"
            },
            "Normal"
        )

    }

    Function Start-TimerInstance {
        <# Create a stopwatch and also a timer object #>
        $SyncHash.Stopwatch = New-Object System.Diagnostics.Stopwatch
        $SyncHash.Timer = New-Object System.Windows.Forms.Timer
        $SyncHash.Timer.Enabled = $true
        $SyncHash.Timer.Interval = 60

        $SyncHash.Stopwatch.Start()
        $SyncHash.Timer.Start()
    }

    Start-TimerInstance

    $SyncHash.Timer.Add_Tick( {

        Invoke-TeamsStatus
        Invoke-OutlookStatus

    })

    Function Invoke-TeamsStart {

        # Restart MS Teams desktop client
        Start-Process -File "$($SyncHash.LoggedOnUserHome)\AppData\Local\Microsoft\Teams\Update.exe" -ArgumentList '--processStart "Teams.exe"'

    }

    Function Invoke-TeamsKill {

        # Get MS Teams process. Only using 'SilentlyContinue' as we test this below
        $SyncHash.TeamsProcess = Get-Process Teams -ErrorAction SilentlyContinue

        # Get MS Teams process. Only using 'SilentlyContinue' as we test this below
        If ($SyncHash.TeamsProcess) {
        
            # If 'Teams' process is running, stop it else do nothing
            $SyncHash.TeamsProcess | Stop-Process -Force

        }
        Else {

            # Do something - removed Write-Host

        }

    }

    Function Invoke-OutlookStart {

        # Restart Outlook
        $SyncHash.OutlookExe = Get-ChildItem -Path 'C:\Program Files\Microsoft Office\root\Office16' -Filter Outlook.exe -Recurse -ErrorAction SilentlyContinue -Force | 
        Where-Object { $_.Directory -notlike "*Updates*" } | 
        Select-Object Name, Directory

        If (!$SyncHash.OutlookExe) {

            $SyncHash.OutlookExe = Get-ChildItem -Path 'C:\Program Files (x86)\Microsoft Office\root\Office16' -Filter Outlook.exe -Recurse -ErrorAction SilentlyContinue -Force | 
            Where-Object { $_.Directory -notlike "*Updates*" } | 
            Select-Object Name, Directory
        }

        Start-Process -File "$($SyncHash.OutlookExe.Directory)\$($SyncHash.OutlookExe.Name)"

    }

    Function Invoke-OutlookKill {

        # Get Outlook process. Only using 'SilentlyContinue' as we test this below
        $SyncHash.OutlookProcess = Get-Process Outlook -ErrorAction SilentlyContinue

        # Get Outlook process. Only using 'SilentlyContinue' as we test this below
        If ($SyncHash.OutlookProcess) {

            # If 'Outlook' process is running, stop it else do nothing
            $SyncHash.OutlookProcess | Stop-Process -Force             

        }
        Else {

            # Do something - removed Write-Host

        }

    }

    Function Invoke-TeamsReProfile {

        [OutputType()]
        [CmdletBinding()]
        Param (

        )

        PROCESS {

            # Get Teams Folder
            If ($SyncHash.TeamsFolder) {

                # If 'Teams' folder exists in %APPDATA%\Microsoft\Teams, rename it
                Rename-Item "$($SyncHash.LoggedOnUserHome)\AppData\Roaming\Microsoft\Teams" "$($SyncHash.LoggedOnUserHome)\AppData\Roaming\Microsoft\$($SyncHash.NewFolderName)"
            }
            Else {

                # If 'Teams' folder does not exist notify user then break
                # Do something - removed Write-Host
                break
            }

        }

    }

    # Label Customisation - Label_TeamsProcessConstant
    $SyncHash.Label_TeamsProcessConstant = $SyncHash.Window.FindName("Label_TeamsProcessConstant")
    $SyncHash.Label_TeamsProcessConstant.Foreground = "Black"
    $SyncHash.Label_TeamsProcessConstant.FontSize = "14" 
    $SyncHash.Label_TeamsProcessConstant.HorizontalContentAlignment = "Center"

    # Label Customisation - Label_OutlookProcessConstant
    $SyncHash.Label_OutlookProcessConstant = $SyncHash.Window.FindName("Label_OutlookProcessConstant")
    $SyncHash.Label_OutlookProcessConstant.Foreground = "Black"
    $SyncHash.Label_OutlookProcessConstant.FontSize = "14" 
    $SyncHash.Label_OutlookProcessConstant.HorizontalContentAlignment = "Center"

    # Button Customisation - Button_TeamsStart
    $SyncHash.Button_TeamsStart = $SyncHash.Window.FindName("Button_TeamsStart")
    $SyncHash.Button_TeamsStart.Content = "START"
    $SyncHash.Button_TeamsStart.Margin = "0 0 0 20"
    $SyncHash.Button_TeamsStart.Width = "100"
    $SyncHash.Button_TeamsStart.Height = "50"
    $SyncHash.Button_TeamsStart.Background = "#7eabfd"
    $SyncHash.Button_TeamsStart.BorderBrush = "White"
    $SyncHash.Button_TeamsStart.Foreground = "White"
    $SyncHash.Button_TeamsStart.Padding = "4"
    $SyncHash.Button_TeamsStart.ToolTip = "Kill Teams"

    $SyncHash.Button_TeamsStart.Add_MouseEnter( {

        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Hand
        $SyncHash.Button_TeamsStart.ForeGround = '#7eabfd'

    })

    $SyncHash.Button_TeamsStart.Add_MouseLeave( {

        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $SyncHash.Button_TeamsStart.ForeGround = '#ffffff'

    })

    $SyncHash.Button_TeamsStart.Add_Click( {

        Invoke-TeamsStart

    })

    # Button Customisation - Button_TeamsKill
    $SyncHash.Button_TeamsKill = $SyncHash.Window.FindName("Button_TeamsKill")
    $SyncHash.Button_TeamsKill.Content = "STOP"
    $SyncHash.Button_TeamsKill.Margin = "0 0 0 20"
    $SyncHash.Button_TeamsKill.Width = "100"
    $SyncHash.Button_TeamsKill.Height = "50"
    $SyncHash.Button_TeamsKill.Background = "#7eabfd"
    $SyncHash.Button_TeamsKill.BorderBrush = "White"
    $SyncHash.Button_TeamsKill.Foreground = "White"
    $SyncHash.Button_TeamsKill.Padding = "4"
    $SyncHash.Button_TeamsKill.ToolTip = "Kill Teams"

    $SyncHash.Button_TeamsKill.Add_MouseEnter( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Hand
        $SyncHash.Button_TeamsKill.ForeGround = '#7eabfd'

    })

    $SyncHash.Button_TeamsKill.Add_MouseLeave( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $SyncHash.Button_TeamsKill.ForeGround = '#ffffff'

    })

    $SyncHash.Button_TeamsKill.Add_Click( {
        
        Invoke-TeamsKill

    })

    # Button Customisation - Button_OutlookStart
    $SyncHash.Button_OutlookStart = $SyncHash.Window.FindName("Button_OutlookStart")
    $SyncHash.Button_OutlookStart.Content = "START"
    $SyncHash.Button_OutlookStart.Margin = "20 0 0 20"
    $SyncHash.Button_OutlookStart.Width = "100"
    $SyncHash.Button_OutlookStart.Height = "50"
    $SyncHash.Button_OutlookStart.Background = "#7eabfd"
    $SyncHash.Button_OutlookStart.BorderBrush = "White"
    $SyncHash.Button_OutlookStart.Foreground = "White"
    $SyncHash.Button_OutlookStart.Padding = "4"
    $SyncHash.Button_OutlookStart.ToolTip = "Kill Teams"

    $SyncHash.Button_OutlookStart.Add_MouseEnter( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Hand
        $SyncHash.Button_OutlookStart.ForeGround = '#7eabfd'

    })

    $SyncHash.Button_OutlookStart.Add_MouseLeave( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $SyncHash.Button_OutlookStart.ForeGround = '#ffffff'

    })

    $SyncHash.Button_OutlookStart.Add_Click( {
        
        Invoke-OutlookStart

    })

    # Button Customisation - Button_OutlookKill
    $SyncHash.Button_OutlookKill = $SyncHash.Window.FindName("Button_OutlookKill")
    $SyncHash.Button_OutlookKill.Content = "STOP"
    $SyncHash.Button_OutlookKill.Margin = "20 0 0 20"
    $SyncHash.Button_OutlookKill.Width = "100"
    $SyncHash.Button_OutlookKill.Height = "50"
    $SyncHash.Button_OutlookKill.Background = "#7eabfd"
    $SyncHash.Button_OutlookKill.BorderBrush = "White"
    $SyncHash.Button_OutlookKill.Foreground = "White"
    $SyncHash.Button_OutlookKill.Padding = "4"
    $SyncHash.Button_OutlookKill.ToolTip = "Kill Outlook"

    $SyncHash.Button_OutlookKill.Add_MouseEnter( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Hand
        $SyncHash.Button_OutlookKill.ForeGround = '#7eabfd'

    })

    $SyncHash.Button_OutlookKill.Add_MouseLeave( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $SyncHash.Button_OutlookKill.ForeGround = '#ffffff'

    })

    $SyncHash.Button_OutlookKill.Add_Click( {
        
        Invoke-OutlookKill

    })

    # Button Customisation - Button_TeamsReprofile
    $SyncHash.Button_TeamsReprofile = $SyncHash.Window.FindName("Button_TeamsReprofile")
    $SyncHash.Button_TeamsReprofile.Content = "REPROFILE"
    $SyncHash.Button_TeamsReprofile.Margin = "120 0 0 20"
    $SyncHash.Button_TeamsReprofile.Width = "100"
    $SyncHash.Button_TeamsReprofile.Height = "50"
    $SyncHash.Button_TeamsReprofile.Background = "#7eabfd"
    $SyncHash.Button_TeamsReprofile.BorderBrush = "White"
    $SyncHash.Button_TeamsReprofile.Foreground = "White"
    $SyncHash.Button_TeamsReprofile.Padding = "4"
    $SyncHash.Button_TeamsReprofile.ToolTip = "Reprofile Teams"

    $SyncHash.Button_TeamsReprofile.Add_MouseEnter( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Hand
        $SyncHash.Button_TeamsReprofile.ForeGround = '#7eabfd'

    })

    $SyncHash.Button_TeamsReprofile.Add_MouseLeave( {
        
        $SyncHash.Window.Cursor = [System.Windows.Input.Cursors]::Arrow
        $SyncHash.Button_TeamsReprofile.ForeGround = '#ffffff'

    })

    $SyncHash.Button_TeamsReprofile.Add_Click( {
        
        Invoke-TeamsKill
        Invoke-OutlookKill
        Start-Sleep -Seconds 2
        $SyncHash.TeamsProcessConstant = $false
        $SyncHash.Label_TeamsProcessConstant.Foreground = "Red"
        $SyncHash.OutlookProcessConstant = $false
        $SyncHash.Label_OutlookProcessConstant.Foreground = "Red"
        Start-Sleep -Seconds 5
        Invoke-TeamsReProfile

        Invoke-TeamsStart
        Invoke-OutlookStart

    })

    $SyncHash.Window.ShowDialog() | Out-Null
    $SyncHash.Close()
    $SyncHash.Dispose()

    $SyncHash.Window.Add_Closing( { 
        $SyncHash.Stopwatch.Stop()
        $SyncHash.Timer.Stop()
    })

    Get-Runspace | Where-Object { $_.RunspaceAvailability -eq 'Available' } | ForEach-Object { $_.dispose() }

}

$PSinstance            = [powershell]::Create().AddScript($Code)
$PSinstance.Runspace   = $Runspace
$Job                   = $PSinstance.BeginInvoke()