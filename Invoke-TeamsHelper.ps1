Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, system.windows.forms

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
    Width="400"
    Height="300"
    ShowInTaskbar="True">
    <Grid Name="Grid_Main">

        <TabControl>

            <TabItem Header="Reprofile">
                
                <Grid>
                    <Grid.ColumnDefinitions>			
                        <ColumnDefinition Width="*" />
                        <ColumnDefinition Width="*" />
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*" />
                    </Grid.RowDefinitions>

                    <StackPanel Name="left" Grid.Column="1" Grid.ColumnSpan="1" Orientation="Vertical">
                        <Image Source="https://mlmho3tq8waj.i.optimole.com/0jKHcqI.QTEh~fb3d/w:70/h:60/q:55/rt:fill/g:ce/https://www.alanps1.io/wp-content/uploads/2021/02/logo-m365-teams-70-60.png" Width="31" Height="27" />
                    </StackPanel>

                    <StackPanel Name="right" Grid.Column="2" Grid.ColumnSpan="1" Orientation="Vertical">
                        <Image Source="https://mlmho3tq8waj.i.optimole.com/0jKHcqI.QTEh~fb3d/w:70/h:60/q:55/rt:fill/g:ce/https://www.alanps1.io/wp-content/uploads/2021/02/logo-m365-powerapps-70-60.png" Width="31" Height="27" />
                    </StackPanel>

                </Grid>

            </TabItem>

            <TabItem Header="Tab 2">
                <Label Content="Content Tab 2 goes here..." />
            </TabItem>

            <TabItem Header="Tab 3">
                <Label Content="Content Tab 3 goes here..." />
            </TabItem>

        </TabControl>

    </Grid>
</Window>
"@

    <# Set up the Synchronized Hash Table as this will allow access to the shared data whilst threading #>
    $Script:SyncHash    = [hashtable]::Synchronized(@{})
    $SyncHash.Host      = $host
    $Reader             = (New-Object System.Xml.XmlNodeReader $Xaml)
    $SyncHash.Window    = [Windows.Markup.XamlReader]::Load( $Reader )

    # Functions

    # Styling Elements

    $SyncHash.Window.ShowDialog() | Out-Null

}

$PSinstance            = [powershell]::Create().AddScript($Code)
$PSinstance.Runspace   = $Runspace
$Job                    = $PSinstance.BeginInvoke()