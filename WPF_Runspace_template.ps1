Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationCore, PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# init synchronized hashtable
$Sync = [HashTable]::Synchronized(@{})

# init runspace
$Runspace = [RunspaceFactory]::CreateRunspace()
$Runspace.ApartmentState = [Threading.ApartmentState]::STA
$Runspace.ThreadOptions = "ReuseThread"         
$Runspace.Open()

# provide the other thread with the synchronized hashtable (variable shared across threads)
$Runspace.SessionStateProxy.SetVariable("Sync", $Sync)

# load GUI from a XAML file
#[Xml]$WpfXml = Get-Content -LiteralPath "$PSScriptRoot\MainWindow.xaml"

# or

# paste XAML here
[Xml]$WpfXml = @"
<Window x:Name="WpfRunspaceTemplate" x:Class="WpfApp1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        mc:Ignorable="d"
        Title="WPF Runspace Template" WindowStartupLocation="CenterScreen" Visibility="Visible" ResizeMode="CanMinimize" Height="500" Width="500">
    <DockPanel>
        <Menu DockPanel.Dock="Top">
            <MenuItem Header="_File">
                <MenuItem
                    x:Name="mnuWindowStart"
                    Header="_Start"/>
                <Separator />
            </MenuItem>
        </Menu>
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition/>
                <RowDefinition Height="25"/>
                <RowDefinition Height="150"/>
                <RowDefinition Height="25"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0">
                <TextBox
                    x:Name="txtOutput"/>
            </Grid>
            <Grid Grid.Row="1">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition/>
                    <ColumnDefinition/>
                </Grid.ColumnDefinitions>
                <Button
                    Grid.Column="0"
                    x:Name="btnStart"
                    Content="Start"
                    IsEnabled="True"/>
                <Button
                    Grid.Column="1"
                    x:Name="btnPause"
                    Content="Pause"
                    IsEnabled="False"/>
                <Button
                    Grid.Column="2"
                    x:Name="btnStop"
                    Content="Stop"
                    IsEnabled="False"/>
            </Grid>
            <GroupBox
                Grid.Row="2"
                Header="Log"
                VerticalAlignment="Stretch">
                <ListBox
                    x:Name="lstLog">
                    <ListBox.ContextMenu>
                        <ContextMenu>
                            <MenuItem
                                x:Name="mnuLogCopy"
                                Header="Copy" />
                        </ContextMenu>
                    </ListBox.ContextMenu>
                </ListBox>
            </GroupBox>
            <ProgressBar
                Grid.Row="3"
                x:Name="progressBar"/>
            <TextBlock
                Grid.Row="3"
                x:Name="txtProgress"
                Text="{Binding ElementName=progressBar, Path=Value, StringFormat={}{0:0}%}"
                HorizontalAlignment="Center"
                VerticalAlignment="Center"/>
        </Grid>
    </DockPanel>
</Window>
"@

# these attributes can disturb powershell's ability to load XAML, so remove them
$WpfXml.Window.RemoveAttribute('x:Class')
$WpfXml.Window.RemoveAttribute('mc:Ignorable')

# add namespaces for later use if needed
$WpfNs = New-Object -TypeName Xml.XmlNamespaceManager -ArgumentList $WpfXml.NameTable
$WpfNs.AddNamespace('x', $WpfXml.DocumentElement.x)
$WpfNs.AddNamespace('d', $WpfXml.DocumentElement.d)
$WpfNs.AddNamespace('mc', $WpfXml.DocumentElement.mc)

$Sync.Gui = @{}
$Sync.UserPause = $false
$Sync.UserStop = $false

# function to update progress bar
$Sync.UpdateProgress = {
    param(
        [int]$ProgressValue
    )

    $Sync.Window.Dispatcher.Invoke([Action]{
        $Sync.Gui.progressBar.Value = $ProgressValue
    })
}

# function to log text to listbox
$Sync.LogText = {
    param(
        [string]$content,
        [string]$color
    )

    $Sync.Window.Dispatcher.Invoke([Action]{
        $lstItem = New-Object System.Windows.Controls.ListBoxItem -Property @{
            Content = $content
            Foreground = if ($color.Length -eq 0) {"Black"} else {$color}
        }

        $Sync.Gui.lstLog.Items.Add($lstItem)

        $Sync.Gui.lstLog.ScrollIntoView($Sync.Gui.lstLog.Items[$Sync.Gui.lstLog.Items.Count-1])
    })
}

# Read XAML markup
try {
    $Sync.Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $WpfXml))
} catch {
    Write-Host $_ -ForegroundColor Red
    Exit
}

#===================================================
# Retrieve a list of all GUI elements
#===================================================
$WpfXml.SelectNodes('//*[@x:Name]', $WpfNs) | ForEach-Object {
    $Sync.Gui.Add($_.Name, $Sync.Window.FindName($_.Name))
}

#===================================================
# Form element event handlers
#===================================================
$Sync.Gui.mnuLogCopy.add_Click({
    # copy listbox selected item text to clipboard
    Set-Clipboard -Value $Sync.Gui.lstLog.SelectedItem.Content
})

$Sync.Gui.btnPause.add_Click({
    # toggle pause/resume
    $Sync.UserPause = -not $Sync.UserPause
    $Sync.Gui.btnPause.Content = if ($Sync.UserPause) {"Resume"} else {"Pause"}
})

$Sync.Gui.btnStop.add_Click({
    # tell the other thread the user would like to stop
    $Sync.UserStop = $true
})

$StartTask = {
    # init states
    $Sync.Gui.btnStart.IsEnabled = $false
    $Sync.Gui.mnuWindowStart.IsEnabled = $false
    $Sync.Gui.btnPause.IsEnabled = $true
    $Sync.Gui.btnStop.IsEnabled = $true

    # add a script to run in the other thread
    $global:Session = [PowerShell]::Create().AddScript({
        # progress starts at 0%
        $Sync.UpdateProgress.Invoke(0)

        # log start date
        $Sync.LogText.Invoke("Started at $(Get-Date)")

        #===================================================
        # Start of long-running task
        #===================================================
        for ([int]$n = 1; (-not $Sync.UserStop) -and ($n -le 10); $n++) {
            # sample simulation of a long-running task
            Start-Sleep -Seconds 1

            # sample update the GUI on the main thread
            # from within the runspace session
            $Sync.Window.Dispatcher.Invoke([Action]{
                $Sync.Gui.txtOutput.Text = $n
            }, "Normal")

            # update progress bar
            $Sync.UpdateProgress.Invoke($n * 10)

            # log item to the listbox
            $Sync.LogText.Invoke("Completed item $n", "Green")

            # handle pause
            while ($Sync.UserPause -and (-not $Sync.UserStop)) {
                Start-Sleep -Seconds 1
            }
        }
        #===================================================
        # End of long-running task
        #===================================================

        # reset states
        $Sync.Window.Dispatcher.Invoke([Action]{
            $Sync.Gui.btnStart.IsEnabled = $true
            $Sync.Gui.mnuWindowStart.IsEnabled = $true
            $Sync.Gui.btnPause.IsEnabled = $false
            $Sync.Gui.btnStop.IsEnabled = $false
            $Sync.UserStop = $false
            $Sync.UserPause = $false
            $Sync.Gui.btnPause.Content = "Pause"
        })

        # log end date
        $Sync.LogText.Invoke("Finished at $(Get-Date)")
    }, $true)

    # invoke the runspace session created above
    $Session.Runspace = $Runspace
    $global:Handle = $Session.BeginInvoke()
}

$Sync.Gui.mnuWindowStart.add_Click($StartTask)

$Sync.Gui.btnStart.add_Click($StartTask)

#===================================================
# Window events
#===================================================
$Sync.Window.add_Closing({
    # if user triggers app close and runspace session not complete
    if (($null -ne $Session) -and ($Handle.IsCompleted -eq $false)) {
        # alert the user the command is still running
        [Windows.MessageBox]::Show('A command is still running.')
        # prevent exit
        $PSItem.Cancel = $true
    }
})

$Sync.Window.add_Closed({
    # end session and close runspace on window exit
    if ($null -ne $Session) {
        $Session.EndInvoke($Handle)
    }
    
    $Runspace.Close()
})

# display the form
[void]$Sync.Window.ShowDialog()