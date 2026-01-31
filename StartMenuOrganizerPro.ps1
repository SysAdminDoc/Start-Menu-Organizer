<#
.SYNOPSIS
    Start Menu Organizer Pro - Comprehensive Windows Start Menu management tool
.DESCRIPTION
    Professional tool to organize Windows Start Menu with features including:
    - Junk detection and removal
    - Broken shortcut detection
    - Duplicate shortcut detection
    - Folder flattening
    - Category-based organization
    - Custom patterns and categories
    - Batch renaming
    - Search/filter
    - Undo support
    - Preview mode
    - Import/Export configuration
.NOTES
    Author: Matt | Maven Imaging
    Requires: Windows 10/11, PowerShell 5.1+
    Run as Administrator for system-wide changes
#>

#Requires -Version 5.1

# ============================================================================
# CONFIGURATION
# ============================================================================

$script:Config = @{
    UserStartMenu   = [Environment]::GetFolderPath('StartMenu') + '\Programs'
    SystemStartMenu = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
    BackupRoot      = "$env:LOCALAPPDATA\StartMenuOrganizerPro\Backups"
    ConfigFile      = "$env:LOCALAPPDATA\StartMenuOrganizerPro\config.json"
    UndoFile        = "$env:LOCALAPPDATA\StartMenuOrganizerPro\undo.json"
    MaxUndoSteps    = 50
}

# Default junk patterns
$script:JunkPatterns = [System.Collections.ObjectModel.ObservableCollection[string]]@(
    '*uninstall*'
    '*readme*'
    '*help*'
    '*documentation*'
    '*manual*'
    '*license*'
    '*website*'
    '*support*'
    '*visit *'
    '*about *'
    '*release notes*'
    '*changelog*'
    '*what''s new*'
    '*getting started*'
    '*user guide*'
    '*online *'
    '*web link*'
    '*url*'
    '*register*'
    '*feedback*'
    '*update*'
    '*check for update*'
)

# Default categories
$script:Categories = [ordered]@{
    'Development'    = @('Visual Studio*', 'VS Code*', 'Code*', 'Git*', 'GitHub*', 'Python*', 'Node*', 'npm*', 'PowerShell*', 'Terminal*', 'Notepad++*', 'Sublime*', 'JetBrains*', 'Android Studio*', 'Eclipse*', 'NetBeans*', 'Arduino*', 'Docker*', 'Postman*', 'Insomnia*', 'MySQL*', 'PostgreSQL*', 'MongoDB*', 'Redis*', 'SQL Server*', 'Azure*', 'AWS*', 'Cursor*', 'Windsurf*', 'SSMS*', 'Management Studio*', 'HeidiSQL*', 'DBeaver*')
    'Browsers'       = @('Google Chrome*', 'Chrome*', 'Firefox*', 'Mozilla*', 'Edge*', 'Opera*', 'Brave*', 'Vivaldi*', 'Tor*', 'Waterfox*', 'LibreWolf*', 'Zen*', 'Arc*', 'Floorp*')
    'Communication'  = @('Discord*', 'Slack*', 'Teams*', 'Zoom*', 'Skype*', 'Telegram*', 'WhatsApp*', 'Signal*', 'Outlook*', 'Thunderbird*', 'Mail*', 'Messages*', 'Element*', 'Webex*', 'GoTo*', 'Beeper*', 'Ferdium*')
    'Media'          = @('VLC*', 'Media Player*', 'Spotify*', 'iTunes*', 'Music*', 'Audacity*', 'OBS*', 'Handbrake*', 'FFmpeg*', 'Plex*', 'Kodi*', 'foobar*', 'AIMP*', 'Winamp*', 'MusicBee*', 'MediaMonkey*', 'Jellyfin*', 'mpv*', 'PotPlayer*', 'MPC-*', 'Stremio*', 'Tidal*', 'Deezer*', 'Amazon Music*', 'Apple Music*')
    'Graphics'       = @('Photoshop*', 'GIMP*', 'Paint*', 'Illustrator*', 'Inkscape*', 'Blender*', 'SketchUp*', 'AutoCAD*', 'Figma*', 'Canva*', 'Affinity*', 'CorelDRAW*', 'Krita*', 'Lightroom*', 'DaVinci*', 'Premiere*', 'After Effects*', 'Final Cut*', 'ShareX*', 'Greenshot*', 'Snagit*', 'IrfanView*', 'XnView*', 'FastStone*', 'Flameshot*', 'Clip Studio*', 'Aseprite*')
    'Office'         = @('Word*', 'Excel*', 'PowerPoint*', 'OneNote*', 'Access*', 'Publisher*', 'Visio*', 'Project*', 'LibreOffice*', 'OpenOffice*', 'WPS*', 'Acrobat*', 'PDF*', 'Foxit*', 'SumatraPDF*', 'Nitro*', 'Notion*', 'Obsidian*', 'Evernote*', 'Joplin*', 'Standard Notes*', 'Logseq*', 'Craft*', 'Coda*')
    'Utilities'      = @('7-Zip*', 'WinRAR*', 'WinZip*', 'PeaZip*', 'CCleaner*', 'Revo*', 'IObit*', 'Malwarebytes*', 'Everything*', 'Wox*', 'PowerToys*', 'AutoHotkey*', 'TreeSize*', 'WizTree*', 'SpaceSniffer*', 'HWiNFO*', 'CPU-Z*', 'GPU-Z*', 'CrystalDisk*', 'Speccy*', 'AIDA64*', 'Rufus*', 'Etcher*', 'Ventoy*', 'ImgBurn*', 'AnyDesk*', 'TeamViewer*', 'RustDesk*', 'Parsec*', 'Barrier*', 'Synergy*', 'KeePass*', 'Bitwarden*', '1Password*', 'LastPass*', 'Dashlane*', 'Flow Launcher*', 'Keypirinha*', 'Listary*', 'Directory Opus*', 'Total Commander*', 'Bulk Rename*', 'Advanced Renamer*')
    'Gaming'         = @('Steam*', 'Epic Games*', 'GOG*', 'Origin*', 'EA *', 'Ubisoft*', 'Battle.net*', 'Riot*', 'Xbox*', 'PlayStation*', 'GeForce*', 'NVIDIA*', 'AMD *', 'Radeon*', 'MSI Afterburner*', 'RTSS*', 'Playnite*', 'LaunchBox*', 'Retroarch*', 'Moonlight*', 'Sunshine*', 'Game Bar*', 'Heroic*', 'Lutris*', 'itch*', 'Prism Launcher*', 'MultiMC*')
    'System'         = @('Control Panel*', 'Settings*', 'Device Manager*', 'Task Manager*', 'Registry*', 'Services*', 'Event Viewer*', 'Computer Management*', 'Disk Management*', 'System Information*', 'Resource Monitor*', 'Performance Monitor*', 'Windows Admin*', 'Administrative*', 'Command Prompt*', 'cmd*', 'Remote Desktop*', 'Hyper-V*', 'VMware*', 'VirtualBox*', 'WSL*', 'Linux*', 'Sysinternals*', 'Process Explorer*', 'Process Monitor*', 'Autoruns*')
    'Security'       = @('Windows Security*', 'Defender*', 'Firewall*', 'Antivirus*', 'Norton*', 'McAfee*', 'Kaspersky*', 'Avast*', 'AVG*', 'ESET*', 'Bitdefender*', 'Webroot*', 'Comodo*', 'Sophos*', 'F-Secure*', 'Trend Micro*', 'Wireshark*', 'Nmap*', 'Burp*', 'GlassWire*', 'Simplewall*', 'PortMaster*')
    'Networking'     = @('PuTTY*', 'WinSCP*', 'FileZilla*', 'Cyberduck*', 'mRemoteNG*', 'MobaXterm*', 'Royal TS*', 'Termius*', 'Angry IP*', 'Advanced IP*', 'NetSetMan*', 'OpenVPN*', 'WireGuard*', 'NordVPN*', 'ExpressVPN*', 'ProtonVPN*', 'Mullvad*', 'Tailscale*', 'ZeroTier*')
}

# Undo stack
$script:UndoStack = [System.Collections.Generic.List[PSObject]]::new()

# ============================================================================
# ELEVATION CHECK
# ============================================================================

$script:IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# ============================================================================
# ASSEMBLIES
# ============================================================================

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Shell COM object for reading shortcut targets
$script:WScriptShell = New-Object -ComObject WScript.Shell

# ============================================================================
# XAML INTERFACE
# ============================================================================

[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Start Menu Organizer Pro"
    Width="1400" Height="900"
    MinWidth="1200" MinHeight="700"
    WindowStartupLocation="CenterScreen"
    Background="#0d1117">
    
    <Window.Resources>
        <!-- Dark Theme Colors -->
        <SolidColorBrush x:Key="PrimaryBg" Color="#0d1117"/>
        <SolidColorBrush x:Key="SecondaryBg" Color="#161b22"/>
        <SolidColorBrush x:Key="TertiaryBg" Color="#21262d"/>
        <SolidColorBrush x:Key="CardBg" Color="#1c2128"/>
        <SolidColorBrush x:Key="AccentColor" Color="#238636"/>
        <SolidColorBrush x:Key="AccentHover" Color="#2ea043"/>
        <SolidColorBrush x:Key="DangerColor" Color="#da3633"/>
        <SolidColorBrush x:Key="DangerHover" Color="#f85149"/>
        <SolidColorBrush x:Key="WarningColor" Color="#d29922"/>
        <SolidColorBrush x:Key="InfoColor" Color="#58a6ff"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#e6edf3"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#8b949e"/>
        <SolidColorBrush x:Key="TextMuted" Color="#6e7681"/>
        <SolidColorBrush x:Key="BorderColor" Color="#30363d"/>
        <SolidColorBrush x:Key="BorderHover" Color="#8b949e"/>
        
        <!-- Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource AccentColor}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{StaticResource AccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#21262d"/>
                                <Setter Property="Foreground" Value="#484f58"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Secondary Button -->
        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="{StaticResource TertiaryBg}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                CornerRadius="6" Padding="{TemplateBinding Padding}"
                                BorderBrush="{StaticResource BorderColor}" BorderThickness="1">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#30363d"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="{StaticResource BorderHover}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#161b22"/>
                                <Setter Property="Foreground" Value="#484f58"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Danger Button -->
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="{StaticResource DangerColor}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{StaticResource DangerHover}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#21262d"/>
                                <Setter Property="Foreground" Value="#484f58"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Small Button -->
        <Style x:Key="SmallButton" TargetType="Button" BasedOn="{StaticResource SecondaryButton}">
            <Setter Property="Padding" Value="10,4"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        
        <!-- TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{StaticResource SecondaryBg}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="CaretBrush" Value="{StaticResource TextPrimary}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- CheckBox Style -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,3"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        
        <!-- ComboBox Full Dark Theme -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="28"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2" Background="{StaticResource SecondaryBg}"
                        BorderBrush="{StaticResource BorderColor}" BorderThickness="1" CornerRadius="6"/>
                <Path x:Name="Arrow" Grid.Column="1" Fill="{StaticResource TextSecondary}"
                      HorizontalAlignment="Center" VerticalAlignment="Center"
                      Data="M 0 0 L 4 4 L 8 0 Z"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="BorderBrush" Value="{StaticResource BorderHover}"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>
        
        <Style TargetType="ComboBox">
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="Background" Value="{StaticResource SecondaryBg}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton" Template="{StaticResource ComboBoxToggleButton}"
                                          Focusable="False" IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"/>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              Margin="10,3,28,3" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Grid x:Name="DropDown" SnapsToDevicePixels="True"
                                      MinWidth="{TemplateBinding ActualWidth}" MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border Background="{StaticResource SecondaryBg}"
                                            BorderBrush="{StaticResource BorderColor}" BorderThickness="1" CornerRadius="6">
                                        <ScrollViewer Margin="4" SnapsToDevicePixels="True">
                                            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                        </ScrollViewer>
                                    </Border>
                                </Grid>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <Style TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="border" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" CornerRadius="4">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{StaticResource TertiaryBg}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{StaticResource AccentColor}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- TabControl Style -->
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
        
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Padding" Value="16,10"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="border" Background="{TemplateBinding Background}" 
                                Padding="{TemplateBinding Padding}" CornerRadius="6,6,0,0"
                                BorderBrush="Transparent" BorderThickness="1,1,1,0">
                            <ContentPresenter ContentSource="Header"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{StaticResource SecondaryBg}"/>
                                <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="{StaticResource BorderColor}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- DataGrid Styles -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{StaticResource SecondaryBg}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowBackground" Value="Transparent"/>
            <Setter Property="AlternatingRowBackground" Value="#1c2128"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="VerticalGridLinesBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="SelectionMode" Value="Extended"/>
            <Setter Property="SelectionUnit" Value="FullRow"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{StaticResource TertiaryBg}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        
        <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1f6feb"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="Transparent"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#1c2128"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#1f6feb"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <!-- ListBox Style -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="{StaticResource SecondaryBg}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="Padding" Value="4"/>
        </Style>
        
        <Style TargetType="ListBoxItem">
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ListBoxItem">
                        <Border x:Name="border" Background="Transparent" Padding="{TemplateBinding Padding}" CornerRadius="4">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{StaticResource TertiaryBg}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1c2128"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- ProgressBar Style -->
        <Style TargetType="ProgressBar">
            <Setter Property="Background" Value="{StaticResource TertiaryBg}"/>
            <Setter Property="Foreground" Value="{StaticResource AccentColor}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Height" Value="4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ProgressBar">
                        <Grid>
                            <Border Background="{TemplateBinding Background}" CornerRadius="2"/>
                            <Border x:Name="PART_Track"/>
                            <Border x:Name="PART_Indicator" Background="{TemplateBinding Foreground}" 
                                    CornerRadius="2" HorizontalAlignment="Left"/>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- Context Menu Style -->
        <Style TargetType="ContextMenu">
            <Setter Property="Background" Value="{StaticResource SecondaryBg}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="Padding" Value="4"/>
        </Style>
        
        <Style TargetType="MenuItem">
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="MenuItem">
                        <Border x:Name="border" Background="Transparent" Padding="{TemplateBinding Padding}" CornerRadius="4">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto" SharedSizeGroup="Icon"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto" SharedSizeGroup="Shortcut"/>
                                </Grid.ColumnDefinitions>
                                <ContentPresenter x:Name="Icon" Grid.Column="0" ContentSource="Icon" Margin="0,0,8,0"/>
                                <ContentPresenter Grid.Column="1" ContentSource="Header"/>
                                <TextBlock Grid.Column="2" Text="{TemplateBinding InputGestureText}" 
                                           Foreground="{StaticResource TextMuted}" Margin="20,0,0,0"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="border" Property="Background" Value="{StaticResource TertiaryBg}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        
        <!-- ScrollBar Style -->
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Width" Value="10"/>
        </Style>
    </Window.Resources>
    
    <Window.InputBindings>
        <KeyBinding Key="Delete" Command="{x:Static ApplicationCommands.Delete}"/>
        <KeyBinding Key="A" Modifiers="Ctrl" Command="{x:Static ApplicationCommands.SelectAll}"/>
        <KeyBinding Key="Z" Modifiers="Ctrl" Command="{x:Static ApplicationCommands.Undo}"/>
        <KeyBinding Key="F" Modifiers="Ctrl" Command="{x:Static ApplicationCommands.Find}"/>
        <KeyBinding Key="F5" Command="{x:Static NavigationCommands.Refresh}"/>
    </Window.InputBindings>
    
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="{StaticResource SecondaryBg}" Padding="20,15" 
                BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Start Menu Organizer Pro" FontSize="22" FontWeight="Bold" 
                               Foreground="{StaticResource TextPrimary}"/>
                    <TextBlock x:Name="txtAdminStatus" FontSize="11" Foreground="{StaticResource TextSecondary}" Margin="0,4,0,0"/>
                </StackPanel>
                
                <!-- Search Bar -->
                <Grid Grid.Column="1" Margin="40,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*" MaxWidth="500"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="txtSearch" VerticalAlignment="Center">
                        <TextBox.Style>
                            <Style TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
                                <Setter Property="Tag" Value="Search items... (Ctrl+F)"/>
                            </Style>
                        </TextBox.Style>
                    </TextBox>
                    <TextBlock x:Name="txtSearchPlaceholder" Text="Search items... (Ctrl+F)" 
                               Foreground="{StaticResource TextMuted}" Padding="12,10"
                               IsHitTestVisible="False" VerticalAlignment="Center"/>
                </Grid>
                
                <!-- Quick Actions -->
                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <Button x:Name="btnUndo" Content="Undo" Style="{StaticResource SecondaryButton}" 
                            Margin="0,0,8,0" IsEnabled="False" ToolTip="Ctrl+Z"/>
                    <Button x:Name="btnBackup" Content="Backup" Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                    <Button x:Name="btnRestore" Content="Restore" Style="{StaticResource SecondaryButton}"/>
                </StackPanel>
            </Grid>
        </Border>
        
        <!-- Main Content -->
        <Grid Grid.Row="1" Margin="20">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" MinWidth="500"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="380"/>
            </Grid.ColumnDefinitions>
            
            <!-- Left Panel - Items Grid -->
            <Border Grid.Column="0" Background="{StaticResource SecondaryBg}" CornerRadius="8" 
                    BorderBrush="{StaticResource BorderColor}" BorderThickness="1">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Toolbar -->
                    <Border Grid.Row="0" Padding="12,10" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,1">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Grid.Column="0" Orientation="Horizontal">
                                <TextBlock Text="Scope:" Foreground="{StaticResource TextSecondary}" 
                                           VerticalAlignment="Center" Margin="0,0,8,0"/>
                                <ComboBox x:Name="cmbScope" Width="150">
                                    <ComboBoxItem Content="User Start Menu"/>
                                    <ComboBoxItem Content="System Start Menu"/>
                                    <ComboBoxItem Content="Both" IsSelected="True"/>
                                </ComboBox>
                                <Button x:Name="btnRefresh" Content="Refresh" Style="{StaticResource SmallButton}" 
                                        Margin="10,0,0,0" ToolTip="F5"/>
                            </StackPanel>
                            
                            <StackPanel Grid.Column="2" Orientation="Horizontal">
                                <TextBlock Text="Sort:" Foreground="{StaticResource TextSecondary}" 
                                           VerticalAlignment="Center" Margin="0,0,8,0"/>
                                <ComboBox x:Name="cmbSort" Width="140">
                                    <ComboBoxItem Content="Name" IsSelected="True"/>
                                    <ComboBoxItem Content="Type"/>
                                    <ComboBoxItem Content="Status"/>
                                    <ComboBoxItem Content="Location"/>
                                    <ComboBoxItem Content="Target"/>
                                </ComboBox>
                                <CheckBox x:Name="chkPreviewMode" Content="Preview Mode" Margin="15,0,0,0" 
                                          VerticalAlignment="Center" ToolTip="Show what actions would do without executing"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    
                    <!-- Filter Bar -->
                    <Border Grid.Row="1" Padding="12,8" Background="{StaticResource TertiaryBg}">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Filter:" Foreground="{StaticResource TextSecondary}" 
                                       VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <CheckBox x:Name="chkShowShortcuts" Content="Shortcuts" IsChecked="True" Margin="0,0,15,0"/>
                            <CheckBox x:Name="chkShowFolders" Content="Folders" IsChecked="True" Margin="0,0,15,0"/>
                            <CheckBox x:Name="chkShowJunk" Content="Junk" IsChecked="True" Margin="0,0,15,0"/>
                            <CheckBox x:Name="chkShowBroken" Content="Broken" IsChecked="True" Margin="0,0,15,0"/>
                            <CheckBox x:Name="chkShowDuplicates" Content="Duplicates" IsChecked="True"/>
                            <TextBlock x:Name="txtItemCount" Foreground="{StaticResource TextMuted}" 
                                       VerticalAlignment="Center" Margin="20,0,0,0"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- Data Grid -->
                    <DataGrid x:Name="dgItems" Grid.Row="2" Margin="0">
                        <DataGrid.ContextMenu>
                            <ContextMenu>
                                <MenuItem x:Name="ctxDelete" Header="Delete" InputGestureText="Del"/>
                                <MenuItem x:Name="ctxRename" Header="Rename"/>
                                <Separator Background="{StaticResource BorderColor}"/>
                                <MenuItem x:Name="ctxOpenLocation" Header="Open File Location"/>
                                <MenuItem x:Name="ctxOpenTarget" Header="Open Target Location"/>
                                <Separator Background="{StaticResource BorderColor}"/>
                                <MenuItem x:Name="ctxMoveToCategory" Header="Move to Category">
                                    <MenuItem x:Name="ctxCatDev" Header="Development"/>
                                    <MenuItem x:Name="ctxCatBrowsers" Header="Browsers"/>
                                    <MenuItem x:Name="ctxCatComm" Header="Communication"/>
                                    <MenuItem x:Name="ctxCatMedia" Header="Media"/>
                                    <MenuItem x:Name="ctxCatGraphics" Header="Graphics"/>
                                    <MenuItem x:Name="ctxCatOffice" Header="Office"/>
                                    <MenuItem x:Name="ctxCatUtils" Header="Utilities"/>
                                    <MenuItem x:Name="ctxCatGaming" Header="Gaming"/>
                                    <MenuItem x:Name="ctxCatSystem" Header="System"/>
                                    <MenuItem x:Name="ctxCatSecurity" Header="Security"/>
                                    <MenuItem x:Name="ctxCatNetwork" Header="Networking"/>
                                </MenuItem>
                                <Separator Background="{StaticResource BorderColor}"/>
                                <MenuItem x:Name="ctxSelectAll" Header="Select All" InputGestureText="Ctrl+A"/>
                                <MenuItem x:Name="ctxSelectNone" Header="Select None"/>
                            </ContextMenu>
                        </DataGrid.ContextMenu>
                        <DataGrid.Columns>
                            <DataGridCheckBoxColumn Binding="{Binding IsSelected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" 
                                                    Width="40" CanUserResize="False"/>
                            <DataGridTextColumn Header="Name" Binding="{Binding DisplayName}" Width="200"/>
                            <DataGridTextColumn Header="Type" Binding="{Binding ItemType}" Width="80"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="90"/>
                            <DataGridTextColumn Header="Location" Binding="{Binding RelativePath}" Width="200"/>
                            <DataGridTextColumn Header="Target" Binding="{Binding TargetPath}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <!-- Selection Bar -->
                    <Border Grid.Row="3" Padding="12,10" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,1,0,0">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="btnSelectAll" Content="All" Style="{StaticResource SmallButton}" Margin="0,0,5,0"/>
                                <Button x:Name="btnSelectNone" Content="None" Style="{StaticResource SmallButton}" Margin="0,0,5,0"/>
                                <Button x:Name="btnSelectJunk" Content="Junk" Style="{StaticResource SmallButton}" Margin="0,0,5,0"/>
                                <Button x:Name="btnSelectBroken" Content="Broken" Style="{StaticResource SmallButton}" Margin="0,0,5,0"/>
                                <Button x:Name="btnSelectDuplicates" Content="Duplicates" Style="{StaticResource SmallButton}" Margin="0,0,5,0"/>
                                <Button x:Name="btnSelectFolders" Content="Folders" Style="{StaticResource SmallButton}" Margin="0,0,5,0"/>
                                <Button x:Name="btnInvertSelection" Content="Invert" Style="{StaticResource SmallButton}"/>
                            </StackPanel>
                            
                            <TextBlock x:Name="txtSelectionCount" Grid.Column="1" Foreground="{StaticResource TextSecondary}" 
                                       VerticalAlignment="Center"/>
                        </Grid>
                    </Border>
                </Grid>
            </Border>
            
            <!-- Right Panel - Actions -->
            <Border Grid.Column="2" Background="{StaticResource SecondaryBg}" CornerRadius="8"
                    BorderBrush="{StaticResource BorderColor}" BorderThickness="1">
                <TabControl>
                    <!-- Actions Tab -->
                    <TabItem Header="Actions">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel Margin="15">
                                <!-- Cleanup Section -->
                                <TextBlock Text="CLEANUP" FontSize="11" FontWeight="Bold" 
                                           Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                                
                                <Button x:Name="btnDeleteSelected" Content="Delete Selected" 
                                        Style="{StaticResource DangerButton}" Margin="0,0,0,8"/>
                                <Button x:Name="btnRemoveAllJunk" Content="Remove All Junk" 
                                        Style="{StaticResource DangerButton}" Margin="0,0,0,8"/>
                                <Button x:Name="btnRemoveBroken" Content="Remove Broken Shortcuts" 
                                        Style="{StaticResource DangerButton}" Margin="0,0,0,8"/>
                                <Button x:Name="btnRemoveDuplicates" Content="Remove Duplicates" 
                                        Style="{StaticResource SecondaryButton}" Margin="0,0,0,8"/>
                                <Button x:Name="btnFlattenFolders" Content="Flatten Single-Item Folders" 
                                        Style="{StaticResource SecondaryButton}" Margin="0,0,0,8"/>
                                <Button x:Name="btnRemoveEmpty" Content="Remove Empty Folders" 
                                        Style="{StaticResource SecondaryButton}" Margin="0,0,0,8"/>
                                <Button x:Name="btnMoveAllToRoot" Content="Move All to Root" 
                                        Style="{StaticResource SecondaryButton}" Margin="0,0,0,20"
                                        ToolTip="Move all shortcuts from folders to the Start Menu root"/>
                                
                                <!-- Organize Section -->
                                <TextBlock Text="ORGANIZE" FontSize="11" FontWeight="Bold" 
                                           Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                                
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <ComboBox x:Name="cmbCategory"/>
                                    <Button x:Name="btnMoveToCategory" Grid.Column="1" Content="Move" 
                                            Style="{StaticResource SmallButton}" Margin="8,0,0,0"/>
                                </Grid>
                                <Button x:Name="btnAutoOrganize" Content="Auto-Organize All" 
                                        Style="{StaticResource ModernButton}" Margin="0,0,0,20"/>
                                
                                <!-- Rename Section -->
                                <TextBlock Text="BATCH RENAME" FontSize="11" FontWeight="Bold" 
                                           Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                                
                                <Button x:Name="btnStripVersions" Content="Strip Version Numbers" 
                                        Style="{StaticResource SecondaryButton}" Margin="0,0,0,8"/>
                                <Button x:Name="btnCleanNames" Content="Clean Up Names" 
                                        Style="{StaticResource SecondaryButton}" Margin="0,0,0,8"/>
                                
                                <Grid Margin="0,0,0,8">
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                    </Grid.RowDefinitions>
                                    <TextBox x:Name="txtFindText" Grid.Row="0" Margin="0,0,0,5"/>
                                    <TextBlock Text="Find:" Foreground="{StaticResource TextMuted}" 
                                               Padding="12,10" IsHitTestVisible="False" Grid.Row="0"
                                               x:Name="txtFindPlaceholder"/>
                                    <TextBox x:Name="txtReplaceText" Grid.Row="1" Margin="0,0,0,5"/>
                                    <TextBlock Text="Replace:" Foreground="{StaticResource TextMuted}" 
                                               Padding="12,10" IsHitTestVisible="False" Grid.Row="1"
                                               x:Name="txtReplacePlaceholder"/>
                                    <Button x:Name="btnFindReplace" Grid.Row="2" Content="Find &amp; Replace in Names" 
                                            Style="{StaticResource SecondaryButton}"/>
                                </Grid>
                                
                                <!-- Quick Open Section -->
                                <TextBlock Text="QUICK OPEN" FontSize="11" FontWeight="Bold" 
                                           Foreground="{StaticResource TextMuted}" Margin="0,15,0,10"/>
                                
                                <StackPanel Orientation="Horizontal">
                                    <Button x:Name="btnOpenUserMenu" Content="User Menu" 
                                            Style="{StaticResource SmallButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnOpenSystemMenu" Content="System Menu" 
                                            Style="{StaticResource SmallButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnOpenBackups" Content="Backups" 
                                            Style="{StaticResource SmallButton}"/>
                                </StackPanel>
                            </StackPanel>
                        </ScrollViewer>
                    </TabItem>
                    
                    <!-- Settings Tab -->
                    <TabItem Header="Settings">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel Margin="15">
                                <!-- Junk Patterns -->
                                <TextBlock Text="JUNK PATTERNS" FontSize="11" FontWeight="Bold" 
                                           Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                                
                                <ListBox x:Name="lstJunkPatterns" Height="150" Margin="0,0,0,8"/>
                                
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="txtNewJunkPattern"/>
                                    <Button x:Name="btnAddJunkPattern" Grid.Column="1" Content="Add" 
                                            Style="{StaticResource SmallButton}" Margin="5,0,0,0"/>
                                    <Button x:Name="btnRemoveJunkPattern" Grid.Column="2" Content="Remove" 
                                            Style="{StaticResource SmallButton}" Margin="5,0,0,0"/>
                                </Grid>
                                
                                <!-- Category Management -->
                                <TextBlock Text="CATEGORIES" FontSize="11" FontWeight="Bold" 
                                           Foreground="{StaticResource TextMuted}" Margin="0,15,0,10"/>
                                
                                <ComboBox x:Name="cmbEditCategory" Margin="0,0,0,8"/>
                                <ListBox x:Name="lstCategoryPatterns" Height="120" Margin="0,0,0,8"/>
                                
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="txtNewCategoryPattern"/>
                                    <Button x:Name="btnAddCategoryPattern" Grid.Column="1" Content="Add" 
                                            Style="{StaticResource SmallButton}" Margin="5,0,0,0"/>
                                    <Button x:Name="btnRemoveCategoryPattern" Grid.Column="2" Content="Remove" 
                                            Style="{StaticResource SmallButton}" Margin="5,0,0,0"/>
                                </Grid>
                                
                                <!-- Import/Export -->
                                <TextBlock Text="CONFIGURATION" FontSize="11" FontWeight="Bold" 
                                           Foreground="{StaticResource TextMuted}" Margin="0,15,0,10"/>
                                
                                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                    <Button x:Name="btnExportConfig" Content="Export Config" 
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnImportConfig" Content="Import Config" 
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnResetConfig" Content="Reset to Defaults" 
                                            Style="{StaticResource SecondaryButton}"/>
                                </StackPanel>
                            </StackPanel>
                        </ScrollViewer>
                    </TabItem>
                    
                    <!-- Log Tab -->
                    <TabItem Header="Log">
                        <Grid Margin="10">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <Border Background="{StaticResource PrimaryBg}" CornerRadius="6" Padding="10">
                                <ScrollViewer x:Name="svLog" VerticalScrollBarVisibility="Auto">
                                    <TextBlock x:Name="txtLog" Foreground="{StaticResource TextSecondary}" 
                                               FontFamily="Cascadia Code,Consolas,monospace" FontSize="11" TextWrapping="Wrap"/>
                                </ScrollViewer>
                            </Border>
                            
                            <Button x:Name="btnClearLog" Grid.Row="1" Content="Clear Log" 
                                    Style="{StaticResource SmallButton}" HorizontalAlignment="Right" Margin="0,10,0,0"/>
                        </Grid>
                    </TabItem>
                </TabControl>
            </Border>
        </Grid>
        
        <!-- Footer -->
        <Border Grid.Row="2" Background="{StaticResource SecondaryBg}" Padding="20,12"
                BorderBrush="{StaticResource BorderColor}" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Orientation="Horizontal">
                    <ProgressBar x:Name="progressBar" Width="200" Visibility="Collapsed" Margin="0,0,15,0"/>
                    <TextBlock x:Name="txtStatus" Foreground="{StaticResource TextSecondary}" 
                               VerticalAlignment="Center" FontSize="12"/>
                </StackPanel>
                
                <TextBlock x:Name="txtStats" Grid.Column="2" Foreground="{StaticResource TextMuted}" 
                           VerticalAlignment="Center" FontSize="11"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ============================================================================
# LOAD WINDOW
# ============================================================================

$reader = [System.Xml.XmlNodeReader]::new($XAML)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Get all named controls
$XAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
    $name = $_.Name
    Set-Variable -Name $name -Value $Window.FindName($name) -Scope Script
}

# ============================================================================
# DATA MODEL
# ============================================================================

$script:AllItems = [System.Collections.Generic.List[PSObject]]::new()
$script:FilteredItems = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
$dgItems.ItemsSource = $script:FilteredItems

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Write-Log {
    param([string]$Message, [ValidateSet('Info','Success','Warning','Error')]$Level = 'Info')
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $color = switch ($Level) {
        'Success' { '#2ea043' }
        'Warning' { '#d29922' }
        'Error'   { '#f85149' }
        default   { '#8b949e' }
    }
    
    $Window.Dispatcher.Invoke([Action]{
        $run = [System.Windows.Documents.Run]::new("[$timestamp] $Message`r`n")
        $run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($color)
        $txtLog.Inlines.Add($run)
        $svLog.ScrollToEnd()
    })
}

function Update-Status {
    param([string]$Message)
    $Window.Dispatcher.Invoke([Action]{ $txtStatus.Text = $Message })
}

function Show-Progress {
    param([int]$Value, [int]$Maximum = 100, [bool]$Visible = $true)
    $Window.Dispatcher.Invoke([Action]{
        $progressBar.Maximum = $Maximum
        $progressBar.Value = $Value
        $progressBar.Visibility = if ($Visible) { 'Visible' } else { 'Collapsed' }
    })
}

function Get-StartMenuPaths {
    $scope = $cmbScope.SelectedIndex
    $paths = @()
    
    if ($scope -eq 0 -or $scope -eq 2) { $paths += $Config.UserStartMenu }
    if ($scope -eq 1 -or $scope -eq 2) { $paths += $Config.SystemStartMenu }
    
    return $paths
}

function Get-ShortcutTarget {
    param([string]$ShortcutPath)
    
    try {
        $shortcut = $script:WScriptShell.CreateShortcut($ShortcutPath)
        return $shortcut.TargetPath
    }
    catch {
        return $null
    }
}

function Test-ShortcutBroken {
    param([string]$ShortcutPath)
    
    $target = Get-ShortcutTarget $ShortcutPath
    if ([string]::IsNullOrEmpty($target)) { return $false }
    
    # Skip UWP apps and special paths
    if ($target -match '^[A-Za-z]:\\Windows\\' -or 
        $target -match '^shell:' -or 
        $target -match '\.exe$' -eq $false) {
        return $false
    }
    
    return -not (Test-Path $target)
}

function Test-IsJunk {
    param([string]$Name)
    
    foreach ($pattern in $script:JunkPatterns) {
        if ($Name -like $pattern) { return $true }
    }
    return $false
}

function Get-ItemCategory {
    param([string]$Name)
    
    foreach ($category in $script:Categories.Keys) {
        foreach ($pattern in $script:Categories[$category]) {
            if ($Name -like $pattern) { return $category }
        }
    }
    return $null
}

function Get-ItemStatus {
    param($Item)
    
    $statuses = @()
    if ($Item.IsJunk) { $statuses += 'Junk' }
    if ($Item.IsBroken) { $statuses += 'Broken' }
    if ($Item.IsDuplicate) { $statuses += 'Duplicate' }
    if ($statuses.Count -eq 0) { return 'OK' }
    return $statuses -join ', '
}

function Add-UndoAction {
    param(
        [string]$ActionType,
        [string]$Description,
        [array]$Items
    )
    
    $undoData = @{
        Type = $ActionType
        Description = $Description
        Timestamp = Get-Date
        Items = $Items | ForEach-Object {
            @{
                OriginalPath = $_.FullPath
                OriginalName = $_.DisplayName
                BasePath = $_.BasePath
                BackupPath = $null
            }
        }
    }
    
    $script:UndoStack.Add($undoData)
    if ($script:UndoStack.Count -gt $Config.MaxUndoSteps) {
        $script:UndoStack.RemoveAt(0)
    }
    
    $Window.Dispatcher.Invoke([Action]{ $btnUndo.IsEnabled = $true })
}

function Invoke-Undo {
    if ($script:UndoStack.Count -eq 0) { return }
    
    $lastAction = $script:UndoStack[$script:UndoStack.Count - 1]
    $script:UndoStack.RemoveAt($script:UndoStack.Count - 1)
    
    Write-Log "Undo: $($lastAction.Description)" 'Warning'
    
    # Undo logic depends on action type
    switch ($lastAction.Type) {
        'Delete' {
            # Restore from backup if available
            foreach ($item in $lastAction.Items) {
                if ($item.BackupPath -and (Test-Path $item.BackupPath)) {
                    $destDir = Split-Path $item.OriginalPath -Parent
                    if (-not (Test-Path $destDir)) {
                        New-Item -Path $destDir -ItemType Directory -Force | Out-Null
                    }
                    Move-Item -Path $item.BackupPath -Destination $item.OriginalPath -Force -ErrorAction SilentlyContinue
                }
            }
        }
        'Move' {
            foreach ($item in $lastAction.Items) {
                if (Test-Path $item.NewPath) {
                    Move-Item -Path $item.NewPath -Destination $item.OriginalPath -Force -ErrorAction SilentlyContinue
                }
            }
        }
        'Rename' {
            foreach ($item in $lastAction.Items) {
                $currentPath = Join-Path (Split-Path $item.OriginalPath -Parent) $item.NewName
                if (Test-Path $currentPath) {
                    Rename-Item -Path $currentPath -NewName $item.OriginalName -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
    
    if ($script:UndoStack.Count -eq 0) {
        $Window.Dispatcher.Invoke([Action]{ $btnUndo.IsEnabled = $false })
    }
    
    Refresh-Items
}

function Refresh-Items {
    $script:AllItems.Clear()
    $script:FilteredItems.Clear()
    
    $paths = Get-StartMenuPaths
    $allShortcuts = @{}
    $duplicateTargets = @{}
    
    Update-Status "Scanning Start Menu..."
    
    # First pass - collect all items and identify duplicates
    foreach ($basePath in $paths) {
        if (-not (Test-Path $basePath)) { continue }
        
        $isSystem = $basePath -eq $Config.SystemStartMenu
        $prefix = if ($isSystem) { "[Sys]" } else { "[User]" }
        
        Get-ChildItem -Path $basePath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $relativePath = $_.FullName.Substring($basePath.Length + 1)
            $isFolder = $_.PSIsContainer
            $isJunk = Test-IsJunk $_.BaseName
            $isBroken = $false
            $targetPath = ''
            
            if (-not $isFolder -and $_.Extension -eq '.lnk') {
                $targetPath = Get-ShortcutTarget $_.FullName
                $isBroken = Test-ShortcutBroken $_.FullName
                
                # Track for duplicate detection
                if (-not [string]::IsNullOrEmpty($targetPath)) {
                    if (-not $allShortcuts.ContainsKey($targetPath)) {
                        $allShortcuts[$targetPath] = @()
                    }
                    $allShortcuts[$targetPath] += $_.FullName
                }
            }
            
            $itemType = if ($isFolder) { 'Folder' } elseif ($isJunk) { 'Junk' } else { 'Shortcut' }
            
            $item = [PSCustomObject]@{
                IsSelected   = $false
                DisplayName  = $_.BaseName
                RelativePath = "$prefix $relativePath"
                FullPath     = $_.FullName
                BasePath     = $basePath
                IsFolder     = $isFolder
                IsJunk       = $isJunk
                IsBroken     = $isBroken
                IsDuplicate  = $false
                ItemType     = $itemType
                TargetPath   = $targetPath
                Status       = 'OK'
                IsSystem     = $isSystem
            }
            
            $script:AllItems.Add($item)
        }
    }
    
    # Second pass - mark duplicates
    $duplicateTargets = $allShortcuts.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
    foreach ($dup in $duplicateTargets) {
        foreach ($path in $dup.Value) {
            $item = $script:AllItems | Where-Object { $_.FullPath -eq $path } | Select-Object -First 1
            if ($item) {
                $item.IsDuplicate = $true
                $item.ItemType = 'Duplicate'
            }
        }
    }
    
    # Update status for all items
    foreach ($item in $script:AllItems) {
        $item.Status = Get-ItemStatus $item
    }
    
    Apply-Filters
    Update-Stats
    Update-Status "Ready"
}

function Apply-Filters {
    $searchText = $txtSearch.Text.ToLower()
    
    $script:FilteredItems.Clear()
    
    foreach ($item in $script:AllItems) {
        # Search filter
        if (-not [string]::IsNullOrEmpty($searchText)) {
            if (-not ($item.DisplayName.ToLower().Contains($searchText) -or 
                      $item.RelativePath.ToLower().Contains($searchText) -or
                      $item.TargetPath.ToLower().Contains($searchText))) {
                continue
            }
        }
        
        # Type filters
        if ($item.IsFolder -and -not $chkShowFolders.IsChecked) { continue }
        if (-not $item.IsFolder -and -not $item.IsJunk -and -not $item.IsBroken -and -not $item.IsDuplicate -and -not $chkShowShortcuts.IsChecked) { continue }
        if ($item.IsJunk -and -not $chkShowJunk.IsChecked) { continue }
        if ($item.IsBroken -and -not $chkShowBroken.IsChecked) { continue }
        if ($item.IsDuplicate -and -not $chkShowDuplicates.IsChecked) { continue }
        
        $script:FilteredItems.Add($item)
    }
    
    # Apply sorting
    $sortBy = $cmbSort.SelectedIndex
    $sorted = switch ($sortBy) {
        0 { $script:FilteredItems | Sort-Object DisplayName }
        1 { $script:FilteredItems | Sort-Object ItemType, DisplayName }
        2 { $script:FilteredItems | Sort-Object Status, DisplayName }
        3 { $script:FilteredItems | Sort-Object RelativePath }
        4 { $script:FilteredItems | Sort-Object TargetPath }
        default { $script:FilteredItems }
    }
    
    $script:FilteredItems.Clear()
    foreach ($item in $sorted) {
        $script:FilteredItems.Add($item)
    }
    
    $txtItemCount.Text = "$($script:FilteredItems.Count) of $($script:AllItems.Count) items"
}

function Update-Stats {
    $total = $script:AllItems.Count
    $junk = ($script:AllItems | Where-Object { $_.IsJunk }).Count
    $broken = ($script:AllItems | Where-Object { $_.IsBroken }).Count
    $duplicates = ($script:AllItems | Where-Object { $_.IsDuplicate }).Count
    $folders = ($script:AllItems | Where-Object { $_.IsFolder }).Count
    
    $txtStats.Text = "Total: $total | Folders: $folders | Junk: $junk | Broken: $broken | Duplicates: $duplicates"
}

function Update-SelectionCount {
    $selected = ($script:FilteredItems | Where-Object { $_.IsSelected }).Count
    $txtSelectionCount.Text = "$selected selected"
}

function Get-SelectedItems {
    return $script:FilteredItems | Where-Object { $_.IsSelected }
}

# ============================================================================
# ACTION FUNCTIONS
# ============================================================================

function Create-Backup {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = Join-Path $Config.BackupRoot $timestamp
    
    if (-not (Test-Path $Config.BackupRoot)) {
        New-Item -Path $Config.BackupRoot -ItemType Directory -Force | Out-Null
    }
    
    New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
    
    $paths = Get-StartMenuPaths
    $i = 0
    foreach ($path in $paths) {
        if (Test-Path $path) {
            $destName = if ($path -eq $Config.SystemStartMenu) { "System" } else { "User" }
            $destPath = Join-Path $backupPath $destName
            Show-Progress -Value ($i++ * 50) -Maximum 100
            Copy-Item -Path $path -Destination $destPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    
    Show-Progress -Visible $false
    Write-Log "Backup created: $backupPath" 'Success'
    [System.Windows.MessageBox]::Show("Backup created successfully!`n`n$backupPath", "Backup Complete", "OK", "Information")
    return $backupPath
}

function Delete-SelectedItems {
    $selected = @(Get-SelectedItems)
    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No items selected.", "Info", "OK", "Information")
        return
    }
    
    $isPreview = $chkPreviewMode.IsChecked
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would delete $($selected.Count) item(s):" 'Warning'
        foreach ($item in $selected) {
            Write-Log "  - $($item.DisplayName) ($($item.RelativePath))" 'Info'
        }
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Delete $($selected.Count) selected item(s)?`n`nThis action can be undone.",
        "Confirm Delete", "YesNo", "Warning"
    )
    
    if ($result -ne 'Yes') { return }
    
    # Create temp backup for undo
    $tempBackup = Join-Path $env:TEMP "StartMenuOrganizer_$(Get-Date -Format 'yyyyMMddHHmmss')"
    New-Item -Path $tempBackup -ItemType Directory -Force | Out-Null
    
    $undoItems = @()
    $deleted = 0
    $failed = 0
    
    Show-Progress -Value 0 -Maximum $selected.Count
    
    foreach ($i in 0..($selected.Count - 1)) {
        $item = $selected[$i]
        Show-Progress -Value ($i + 1) -Maximum $selected.Count
        
        try {
            if ($item.IsSystem -and -not $script:IsAdmin) {
                Write-Log "Skipped (need admin): $($item.DisplayName)" 'Warning'
                $failed++
                continue
            }
            
            if (Test-Path $item.FullPath) {
                # Backup for undo
                $backupDest = Join-Path $tempBackup ([System.IO.Path]::GetRandomFileName())
                Copy-Item -Path $item.FullPath -Destination $backupDest -Recurse -Force
                
                Remove-Item -Path $item.FullPath -Recurse -Force -ErrorAction Stop
                Write-Log "Deleted: $($item.DisplayName)" 'Success'
                
                $undoItems += @{
                    OriginalPath = $item.FullPath
                    OriginalName = $item.DisplayName
                    BasePath = $item.BasePath
                    BackupPath = $backupDest
                }
                $deleted++
            }
        }
        catch {
            Write-Log "Failed to delete: $($item.DisplayName) - $_" 'Error'
            $failed++
        }
    }
    
    if ($undoItems.Count -gt 0) {
        $script:UndoStack.Add(@{
            Type = 'Delete'
            Description = "Delete $deleted item(s)"
            Timestamp = Get-Date
            Items = $undoItems
        })
        $btnUndo.IsEnabled = $true
    }
    
    Show-Progress -Visible $false
    Write-Log "Deletion complete: $deleted deleted, $failed failed" 'Info'
    Refresh-Items
}

function Remove-AllJunk {
    $isPreview = $chkPreviewMode.IsChecked
    $junkItems = @($script:AllItems | Where-Object { $_.IsJunk })
    
    if ($junkItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No junk items found.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would delete $($junkItems.Count) junk item(s):" 'Warning'
        foreach ($item in $junkItems | Select-Object -First 20) {
            Write-Log "  - $($item.DisplayName)" 'Info'
        }
        if ($junkItems.Count -gt 20) {
            Write-Log "  ... and $($junkItems.Count - 20) more" 'Info'
        }
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Delete ALL $($junkItems.Count) junk items?`n`nThis action can be undone.",
        "Confirm Junk Removal", "YesNo", "Warning"
    )
    
    if ($result -ne 'Yes') { return }
    
    foreach ($item in $junkItems) { $item.IsSelected = $true }
    $dgItems.Items.Refresh()
    Delete-SelectedItems
}

function Remove-BrokenShortcuts {
    $isPreview = $chkPreviewMode.IsChecked
    $brokenItems = @($script:AllItems | Where-Object { $_.IsBroken })
    
    if ($brokenItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No broken shortcuts found.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would delete $($brokenItems.Count) broken shortcut(s):" 'Warning'
        foreach ($item in $brokenItems) {
            Write-Log "  - $($item.DisplayName) -> $($item.TargetPath)" 'Info'
        }
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Delete $($brokenItems.Count) broken shortcuts?`n`nThese point to files/folders that no longer exist.",
        "Confirm Removal", "YesNo", "Warning"
    )
    
    if ($result -ne 'Yes') { return }
    
    foreach ($item in $brokenItems) { $item.IsSelected = $true }
    $dgItems.Items.Refresh()
    Delete-SelectedItems
}

function Remove-Duplicates {
    $isPreview = $chkPreviewMode.IsChecked
    
    # Group by target path and keep only the first (usually in root)
    $duplicateGroups = $script:AllItems | 
        Where-Object { $_.IsDuplicate } | 
        Group-Object TargetPath
    
    if ($duplicateGroups.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No duplicate shortcuts found.", "Info", "OK", "Information")
        return
    }
    
    $toDelete = @()
    foreach ($group in $duplicateGroups) {
        # Keep the shortest path (usually root level), delete the rest
        $sorted = $group.Group | Sort-Object { $_.RelativePath.Length }
        $toDelete += $sorted | Select-Object -Skip 1
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would delete $($toDelete.Count) duplicate(s), keeping 1 of each:" 'Warning'
        foreach ($item in $toDelete) {
            Write-Log "  - $($item.DisplayName) ($($item.RelativePath))" 'Info'
        }
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Delete $($toDelete.Count) duplicate shortcuts?`n`nOne copy of each will be kept.",
        "Confirm Removal", "YesNo", "Warning"
    )
    
    if ($result -ne 'Yes') { return }
    
    foreach ($item in $script:AllItems) { $item.IsSelected = $false }
    foreach ($item in $toDelete) { 
        $match = $script:AllItems | Where-Object { $_.FullPath -eq $item.FullPath } | Select-Object -First 1
        if ($match) { $match.IsSelected = $true }
    }
    $dgItems.Items.Refresh()
    Delete-SelectedItems
}

function Flatten-SingleItemFolders {
    $isPreview = $chkPreviewMode.IsChecked
    $paths = Get-StartMenuPaths
    $toFlatten = @()
    
    foreach ($basePath in $paths) {
        $isSystem = $basePath -eq $Config.SystemStartMenu
        if ($isSystem -and -not $script:IsAdmin) { continue }
        
        $folders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue
        
        foreach ($folder in $folders) {
            $contents = Get-ChildItem -Path $folder.FullName -File -Filter "*.lnk" -ErrorAction SilentlyContinue
            $subfolders = Get-ChildItem -Path $folder.FullName -Directory -ErrorAction SilentlyContinue
            
            if ($contents.Count -eq 1 -and $subfolders.Count -eq 0) {
                $toFlatten += @{
                    Folder = $folder
                    Shortcut = $contents[0]
                    BasePath = $basePath
                }
            }
        }
    }
    
    if ($toFlatten.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No single-item folders found to flatten.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would flatten $($toFlatten.Count) folder(s):" 'Warning'
        foreach ($item in $toFlatten) {
            Write-Log "  - $($item.Folder.Name) -> $($item.Shortcut.Name)" 'Info'
        }
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Flatten $($toFlatten.Count) single-item folders?`n`nShortcuts will be moved to the parent level.",
        "Confirm Flatten", "YesNo", "Question"
    )
    
    if ($result -ne 'Yes') { return }
    
    $flattened = 0
    foreach ($item in $toFlatten) {
        try {
            $destPath = Join-Path $item.BasePath $item.Shortcut.Name
            
            if (Test-Path $destPath) {
                $destPath = Join-Path $item.BasePath "$($item.Folder.Name) - $($item.Shortcut.Name)"
            }
            
            Move-Item -Path $item.Shortcut.FullName -Destination $destPath -Force
            Remove-Item -Path $item.Folder.FullName -Force -Recurse
            
            Write-Log "Flattened: $($item.Folder.Name)" 'Success'
            $flattened++
        }
        catch {
            Write-Log "Failed to flatten: $($item.Folder.Name) - $_" 'Error'
        }
    }
    
    Write-Log "Flattening complete: $flattened folders flattened" 'Info'
    Refresh-Items
}

function Remove-EmptyFolders {
    $isPreview = $chkPreviewMode.IsChecked
    $paths = Get-StartMenuPaths
    $emptyFolders = @()
    
    foreach ($basePath in $paths) {
        $isSystem = $basePath -eq $Config.SystemStartMenu
        if ($isSystem -and -not $script:IsAdmin) { continue }
        
        $folders = Get-ChildItem -Path $basePath -Directory -Recurse -ErrorAction SilentlyContinue | 
                   Sort-Object { $_.FullName.Length } -Descending
        
        foreach ($folder in $folders) {
            $contents = Get-ChildItem -Path $folder.FullName -Force -ErrorAction SilentlyContinue
            if ($contents.Count -eq 0) {
                $emptyFolders += $folder
            }
        }
    }
    
    if ($emptyFolders.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No empty folders found.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would remove $($emptyFolders.Count) empty folder(s):" 'Warning'
        foreach ($folder in $emptyFolders) {
            Write-Log "  - $($folder.Name)" 'Info'
        }
        return
    }
    
    $removed = 0
    foreach ($folder in $emptyFolders) {
        try {
            Remove-Item -Path $folder.FullName -Force
            Write-Log "Removed empty: $($folder.Name)" 'Success'
            $removed++
        }
        catch {
            Write-Log "Failed to remove: $($folder.Name) - $_" 'Error'
        }
    }
    
    Write-Log "Removed $removed empty folders" 'Info'
    Refresh-Items
}

function Move-AllToRoot {
    $isPreview = $chkPreviewMode.IsChecked
    $paths = Get-StartMenuPaths
    $toMove = @()
    
    foreach ($basePath in $paths) {
        $isSystem = $basePath -eq $Config.SystemStartMenu
        if ($isSystem -and -not $script:IsAdmin) { continue }
        
        # Get all shortcuts that are NOT in the root directory
        Get-ChildItem -Path $basePath -Recurse -Filter "*.lnk" -ErrorAction SilentlyContinue | ForEach-Object {
            $parentDir = Split-Path $_.FullName -Parent
            if ($parentDir -ne $basePath) {
                $toMove += @{
                    Item = $_
                    BasePath = $basePath
                    IsSystem = $isSystem
                }
            }
        }
    }
    
    if ($toMove.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No shortcuts found in folders to move.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would move $($toMove.Count) shortcut(s) to root:" 'Warning'
        foreach ($entry in $toMove | Select-Object -First 30) {
            Write-Log "  - $($entry.Item.Name)" 'Info'
        }
        if ($toMove.Count -gt 30) {
            Write-Log "  ... and $($toMove.Count - 30) more" 'Info'
        }
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Move $($toMove.Count) shortcuts from folders to the Start Menu root?`n`nEmpty folders will be removed afterward.",
        "Confirm Move All to Root", "YesNo", "Question"
    )
    
    if ($result -ne 'Yes') { return }
    
    $moved = 0
    $skipped = 0
    Show-Progress -Value 0 -Maximum $toMove.Count
    
    foreach ($i in 0..($toMove.Count - 1)) {
        $entry = $toMove[$i]
        Show-Progress -Value ($i + 1) -Maximum $toMove.Count
        
        try {
            $destPath = Join-Path $entry.BasePath $entry.Item.Name
            
            # Handle name collision
            if (Test-Path $destPath) {
                # Check if it's pointing to the same target
                $existingTarget = Get-ShortcutTarget $destPath
                $newTarget = Get-ShortcutTarget $entry.Item.FullName
                
                if ($existingTarget -eq $newTarget) {
                    # Same target, just delete the duplicate
                    Remove-Item -Path $entry.Item.FullName -Force
                    Write-Log "Removed duplicate: $($entry.Item.BaseName)" 'Info'
                    $skipped++
                    continue
                }
                
                # Different target, create unique name
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($entry.Item.Name)
                $ext = [System.IO.Path]::GetExtension($entry.Item.Name)
                $counter = 2
                do {
                    $destPath = Join-Path $entry.BasePath "$baseName ($counter)$ext"
                    $counter++
                } while (Test-Path $destPath)
            }
            
            Move-Item -Path $entry.Item.FullName -Destination $destPath -Force
            Write-Log "Moved to root: $($entry.Item.BaseName)" 'Success'
            $moved++
        }
        catch {
            Write-Log "Failed to move: $($entry.Item.BaseName) - $_" 'Error'
        }
    }
    
    Show-Progress -Visible $false
    
    # Clean up empty folders
    Write-Log "Cleaning up empty folders..." 'Info'
    Remove-EmptyFolders
    
    Write-Log "Move to root complete: $moved moved, $skipped duplicates removed" 'Info'
    Refresh-Items
}

function Move-ToCategory {
    param([string]$CategoryName)
    
    $selected = @(Get-SelectedItems | Where-Object { -not $_.IsFolder })
    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No shortcuts selected.", "Info", "OK", "Information")
        return
    }
    
    $isPreview = $chkPreviewMode.IsChecked
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would move $($selected.Count) item(s) to '$CategoryName':" 'Warning'
        foreach ($item in $selected) {
            Write-Log "  - $($item.DisplayName)" 'Info'
        }
        return
    }
    
    $moved = 0
    foreach ($item in $selected) {
        if ($item.IsSystem -and -not $script:IsAdmin) {
            Write-Log "Skipped (need admin): $($item.DisplayName)" 'Warning'
            continue
        }
        
        try {
            $categoryPath = Join-Path $item.BasePath $CategoryName
            if (-not (Test-Path $categoryPath)) {
                New-Item -Path $categoryPath -ItemType Directory -Force | Out-Null
            }
            
            $destPath = Join-Path $categoryPath (Split-Path $item.FullPath -Leaf)
            Move-Item -Path $item.FullPath -Destination $destPath -Force
            Write-Log "Moved to $CategoryName`: $($item.DisplayName)" 'Success'
            $moved++
        }
        catch {
            Write-Log "Failed to move: $($item.DisplayName) - $_" 'Error'
        }
    }
    
    Write-Log "Moved $moved items to $CategoryName" 'Info'
    Refresh-Items
}

function Auto-OrganizeAll {
    $isPreview = $chkPreviewMode.IsChecked
    $paths = Get-StartMenuPaths
    $toOrganize = @()
    
    foreach ($basePath in $paths) {
        $isSystem = $basePath -eq $Config.SystemStartMenu
        if ($isSystem -and -not $script:IsAdmin) { continue }
        
        Get-ChildItem -Path $basePath -Recurse -Filter "*.lnk" -ErrorAction SilentlyContinue | ForEach-Object {
            $category = Get-ItemCategory $_.BaseName
            if ($category) {
                $parentName = Split-Path (Split-Path $_.FullName -Parent) -Leaf
                if ($parentName -ne $category) {
                    $toOrganize += @{
                        Item = $_
                        Category = $category
                        BasePath = $basePath
                    }
                }
            }
        }
    }
    
    if ($toOrganize.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No items found to organize.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would organize $($toOrganize.Count) item(s):" 'Warning'
        $grouped = $toOrganize | Group-Object { $_.Category }
        foreach ($group in $grouped) {
            Write-Log "  $($group.Name): $($group.Count) items" 'Info'
        }
        return
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Organize $($toOrganize.Count) items into category folders?",
        "Confirm Auto-Organize", "YesNo", "Question"
    )
    
    if ($result -ne 'Yes') { return }
    
    $organized = 0
    Show-Progress -Value 0 -Maximum $toOrganize.Count
    
    foreach ($i in 0..($toOrganize.Count - 1)) {
        $entry = $toOrganize[$i]
        Show-Progress -Value ($i + 1) -Maximum $toOrganize.Count
        
        try {
            $categoryPath = Join-Path $entry.BasePath $entry.Category
            if (-not (Test-Path $categoryPath)) {
                New-Item -Path $categoryPath -ItemType Directory -Force | Out-Null
            }
            
            $destPath = Join-Path $categoryPath $entry.Item.Name
            if (-not (Test-Path $destPath)) {
                Move-Item -Path $entry.Item.FullName -Destination $destPath -Force
                Write-Log "Organized: $($entry.Item.BaseName) -> $($entry.Category)" 'Success'
                $organized++
            }
        }
        catch {
            Write-Log "Failed: $($entry.Item.BaseName) - $_" 'Error'
        }
    }
    
    Show-Progress -Visible $false
    
    # Clean up empty folders
    Remove-EmptyFolders
    
    Write-Log "Auto-organize complete: $organized items organized" 'Info'
    Refresh-Items
}

function Strip-VersionNumbers {
    $selected = @(Get-SelectedItems | Where-Object { -not $_.IsFolder })
    if ($selected.Count -eq 0) {
        # If nothing selected, operate on all
        $selected = @($script:AllItems | Where-Object { -not $_.IsFolder })
    }
    
    $isPreview = $chkPreviewMode.IsChecked
    $toRename = @()
    
    # Pattern to match version numbers
    $versionPatterns = @(
        '\s*v?\d+\.\d+(\.\d+)*\s*$',
        '\s*\(\d+\.\d+(\.\d+)*\)\s*$',
        '\s*\[\d+\.\d+(\.\d+)*\]\s*$',
        '\s*-\s*v?\d+\.\d+(\.\d+)*\s*$',
        '\s+\d{4}\s*$',
        '\s*v\d+\s*$'
    )
    
    foreach ($item in $selected) {
        $newName = $item.DisplayName
        foreach ($pattern in $versionPatterns) {
            $newName = $newName -replace $pattern, ''
        }
        $newName = $newName.Trim()
        
        if ($newName -ne $item.DisplayName -and -not [string]::IsNullOrEmpty($newName)) {
            $toRename += @{
                Item = $item
                NewName = $newName
            }
        }
    }
    
    if ($toRename.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No version numbers found to strip.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would rename $($toRename.Count) item(s):" 'Warning'
        foreach ($entry in $toRename) {
            Write-Log "  - '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Info'
        }
        return
    }
    
    $renamed = 0
    foreach ($entry in $toRename) {
        if ($entry.Item.IsSystem -and -not $script:IsAdmin) { continue }
        
        try {
            $ext = [System.IO.Path]::GetExtension($entry.Item.FullPath)
            $newPath = Join-Path (Split-Path $entry.Item.FullPath -Parent) "$($entry.NewName)$ext"
            
            if (-not (Test-Path $newPath)) {
                Rename-Item -Path $entry.Item.FullPath -NewName "$($entry.NewName)$ext" -Force
                Write-Log "Renamed: '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Success'
                $renamed++
            }
        }
        catch {
            Write-Log "Failed to rename: $($entry.Item.DisplayName) - $_" 'Error'
        }
    }
    
    Write-Log "Renamed $renamed items" 'Info'
    Refresh-Items
}

function Clean-Names {
    $selected = @(Get-SelectedItems | Where-Object { -not $_.IsFolder })
    if ($selected.Count -eq 0) {
        $selected = @($script:AllItems | Where-Object { -not $_.IsFolder })
    }
    
    $isPreview = $chkPreviewMode.IsChecked
    $toRename = @()
    
    # Patterns to clean
    $cleanPatterns = @(
        @{ Pattern = '^\s+|\s+$'; Replace = '' },
        @{ Pattern = '\s{2,}'; Replace = ' ' },
        @{ Pattern = '\s*-\s*Shortcut\s*$'; Replace = '' },
        @{ Pattern = '\s*\(x64\)\s*$'; Replace = '' },
        @{ Pattern = '\s*\(x86\)\s*$'; Replace = '' },
        @{ Pattern = '\s*\(64-bit\)\s*$'; Replace = '' },
        @{ Pattern = '\s*\(32-bit\)\s*$'; Replace = '' },
        @{ Pattern = '\s*64-bit\s*$'; Replace = '' },
        @{ Pattern = '^\s*Microsoft\s+'; Replace = '' }
    )
    
    foreach ($item in $selected) {
        $newName = $item.DisplayName
        foreach ($p in $cleanPatterns) {
            $newName = $newName -replace $p.Pattern, $p.Replace
        }
        $newName = $newName.Trim()
        
        if ($newName -ne $item.DisplayName -and -not [string]::IsNullOrEmpty($newName)) {
            $toRename += @{
                Item = $item
                NewName = $newName
            }
        }
    }
    
    if ($toRename.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No names found to clean.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would clean $($toRename.Count) name(s):" 'Warning'
        foreach ($entry in $toRename) {
            Write-Log "  - '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Info'
        }
        return
    }
    
    $renamed = 0
    foreach ($entry in $toRename) {
        if ($entry.Item.IsSystem -and -not $script:IsAdmin) { continue }
        
        try {
            $ext = [System.IO.Path]::GetExtension($entry.Item.FullPath)
            $newPath = Join-Path (Split-Path $entry.Item.FullPath -Parent) "$($entry.NewName)$ext"
            
            if (-not (Test-Path $newPath)) {
                Rename-Item -Path $entry.Item.FullPath -NewName "$($entry.NewName)$ext" -Force
                Write-Log "Cleaned: '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Success'
                $renamed++
            }
        }
        catch {
            Write-Log "Failed to clean: $($entry.Item.DisplayName) - $_" 'Error'
        }
    }
    
    Write-Log "Cleaned $renamed names" 'Info'
    Refresh-Items
}

function Find-Replace-Names {
    $findText = $txtFindText.Text
    $replaceText = $txtReplaceText.Text
    
    if ([string]::IsNullOrEmpty($findText)) {
        [System.Windows.MessageBox]::Show("Enter text to find.", "Info", "OK", "Information")
        return
    }
    
    $selected = @(Get-SelectedItems | Where-Object { -not $_.IsFolder })
    if ($selected.Count -eq 0) {
        $selected = @($script:AllItems | Where-Object { -not $_.IsFolder -and $_.DisplayName -like "*$findText*" })
    }
    
    $isPreview = $chkPreviewMode.IsChecked
    $toRename = @()
    
    foreach ($item in $selected) {
        if ($item.DisplayName -like "*$findText*") {
            $newName = $item.DisplayName -replace [regex]::Escape($findText), $replaceText
            if ($newName -ne $item.DisplayName -and -not [string]::IsNullOrEmpty($newName)) {
                $toRename += @{
                    Item = $item
                    NewName = $newName
                }
            }
        }
    }
    
    if ($toRename.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No matches found.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        Write-Log "PREVIEW: Would rename $($toRename.Count) item(s):" 'Warning'
        foreach ($entry in $toRename) {
            Write-Log "  - '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Info'
        }
        return
    }
    
    $renamed = 0
    foreach ($entry in $toRename) {
        if ($entry.Item.IsSystem -and -not $script:IsAdmin) { continue }
        
        try {
            $ext = [System.IO.Path]::GetExtension($entry.Item.FullPath)
            Rename-Item -Path $entry.Item.FullPath -NewName "$($entry.NewName)$ext" -Force
            Write-Log "Renamed: '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Success'
            $renamed++
        }
        catch {
            Write-Log "Failed: $($entry.Item.DisplayName) - $_" 'Error'
        }
    }
    
    Write-Log "Renamed $renamed items" 'Info'
    Refresh-Items
}

function Restore-Backup {
    if (-not (Test-Path $Config.BackupRoot)) {
        [System.Windows.MessageBox]::Show("No backups found.", "Info", "OK", "Information")
        return
    }
    
    $backups = Get-ChildItem -Path $Config.BackupRoot -Directory | Sort-Object Name -Descending
    if ($backups.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No backups found.", "Info", "OK", "Information")
        return
    }
    
    # Show selection dialog
    $backupList = $backups | ForEach-Object { 
        $date = [DateTime]::ParseExact($_.Name, 'yyyyMMdd_HHmmss', $null)
        "$($date.ToString('yyyy-MM-dd HH:mm:ss')) - $($_.Name)"
    }
    
    $result = [System.Windows.MessageBox]::Show(
        "Restore from latest backup?`n`n$($backups[0].Name)`n`nThis will replace current Start Menu contents!",
        "Confirm Restore", "YesNo", "Warning"
    )
    
    if ($result -ne 'Yes') { return }
    
    $latestBackup = $backups[0]
    
    try {
        $userBackup = Join-Path $latestBackup.FullName "User"
        $systemBackup = Join-Path $latestBackup.FullName "System"
        
        if (Test-Path $userBackup) {
            Remove-Item -Path "$($Config.UserStartMenu)\*" -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item -Path "$userBackup\*" -Destination $Config.UserStartMenu -Recurse -Force
            Write-Log "Restored user Start Menu" 'Success'
        }
        
        if ((Test-Path $systemBackup) -and $script:IsAdmin) {
            Remove-Item -Path "$($Config.SystemStartMenu)\*" -Recurse -Force -ErrorAction SilentlyContinue
            Copy-Item -Path "$systemBackup\*" -Destination $Config.SystemStartMenu -Recurse -Force
            Write-Log "Restored system Start Menu" 'Success'
        }
        
        Write-Log "Restore complete!" 'Success'
        Refresh-Items
        [System.Windows.MessageBox]::Show("Restore complete!", "Success", "OK", "Information")
    }
    catch {
        Write-Log "Restore failed: $_" 'Error'
        [System.Windows.MessageBox]::Show("Restore failed: $_", "Error", "OK", "Error")
    }
}

function Export-Configuration {
    $saveDialog = [System.Windows.Forms.SaveFileDialog]::new()
    $saveDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $saveDialog.FileName = "StartMenuOrganizerConfig.json"
    
    if ($saveDialog.ShowDialog() -eq 'OK') {
        $config = @{
            JunkPatterns = @($script:JunkPatterns)
            Categories = $script:Categories
        }
        
        $config | ConvertTo-Json -Depth 5 | Set-Content -Path $saveDialog.FileName -Encoding UTF8
        Write-Log "Configuration exported: $($saveDialog.FileName)" 'Success'
    }
}

function Import-Configuration {
    $openDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $openDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    
    if ($openDialog.ShowDialog() -eq 'OK') {
        try {
            $config = Get-Content -Path $openDialog.FileName -Raw | ConvertFrom-Json
            
            if ($config.JunkPatterns) {
                $script:JunkPatterns.Clear()
                foreach ($pattern in $config.JunkPatterns) {
                    $script:JunkPatterns.Add($pattern)
                }
            }
            
            if ($config.Categories) {
                $script:Categories = [ordered]@{}
                foreach ($prop in $config.Categories.PSObject.Properties) {
                    $script:Categories[$prop.Name] = @($prop.Value)
                }
                Refresh-CategoryUI
            }
            
            Refresh-JunkPatternsUI
            Write-Log "Configuration imported: $($openDialog.FileName)" 'Success'
            Refresh-Items
        }
        catch {
            Write-Log "Failed to import configuration: $_" 'Error'
            [System.Windows.MessageBox]::Show("Failed to import: $_", "Error", "OK", "Error")
        }
    }
}

function Reset-Configuration {
    $result = [System.Windows.MessageBox]::Show(
        "Reset all junk patterns and categories to defaults?",
        "Confirm Reset", "YesNo", "Question"
    )
    
    if ($result -ne 'Yes') { return }
    
    # Reset junk patterns
    $script:JunkPatterns.Clear()
    @(
        '*uninstall*', '*readme*', '*help*', '*documentation*', '*manual*',
        '*license*', '*website*', '*support*', '*visit *', '*about *',
        '*release notes*', '*changelog*', '*what''s new*', '*getting started*',
        '*user guide*', '*online *', '*web link*', '*url*', '*register*',
        '*feedback*', '*update*', '*check for update*'
    ) | ForEach-Object { $script:JunkPatterns.Add($_) }
    
    Refresh-JunkPatternsUI
    Write-Log "Configuration reset to defaults" 'Success'
    Refresh-Items
}

function Refresh-JunkPatternsUI {
    $lstJunkPatterns.ItemsSource = $null
    $lstJunkPatterns.ItemsSource = $script:JunkPatterns
}

function Refresh-CategoryUI {
    $cmbCategory.Items.Clear()
    $cmbEditCategory.Items.Clear()
    
    foreach ($category in $script:Categories.Keys) {
        $cmbCategory.Items.Add($category) | Out-Null
        $cmbEditCategory.Items.Add($category) | Out-Null
    }
    
    if ($cmbCategory.Items.Count -gt 0) { $cmbCategory.SelectedIndex = 0 }
    if ($cmbEditCategory.Items.Count -gt 0) { $cmbEditCategory.SelectedIndex = 0 }
}

function Refresh-CategoryPatternsUI {
    $selectedCategory = $cmbEditCategory.SelectedItem
    if (-not $selectedCategory) { return }
    
    $lstCategoryPatterns.ItemsSource = $null
    $lstCategoryPatterns.ItemsSource = $script:Categories[$selectedCategory]
}

# ============================================================================
# EVENT HANDLERS
# ============================================================================

# Search box placeholder
$txtSearch.Add_GotFocus({
    $txtSearchPlaceholder.Visibility = 'Collapsed'
})

$txtSearch.Add_LostFocus({
    if ([string]::IsNullOrEmpty($txtSearch.Text)) {
        $txtSearchPlaceholder.Visibility = 'Visible'
    }
})

$txtSearch.Add_TextChanged({
    Apply-Filters
})

# Find/Replace placeholders
$txtFindText.Add_GotFocus({ $txtFindPlaceholder.Visibility = 'Collapsed' })
$txtFindText.Add_LostFocus({ if ([string]::IsNullOrEmpty($txtFindText.Text)) { $txtFindPlaceholder.Visibility = 'Visible' } })
$txtReplaceText.Add_GotFocus({ $txtReplacePlaceholder.Visibility = 'Collapsed' })
$txtReplaceText.Add_LostFocus({ if ([string]::IsNullOrEmpty($txtReplaceText.Text)) { $txtReplacePlaceholder.Visibility = 'Visible' } })

# Refresh and scope
$btnRefresh.Add_Click({ Refresh-Items })
$cmbScope.Add_SelectionChanged({ Refresh-Items })
$cmbSort.Add_SelectionChanged({ Apply-Filters })

# Filter checkboxes
$chkShowShortcuts.Add_Checked({ Apply-Filters })
$chkShowShortcuts.Add_Unchecked({ Apply-Filters })
$chkShowFolders.Add_Checked({ Apply-Filters })
$chkShowFolders.Add_Unchecked({ Apply-Filters })
$chkShowJunk.Add_Checked({ Apply-Filters })
$chkShowJunk.Add_Unchecked({ Apply-Filters })
$chkShowBroken.Add_Checked({ Apply-Filters })
$chkShowBroken.Add_Unchecked({ Apply-Filters })
$chkShowDuplicates.Add_Checked({ Apply-Filters })
$chkShowDuplicates.Add_Unchecked({ Apply-Filters })

# Selection buttons
$btnSelectAll.Add_Click({
    foreach ($item in $script:FilteredItems) { $item.IsSelected = $true }
    $dgItems.Items.Refresh()
    Update-SelectionCount
})

$btnSelectNone.Add_Click({
    foreach ($item in $script:FilteredItems) { $item.IsSelected = $false }
    $dgItems.Items.Refresh()
    Update-SelectionCount
})

$btnSelectJunk.Add_Click({
    foreach ($item in $script:FilteredItems) { $item.IsSelected = $item.IsJunk }
    $dgItems.Items.Refresh()
    Update-SelectionCount
})

$btnSelectBroken.Add_Click({
    foreach ($item in $script:FilteredItems) { $item.IsSelected = $item.IsBroken }
    $dgItems.Items.Refresh()
    Update-SelectionCount
})

$btnSelectDuplicates.Add_Click({
    foreach ($item in $script:FilteredItems) { $item.IsSelected = $item.IsDuplicate }
    $dgItems.Items.Refresh()
    Update-SelectionCount
})

$btnSelectFolders.Add_Click({
    foreach ($item in $script:FilteredItems) { $item.IsSelected = $item.IsFolder }
    $dgItems.Items.Refresh()
    Update-SelectionCount
})

$btnInvertSelection.Add_Click({
    foreach ($item in $script:FilteredItems) { $item.IsSelected = -not $item.IsSelected }
    $dgItems.Items.Refresh()
    Update-SelectionCount
})

# Action buttons
$btnDeleteSelected.Add_Click({ Delete-SelectedItems })
$btnRemoveAllJunk.Add_Click({ Remove-AllJunk })
$btnRemoveBroken.Add_Click({ Remove-BrokenShortcuts })
$btnRemoveDuplicates.Add_Click({ Remove-Duplicates })
$btnFlattenFolders.Add_Click({ Flatten-SingleItemFolders })
$btnRemoveEmpty.Add_Click({ Remove-EmptyFolders })
$btnMoveAllToRoot.Add_Click({ Move-AllToRoot })
$btnMoveToCategory.Add_Click({ Move-ToCategory $cmbCategory.SelectedItem })
$btnAutoOrganize.Add_Click({ Auto-OrganizeAll })
$btnStripVersions.Add_Click({ Strip-VersionNumbers })
$btnCleanNames.Add_Click({ Clean-Names })
$btnFindReplace.Add_Click({ Find-Replace-Names })

# Backup/Restore
$btnBackup.Add_Click({ Create-Backup })
$btnRestore.Add_Click({ Restore-Backup })
$btnUndo.Add_Click({ Invoke-Undo })

# Quick open buttons
$btnOpenUserMenu.Add_Click({ Start-Process explorer.exe -ArgumentList $Config.UserStartMenu })
$btnOpenSystemMenu.Add_Click({ Start-Process explorer.exe -ArgumentList $Config.SystemStartMenu })
$btnOpenBackups.Add_Click({
    if (-not (Test-Path $Config.BackupRoot)) {
        New-Item -Path $Config.BackupRoot -ItemType Directory -Force | Out-Null
    }
    Start-Process explorer.exe -ArgumentList $Config.BackupRoot
})

# Settings - Junk Patterns
$btnAddJunkPattern.Add_Click({
    $pattern = $txtNewJunkPattern.Text.Trim()
    if (-not [string]::IsNullOrEmpty($pattern) -and -not $script:JunkPatterns.Contains($pattern)) {
        $script:JunkPatterns.Add($pattern)
        Refresh-JunkPatternsUI
        $txtNewJunkPattern.Text = ''
        Write-Log "Added junk pattern: $pattern" 'Success'
        Refresh-Items
    }
})

$btnRemoveJunkPattern.Add_Click({
    $selected = $lstJunkPatterns.SelectedItem
    if ($selected) {
        $script:JunkPatterns.Remove($selected)
        Refresh-JunkPatternsUI
        Write-Log "Removed junk pattern: $selected" 'Success'
        Refresh-Items
    }
})

# Settings - Categories
$cmbEditCategory.Add_SelectionChanged({ Refresh-CategoryPatternsUI })

$btnAddCategoryPattern.Add_Click({
    $category = $cmbEditCategory.SelectedItem
    $pattern = $txtNewCategoryPattern.Text.Trim()
    
    if ($category -and -not [string]::IsNullOrEmpty($pattern)) {
        if ($script:Categories[$category] -notcontains $pattern) {
            $script:Categories[$category] += $pattern
            Refresh-CategoryPatternsUI
            $txtNewCategoryPattern.Text = ''
            Write-Log "Added pattern '$pattern' to category '$category'" 'Success'
        }
    }
})

$btnRemoveCategoryPattern.Add_Click({
    $category = $cmbEditCategory.SelectedItem
    $selected = $lstCategoryPatterns.SelectedItem
    
    if ($category -and $selected) {
        $script:Categories[$category] = @($script:Categories[$category] | Where-Object { $_ -ne $selected })
        Refresh-CategoryPatternsUI
        Write-Log "Removed pattern '$selected' from category '$category'" 'Success'
    }
})

$btnExportConfig.Add_Click({ Export-Configuration })
$btnImportConfig.Add_Click({ Import-Configuration })
$btnResetConfig.Add_Click({ Reset-Configuration })

$btnClearLog.Add_Click({
    $txtLog.Inlines.Clear()
})

# Context menu handlers
$ctxDelete.Add_Click({ Delete-SelectedItems })
$ctxSelectAll.Add_Click({
    foreach ($item in $script:FilteredItems) { $item.IsSelected = $true }
    $dgItems.Items.Refresh()
    Update-SelectionCount
})
$ctxSelectNone.Add_Click({
    foreach ($item in $script:FilteredItems) { $item.IsSelected = $false }
    $dgItems.Items.Refresh()
    Update-SelectionCount
})

$ctxOpenLocation.Add_Click({
    $selected = $dgItems.SelectedItem
    if ($selected) {
        $folder = Split-Path $selected.FullPath -Parent
        Start-Process explorer.exe -ArgumentList "/select,`"$($selected.FullPath)`""
    }
})

$ctxOpenTarget.Add_Click({
    $selected = $dgItems.SelectedItem
    if ($selected -and -not [string]::IsNullOrEmpty($selected.TargetPath)) {
        if (Test-Path $selected.TargetPath) {
            Start-Process explorer.exe -ArgumentList "/select,`"$($selected.TargetPath)`""
        }
        else {
            [System.Windows.MessageBox]::Show("Target path does not exist.", "Info", "OK", "Warning")
        }
    }
})

$ctxRename.Add_Click({
    $selected = $dgItems.SelectedItem
    if ($selected) {
        $newName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter new name:", "Rename", $selected.DisplayName)
        if (-not [string]::IsNullOrEmpty($newName) -and $newName -ne $selected.DisplayName) {
            try {
                $ext = [System.IO.Path]::GetExtension($selected.FullPath)
                Rename-Item -Path $selected.FullPath -NewName "$newName$ext" -Force
                Write-Log "Renamed: '$($selected.DisplayName)' -> '$newName'" 'Success'
                Refresh-Items
            }
            catch {
                Write-Log "Failed to rename: $_" 'Error'
            }
        }
    }
})

# Category context menu handlers
$ctxCatDev.Add_Click({ Move-ToCategory 'Development' })
$ctxCatBrowsers.Add_Click({ Move-ToCategory 'Browsers' })
$ctxCatComm.Add_Click({ Move-ToCategory 'Communication' })
$ctxCatMedia.Add_Click({ Move-ToCategory 'Media' })
$ctxCatGraphics.Add_Click({ Move-ToCategory 'Graphics' })
$ctxCatOffice.Add_Click({ Move-ToCategory 'Office' })
$ctxCatUtils.Add_Click({ Move-ToCategory 'Utilities' })
$ctxCatGaming.Add_Click({ Move-ToCategory 'Gaming' })
$ctxCatSystem.Add_Click({ Move-ToCategory 'System' })
$ctxCatSecurity.Add_Click({ Move-ToCategory 'Security' })
$ctxCatNetwork.Add_Click({ Move-ToCategory 'Networking' })

# DataGrid selection changed
$dgItems.Add_SelectionChanged({ Update-SelectionCount })

# Keyboard shortcuts
$Window.Add_KeyDown({
    param($sender, $e)
    
    if ($e.Key -eq 'Delete') {
        Delete-SelectedItems
        $e.Handled = $true
    }
    elseif ($e.Key -eq 'A' -and $e.KeyboardDevice.Modifiers -eq 'Control') {
        foreach ($item in $script:FilteredItems) { $item.IsSelected = $true }
        $dgItems.Items.Refresh()
        Update-SelectionCount
        $e.Handled = $true
    }
    elseif ($e.Key -eq 'Z' -and $e.KeyboardDevice.Modifiers -eq 'Control') {
        Invoke-Undo
        $e.Handled = $true
    }
    elseif ($e.Key -eq 'F' -and $e.KeyboardDevice.Modifiers -eq 'Control') {
        $txtSearch.Focus()
        $e.Handled = $true
    }
    elseif ($e.Key -eq 'F5') {
        Refresh-Items
        $e.Handled = $true
    }
})

# ============================================================================
# INITIALIZATION
# ============================================================================

# Load VB assembly for InputBox
Add-Type -AssemblyName Microsoft.VisualBasic

# Set admin status
if ($script:IsAdmin) {
    $txtAdminStatus.Text = "Running as Administrator - Full access to User and System Start Menu"
    $txtAdminStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
}
else {
    $txtAdminStatus.Text = "Standard User - Limited to User Start Menu (Run as Admin for full access)"
    $txtAdminStatus.Foreground = [System.Windows.Media.Brushes]::Orange
}

# Ensure config directory exists
$configDir = Split-Path $Config.ConfigFile -Parent
if (-not (Test-Path $configDir)) {
    New-Item -Path $configDir -ItemType Directory -Force | Out-Null
}

# Populate UI
Refresh-CategoryUI
Refresh-JunkPatternsUI
Refresh-CategoryPatternsUI

# Initial scan
Refresh-Items
Write-Log "Start Menu Organizer Pro initialized" 'Success'
Write-Log "Keyboard shortcuts: Del=Delete, Ctrl+A=Select All, Ctrl+Z=Undo, Ctrl+F=Search, F5=Refresh" 'Info'

# Show window
$Window.ShowDialog() | Out-Null

# Cleanup COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:WScriptShell) | Out-Null
