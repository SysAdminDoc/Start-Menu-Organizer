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
    Version         = "0.13.0"
    SettingsSchema  = 1
    UserStartMenu   = [Environment]::GetFolderPath('StartMenu') + '\Programs'
    SystemStartMenu = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs"
    ProfileRoot     = ''
    DefaultProfileRoot = "$env:SystemDrive\Users\Default"
    BackupRoot      = "$env:LOCALAPPDATA\StartMenuOrganizerPro\Backups"
    ConfigFile      = "$env:LOCALAPPDATA\StartMenuOrganizerPro\config.json"
    UndoFile        = "$env:LOCALAPPDATA\StartMenuOrganizerPro\undo.json"
    UndoBackupRoot  = "$env:LOCALAPPDATA\StartMenuOrganizerPro\UndoBackups"
    LogRoot         = "$env:LOCALAPPDATA\StartMenuOrganizerPro\Logs"
    LocalizationRoot = "$env:LOCALAPPDATA\StartMenuOrganizerPro\Localization"
    MaxUndoSteps    = 50
    MaxLogFiles     = 14
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

$script:ProtectedFolders = [System.Collections.ObjectModel.ObservableCollection[string]]@(
    'Accessories'
    'Administrative Tools'
    'Maintenance'
    'Startup'
    'System Tools'
    'Windows Administrative Tools'
    'Windows PowerShell'
    'Windows System'
    'Windows Tools'
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
    Title="Start Menu Organizer v$($Config.Version)"
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
        <SolidColorBrush x:Key="TextMuted" Color="#9da7b3"/>
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
                    <TextBlock x:Name="txtAppTitle" Text="Start Menu Organizer Pro" FontSize="22" FontWeight="Bold"
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
                    <Button x:Name="btnCancelWork" Content="Cancel" Style="{StaticResource DangerButton}"
                            Margin="0,0,8,0" IsEnabled="False"/>
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
                                <TextBlock x:Name="lblScope" Text="Scope:" Foreground="{StaticResource TextSecondary}"
                                           VerticalAlignment="Center" Margin="0,0,8,0"/>
                                <ComboBox x:Name="cmbScope" Width="150">
                                    <ComboBoxItem x:Name="cmbScopeUser" Content="User Start Menu"/>
                                    <ComboBoxItem x:Name="cmbScopeSystem" Content="System Start Menu"/>
                                    <ComboBoxItem x:Name="cmbScopeBoth" Content="Both" IsSelected="True"/>
                                    <ComboBoxItem x:Name="cmbScopeProfile" Content="Selected Profile"/>
                                    <ComboBoxItem x:Name="cmbScopeDefaultUser" Content="Default User"/>
                                </ComboBox>
                                <Button x:Name="btnRefresh" Content="Refresh" Style="{StaticResource SmallButton}" 
                                        Margin="10,0,0,0" ToolTip="F5"/>
                            </StackPanel>
                            
                            <StackPanel Grid.Column="2" Orientation="Horizontal">
                                <TextBlock x:Name="lblSort" Text="Sort:" Foreground="{StaticResource TextSecondary}"
                                           VerticalAlignment="Center" Margin="0,0,8,0"/>
                                <ComboBox x:Name="cmbSort" Width="140">
                                    <ComboBoxItem x:Name="cmbSortName" Content="Name" IsSelected="True"/>
                                    <ComboBoxItem x:Name="cmbSortType" Content="Type"/>
                                    <ComboBoxItem x:Name="cmbSortStatus" Content="Status"/>
                                    <ComboBoxItem x:Name="cmbSortLocation" Content="Location"/>
                                    <ComboBoxItem x:Name="cmbSortTarget" Content="Target"/>
                                </ComboBox>
                                <CheckBox x:Name="chkPreviewMode" Content="Preview Mode" Margin="15,0,0,0" 
                                          VerticalAlignment="Center" ToolTip="Show what actions would do without executing"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    
                    <!-- Filter Bar -->
                    <Border Grid.Row="1" Padding="12,8" Background="{StaticResource TertiaryBg}">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock x:Name="lblFilter" Text="Filter:" Foreground="{StaticResource TextSecondary}"
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
                            <DataGridTextColumn x:Name="colName" Header="Name" Binding="{Binding DisplayName}" Width="200"/>
                            <DataGridTextColumn x:Name="colType" Header="Type" Binding="{Binding ItemType}" Width="80"/>
                            <DataGridTextColumn x:Name="colStatus" Header="Status" Binding="{Binding Status}" Width="90"/>
                            <DataGridTextColumn x:Name="colLocation" Header="Location" Binding="{Binding RelativePath}" Width="200"/>
                            <DataGridTextColumn x:Name="colTarget" Header="Target" Binding="{Binding TargetPath}" Width="*"/>
                            <DataGridTextColumn x:Name="colRisk" Header="Risk" Binding="{Binding RiskFlags}" Width="120">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="Foreground" Value="#d29922"/>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
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
                    <TabItem x:Name="tabActions" Header="Actions">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel Margin="15">
                                <!-- Cleanup Section -->
                                <TextBlock x:Name="txtCleanupHeader" Text="CLEANUP" FontSize="11" FontWeight="Bold"
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
                                <TextBlock x:Name="txtOrganizeHeader" Text="ORGANIZE" FontSize="11" FontWeight="Bold"
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
                                <TextBlock x:Name="txtBatchRenameHeader" Text="BATCH RENAME" FontSize="11" FontWeight="Bold"
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

                                <!-- Transaction Plan Section -->
                                <TextBlock x:Name="txtTransactionPlanHeader" Text="TRANSACTION PLAN" FontSize="11" FontWeight="Bold"
                                           Foreground="{StaticResource TextMuted}" Margin="0,15,0,10"/>
                                <TextBlock x:Name="txtPlanStatus" Text="No plan loaded"
                                           Foreground="{StaticResource TextSecondary}" TextWrapping="Wrap"
                                           Margin="0,0,0,8"/>
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Button x:Name="btnExportPlan" Grid.Column="0" Content="Export Plan"
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,4,8" IsEnabled="False"/>
                                    <Button x:Name="btnImportPlan" Grid.Column="1" Content="Import Plan"
                                            Style="{StaticResource SecondaryButton}" Margin="4,0,0,8"/>
                                </Grid>
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Button x:Name="btnExecutePlan" Grid.Column="0" Content="Execute Plan"
                                            Style="{StaticResource ModernButton}" Margin="0,0,4,0" IsEnabled="False"/>
                                    <Button x:Name="btnClearPlan" Grid.Column="1" Content="Clear Plan"
                                            Style="{StaticResource SecondaryButton}" Margin="4,0,0,0" IsEnabled="False"/>
                                </Grid>

                                <!-- Quick Open Section -->
                                <TextBlock x:Name="txtQuickOpenHeader" Text="QUICK OPEN" FontSize="11" FontWeight="Bold"
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
                    <TabItem x:Name="tabSettings" Header="Settings">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel Margin="15">
                                <!-- Junk Patterns -->
                                <TextBlock x:Name="txtJunkPatternsHeader" Text="JUNK PATTERNS" FontSize="11" FontWeight="Bold"
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
                                <TextBlock x:Name="txtCategoriesHeader" Text="CATEGORIES" FontSize="11" FontWeight="Bold"
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
                                <TextBlock x:Name="txtConfigurationHeader" Text="CONFIGURATION" FontSize="11" FontWeight="Bold"
                                           Foreground="{StaticResource TextMuted}" Margin="0,15,0,10"/>
                                
                                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                    <Button x:Name="btnExportConfig" Content="Export Config" 
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnImportConfig" Content="Import Config" 
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="btnResetConfig" Content="Reset to Defaults" 
                                            Style="{StaticResource SecondaryButton}"/>
                                </StackPanel>

                                <!-- Profile Target -->
                                <TextBlock x:Name="txtProfileTargetHeader" Text="PROFILE TARGET" FontSize="11" FontWeight="Bold"
                                           Foreground="{StaticResource TextMuted}" Margin="0,15,0,10"/>
                                <TextBlock x:Name="txtProfileTargetHint"
                                           Text="Choose a local or offline Windows profile root. Selected/default profile changes require Administrator rights."
                                           Foreground="{StaticResource TextSecondary}" TextWrapping="Wrap" Margin="0,0,0,8"/>
                                <Grid Margin="0,0,0,8">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="txtProfileRoot"/>
                                    <Button x:Name="btnBrowseProfileRoot" Grid.Column="1" Content="Browse"
                                            Style="{StaticResource SmallButton}" Margin="5,0,0,0"/>
                                </Grid>
                                <Button x:Name="btnUseDefaultProfileRoot" Content="Use Default User Profile"
                                        Style="{StaticResource SecondaryButton}" HorizontalAlignment="Left"/>
                            </StackPanel>
                        </ScrollViewer>
                    </TabItem>
                    
                    <!-- Log Tab -->
                    <TabItem x:Name="tabLog" Header="Log">
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
$script:CurrentOperationPlan = $null
$script:ActiveWorker = $null
$script:IsLoadingConfiguration = $false
$script:SessionId = [System.Guid]::NewGuid().ToString('N')
$script:LogFilePath = $null
$script:UnhandledExceptionHandlersRegistered = $false
$script:UiStrings = [ordered]@{}
$script:DefaultUiStrings = [ordered]@{
    'Window.Title' = 'Start Menu Organizer v{0}'
    'txtAppTitle.Text' = 'Start Menu Organizer Pro'
    'txtSearchPlaceholder.Text' = 'Search items... (Ctrl+F)'
    'btnCancelWork.Content' = 'Cancel'
    'btnUndo.Content' = 'Undo'
    'btnUndo.ToolTip' = 'Undo the last reversible operation'
    'btnBackup.Content' = 'Backup'
    'btnRestore.Content' = 'Restore'
    'lblScope.Text' = 'Scope:'
    'cmbScopeUser.Content' = 'User Start Menu'
    'cmbScopeSystem.Content' = 'System Start Menu'
    'cmbScopeBoth.Content' = 'Both'
    'cmbScopeProfile.Content' = 'Selected Profile'
    'cmbScopeDefaultUser.Content' = 'Default User'
    'btnRefresh.Content' = 'Refresh'
    'btnRefresh.ToolTip' = 'Refresh the Start Menu scan'
    'lblSort.Text' = 'Sort:'
    'cmbSortName.Content' = 'Name'
    'cmbSortType.Content' = 'Type'
    'cmbSortStatus.Content' = 'Status'
    'cmbSortLocation.Content' = 'Location'
    'cmbSortTarget.Content' = 'Target'
    'chkPreviewMode.Content' = 'Preview Mode'
    'chkPreviewMode.ToolTip' = 'Show what actions would do without executing'
    'lblFilter.Text' = 'Filter:'
    'chkShowShortcuts.Content' = 'Shortcuts'
    'chkShowFolders.Content' = 'Folders'
    'chkShowJunk.Content' = 'Junk'
    'chkShowBroken.Content' = 'Broken'
    'chkShowDuplicates.Content' = 'Duplicates'
    'ctxDelete.Header' = 'Delete'
    'ctxRename.Header' = 'Rename'
    'ctxOpenLocation.Header' = 'Open File Location'
    'ctxOpenTarget.Header' = 'Open Target Location'
    'ctxMoveToCategory.Header' = 'Move to Category'
    'ctxCatDev.Header' = 'Development'
    'ctxCatBrowsers.Header' = 'Browsers'
    'ctxCatComm.Header' = 'Communication'
    'ctxCatMedia.Header' = 'Media'
    'ctxCatGraphics.Header' = 'Graphics'
    'ctxCatOffice.Header' = 'Office'
    'ctxCatUtils.Header' = 'Utilities'
    'ctxCatGaming.Header' = 'Gaming'
    'ctxCatSystem.Header' = 'System'
    'ctxCatSecurity.Header' = 'Security'
    'ctxCatNetwork.Header' = 'Networking'
    'ctxSelectAll.Header' = 'Select All'
    'ctxSelectNone.Header' = 'Select None'
    'colName.Header' = 'Name'
    'colType.Header' = 'Type'
    'colStatus.Header' = 'Status'
    'colLocation.Header' = 'Location'
    'colTarget.Header' = 'Target'
    'btnSelectAll.Content' = 'All'
    'btnSelectNone.Content' = 'None'
    'btnSelectJunk.Content' = 'Junk'
    'btnSelectBroken.Content' = 'Broken'
    'btnSelectDuplicates.Content' = 'Duplicates'
    'btnSelectFolders.Content' = 'Folders'
    'btnInvertSelection.Content' = 'Invert'
    'tabActions.Header' = 'Actions'
    'txtCleanupHeader.Text' = 'CLEANUP'
    'btnDeleteSelected.Content' = 'Delete Selected'
    'btnRemoveAllJunk.Content' = 'Remove All Junk'
    'btnRemoveBroken.Content' = 'Remove Broken Shortcuts'
    'btnRemoveDuplicates.Content' = 'Remove Duplicates'
    'btnFlattenFolders.Content' = 'Flatten Single-Item Folders'
    'btnRemoveEmpty.Content' = 'Remove Empty Folders'
    'btnMoveAllToRoot.Content' = 'Move All to Root'
    'btnMoveAllToRoot.ToolTip' = 'Move all shortcuts from folders to the Start Menu root'
    'txtOrganizeHeader.Text' = 'ORGANIZE'
    'btnMoveToCategory.Content' = 'Move'
    'btnAutoOrganize.Content' = 'Auto-Organize All'
    'txtBatchRenameHeader.Text' = 'BATCH RENAME'
    'btnStripVersions.Content' = 'Strip Version Numbers'
    'btnCleanNames.Content' = 'Clean Up Names'
    'txtFindPlaceholder.Text' = 'Find:'
    'txtReplacePlaceholder.Text' = 'Replace:'
    'btnFindReplace.Content' = 'Find & Replace in Names'
    'txtTransactionPlanHeader.Text' = 'TRANSACTION PLAN'
    'txtPlanStatus.Text' = 'No plan loaded'
    'btnExportPlan.Content' = 'Export Plan'
    'btnImportPlan.Content' = 'Import Plan'
    'btnExecutePlan.Content' = 'Execute Plan'
    'btnClearPlan.Content' = 'Clear Plan'
    'txtQuickOpenHeader.Text' = 'QUICK OPEN'
    'btnOpenUserMenu.Content' = 'User Menu'
    'btnOpenSystemMenu.Content' = 'System Menu'
    'btnOpenBackups.Content' = 'Backups'
    'tabSettings.Header' = 'Settings'
    'txtJunkPatternsHeader.Text' = 'JUNK PATTERNS'
    'btnAddJunkPattern.Content' = 'Add'
    'btnRemoveJunkPattern.Content' = 'Remove'
    'txtCategoriesHeader.Text' = 'CATEGORIES'
    'btnAddCategoryPattern.Content' = 'Add'
    'btnRemoveCategoryPattern.Content' = 'Remove'
    'txtConfigurationHeader.Text' = 'CONFIGURATION'
    'btnExportConfig.Content' = 'Export Config'
    'btnImportConfig.Content' = 'Import Config'
    'btnResetConfig.Content' = 'Reset to Defaults'
    'txtProfileTargetHeader.Text' = 'PROFILE TARGET'
    'txtProfileTargetHint.Text' = 'Choose a local or offline Windows profile root. Selected/default profile changes require Administrator rights.'
    'btnBrowseProfileRoot.Content' = 'Browse'
    'btnUseDefaultProfileRoot.Content' = 'Use Default User Profile'
    'tabLog.Header' = 'Log'
    'btnClearLog.Content' = 'Clear Log'
    'Status.AdminFullAccess' = 'Running as Administrator - Full access to User and System Start Menu'
    'Status.StandardUser' = 'Standard User - Limited to User Start Menu (Run as Admin for full access)'
    'txtSearch.AutomationName' = 'Search Start Menu items'
    'txtSearch.HelpText' = 'Filters the item list by name, path, or target.'
    'btnCancelWork.AutomationName' = 'Cancel active work'
    'btnCancelWork.HelpText' = 'Stops the active scan or operation plan worker.'
    'btnUndo.AutomationName' = 'Undo last action'
    'btnUndo.HelpText' = 'Restores the latest reversible operation from the undo journal.'
    'btnBackup.AutomationName' = 'Create backup'
    'btnBackup.HelpText' = 'Creates a timestamped backup of selected Start Menu scopes.'
    'btnRestore.AutomationName' = 'Restore backup'
    'btnRestore.HelpText' = 'Restores a previous Start Menu backup after staging and validation.'
    'cmbScope.AutomationName' = 'Start Menu scope'
    'cmbScope.HelpText' = 'Selects whether scans target the user Start Menu, system Start Menu, both, a selected profile, or the default user profile.'
    'btnRefresh.AutomationName' = 'Refresh items'
    'btnRefresh.HelpText' = 'Scans the selected Start Menu scope again.'
    'cmbSort.AutomationName' = 'Sort items'
    'cmbSort.HelpText' = 'Selects the column used to sort the item list.'
    'chkPreviewMode.AutomationName' = 'Preview mode'
    'chkPreviewMode.HelpText' = 'Builds a reviewed operation plan instead of changing files immediately.'
    'dgItems.AutomationName' = 'Start Menu items'
    'dgItems.HelpText' = 'Displays detected shortcuts and folders with status, location, and target path.'
    'btnDeleteSelected.AutomationName' = 'Delete selected items'
    'btnRemoveAllJunk.AutomationName' = 'Remove all junk items'
    'btnRemoveBroken.AutomationName' = 'Remove broken shortcuts'
    'btnRemoveDuplicates.AutomationName' = 'Remove duplicate shortcuts'
    'btnFlattenFolders.AutomationName' = 'Flatten single item folders'
    'btnRemoveEmpty.AutomationName' = 'Remove empty folders'
    'btnMoveAllToRoot.AutomationName' = 'Move all shortcuts to root'
    'cmbCategory.AutomationName' = 'Target category'
    'cmbCategory.HelpText' = 'Selects the category folder for selected shortcuts.'
    'btnMoveToCategory.AutomationName' = 'Move selected items to category'
    'btnAutoOrganize.AutomationName' = 'Auto organize all items'
    'btnStripVersions.AutomationName' = 'Strip version numbers'
    'btnCleanNames.AutomationName' = 'Clean names'
    'txtFindText.AutomationName' = 'Find text'
    'txtFindText.HelpText' = 'Text to find in selected shortcut names.'
    'txtReplaceText.AutomationName' = 'Replacement text'
    'txtReplaceText.HelpText' = 'Replacement text for matching shortcut names.'
    'btnFindReplace.AutomationName' = 'Find and replace in names'
    'btnExportPlan.AutomationName' = 'Export operation plan'
    'btnImportPlan.AutomationName' = 'Import operation plan'
    'btnExecutePlan.AutomationName' = 'Execute operation plan'
    'btnClearPlan.AutomationName' = 'Clear operation plan'
    'lstJunkPatterns.AutomationName' = 'Junk patterns'
    'txtNewJunkPattern.AutomationName' = 'New junk pattern'
    'btnAddJunkPattern.AutomationName' = 'Add junk pattern'
    'btnRemoveJunkPattern.AutomationName' = 'Remove junk pattern'
    'cmbEditCategory.AutomationName' = 'Category to edit'
    'lstCategoryPatterns.AutomationName' = 'Category patterns'
    'txtNewCategoryPattern.AutomationName' = 'New category pattern'
    'btnAddCategoryPattern.AutomationName' = 'Add category pattern'
    'btnRemoveCategoryPattern.AutomationName' = 'Remove category pattern'
    'btnExportConfig.AutomationName' = 'Export configuration'
    'btnImportConfig.AutomationName' = 'Import configuration'
    'btnResetConfig.AutomationName' = 'Reset configuration to defaults'
    'txtProfileRoot.AutomationName' = 'Profile root'
    'txtProfileRoot.HelpText' = 'Local or offline Windows profile root. The app resolves its Start Menu Programs folder after validation.'
    'btnBrowseProfileRoot.AutomationName' = 'Browse for profile root'
    'btnBrowseProfileRoot.HelpText' = 'Select a local or offline Windows profile folder.'
    'btnUseDefaultProfileRoot.AutomationName' = 'Use default user profile'
    'btnUseDefaultProfileRoot.HelpText' = 'Fills the profile root field with the Windows Default user profile path.'
    'txtLog.AutomationName' = 'Activity log'
    'txtLog.HelpText' = 'Shows session activity; persistent logs are written to the local app data logs folder.'
    'btnClearLog.AutomationName' = 'Clear activity log'
}
$script:UiTabOrder = @(
    'txtSearch',
    'btnCancelWork',
    'btnUndo',
    'btnBackup',
    'btnRestore',
    'cmbScope',
    'btnRefresh',
    'cmbSort',
    'chkPreviewMode',
    'chkShowShortcuts',
    'chkShowFolders',
    'chkShowJunk',
    'chkShowBroken',
    'chkShowDuplicates',
    'dgItems',
    'btnSelectAll',
    'btnSelectNone',
    'btnSelectJunk',
    'btnSelectBroken',
    'btnSelectDuplicates',
    'btnSelectFolders',
    'btnInvertSelection',
    'btnDeleteSelected',
    'btnRemoveAllJunk',
    'btnRemoveBroken',
    'btnRemoveDuplicates',
    'btnFlattenFolders',
    'btnRemoveEmpty',
    'btnMoveAllToRoot',
    'cmbCategory',
    'btnMoveToCategory',
    'btnAutoOrganize',
    'btnStripVersions',
    'btnCleanNames',
    'txtFindText',
    'txtReplaceText',
    'btnFindReplace',
    'btnExportPlan',
    'btnImportPlan',
    'btnExecutePlan',
    'btnClearPlan',
    'btnOpenUserMenu',
    'btnOpenSystemMenu',
    'btnOpenBackups',
    'lstJunkPatterns',
    'txtNewJunkPattern',
    'btnAddJunkPattern',
    'btnRemoveJunkPattern',
    'cmbEditCategory',
    'lstCategoryPatterns',
    'txtNewCategoryPattern',
    'btnAddCategoryPattern',
    'btnRemoveCategoryPattern',
    'btnExportConfig',
    'btnImportConfig',
    'btnResetConfig',
    'txtProfileRoot',
    'btnBrowseProfileRoot',
    'btnUseDefaultProfileRoot',
    'btnClearLog'
)
$script:ThemeContrastPairs = @(
    @{ Name = 'Primary text on primary background'; Foreground = '#e6edf3'; Background = '#0d1117'; Minimum = 4.5 },
    @{ Name = 'Secondary text on secondary background'; Foreground = '#8b949e'; Background = '#161b22'; Minimum = 4.5 },
    @{ Name = 'Muted text on secondary background'; Foreground = '#9da7b3'; Background = '#161b22'; Minimum = 4.5 },
    @{ Name = 'White text on accent button'; Foreground = '#ffffff'; Background = '#238636'; Minimum = 4.5 },
    @{ Name = 'White text on danger button'; Foreground = '#ffffff'; Background = '#da3633'; Minimum = 4.5 }
)

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

function Remove-OldLogFiles {
    param(
        [string]$LogRoot,
        [string]$Filter,
        [int]$MaxLogFiles
    )

    if ($MaxLogFiles -lt 1 -or -not (Test-Path -LiteralPath $LogRoot)) { return }

    $oldFiles = @(Get-ChildItem -LiteralPath $LogRoot -Filter $Filter -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -Skip $MaxLogFiles)
    foreach ($file in $oldFiles) {
        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
    }
}

function Initialize-FileLogging {
    param(
        [string]$LogRoot = $Config.LogRoot,
        [int]$MaxLogFiles = $Config.MaxLogFiles
    )

    [System.IO.Directory]::CreateDirectory($LogRoot) | Out-Null
    $script:LogFilePath = Join-Path $LogRoot "StartMenuOrganizer-$(Get-Date -Format 'yyyyMMdd').jsonl"
    if (-not (Test-Path -LiteralPath $script:LogFilePath)) {
        [System.IO.File]::WriteAllText($script:LogFilePath, '', [System.Text.UTF8Encoding]::new($false))
    }
    Remove-OldLogFiles -LogRoot $LogRoot -Filter 'StartMenuOrganizer-*.jsonl' -MaxLogFiles $MaxLogFiles
    Remove-OldLogFiles -LogRoot $LogRoot -Filter 'StartMenuOrganizer-Crash-*.log' -MaxLogFiles $MaxLogFiles
    return $script:LogFilePath
}

function Get-ExceptionDetails {
    param([object]$Exception)

    $exceptionObject = $Exception
    $scriptStackTrace = $null
    if ($Exception -is [System.Management.Automation.ErrorRecord]) {
        $exceptionObject = $Exception.Exception
        $scriptStackTrace = $Exception.ScriptStackTrace
    }

    if (-not $exceptionObject) {
        return [ordered]@{
            Type = 'Unknown'
            Message = ''
            StackTrace = ''
            ScriptStackTrace = $scriptStackTrace
        }
    }

    return [ordered]@{
        Type = $exceptionObject.GetType().FullName
        Message = $exceptionObject.Message
        StackTrace = $exceptionObject.StackTrace
        ScriptStackTrace = $scriptStackTrace
    }
}

function Write-StructuredLogEntry {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Level = 'Info',
        [string]$OperationId = $null,
        [object]$Exception = $null,
        [string]$CrashLogPath = $null
    )

    try {
        if (-not $script:LogFilePath) {
            Initialize-FileLogging | Out-Null
        }

        $entry = [ordered]@{
            Timestamp = (Get-Date).ToString('o')
            Level = $Level
            Message = $Message
            OperationId = $OperationId
            SessionId = $script:SessionId
            Version = $Config.Version
            IsAdmin = [bool]$script:IsAdmin
            ProcessId = $PID
        }

        if ($CrashLogPath) {
            $entry.CrashLogPath = $CrashLogPath
        }
        if ($Exception) {
            $entry.Exception = Get-ExceptionDetails -Exception $Exception
        }

        $json = ($entry | ConvertTo-Json -Depth 8 -Compress)
        [System.IO.File]::AppendAllText($script:LogFilePath, "$json`r`n", [System.Text.UTF8Encoding]::new($false))
    }
    catch {
        Write-Verbose "File logging failed: $($_.Exception.Message)"
    }
}

function Write-CrashLog {
    param(
        [object]$Exception,
        [string]$Context = 'Unhandled exception',
        [string]$OperationId = $null
    )

    try {
        if (-not $script:LogFilePath) {
            Initialize-FileLogging | Out-Null
        }

        $details = Get-ExceptionDetails -Exception $Exception
        $crashPath = Join-Path $Config.LogRoot "StartMenuOrganizer-Crash-$(Get-Date -Format 'yyyyMMddHHmmssfff').log"
        $lines = @(
            "Context: $Context",
            "Timestamp: $((Get-Date).ToString('o'))",
            "Version: $($Config.Version)",
            "SessionId: $script:SessionId",
            "OperationId: $OperationId",
            "IsAdmin: $([bool]$script:IsAdmin)",
            "ProcessId: $PID",
            '',
            "ExceptionType: $($details.Type)",
            "Message: $($details.Message)",
            '',
            'StackTrace:',
            "$($details.StackTrace)",
            '',
            'ScriptStackTrace:',
            "$($details.ScriptStackTrace)"
        )
        [System.IO.File]::WriteAllLines($crashPath, $lines, [System.Text.UTF8Encoding]::new($false))
        Write-StructuredLogEntry -Message "$Context. Crash log: $crashPath" -Level 'Error' -OperationId $OperationId -Exception $Exception -CrashLogPath $crashPath
        return $crashPath
    }
    catch {
        return $null
    }
}

function Register-UnhandledExceptionHandlers {
    if ($script:UnhandledExceptionHandlersRegistered) { return }

    [System.AppDomain]::CurrentDomain.add_UnhandledException({
        param($eventSource, $eventData)
        $null = $eventSource
        Write-CrashLog -Exception $eventData.ExceptionObject -Context 'Unhandled AppDomain exception' | Out-Null
    })

    if ($Window -and $Window.Dispatcher) {
        $Window.Dispatcher.add_UnhandledException({
            param($eventSource, $eventData)
            $null = $eventSource
            $crashPath = Write-CrashLog -Exception $eventData.Exception -Context 'Unhandled WPF dispatcher exception'
            $message = "An unexpected error occurred."
            if ($crashPath) {
                $message = "$message`n`nCrash log:`n$crashPath"
            }
            [System.Windows.MessageBox]::Show($message, "Start Menu Organizer Error", "OK", "Error") | Out-Null
            $eventData.Handled = $true
        })
    }

    $script:UnhandledExceptionHandlersRegistered = $true
}

function Import-UiStringFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return }

    try {
        $overrides = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
        foreach ($property in $overrides.PSObject.Properties) {
            if ($null -ne $property.Value) {
                $script:UiStrings[$property.Name] = [string]$property.Value
            }
        }
        Write-StructuredLogEntry -Message "Loaded UI strings: $Path" -Level 'Info'
    }
    catch {
        Write-StructuredLogEntry -Message "Failed to load UI strings: $Path" -Level 'Warning' -Exception $_
    }
}

function Load-UiStrings {
    $script:UiStrings = [ordered]@{}
    foreach ($entry in $script:DefaultUiStrings.GetEnumerator()) {
        $script:UiStrings[$entry.Key] = $entry.Value
    }

    if ([string]::IsNullOrWhiteSpace($Config.LocalizationRoot)) { return }

    $culture = [System.Globalization.CultureInfo]::CurrentUICulture
    $neutralCulture = $culture.TwoLetterISOLanguageName
    $candidateNames = @(
        "strings.$($culture.Name).json",
        "strings.$neutralCulture.json",
        'strings.json'
    ) | Select-Object -Unique

    foreach ($candidateName in $candidateNames) {
        Import-UiStringFile -Path (Join-Path $Config.LocalizationRoot $candidateName)
    }
}

function Get-UiString {
    param(
        [string]$Key,
        [object[]]$FormatArgs = @()
    )

    $value = $null
    if ($script:UiStrings.Contains($Key)) {
        $value = $script:UiStrings[$Key]
    }
    elseif ($script:DefaultUiStrings.Contains($Key)) {
        $value = $script:DefaultUiStrings[$Key]
    }
    else {
        return $Key
    }

    if ($FormatArgs.Count -gt 0) {
        return [string]::Format($value, $FormatArgs)
    }

    return $value
}

function Get-NamedUiControl {
    param([string]$Name)

    $variable = Get-Variable -Name $Name -Scope Script -ErrorAction SilentlyContinue
    if ($variable -and $variable.Value) {
        return $variable.Value
    }

    if ($dgItems) {
        $columnMap = @{
            colName = 1
            colType = 2
            colStatus = 3
            colLocation = 4
            colTarget = 5
        }
        if ($columnMap.ContainsKey($Name) -and $dgItems.Columns.Count -gt $columnMap[$Name]) {
            return $dgItems.Columns[$columnMap[$Name]]
        }
    }

    return $null
}

function Set-NamedControlProperty {
    param(
        [string]$ControlName,
        [string]$PropertyName,
        [string]$Value
    )

    $control = Get-NamedUiControl -Name $ControlName
    if (-not $control) { return $false }

    try {
        switch ($PropertyName) {
            'Content' { $control.Content = $Value }
            'Header' { $control.Header = $Value }
            'Text' { $control.Text = $Value }
            'ToolTip' { $control.ToolTip = $Value }
            default { return $false }
        }
        return $true
    }
    catch {
        Write-StructuredLogEntry -Message "Failed to set $ControlName.$PropertyName" -Level 'Warning' -Exception $_
        return $false
    }
}

function Apply-UiText {
    $Window.Title = Get-UiString -Key 'Window.Title' -FormatArgs @($Config.Version)

    foreach ($entry in $script:UiStrings.GetEnumerator()) {
        if ($entry.Key -match '^(?<ControlName>[^.]+)\.(?<PropertyName>Content|Header|Text|ToolTip)$') {
            Set-NamedControlProperty -ControlName $Matches.ControlName -PropertyName $Matches.PropertyName -Value $entry.Value | Out-Null
        }
    }
}

function Apply-UiAccessibilityMetadata {
    foreach ($entry in $script:UiStrings.GetEnumerator()) {
        if ($entry.Key -notmatch '^(?<ControlName>[^.]+)\.(?<PropertyName>AutomationName|HelpText)$') {
            continue
        }

        $control = Get-NamedUiControl -Name $Matches.ControlName
        if (-not $control) { continue }

        if ($Matches.PropertyName -eq 'AutomationName') {
            [System.Windows.Automation.AutomationProperties]::SetName($control, $entry.Value)
        }
        elseif ($Matches.PropertyName -eq 'HelpText') {
            [System.Windows.Automation.AutomationProperties]::SetHelpText($control, $entry.Value)
            if ($control.PSObject.Properties.Name -contains 'ToolTip' -and -not $control.ToolTip) {
                $control.ToolTip = $entry.Value
            }
        }
    }
}

function Set-UiTabOrder {
    $tabIndex = 0
    foreach ($controlName in $script:UiTabOrder) {
        $control = Get-NamedUiControl -Name $controlName
        if (-not $control) { continue }

        try {
            $control.TabIndex = $tabIndex
            if ($control.PSObject.Properties.Name -contains 'IsTabStop') {
                $control.IsTabStop = $true
            }
            $tabIndex++
        }
        catch {
            Write-StructuredLogEntry -Message "Failed to set tab order for $controlName" -Level 'Warning' -Exception $_
        }
    }
}

function Initialize-UiLocalizationAndAccessibility {
    Load-UiStrings
    Apply-UiText
    Apply-UiAccessibilityMetadata
    Set-UiTabOrder
}

function Convert-HexColorToRgb {
    param([string]$Color)

    $hex = $Color.TrimStart('#')
    if ($hex.Length -ne 6) {
        throw "Expected #RRGGBB color: $Color"
    }

    return [PSCustomObject]@{
        R = [Convert]::ToInt32($hex.Substring(0, 2), 16)
        G = [Convert]::ToInt32($hex.Substring(2, 2), 16)
        B = [Convert]::ToInt32($hex.Substring(4, 2), 16)
    }
}

function Convert-SrgbChannelToLinear {
    param([double]$Channel)

    $normalized = $Channel / 255
    if ($normalized -le 0.03928) {
        return $normalized / 12.92
    }

    return [Math]::Pow((($normalized + 0.055) / 1.055), 2.4)
}

function Get-RelativeLuminance {
    param([string]$Color)

    $rgb = Convert-HexColorToRgb -Color $Color
    $red = Convert-SrgbChannelToLinear -Channel $rgb.R
    $green = Convert-SrgbChannelToLinear -Channel $rgb.G
    $blue = Convert-SrgbChannelToLinear -Channel $rgb.B
    return (0.2126 * $red) + (0.7152 * $green) + (0.0722 * $blue)
}

function Get-ContrastRatio {
    param(
        [string]$Foreground,
        [string]$Background
    )

    $foregroundLuminance = Get-RelativeLuminance -Color $Foreground
    $backgroundLuminance = Get-RelativeLuminance -Color $Background
    $lighter = [Math]::Max($foregroundLuminance, $backgroundLuminance)
    $darker = [Math]::Min($foregroundLuminance, $backgroundLuminance)
    return [Math]::Round((($lighter + 0.05) / ($darker + 0.05)), 2)
}

function Test-ThemeContrast {
    $script:ThemeContrastPairs | ForEach-Object {
        $ratio = Get-ContrastRatio -Foreground $_.Foreground -Background $_.Background
        [PSCustomObject]@{
            Name = $_.Name
            Ratio = $ratio
            Minimum = [double]$_.Minimum
            Pass = ($ratio -ge [double]$_.Minimum)
        }
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Level = 'Info',
        [string]$OperationId = $null
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $color = switch ($Level) {
        'Success' { '#2ea043' }
        'Warning' { '#d29922' }
        'Error'   { '#f85149' }
        default   { '#8b949e' }
    }

    Write-StructuredLogEntry -Message $Message -Level $Level -OperationId $OperationId
    
    if ($Window -and $txtLog -and $svLog) {
        $Window.Dispatcher.Invoke([Action]{
            $run = [System.Windows.Documents.Run]::new("[$timestamp] $Message`r`n")
            $run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($color)
            $txtLog.Inlines.Add($run)
            $svLog.ScrollToEnd()
        })
    }
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

function Set-WorkerUiState {
    param(
        [bool]$IsBusy,
        [string]$Message = 'Ready'
    )

    $Window.Dispatcher.Invoke([Action]{
        $btnCancelWork.IsEnabled = $IsBusy
        $btnRefresh.IsEnabled = -not $IsBusy
        $txtStatus.Text = $Message
    })
}

function Start-BackgroundWorker {
    param(
        [string]$Name,
        [scriptblock]$ScriptBlock,
        [object[]]$Arguments,
        [scriptblock]$OnCompleted
    )

    if ($script:ActiveWorker) {
        Write-Log "Worker already running: $($script:ActiveWorker.Name)" 'Warning'
        return $false
    }

    $powerShell = [PowerShell]::Create()
    [void]$powerShell.AddScript($ScriptBlock)
    foreach ($argument in $Arguments) {
        [void]$powerShell.AddArgument($argument)
    }

    $asyncResult = $powerShell.BeginInvoke()
    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(150)

    $worker = [PSCustomObject]@{
        Name = $Name
        PowerShell = $powerShell
        AsyncResult = $asyncResult
        Timer = $timer
        OnCompleted = $OnCompleted
    }
    $script:ActiveWorker = $worker

    $timer.Add_Tick({
        if (-not $script:ActiveWorker) { return }
        if (-not $script:ActiveWorker.AsyncResult.IsCompleted) { return }

        $completedWorker = $script:ActiveWorker
        $script:ActiveWorker = $null
        $completedWorker.Timer.Stop()

        try {
            $output = $completedWorker.PowerShell.EndInvoke($completedWorker.AsyncResult)
            & $completedWorker.OnCompleted $output
        }
        catch {
            Write-Log "$($completedWorker.Name) failed: $($_.Exception.Message)" 'Error'
            Update-Status "Ready"
        }
        finally {
            $completedWorker.PowerShell.Dispose()
            Set-WorkerUiState -IsBusy $false
            Show-Progress -Visible $false
        }
    })

    Set-WorkerUiState -IsBusy $true -Message $Name
    $timer.Start()
    return $true
}

function Stop-BackgroundWorker {
    if (-not $script:ActiveWorker) { return }

    $worker = $script:ActiveWorker
    $script:ActiveWorker = $null
    $worker.Timer.Stop()

    try {
        if ($worker.PowerShell) {
            $worker.PowerShell.Stop()
        }
        if ($worker.PSObject.Properties.Name -contains 'OnCanceled' -and $worker.OnCanceled) {
            & $worker.OnCanceled
        }
    }
    catch {
        Write-Log "Failed to stop $($worker.Name): $($_.Exception.Message)" 'Error'
    }
    finally {
        if ($worker.PowerShell) {
            $worker.PowerShell.Dispose()
        }
        Set-WorkerUiState -IsBusy $false
        Show-Progress -Visible $false
        Write-Log "$($worker.Name) canceled." 'Warning'
    }
}

function Get-DefaultUserProfileRoot {
    if (-not [string]::IsNullOrWhiteSpace($Config.DefaultProfileRoot)) {
        return $Config.DefaultProfileRoot
    }

    return (Join-Path $env:SystemDrive 'Users\Default')
}

function Get-ProfileProgramsPath {
    param([string]$ProfileRootOrProgramsPath)

    if ([string]::IsNullOrWhiteSpace($ProfileRootOrProgramsPath)) {
        throw "Profile root is required."
    }

    $candidate = Get-NormalizedPath ([Environment]::ExpandEnvironmentVariables($ProfileRootOrProgramsPath.Trim()))
    if (-not [System.IO.Path]::IsPathRooted($candidate)) {
        throw "Profile path must be an absolute local path: $ProfileRootOrProgramsPath"
    }

    $programsSuffix = 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs'
    if ($candidate.EndsWith($programsSuffix, [System.StringComparison]::OrdinalIgnoreCase)) {
        $profileRoot = $candidate.Substring(0, $candidate.Length - $programsSuffix.Length).TrimEnd('\', '/')
        $programsPath = $candidate
    }
    else {
        $profileRoot = $candidate
        $programsPath = Join-Path $profileRoot $programsSuffix
    }

    Test-ProfileTargetPath -ProfileRoot $profileRoot -ProgramsPath $programsPath | Out-Null

    return [PSCustomObject]@{
        ProfileRoot = $profileRoot
        ProgramsPath = (Get-NormalizedPath $programsPath)
    }
}

function Test-ProfileTargetPath {
    param(
        [string]$ProfileRoot,
        [string]$ProgramsPath
    )

    $profileRootPath = Get-NormalizedPath $ProfileRoot
    $programsRootPath = Get-NormalizedPath $ProgramsPath
    $driveRoot = ([System.IO.Path]::GetPathRoot($profileRootPath)).TrimEnd('\', '/')

    if ($profileRootPath -eq $driveRoot) {
        throw "Profile root cannot be a drive root: $profileRootPath"
    }
    if (-not (Test-Path -LiteralPath $profileRootPath -PathType Container)) {
        throw "Profile root does not exist: $profileRootPath"
    }
    if (-not (Test-PathWithinRoot -Path $programsRootPath -Root $profileRootPath)) {
        throw "Profile Start Menu path must be inside the profile root: $programsRootPath"
    }

    $windowsRoot = if ($env:windir) { Get-NormalizedPath $env:windir } else { $null }
    if ($windowsRoot -and (Test-PathWithinRoot -Path $profileRootPath -Root $windowsRoot)) {
        throw "Profile root cannot be inside Windows: $profileRootPath"
    }

    $programDataRoot = if ($env:ProgramData) { Get-NormalizedPath $env:ProgramData } else { $null }
    if ($programDataRoot -and (Test-PathWithinRoot -Path $profileRootPath -Root $programDataRoot)) {
        throw "Profile root cannot be inside ProgramData: $profileRootPath"
    }

    return $true
}

function New-StartMenuScope {
    param(
        [string]$ScopeName,
        [string]$Path,
        [string]$Prefix,
        [bool]$RequiresAdmin = $false,
        [string]$BackupName = $ScopeName,
        [string]$ProfileRoot = $null
    )

    return [PSCustomObject]@{
        ScopeName = $ScopeName
        Path = (Get-NormalizedPath $Path)
        Prefix = $Prefix
        RequiresAdmin = $RequiresAdmin
        BackupName = $BackupName
        ProfileRoot = $ProfileRoot
    }
}

function Get-ConfiguredProfileScope {
    $profileTarget = Get-ProfileProgramsPath -ProfileRootOrProgramsPath $Config.ProfileRoot
    return New-StartMenuScope -ScopeName 'Profile' -Path $profileTarget.ProgramsPath -Prefix '[Profile]' -RequiresAdmin $true -BackupName 'Profile' -ProfileRoot $profileTarget.ProfileRoot
}

function Get-DefaultUserScope {
    $profileTarget = Get-ProfileProgramsPath -ProfileRootOrProgramsPath (Get-DefaultUserProfileRoot)
    return New-StartMenuScope -ScopeName 'DefaultUser' -Path $profileTarget.ProgramsPath -Prefix '[Default]' -RequiresAdmin $true -BackupName 'DefaultUser' -ProfileRoot $profileTarget.ProfileRoot
}

function Get-StartMenuScopes {
    $scope = if ($cmbScope) { $cmbScope.SelectedIndex } else { 2 }
    $scopes = @()

    if ($scope -eq 0 -or $scope -eq 2) {
        $scopes += New-StartMenuScope -ScopeName 'User' -Path $Config.UserStartMenu -Prefix '[User]' -RequiresAdmin $false -BackupName 'User'
    }
    if ($scope -eq 1 -or $scope -eq 2) {
        $scopes += New-StartMenuScope -ScopeName 'System' -Path $Config.SystemStartMenu -Prefix '[Sys]' -RequiresAdmin $true -BackupName 'System'
    }
    if ($scope -eq 3) {
        $scopes += Get-ConfiguredProfileScope
    }
    if ($scope -eq 4) {
        $scopes += Get-DefaultUserScope
    }

    return $scopes
}

function Get-StartMenuPaths {
    return @((Get-StartMenuScopes) | ForEach-Object { $_.Path })
}

function Get-ApprovedMutationRoots {
    $roots = @($Config.UserStartMenu, $Config.SystemStartMenu)

    foreach ($scopeResolver in @({ Get-ConfiguredProfileScope }, { Get-DefaultUserScope })) {
        try {
            $scope = & $scopeResolver
            if ($scope -and $scope.Path) {
                $roots += $scope.Path
            }
        }
        catch {
            Write-Verbose "Skipping invalid mutation root: $($_.Exception.Message)"
        }
    }

    return @($roots | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { Get-NormalizedPath $_ } | Select-Object -Unique)
}

function Test-MutationRootRequiresAdministrator {
    param([string]$Root)

    $rootPath = Get-NormalizedPath $Root
    if ($rootPath -eq (Get-NormalizedPath $Config.SystemStartMenu)) {
        return $true
    }

    foreach ($scopeResolver in @({ Get-ConfiguredProfileScope }, { Get-DefaultUserScope })) {
        try {
            $scope = & $scopeResolver
            if ($scope -and $scope.RequiresAdmin -and $rootPath -eq (Get-NormalizedPath $scope.Path)) {
                return $true
            }
        }
        catch {
            continue
        }
    }

    return $false
}

function Get-ShortcutTarget {
    param([string]$ShortcutPath)

    $extension = [System.IO.Path]::GetExtension($ShortcutPath).ToLowerInvariant()

    try {
        if ($extension -eq '.url') {
            $urlLine = Get-Content -LiteralPath $ShortcutPath -ErrorAction Stop |
                Where-Object { $_ -match '^URL=' } |
                Select-Object -First 1
            if ($urlLine) {
                return ($urlLine -replace '^URL=', '').Trim()
            }
            return $null
        }

        if ($extension -eq '.appref-ms') {
            return $ShortcutPath
        }

        $shortcut = $script:WScriptShell.CreateShortcut($ShortcutPath)
        return $shortcut.TargetPath
    }
    catch {
        return $null
    }
}

function Test-ShortcutBroken {
    param([string]$ShortcutPath)

    if (-not (Test-Path -LiteralPath $ShortcutPath)) { return $true }

    $extension = [System.IO.Path]::GetExtension($ShortcutPath).ToLowerInvariant()
    if ($extension -eq '.appref-ms') { return $false }

    $target = Get-ShortcutTarget $ShortcutPath
    if ([string]::IsNullOrEmpty($target)) { return $false }

    if ($target -match '^(https?|ftp)://') { return $false }
    if ($target -match '^file://') {
        $target = ([System.Uri]$target).LocalPath
    }

    $expandedTarget = [Environment]::ExpandEnvironmentVariables($target)

    # Skip UWP apps, shell links, and Windows-managed targets.
    if ($expandedTarget -match '^[A-Za-z]:\\Windows\\' -or
        $expandedTarget -match '^shell:' -or
        $expandedTarget -match '^ms-[A-Za-z-]+:') {
        return $false
    }

    if ($expandedTarget -match '^[A-Za-z]:\\' -or $expandedTarget.StartsWith('\\')) {
        return -not (Test-Path -LiteralPath $expandedTarget)
    }

    return $false
}

function Get-ShortcutMetadata {
    param(
        [string]$ShortcutPath,
        $Shell = $script:WScriptShell
    )

    $meta = [ordered]@{
        Target       = ''
        Arguments    = ''
        WorkingDir   = ''
        Description  = ''
        Hotkey       = ''
        IconLocation = ''
        WindowStyle  = 0
    }

    $extension = [System.IO.Path]::GetExtension($ShortcutPath).ToLowerInvariant()

    try {
        if ($extension -eq '.lnk') {
            $lnk = $Shell.CreateShortcut($ShortcutPath)
            $meta.Target = $lnk.TargetPath
            $meta.Arguments = $lnk.Arguments
            $meta.WorkingDir = $lnk.WorkingDirectory
            $meta.Description = $lnk.Description
            $meta.Hotkey = $lnk.Hotkey
            $meta.IconLocation = $lnk.IconLocation
            $meta.WindowStyle = $lnk.WindowStyle
        }
        elseif ($extension -eq '.url') {
            $lines = @(Get-Content -LiteralPath $ShortcutPath -ErrorAction Stop)
            foreach ($line in $lines) {
                if ($line -match '^URL=(.*)') { $meta.Target = $Matches[1].Trim() }
                elseif ($line -match '^IconFile=(.*)') { $meta.IconLocation = $Matches[1].Trim() }
                elseif ($line -match '^HotKey=(.*)') { $meta.Hotkey = $Matches[1].Trim() }
                elseif ($line -match '^WorkingDirectory=(.*)') { $meta.WorkingDir = $Matches[1].Trim() }
            }
        }
        elseif ($extension -eq '.appref-ms') {
            $meta.Target = $ShortcutPath
        }
    }
    catch {
        $meta.Target = ''
    }

    return [PSCustomObject]$meta
}

function Get-ShortcutRiskFlags {
    param([PSCustomObject]$Metadata)

    $flags = @()

    $target = if ($Metadata.Target) { $Metadata.Target } else { '' }
    $metaArgs = if ($Metadata.Arguments) { $Metadata.Arguments } else { '' }

    $scriptHosts = @('cmd.exe', 'cmd', 'powershell.exe', 'powershell', 'pwsh.exe', 'pwsh',
                     'wscript.exe', 'wscript', 'cscript.exe', 'cscript', 'mshta.exe', 'mshta',
                     'rundll32.exe', 'rundll32', 'regsvr32.exe', 'regsvr32', 'msiexec.exe', 'msiexec')
    $targetLeaf = [System.IO.Path]::GetFileNameWithoutExtension($target)
    $targetFile = [System.IO.Path]::GetFileName($target)
    if ($scriptHosts -contains $targetLeaf -or $scriptHosts -contains $targetFile) {
        $flags += 'ScriptHost'
    }

    if ($target -match '\\\\' -or $target -match '^\\\\') {
        $flags += 'NetworkTarget'
    }

    if ($metaArgs.Length -gt 260) {
        $flags += 'LongArguments'
    }

    if ($metaArgs -match '(^|\s)-[Ee]nc(odedCommand)?\s' -or $metaArgs -match '-[Ww]indowStyle\s+[Hh]idden') {
        $flags += 'HiddenExecution'
    }

    if ($target -match '^(https?|ftp)://' -or ($target -match '^file://')) {
        $flags += 'WebTarget'
    }

    if ($target -match '\.(bat|cmd|vbs|vbe|js|jse|wsf|wsh|ps1|psm1|psd1)$') {
        $flags += 'ScriptTarget'
    }

    return $flags
}

function Test-ProtectedFolder {
    param(
        [string]$Path,
        [string]$BasePath
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($BasePath)) {
        return $false
    }

    $base = Get-NormalizedPath $BasePath
    $candidate = Get-NormalizedPath $Path
    if (-not $candidate.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    $relative = $candidate.Substring($base.Length).TrimStart('\', '/')
    if ([string]::IsNullOrWhiteSpace($relative)) { return $false }

    $topFolder = $relative.Split([char[]]@('\', '/'), 2)[0]
    return @($script:ProtectedFolders) -contains $topFolder
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
    if ($Item.IsProtected) { $statuses += 'Protected' }
    if ($Item.RiskFlags -and $Item.RiskFlags -ne '') { $statuses += 'Risk' }
    if ($statuses.Count -eq 0) { return 'OK' }
    return $statuses -join ', '
}

function Ensure-JournalStorage {
    $journalDir = Split-Path $Config.UndoFile -Parent
    [System.IO.Directory]::CreateDirectory($journalDir) | Out-Null
    [System.IO.Directory]::CreateDirectory($Config.UndoBackupRoot) | Out-Null
}

function New-OperationId {
    return [System.Guid]::NewGuid().ToString('N')
}

function New-UndoBackupCopy {
    param(
        [string]$SourcePath,
        [string]$OperationId
    )

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        return $null
    }

    Ensure-JournalStorage
    $backupDir = Join-Path $Config.UndoBackupRoot $OperationId
    [System.IO.Directory]::CreateDirectory($backupDir) | Out-Null

    $leafName = Split-Path $SourcePath -Leaf
    if ([string]::IsNullOrWhiteSpace($leafName)) {
        $leafName = 'item'
    }

    $backupPath = Join-Path $backupDir "$([System.Guid]::NewGuid().ToString('N'))_$leafName"
    Copy-Item -LiteralPath $SourcePath -Destination $backupPath -Recurse -Force -ErrorAction Stop
    return $backupPath
}

function Test-PathWithinRoot {
    param(
        [string]$Path,
        [string]$Root
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Root)) {
        return $false
    }

    $fullPath = Get-NormalizedPath $Path
    $rootPath = Get-NormalizedPath $Root
    if ($fullPath -eq $rootPath) { return $true }

    $rootPrefix = $rootPath + [System.IO.Path]::DirectorySeparatorChar
    return $fullPath.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-ReparsePoint {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    try {
        $info = [System.IO.FileInfo]::new($Path)
        if (-not $info.Exists) {
            $dirInfo = [System.IO.DirectoryInfo]::new($Path)
            if (-not $dirInfo.Exists) { return $false }
            return ($dirInfo.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
        }
        return ($info.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
    }
    catch {
        return $false
    }
}

function Test-PathContainsReparsePoint {
    param(
        [string]$Path,
        [string]$Root
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Root)) {
        return $false
    }

    $fullPath = Get-NormalizedPath $Path
    $rootPath = Get-NormalizedPath $Root

    $current = $fullPath
    while ($current.Length -gt $rootPath.Length) {
        if (Test-ReparsePoint $current) { return $true }
        $parent = [System.IO.Path]::GetDirectoryName($current)
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }

    return $false
}

function Get-ApprovedMutationRoot {
    param(
        [string]$Path,
        [string[]]$AllowedRoots
    )

    foreach ($root in $AllowedRoots) {
        if (Test-PathWithinRoot -Path $Path -Root $root) {
            return (Get-NormalizedPath $root)
        }
    }

    return $null
}

function Invoke-GuardedFileOperation {
    param(
        [ValidateSet('Delete','Move','Rename')]
        [string]$Action,
        [string]$SourcePath,
        [string]$DestinationPath = $null,
        [string]$NewName = $null,
        [string[]]$AllowedRoots = (Get-ApprovedMutationRoots),
        [ValidateSet('Fail','Overwrite')]
        [string]$CollisionPolicy = 'Fail',
        [bool]$Preview = $false,
        [switch]$RegisterRollback,
        [string]$OperationId = (New-OperationId)
    )

    $result = [ordered]@{
        Action = $Action
        OperationId = $OperationId
        SourcePath = $SourcePath
        DestinationPath = $DestinationPath
        BackupPath = $null
        Result = 'Failed'
        Preview = $Preview
        Message = ''
        Timestamp = (Get-Date).ToString('o')
    }

    try {
        $sourceRoot = Get-ApprovedMutationRoot -Path $SourcePath -AllowedRoots $AllowedRoots
        if (-not $sourceRoot) {
            throw "Source path is outside approved roots: $SourcePath"
        }
        if ((Test-MutationRootRequiresAdministrator -Root $sourceRoot) -and -not $script:IsAdmin) {
            throw "Administrator privileges required for this Start Menu target: $sourceRoot"
        }
        if (Test-ProtectedFolder -Path $SourcePath -BasePath $sourceRoot) {
            throw "Source path is protected: $SourcePath"
        }
        if (Test-ReparsePoint $SourcePath) {
            throw "Source path is a reparse point (junction/symlink): $SourcePath"
        }
        if (Test-PathContainsReparsePoint -Path $SourcePath -Root $sourceRoot) {
            throw "Source path traverses a reparse point: $SourcePath"
        }

        if ($Action -eq 'Rename') {
            if ([string]::IsNullOrWhiteSpace($NewName)) {
                throw "Rename requires a new name."
            }
            $DestinationPath = Join-Path (Split-Path $SourcePath -Parent) $NewName
            $result.DestinationPath = $DestinationPath
        }

        if ($Action -in @('Move','Rename')) {
            if ([string]::IsNullOrWhiteSpace($DestinationPath)) {
                throw "$Action requires a destination path."
            }

            $destinationRoot = Get-ApprovedMutationRoot -Path $DestinationPath -AllowedRoots $AllowedRoots
            if (-not $destinationRoot) {
                throw "Destination path is outside approved roots: $DestinationPath"
            }
            if ($sourceRoot -ne $destinationRoot) {
                throw "$Action must stay within the same approved root."
            }
            if (Test-ProtectedFolder -Path $DestinationPath -BasePath $destinationRoot) {
                throw "Destination path is protected: $DestinationPath"
            }
            if (Test-PathContainsReparsePoint -Path $DestinationPath -Root $destinationRoot) {
                throw "Destination path traverses a reparse point: $DestinationPath"
            }

            if ((Test-Path -LiteralPath $DestinationPath) -and ((Get-NormalizedPath $SourcePath) -ne (Get-NormalizedPath $DestinationPath)) -and $CollisionPolicy -eq 'Fail') {
                throw "Destination already exists: $DestinationPath"
            }
        }

        if ($Preview) {
            $result.Result = 'Preview'
            $result.Message = 'Preview only; no filesystem change was made.'
            return [PSCustomObject]$result
        }

        if ($Action -eq 'Move') {
            $destinationParent = Split-Path $DestinationPath -Parent
            if (-not (Test-Path -LiteralPath $destinationParent)) {
                [System.IO.Directory]::CreateDirectory($destinationParent) | Out-Null
            }
        }

        if ($RegisterRollback) {
            $result.BackupPath = New-UndoBackupCopy -SourcePath $SourcePath -OperationId $OperationId
        }

        switch ($Action) {
            'Delete' {
                Remove-Item -LiteralPath $SourcePath -Recurse -Force -ErrorAction Stop
            }
            'Move' {
                Move-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force -ErrorAction Stop
            }
            'Rename' {
                Rename-Item -LiteralPath $SourcePath -NewName $NewName -Force -ErrorAction Stop
            }
        }

        $result.Result = 'Success'
        $result.Message = "$Action completed."
        return [PSCustomObject]$result
    }
    catch {
        $result.Message = $_.Exception.Message
        return [PSCustomObject]$result
    }
}

function New-JournalItem {
    param(
        [string]$ActionType,
        [string]$OriginalPath,
        [string]$NewPath = $null,
        [string]$BackupPath = $null,
        [string]$Result = 'Success',
        [string]$Message = '',
        [string]$OperationId = $null
    )

    return [PSCustomObject]@{
        ActionType   = $ActionType
        OperationId  = $OperationId
        OriginalPath = $OriginalPath
        NewPath      = $NewPath
        BackupPath   = $BackupPath
        Timestamp    = (Get-Date).ToString('o')
        Result       = $Result
        Message      = $Message
    }
}

function Write-AtomicFile {
    param(
        [string]$Path,
        [string]$Content
    )

    $dir = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $dir)) {
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
    }

    $tmpPath = "$Path.tmp"
    [System.IO.File]::WriteAllText($tmpPath, $Content, [System.Text.UTF8Encoding]::new($false))

    if (Test-Path -LiteralPath $Path) {
        $bakPath = "$Path.bak"
        Copy-Item -LiteralPath $Path -Destination $bakPath -Force -ErrorAction SilentlyContinue
    }

    Move-Item -LiteralPath $tmpPath -Destination $Path -Force
}

function Read-WithBackupFallback {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        try {
            $raw = Get-Content -LiteralPath $Path -Raw
            $null = $raw | ConvertFrom-Json
            return $raw
        }
        catch {
            $badPath = "$Path.bad.$(Get-Date -Format 'yyyyMMddHHmmss')"
            Move-Item -LiteralPath $Path -Destination $badPath -Force -ErrorAction SilentlyContinue
        }
    }

    $bakPath = "$Path.bak"
    if (Test-Path -LiteralPath $bakPath) {
        try {
            $raw = Get-Content -LiteralPath $bakPath -Raw
            $null = $raw | ConvertFrom-Json
            Copy-Item -LiteralPath $bakPath -Destination $Path -Force
            return $raw
        }
        catch {
            $badBak = "$bakPath.bad.$(Get-Date -Format 'yyyyMMddHHmmss')"
            Move-Item -LiteralPath $bakPath -Destination $badBak -Force -ErrorAction SilentlyContinue
        }
    }

    return $null
}

function Save-UndoJournal {
    Ensure-JournalStorage
    if ($script:UndoStack.Count -eq 0) {
        $json = '[]'
    }
    else {
        $json = @($script:UndoStack) | ConvertTo-Json -Depth 12
    }
    Write-AtomicFile -Path $Config.UndoFile -Content $json
}

function Load-UndoJournal {
    Ensure-JournalStorage
    $script:UndoStack.Clear()

    $raw = Read-WithBackupFallback -Path $Config.UndoFile
    if ($raw) {
        try {
            $loaded = @($raw | ConvertFrom-Json)
            foreach ($entry in $loaded) {
                if ($entry) {
                    $script:UndoStack.Add($entry)
                }
            }

            while ($script:UndoStack.Count -gt $Config.MaxUndoSteps) {
                $script:UndoStack.RemoveAt(0)
            }
        }
        catch {
            Write-Log "Undo journal could not be parsed after recovery attempt." 'Warning'
        }
    }

    $Window.Dispatcher.Invoke([Action]{ $btnUndo.IsEnabled = ($script:UndoStack.Count -gt 0) })
}

function Add-JournalEntry {
    param(
        [string]$Type,
        [string]$Description,
        [array]$Items,
        [string]$OperationId = (New-OperationId)
    )

    $journalItems = @($Items | Where-Object { $_ })
    if ($journalItems.Count -eq 0) { return }

    $hasFailure = @($journalItems | Where-Object { $_.Result -ne 'Success' }).Count -gt 0
    $undoData = [PSCustomObject]@{
        OperationId = $OperationId
        Type = $Type
        ActionType = $Type
        Description = $Description
        Timestamp = (Get-Date).ToString('o')
        Result = if ($hasFailure) { 'Partial' } else { 'Success' }
        Items = $journalItems
    }

    $script:UndoStack.Add($undoData)
    if ($script:UndoStack.Count -gt $Config.MaxUndoSteps) {
        $script:UndoStack.RemoveAt(0)
    }

    Save-UndoJournal
    Write-Log "Journal recorded: $Description" 'Info' -OperationId $OperationId
    $Window.Dispatcher.Invoke([Action]{ $btnUndo.IsEnabled = $true })
}

function Add-UndoAction {
    param(
        [string]$ActionType,
        [string]$Description,
        [array]$Items
    )

    $records = @($Items | ForEach-Object {
        New-JournalItem -ActionType $ActionType -OriginalPath $_.FullPath -BackupPath $null -Result 'Success'
    })
    Add-JournalEntry -Type $ActionType -Description $Description -Items $records
}

function Invoke-JournaledDelete {
    param(
        [string]$Path,
        [string]$OperationId
    )

    $operation = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $Path -OperationId $OperationId -RegisterRollback
    if ($operation.Result -ne 'Success') {
        throw $operation.Message
    }

    return New-JournalItem -ActionType 'Delete' -OriginalPath $Path -BackupPath $operation.BackupPath -Result $operation.Result -Message $operation.Message -OperationId $OperationId
}

function Invoke-JournaledMove {
    param(
        [string]$SourcePath,
        [string]$DestinationPath,
        [string]$OperationId
    )

    $operation = Invoke-GuardedFileOperation -Action 'Move' -SourcePath $SourcePath -DestinationPath $DestinationPath -OperationId $OperationId -RegisterRollback
    if ($operation.Result -ne 'Success') {
        throw $operation.Message
    }

    return New-JournalItem -ActionType 'Move' -OriginalPath $SourcePath -NewPath $DestinationPath -BackupPath $operation.BackupPath -Result $operation.Result -Message $operation.Message -OperationId $OperationId
}

function Invoke-JournaledRename {
    param(
        [string]$SourcePath,
        [string]$NewName,
        [string]$OperationId
    )

    $newPath = Join-Path (Split-Path $SourcePath -Parent) $NewName
    $operation = Invoke-GuardedFileOperation -Action 'Rename' -SourcePath $SourcePath -NewName $NewName -OperationId $OperationId -RegisterRollback
    if ($operation.Result -ne 'Success') {
        throw $operation.Message
    }

    return New-JournalItem -ActionType 'Rename' -OriginalPath $SourcePath -NewPath $newPath -BackupPath $operation.BackupPath -Result $operation.Result -Message $operation.Message -OperationId $OperationId
}

function Restore-JournalBackup {
    param($Item)

    if ($Item.BackupPath -and (Test-Path -LiteralPath $Item.BackupPath)) {
        $destDir = Split-Path $Item.OriginalPath -Parent
        if (-not (Test-Path -LiteralPath $destDir)) {
            [System.IO.Directory]::CreateDirectory($destDir) | Out-Null
        }
        Copy-Item -LiteralPath $Item.BackupPath -Destination $Item.OriginalPath -Recurse -Force -ErrorAction Stop
        return $true
    }

    return $false
}

function Invoke-Undo {
    if ($script:UndoStack.Count -eq 0) { return }

    $lastAction = $script:UndoStack[$script:UndoStack.Count - 1]
    $script:UndoStack.RemoveAt($script:UndoStack.Count - 1)

    Write-Log "Undo: $($lastAction.Description)" 'Warning' -OperationId $lastAction.OperationId

    $items = @($lastAction.Items)
    for ($i = $items.Count - 1; $i -ge 0; $i--) {
        $item = $items[$i]
        if ($item.Result -ne 'Success') { continue }

        try {
            switch ($item.ActionType) {
                'Delete' {
                    Restore-JournalBackup -Item $item | Out-Null
                }
                'Move' {
                    if ($item.NewPath -and (Test-Path -LiteralPath $item.NewPath)) {
                        $operation = Invoke-GuardedFileOperation -Action 'Move' -SourcePath $item.NewPath -DestinationPath $item.OriginalPath -CollisionPolicy 'Overwrite'
                        if ($operation.Result -ne 'Success') { throw $operation.Message }
                    }
                    else {
                        Restore-JournalBackup -Item $item | Out-Null
                    }
                }
                'Rename' {
                    if ($item.NewPath -and (Test-Path -LiteralPath $item.NewPath)) {
                        $operation = Invoke-GuardedFileOperation -Action 'Move' -SourcePath $item.NewPath -DestinationPath $item.OriginalPath -CollisionPolicy 'Overwrite'
                        if ($operation.Result -ne 'Success') { throw $operation.Message }
                    }
                    else {
                        Restore-JournalBackup -Item $item | Out-Null
                    }
                }
                'Restore' {
                    if ($item.BackupPath -and (Test-Path -LiteralPath $item.BackupPath -PathType Container)) {
                        Clear-DirectoryContents -Path $item.OriginalPath
                        Copy-DirectoryContents -SourcePath $item.BackupPath -DestinationPath $item.OriginalPath | Out-Null
                    }
                }
            }
        }
        catch {
            Write-Log "Undo failed for $($item.ActionType) $($item.OriginalPath): $($_.Exception.Message)" 'Error'
        }
    }

    Save-UndoJournal
    if ($script:UndoStack.Count -eq 0) {
        $Window.Dispatcher.Invoke([Action]{ $btnUndo.IsEnabled = $false })
    }

    Refresh-Items
}

function New-PlanOperation {
    param(
        [ValidateSet('Delete','Move','Rename')]
        [string]$Action,
        [string]$SourcePath,
        [string]$DestinationPath = $null,
        [string]$NewName = $null,
        [string]$DisplayName = $null
    )

    return [PSCustomObject]@{
        Action = $Action
        SourcePath = $SourcePath
        DestinationPath = $DestinationPath
        NewName = $NewName
        DisplayName = $DisplayName
        CreatedAt = (Get-Date).ToString('o')
    }
}

function Set-CurrentOperationPlan {
    param(
        [string]$Name,
        [array]$Operations
    )

    $planOperations = @($Operations | Where-Object { $_ })
    $script:CurrentOperationPlan = [PSCustomObject]@{
        PlanId = [System.Guid]::NewGuid().ToString('N')
        Name = $Name
        CreatedAt = (Get-Date).ToString('o')
        Version = $Config.Version
        Operations = $planOperations
        Counts = @{}
    }

    foreach ($group in ($planOperations | Group-Object Action)) {
        $script:CurrentOperationPlan.Counts[$group.Name] = $group.Count
    }

    Update-PlanStatus
    Show-OperationPlanSummary -Plan $script:CurrentOperationPlan
}

function Update-PlanStatus {
    if (-not $txtPlanStatus) { return }

    if ($script:CurrentOperationPlan -and @($script:CurrentOperationPlan.Operations).Count -gt 0) {
        $count = @($script:CurrentOperationPlan.Operations).Count
        $txtPlanStatus.Text = "$($script:CurrentOperationPlan.Name) - $count operation(s) loaded"
        $btnExportPlan.IsEnabled = $true
        $btnExecutePlan.IsEnabled = $true
        $btnClearPlan.IsEnabled = $true
    }
    else {
        $txtPlanStatus.Text = "No plan loaded"
        $btnExportPlan.IsEnabled = $false
        $btnExecutePlan.IsEnabled = $false
        $btnClearPlan.IsEnabled = $false
    }
}

function Show-OperationPlanSummary {
    param($Plan)

    $operations = @($Plan.Operations)
    Write-Log "PLAN: $($Plan.Name) - $($operations.Count) operation(s)" 'Warning'
    foreach ($group in ($operations | Group-Object Action)) {
        Write-Log "  $($group.Name): $($group.Count)" 'Info'
    }

    foreach ($operation in ($operations | Select-Object -First 20)) {
        $after = if ($operation.Action -eq 'Rename') { $operation.NewName } else { $operation.DestinationPath }
        if ([string]::IsNullOrWhiteSpace($after)) { $after = '(none)' }
        Write-Log "  $($operation.Action): $($operation.SourcePath) -> $after" 'Info'
    }

    if ($operations.Count -gt 20) {
        Write-Log "  ... and $($operations.Count - 20) more" 'Info'
    }
}

function Export-CurrentOperationPlan {
    if (-not $script:CurrentOperationPlan) {
        [System.Windows.MessageBox]::Show("No operation plan is loaded.", "Info", "OK", "Information")
        return
    }

    $saveDialog = [System.Windows.Forms.SaveFileDialog]::new()
    $saveDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $saveDialog.FileName = "StartMenuOperationPlan.json"

    if ($saveDialog.ShowDialog() -eq 'OK') {
        $json = $script:CurrentOperationPlan | ConvertTo-Json -Depth 8
        [System.IO.File]::WriteAllText($saveDialog.FileName, $json, [System.Text.UTF8Encoding]::new($false))
        Write-Log "Operation plan exported: $($saveDialog.FileName)" 'Success'
    }
}

function Import-OperationPlan {
    $openDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $openDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"

    if ($openDialog.ShowDialog() -eq 'OK') {
        try {
            $plan = Get-Content -LiteralPath $openDialog.FileName -Raw | ConvertFrom-Json
            $operations = @($plan.Operations)
            foreach ($operation in $operations) {
                if ($operation.Action -notin @('Delete','Move','Rename')) {
                    throw "Unsupported operation action: $($operation.Action)"
                }
                if ([string]::IsNullOrWhiteSpace($operation.SourcePath)) {
                    throw "Operation is missing SourcePath."
                }
            }

            $script:CurrentOperationPlan = $plan
            Update-PlanStatus
            Show-OperationPlanSummary -Plan $script:CurrentOperationPlan
            Write-Log "Operation plan imported: $($openDialog.FileName)" 'Success'
        }
        catch {
            Write-Log "Failed to import operation plan: $_" 'Error'
            [System.Windows.MessageBox]::Show("Failed to import operation plan: $_", "Error", "OK", "Error")
        }
    }
}

function Clear-OperationPlan {
    $script:CurrentOperationPlan = $null
    Update-PlanStatus
    Write-Log "Operation plan cleared." 'Info'
}

function Invoke-OperationPlanSynchronously {
    param(
        [array]$Operations,
        [string]$PlanName,
        [string]$OperationId = (New-OperationId)
    )

    $journalItems = @()
    $succeeded = 0
    $failed = 0

    Show-Progress -Value 0 -Maximum $Operations.Count
    for ($i = 0; $i -lt $Operations.Count; $i++) {
        $operation = $Operations[$i]
        Show-Progress -Value ($i + 1) -Maximum $Operations.Count

        try {
            switch ($operation.Action) {
                'Delete' {
                    $journalItems += Invoke-JournaledDelete -Path $operation.SourcePath -OperationId $operationId
                }
                'Move' {
                    $journalItems += Invoke-JournaledMove -SourcePath $operation.SourcePath -DestinationPath $operation.DestinationPath -OperationId $operationId
                }
                'Rename' {
                    $journalItems += Invoke-JournaledRename -SourcePath $operation.SourcePath -NewName $operation.NewName -OperationId $operationId
                }
            }
            $succeeded++
        }
        catch {
            $failed++
            $journalItems += New-JournalItem -ActionType $operation.Action -OriginalPath $operation.SourcePath -NewPath $operation.DestinationPath -Result 'Failed' -Message $_.Exception.Message
            Write-Log "Plan operation failed: $($operation.Action) $($operation.SourcePath) - $($_.Exception.Message)" 'Error'
        }
    }

    if ($journalItems.Count -gt 0) {
        Add-JournalEntry -Type 'Plan' -Description "Execute plan: $PlanName" -Items $journalItems -OperationId $OperationId
    }

    Show-Progress -Visible $false
    Write-Log "Operation plan executed: $succeeded succeeded, $failed failed" 'Info'

    return [PSCustomObject]@{
        Succeeded = $succeeded
        Failed = $failed
        JournalItems = $journalItems
    }
}

function Start-OperationPlanWorker {
    param(
        [array]$Operations,
        [string]$PlanName
    )

    if ($script:ActiveWorker) {
        Write-Log "Worker already running: $($script:ActiveWorker.Name)" 'Warning'
        return
    }

    $operationId = New-OperationId
    $state = [PSCustomObject]@{
        Index = 0
        Operations = $Operations
        PlanName = $PlanName
        OperationId = $operationId
        JournalItems = @()
        Succeeded = 0
        Failed = 0
    }

    $timer = [System.Windows.Threading.DispatcherTimer]::new()
    $timer.Interval = [TimeSpan]::FromMilliseconds(25)

    $finalize = {
        if ($state.JournalItems.Count -gt 0) {
            Add-JournalEntry -Type 'Plan' -Description "Execute plan: $($state.PlanName)" -Items $state.JournalItems -OperationId $state.OperationId
        }
        Write-Log "Operation plan executed: $($state.Succeeded) succeeded, $($state.Failed) failed" 'Info'
        Clear-OperationPlan
        Refresh-Items
    }

    $timer.Add_Tick({
        if (-not $script:ActiveWorker) { return }

        if ($state.Index -ge $state.Operations.Count) {
            $script:ActiveWorker = $null
            $timer.Stop()
            & $finalize
            Set-WorkerUiState -IsBusy $false
            Show-Progress -Visible $false
            return
        }

        $operation = $state.Operations[$state.Index]
        $state.Index++
        Show-Progress -Value $state.Index -Maximum $state.Operations.Count

        try {
            switch ($operation.Action) {
                'Delete' {
                    $state.JournalItems += Invoke-JournaledDelete -Path $operation.SourcePath -OperationId $state.OperationId
                }
                'Move' {
                    $state.JournalItems += Invoke-JournaledMove -SourcePath $operation.SourcePath -DestinationPath $operation.DestinationPath -OperationId $state.OperationId
                }
                'Rename' {
                    $state.JournalItems += Invoke-JournaledRename -SourcePath $operation.SourcePath -NewName $operation.NewName -OperationId $state.OperationId
                }
            }
            $state.Succeeded++
        }
        catch {
            $state.Failed++
            $state.JournalItems += New-JournalItem -ActionType $operation.Action -OriginalPath $operation.SourcePath -NewPath $operation.DestinationPath -Result 'Failed' -Message $_.Exception.Message
            Write-Log "Plan operation failed: $($operation.Action) $($operation.SourcePath) - $($_.Exception.Message)" 'Error'
        }
    })

    $script:ActiveWorker = [PSCustomObject]@{
        Name = "Executing operation plan..."
        Timer = $timer
        PowerShell = $null
        OnCanceled = $finalize
    }

    Show-Progress -Value 0 -Maximum $Operations.Count
    Set-WorkerUiState -IsBusy $true -Message "Executing operation plan..."
    $timer.Start()
}

function Invoke-CurrentOperationPlan {
    if (-not $script:CurrentOperationPlan) {
        [System.Windows.MessageBox]::Show("No operation plan is loaded.", "Info", "OK", "Information")
        return
    }

    $operations = @($script:CurrentOperationPlan.Operations)
    if ($operations.Count -eq 0) {
        [System.Windows.MessageBox]::Show("The operation plan is empty.", "Info", "OK", "Information")
        return
    }

    if ($script:UseSynchronousPlanExecution) {
        Invoke-OperationPlanSynchronously -Operations $operations -PlanName $script:CurrentOperationPlan.Name | Out-Null
        Clear-OperationPlan
        Refresh-Items
        return
    }

    Start-OperationPlanWorker -Operations $operations -PlanName $script:CurrentOperationPlan.Name
}

function Refresh-Items {
    try {
        $scopes = @(Get-StartMenuScopes)
    }
    catch {
        Write-Log "Unable to resolve Start Menu scope: $($_.Exception.Message)" 'Error'
        Update-Status "Scope is not valid"
        [System.Windows.MessageBox]::Show("Unable to resolve Start Menu scope:`n`n$($_.Exception.Message)", "Invalid Scope", "OK", "Error")
        return
    }
    if (($scopes | Where-Object { $_.RequiresAdmin }).Count -gt 0 -and -not $script:IsAdmin) {
        Write-Log "Selected scope can be scanned, but changes require Administrator rights." 'Warning'
    }

    $junkPatterns = @($script:JunkPatterns)
    $protectedFolders = @($script:ProtectedFolders)

    $scanScript = {
        param(
            [object[]]$Scopes,
            [string[]]$JunkPatterns,
            [string[]]$ProtectedFolders
        )

        function Get-WorkerShortcutTarget {
            param(
                [string]$ShortcutPath,
                $Shell
            )

            try {
                $extension = [System.IO.Path]::GetExtension($ShortcutPath).ToLowerInvariant()
                if ($extension -eq '.url') {
                    $urlLine = Get-Content -LiteralPath $ShortcutPath -ErrorAction Stop |
                        Where-Object { $_ -match '^URL=' } |
                        Select-Object -First 1
                    if ($urlLine) {
                        return ($urlLine -replace '^URL=', '').Trim()
                    }
                    return $null
                }
                if ($extension -eq '.appref-ms') {
                    return $ShortcutPath
                }

                $shortcut = $Shell.CreateShortcut($ShortcutPath)
                return $shortcut.TargetPath
            }
            catch {
                return $null
            }
        }

        function Test-WorkerShortcutBroken {
            param(
                [string]$ShortcutPath,
                $Shell
            )

            if (-not (Test-Path -LiteralPath $ShortcutPath)) { return $true }

            $extension = [System.IO.Path]::GetExtension($ShortcutPath).ToLowerInvariant()
            if ($extension -eq '.appref-ms') { return $false }

            $target = Get-WorkerShortcutTarget -ShortcutPath $ShortcutPath -Shell $Shell
            if ([string]::IsNullOrEmpty($target)) { return $false }
            if ($target -match '^(https?|ftp)://') { return $false }
            if ($target -match '^file://') {
                $target = ([System.Uri]$target).LocalPath
            }

            $expandedTarget = [Environment]::ExpandEnvironmentVariables($target)
            if ($expandedTarget -match '^[A-Za-z]:\\Windows\\' -or
                $expandedTarget -match '^shell:' -or
                $expandedTarget -match '^ms-[A-Za-z-]+:') {
                return $false
            }

            if ($expandedTarget -match '^[A-Za-z]:\\' -or $expandedTarget.StartsWith('\\')) {
                return -not (Test-Path -LiteralPath $expandedTarget)
            }

            return $false
        }

        function Test-WorkerJunk {
            param([string]$Name)
            foreach ($pattern in $JunkPatterns) {
                if ($Name -like $pattern) { return $true }
            }
            return $false
        }

        function Test-WorkerProtected {
            param(
                [string]$Path,
                [string]$BasePath
            )

            $base = ([System.IO.Path]::GetFullPath($BasePath)).TrimEnd('\', '/')
            $candidate = ([System.IO.Path]::GetFullPath($Path)).TrimEnd('\', '/')
            if (-not $candidate.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $false
            }

            $relative = $candidate.Substring($base.Length).TrimStart('\', '/')
            if ([string]::IsNullOrWhiteSpace($relative)) { return $false }

            $topFolder = $relative.Split([char[]]@('\', '/'), 2)[0]
            return $ProtectedFolders -contains $topFolder
        }

        function Get-WorkerShortcutMetadata {
            param(
                [string]$ShortcutPath,
                $Shell
            )

            $meta = [ordered]@{
                Target       = ''
                Arguments    = ''
                WorkingDir   = ''
                Description  = ''
                Hotkey       = ''
                IconLocation = ''
                WindowStyle  = 0
            }

            $ext = [System.IO.Path]::GetExtension($ShortcutPath).ToLowerInvariant()
            try {
                if ($ext -eq '.lnk') {
                    $lnk = $Shell.CreateShortcut($ShortcutPath)
                    $meta.Target = $lnk.TargetPath
                    $meta.Arguments = $lnk.Arguments
                    $meta.WorkingDir = $lnk.WorkingDirectory
                    $meta.Description = $lnk.Description
                    $meta.Hotkey = $lnk.Hotkey
                    $meta.IconLocation = $lnk.IconLocation
                    $meta.WindowStyle = $lnk.WindowStyle
                }
                elseif ($ext -eq '.url') {
                    $lines = @(Get-Content -LiteralPath $ShortcutPath -ErrorAction Stop)
                    foreach ($line in $lines) {
                        if ($line -match '^URL=(.*)') { $meta.Target = $Matches[1].Trim() }
                        elseif ($line -match '^IconFile=(.*)') { $meta.IconLocation = $Matches[1].Trim() }
                        elseif ($line -match '^HotKey=(.*)') { $meta.Hotkey = $Matches[1].Trim() }
                        elseif ($line -match '^WorkingDirectory=(.*)') { $meta.WorkingDir = $Matches[1].Trim() }
                    }
                }
                elseif ($ext -eq '.appref-ms') {
                    $meta.Target = $ShortcutPath
                }
            }
            catch {
                $meta.Target = ''
            }

            return [PSCustomObject]$meta
        }

        function Get-WorkerRiskFlags {
            param([PSCustomObject]$Metadata)

            $flags = @()
            $target = if ($Metadata.Target) { $Metadata.Target } else { '' }
            $mArgs = if ($Metadata.Arguments) { $Metadata.Arguments } else { '' }

            $scriptHosts = @('cmd.exe', 'cmd', 'powershell.exe', 'powershell', 'pwsh.exe', 'pwsh',
                             'wscript.exe', 'wscript', 'cscript.exe', 'cscript', 'mshta.exe', 'mshta',
                             'rundll32.exe', 'rundll32', 'regsvr32.exe', 'regsvr32', 'msiexec.exe', 'msiexec')
            $targetLeaf = [System.IO.Path]::GetFileNameWithoutExtension($target)
            $targetFile = [System.IO.Path]::GetFileName($target)
            if ($scriptHosts -contains $targetLeaf -or $scriptHosts -contains $targetFile) {
                $flags += 'ScriptHost'
            }

            if ($target -match '\\\\' -or $target -match '^\\\\') {
                $flags += 'NetworkTarget'
            }

            if ($mArgs.Length -gt 260) {
                $flags += 'LongArguments'
            }

            if ($mArgs -match '(^|\s)-[Ee]nc(odedCommand)?\s' -or $mArgs -match '-[Ww]indowStyle\s+[Hh]idden') {
                $flags += 'HiddenExecution'
            }

            if ($target -match '^(https?|ftp)://' -or ($target -match '^file://')) {
                $flags += 'WebTarget'
            }

            if ($target -match '\.(bat|cmd|vbs|vbe|js|jse|wsf|wsh|ps1|psm1|psd1)$') {
                $flags += 'ScriptTarget'
            }

            return $flags
        }

        $items = [System.Collections.Generic.List[object]]::new()
        $allShortcuts = @{}
        $shell = New-Object -ComObject WScript.Shell

        try {
            foreach ($scope in $Scopes) {
                $basePath = $scope.Path
                if (-not (Test-Path -LiteralPath $basePath)) { continue }

                $prefix = $scope.Prefix
                $requiresAdmin = [bool]$scope.RequiresAdmin
                $scopeName = $scope.ScopeName

                $scanQueue = [System.Collections.Generic.Queue[string]]::new()
                $scanQueue.Enqueue($basePath)
                while ($scanQueue.Count -gt 0) {
                    $currentDir = $scanQueue.Dequeue()
                    $children = @(Get-ChildItem -LiteralPath $currentDir -Force -ErrorAction SilentlyContinue)
                    foreach ($dirChild in $children) {
                        if ($dirChild.PSIsContainer -and ($dirChild.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0) {
                            $scanQueue.Enqueue($dirChild.FullName)
                        }
                    }
                    foreach ($entry in $children) {
                        $isReparsePoint = ($entry.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0
                        if ($isReparsePoint -and $entry.PSIsContainer) { continue }
                        $relativePath = $entry.FullName.Substring($basePath.Length + 1)
                        $isFolder = $entry.PSIsContainer
                        $isJunk = Test-WorkerJunk $entry.BaseName
                        $isProtected = Test-WorkerProtected -Path $entry.FullName -BasePath $basePath
                        $isBroken = $false
                        $targetPath = ''
                        $shortcutArgs = ''
                        $shortcutWorkingDir = ''
                        $shortcutDescription = ''
                        $riskFlags = @()
                        $extension = $entry.Extension.ToLowerInvariant()

                        if (-not $isFolder -and $extension -in @('.lnk', '.url', '.appref-ms')) {
                            $metadata = Get-WorkerShortcutMetadata -ShortcutPath $entry.FullName -Shell $shell
                            $targetPath = $metadata.Target
                            $shortcutArgs = $metadata.Arguments
                            $shortcutWorkingDir = $metadata.WorkingDir
                            $shortcutDescription = $metadata.Description
                            $isBroken = Test-WorkerShortcutBroken -ShortcutPath $entry.FullName -Shell $shell
                            $riskFlags = @(Get-WorkerRiskFlags -Metadata $metadata)

                            if (-not [string]::IsNullOrEmpty($targetPath)) {
                                if (-not $allShortcuts.ContainsKey($targetPath)) {
                                    $allShortcuts[$targetPath] = @()
                                }
                                $allShortcuts[$targetPath] += $entry.FullName
                            }
                        }

                        $itemType = if ($isFolder) {
                            'Folder'
                        }
                        elseif ($extension -eq '.url') {
                            'URL'
                        }
                        elseif ($extension -eq '.appref-ms') {
                            'App Reference'
                        }
                        elseif ($isJunk) {
                            'Junk'
                        }
                        else {
                            'Shortcut'
                        }
                        $riskDisplay = if ($riskFlags.Count -gt 0) { $riskFlags -join ', ' } else { '' }
                        $item = [PSCustomObject]@{
                            IsSelected   = $false
                            DisplayName  = $entry.BaseName
                            RelativePath = "$prefix $relativePath"
                            FullPath     = $entry.FullName
                            BasePath     = $basePath
                            IsFolder     = $isFolder
                            IsJunk       = $isJunk
                            IsBroken     = $isBroken
                            IsDuplicate  = $false
                            IsProtected  = $isProtected
                            ItemType     = $itemType
                            TargetPath   = $targetPath
                            Arguments    = $shortcutArgs
                            WorkingDir   = $shortcutWorkingDir
                            Description  = $shortcutDescription
                            RiskFlags    = $riskDisplay
                            Status       = 'OK'
                            IsSystem     = $requiresAdmin
                            RequiresAdmin = $requiresAdmin
                            ScopeName    = $scopeName
                        }
                        $items.Add($item)
                    }
                }
            }

            $duplicateTargets = $allShortcuts.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
            foreach ($dup in $duplicateTargets) {
                foreach ($path in $dup.Value) {
                    $item = $items | Where-Object { $_.FullPath -eq $path } | Select-Object -First 1
                    if ($item) {
                        $item.IsDuplicate = $true
                        $item.ItemType = 'Duplicate'
                    }
                }
            }

            foreach ($item in $items) {
                $statuses = @()
                if ($item.IsJunk) { $statuses += 'Junk' }
                if ($item.IsBroken) { $statuses += 'Broken' }
                if ($item.IsDuplicate) { $statuses += 'Duplicate' }
                if ($item.IsProtected) { $statuses += 'Protected' }
                $item.Status = if ($statuses.Count -eq 0) { 'OK' } else { $statuses -join ', ' }
            }

            return [PSCustomObject]@{ Items = @($items) }
        }
        finally {
            if ($shell) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
            }
        }
    }

    $completed = {
        param($Output)

        $result = $Output | Select-Object -First 1
        $script:AllItems.Clear()
        $script:FilteredItems.Clear()

        foreach ($item in @($result.Items)) {
            $script:AllItems.Add($item)
        }

        Apply-Filters
        Update-Stats
        Update-Status "Ready"
        Write-Log "Scan complete: $($script:AllItems.Count) item(s)" 'Success'
    }

    Show-Progress -Value 0 -Maximum 100
    Start-BackgroundWorker -Name "Scanning Start Menu..." -ScriptBlock $scanScript -Arguments @($scopes, $junkPatterns, $protectedFolders) -OnCompleted $completed | Out-Null
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

    if (-not (Test-Path -LiteralPath $Config.BackupRoot)) {
        [System.IO.Directory]::CreateDirectory($Config.BackupRoot) | Out-Null
    }

    [System.IO.Directory]::CreateDirectory($backupPath) | Out-Null

    try {
        $scopes = @(Get-StartMenuScopes)
    }
    catch {
        Write-Log "Backup failed before start: $($_.Exception.Message)" 'Error'
        [System.Windows.MessageBox]::Show("Unable to resolve backup scope:`n`n$($_.Exception.Message)", "Backup Error", "OK", "Error")
        return $null
    }

    $i = 0
    $failed = @()
    foreach ($scope in $scopes) {
        $path = $scope.Path
        if ($scope.RequiresAdmin -and -not $script:IsAdmin) {
            $failed += "$($scope.BackupName): Administrator privileges required"
            Write-Log "Backup skipped for $($scope.ScopeName) Start Menu because the app is not elevated." 'Warning'
            continue
        }
        if (Test-Path -LiteralPath $path) {
            $destName = $scope.BackupName
            $destPath = Join-Path $backupPath $destName
            Show-Progress -Value ($i++ * [Math]::Max(1, [Math]::Floor(100 / [Math]::Max(1, $scopes.Count)))) -Maximum 100
            try {
                Copy-Item -LiteralPath $path -Destination $destPath -Recurse -Force -ErrorAction Stop
            }
            catch {
                $failed += "$destName`: $($_.Exception.Message)"
                Write-Log "Backup failed for $destName Start Menu: $($_.Exception.Message)" 'Error'
            }
        }
    }

    Show-Progress -Visible $false
    if ($failed.Count -gt 0) {
        Write-Log "Backup completed with $($failed.Count) failure(s): $backupPath" 'Warning'
        [System.Windows.MessageBox]::Show("Backup completed with warnings.`n`n$backupPath`n`n$($failed -join "`n")", "Backup Warning", "OK", "Warning")
    }
    else {
        Write-Log "Backup created: $backupPath" 'Success'
        [System.Windows.MessageBox]::Show("Backup created successfully!`n`n$backupPath", "Backup Complete", "OK", "Information")
    }
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
        $operations = @($selected | ForEach-Object {
            New-PlanOperation -Action 'Delete' -SourcePath $_.FullPath -DisplayName $_.DisplayName
        })
        Set-CurrentOperationPlan -Name "Delete selected items" -Operations $operations
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
    
    $operationId = New-OperationId
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
                $undoItems += New-JournalItem -ActionType 'Delete' -OriginalPath $item.FullPath -Result 'Skipped' -Message 'Administrator privileges required'
                $failed++
                continue
            }
            if ($item.IsProtected) {
                Write-Log "Skipped protected item: $($item.DisplayName)" 'Warning'
                $undoItems += New-JournalItem -ActionType 'Delete' -OriginalPath $item.FullPath -Result 'Skipped' -Message 'Protected Start Menu item'
                $failed++
                continue
            }
            
            if (Test-Path $item.FullPath) {
                $undoItems += Invoke-JournaledDelete -Path $item.FullPath -OperationId $operationId
                Write-Log "Deleted: $($item.DisplayName)" 'Success'
                $deleted++
            }
        }
        catch {
            Write-Log "Failed to delete: $($item.DisplayName) - $_" 'Error'
            $undoItems += New-JournalItem -ActionType 'Delete' -OriginalPath $item.FullPath -Result 'Failed' -Message $_.Exception.Message
            $failed++
        }
    }

    if ($undoItems.Count -gt 0) {
        Add-JournalEntry -Type 'Delete' -Description "Delete $deleted item(s)" -Items $undoItems -OperationId $operationId
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
        $operations = @($junkItems | ForEach-Object {
            New-PlanOperation -Action 'Delete' -SourcePath $_.FullPath -DisplayName $_.DisplayName
        })
        Set-CurrentOperationPlan -Name "Remove all junk" -Operations $operations
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
        $operations = @($brokenItems | ForEach-Object {
            New-PlanOperation -Action 'Delete' -SourcePath $_.FullPath -DisplayName $_.DisplayName
        })
        Set-CurrentOperationPlan -Name "Remove broken shortcuts" -Operations $operations
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
        $operations = @($toDelete | ForEach-Object {
            New-PlanOperation -Action 'Delete' -SourcePath $_.FullPath -DisplayName $_.DisplayName
        })
        Set-CurrentOperationPlan -Name "Remove duplicates" -Operations $operations
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
    $scopes = Get-StartMenuScopes
    $toFlatten = @()
    
    foreach ($scope in $scopes) {
        $basePath = $scope.Path
        if ($scope.RequiresAdmin -and -not $script:IsAdmin) { continue }

        $folders = Get-ChildItem -Path $basePath -Directory -ErrorAction SilentlyContinue

        foreach ($folder in $folders) {
            if (Test-ProtectedFolder -Path $folder.FullName -BasePath $basePath) {
                Write-Log "Skipped protected folder: $($folder.Name)" 'Warning'
                continue
            }

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
        $operations = @()
        foreach ($item in $toFlatten) {
            $destPath = Join-Path $item.BasePath $item.Shortcut.Name
            if (Test-Path $destPath) {
                $destPath = Join-Path $item.BasePath "$($item.Folder.Name) - $($item.Shortcut.Name)"
            }
            $operations += New-PlanOperation -Action 'Move' -SourcePath $item.Shortcut.FullName -DestinationPath $destPath -DisplayName $item.Shortcut.Name
            $operations += New-PlanOperation -Action 'Delete' -SourcePath $item.Folder.FullName -DisplayName $item.Folder.Name
        }
        Set-CurrentOperationPlan -Name "Flatten single-item folders" -Operations $operations
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
    
    $operationId = New-OperationId
    $journalItems = @()
    $flattened = 0
    foreach ($item in $toFlatten) {
        try {
            $destPath = Join-Path $item.BasePath $item.Shortcut.Name

            if (Test-Path $destPath) {
                $destPath = Join-Path $item.BasePath "$($item.Folder.Name) - $($item.Shortcut.Name)"
            }

            $journalItems += Invoke-JournaledMove -SourcePath $item.Shortcut.FullName -DestinationPath $destPath -OperationId $operationId
            $journalItems += Invoke-JournaledDelete -Path $item.Folder.FullName -OperationId $operationId

            Write-Log "Flattened: $($item.Folder.Name)" 'Success'
            $flattened++
        }
        catch {
            Write-Log "Failed to flatten: $($item.Folder.Name) - $_" 'Error'
            $journalItems += New-JournalItem -ActionType 'Flatten' -OriginalPath $item.Folder.FullName -Result 'Failed' -Message $_.Exception.Message
        }
    }

    if ($journalItems.Count -gt 0) {
        Add-JournalEntry -Type 'Flatten' -Description "Flatten $flattened folder(s)" -Items $journalItems -OperationId $operationId
    }

    Write-Log "Flattening complete: $flattened folders flattened" 'Info'
    Refresh-Items
}

function Remove-EmptyFolders {
    $isPreview = $chkPreviewMode.IsChecked
    $scopes = Get-StartMenuScopes
    $emptyFolders = @()
    
    foreach ($scope in $scopes) {
        $basePath = $scope.Path
        if ($scope.RequiresAdmin -and -not $script:IsAdmin) { continue }
        
        $folders = Get-ChildItem -Path $basePath -Directory -Recurse -ErrorAction SilentlyContinue | 
                   Sort-Object { $_.FullName.Length } -Descending
        
        foreach ($folder in $folders) {
            if (Test-ProtectedFolder -Path $folder.FullName -BasePath $basePath) {
                Write-Log "Skipped protected folder: $($folder.Name)" 'Warning'
                continue
            }

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
        $operations = @($emptyFolders | ForEach-Object {
            New-PlanOperation -Action 'Delete' -SourcePath $_.FullName -DisplayName $_.Name
        })
        Set-CurrentOperationPlan -Name "Remove empty folders" -Operations $operations
        Write-Log "PREVIEW: Would remove $($emptyFolders.Count) empty folder(s):" 'Warning'
        foreach ($folder in $emptyFolders) {
            Write-Log "  - $($folder.Name)" 'Info'
        }
        return
    }
    
    $operationId = New-OperationId
    $journalItems = @()
    $removed = 0
    foreach ($folder in $emptyFolders) {
        try {
            $journalItems += Invoke-JournaledDelete -Path $folder.FullName -OperationId $operationId
            Write-Log "Removed empty: $($folder.Name)" 'Success'
            $removed++
        }
        catch {
            Write-Log "Failed to remove: $($folder.Name) - $_" 'Error'
            $journalItems += New-JournalItem -ActionType 'Delete' -OriginalPath $folder.FullName -Result 'Failed' -Message $_.Exception.Message
        }
    }

    if ($journalItems.Count -gt 0) {
        Add-JournalEntry -Type 'Delete' -Description "Remove $removed empty folder(s)" -Items $journalItems -OperationId $operationId
    }

    Write-Log "Removed $removed empty folders" 'Info'
    Refresh-Items
}

function Move-AllToRoot {
    $isPreview = $chkPreviewMode.IsChecked
    $scopes = Get-StartMenuScopes
    $toMove = @()
    
    foreach ($scope in $scopes) {
        $basePath = $scope.Path
        if ($scope.RequiresAdmin -and -not $script:IsAdmin) { continue }
        
        # Get all shortcuts that are NOT in the root directory
        Get-ChildItem -Path $basePath -Recurse -Filter "*.lnk" -ErrorAction SilentlyContinue | ForEach-Object {
            if (Test-ProtectedFolder -Path $_.FullName -BasePath $basePath) {
                Write-Log "Skipped protected shortcut: $($_.BaseName)" 'Warning'
                return
            }

            $parentDir = Split-Path $_.FullName -Parent
            if ($parentDir -ne $basePath) {
                $toMove += @{
                    Item = $_
                    BasePath = $basePath
                    IsSystem = $scope.RequiresAdmin
                }
            }
        }
    }
    
    if ($toMove.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No shortcuts found in folders to move.", "Info", "OK", "Information")
        return
    }
    
    if ($isPreview) {
        $operations = @()
        foreach ($entry in $toMove) {
            $destPath = Join-Path $entry.BasePath $entry.Item.Name
            if (Test-Path $destPath) {
                $existingTarget = Get-ShortcutTarget $destPath
                $newTarget = Get-ShortcutTarget $entry.Item.FullName
                if ($existingTarget -eq $newTarget) {
                    $operations += New-PlanOperation -Action 'Delete' -SourcePath $entry.Item.FullName -DisplayName $entry.Item.BaseName
                    continue
                }

                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($entry.Item.Name)
                $ext = [System.IO.Path]::GetExtension($entry.Item.Name)
                $counter = 2
                do {
                    $destPath = Join-Path $entry.BasePath "$baseName ($counter)$ext"
                    $counter++
                } while (Test-Path $destPath)
            }

            $operations += New-PlanOperation -Action 'Move' -SourcePath $entry.Item.FullName -DestinationPath $destPath -DisplayName $entry.Item.BaseName
        }
        Set-CurrentOperationPlan -Name "Move all to root" -Operations $operations
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
    
    $operationId = New-OperationId
    $journalItems = @()
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
                    $journalItems += Invoke-JournaledDelete -Path $entry.Item.FullName -OperationId $operationId
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

            $journalItems += Invoke-JournaledMove -SourcePath $entry.Item.FullName -DestinationPath $destPath -OperationId $operationId
            Write-Log "Moved to root: $($entry.Item.BaseName)" 'Success'
            $moved++
        }
        catch {
            Write-Log "Failed to move: $($entry.Item.BaseName) - $_" 'Error'
            $journalItems += New-JournalItem -ActionType 'Move' -OriginalPath $entry.Item.FullName -NewPath $destPath -Result 'Failed' -Message $_.Exception.Message
        }
    }

    if ($journalItems.Count -gt 0) {
        Add-JournalEntry -Type 'Move' -Description "Move $moved shortcut(s) to root" -Items $journalItems -OperationId $operationId
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
        $operations = @($selected | ForEach-Object {
            $categoryPath = Join-Path $_.BasePath $CategoryName
            $destPath = Join-Path $categoryPath (Split-Path $_.FullPath -Leaf)
            New-PlanOperation -Action 'Move' -SourcePath $_.FullPath -DestinationPath $destPath -DisplayName $_.DisplayName
        })
        Set-CurrentOperationPlan -Name "Move to $CategoryName" -Operations $operations
        Write-Log "PREVIEW: Would move $($selected.Count) item(s) to '$CategoryName':" 'Warning'
        foreach ($item in $selected) {
            Write-Log "  - $($item.DisplayName)" 'Info'
        }
        return
    }
    
    $operationId = New-OperationId
    $journalItems = @()
    $moved = 0
    foreach ($item in $selected) {
        if ($item.IsSystem -and -not $script:IsAdmin) {
            Write-Log "Skipped (need admin): $($item.DisplayName)" 'Warning'
            $journalItems += New-JournalItem -ActionType 'Move' -OriginalPath $item.FullPath -Result 'Skipped' -Message 'Administrator privileges required'
            continue
        }
        if ($item.IsProtected) {
            Write-Log "Skipped protected item: $($item.DisplayName)" 'Warning'
            $journalItems += New-JournalItem -ActionType 'Move' -OriginalPath $item.FullPath -Result 'Skipped' -Message 'Protected Start Menu item'
            continue
        }

        $destPath = $null
        try {
            $categoryPath = Join-Path $item.BasePath $CategoryName
            if (-not (Test-Path $categoryPath)) {
                New-Item -Path $categoryPath -ItemType Directory -Force | Out-Null
            }

            $destPath = Join-Path $categoryPath (Split-Path $item.FullPath -Leaf)
            $journalItems += Invoke-JournaledMove -SourcePath $item.FullPath -DestinationPath $destPath -OperationId $operationId
            Write-Log "Moved to $CategoryName`: $($item.DisplayName)" 'Success'
            $moved++
        }
        catch {
            Write-Log "Failed to move: $($item.DisplayName) - $_" 'Error'
            $journalItems += New-JournalItem -ActionType 'Move' -OriginalPath $item.FullPath -NewPath $destPath -Result 'Failed' -Message $_.Exception.Message
        }
    }

    if ($journalItems.Count -gt 0) {
        Add-JournalEntry -Type 'Move' -Description "Move $moved item(s) to $CategoryName" -Items $journalItems -OperationId $operationId
    }

    Write-Log "Moved $moved items to $CategoryName" 'Info'
    Refresh-Items
}

function Auto-OrganizeAll {
    $isPreview = $chkPreviewMode.IsChecked
    $scopes = Get-StartMenuScopes
    $toOrganize = @()
    
    foreach ($scope in $scopes) {
        $basePath = $scope.Path
        if ($scope.RequiresAdmin -and -not $script:IsAdmin) { continue }
        
        Get-ChildItem -Path $basePath -Recurse -Filter "*.lnk" -ErrorAction SilentlyContinue | ForEach-Object {
            if (Test-ProtectedFolder -Path $_.FullName -BasePath $basePath) {
                Write-Log "Skipped protected shortcut: $($_.BaseName)" 'Warning'
                return
            }

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
        $operations = @($toOrganize | ForEach-Object {
            $categoryPath = Join-Path $_.BasePath $_.Category
            $destPath = Join-Path $categoryPath $_.Item.Name
            if (-not (Test-Path -LiteralPath $destPath)) {
                New-PlanOperation -Action 'Move' -SourcePath $_.Item.FullName -DestinationPath $destPath -DisplayName $_.Item.BaseName
            }
        })
        Set-CurrentOperationPlan -Name "Auto-organize all" -Operations $operations
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
    
    $operationId = New-OperationId
    $journalItems = @()
    $organized = 0
    Show-Progress -Value 0 -Maximum $toOrganize.Count
    
    foreach ($i in 0..($toOrganize.Count - 1)) {
        $entry = $toOrganize[$i]
        Show-Progress -Value ($i + 1) -Maximum $toOrganize.Count
        
        $destPath = $null
        try {
            $categoryPath = Join-Path $entry.BasePath $entry.Category
            if (-not (Test-Path $categoryPath)) {
                New-Item -Path $categoryPath -ItemType Directory -Force | Out-Null
            }

            $destPath = Join-Path $categoryPath $entry.Item.Name
            if (-not (Test-Path $destPath)) {
                $journalItems += Invoke-JournaledMove -SourcePath $entry.Item.FullName -DestinationPath $destPath -OperationId $operationId
                Write-Log "Organized: $($entry.Item.BaseName) -> $($entry.Category)" 'Success'
                $organized++
            }
        }
        catch {
            Write-Log "Failed: $($entry.Item.BaseName) - $_" 'Error'
            $journalItems += New-JournalItem -ActionType 'Move' -OriginalPath $entry.Item.FullName -NewPath $destPath -Result 'Failed' -Message $_.Exception.Message
        }
    }

    if ($journalItems.Count -gt 0) {
        Add-JournalEntry -Type 'Move' -Description "Auto-organize $organized item(s)" -Items $journalItems -OperationId $operationId
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
        $operations = @($toRename | ForEach-Object {
            $ext = [System.IO.Path]::GetExtension($_.Item.FullPath)
            New-PlanOperation -Action 'Rename' -SourcePath $_.Item.FullPath -NewName "$($_.NewName)$ext" -DisplayName $_.Item.DisplayName
        })
        Set-CurrentOperationPlan -Name "Strip version numbers" -Operations $operations
        Write-Log "PREVIEW: Would rename $($toRename.Count) item(s):" 'Warning'
        foreach ($entry in $toRename) {
            Write-Log "  - '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Info'
        }
        return
    }
    
    $operationId = New-OperationId
    $journalItems = @()
    $renamed = 0
    foreach ($entry in $toRename) {
        if ($entry.Item.IsSystem -and -not $script:IsAdmin) {
            $journalItems += New-JournalItem -ActionType 'Rename' -OriginalPath $entry.Item.FullPath -Result 'Skipped' -Message 'Administrator privileges required'
            continue
        }
        if ($entry.Item.IsProtected) {
            $journalItems += New-JournalItem -ActionType 'Rename' -OriginalPath $entry.Item.FullPath -Result 'Skipped' -Message 'Protected Start Menu item'
            Write-Log "Skipped protected item: $($entry.Item.DisplayName)" 'Warning'
            continue
        }

        $newPath = $null
        try {
            $ext = [System.IO.Path]::GetExtension($entry.Item.FullPath)
            $newPath = Join-Path (Split-Path $entry.Item.FullPath -Parent) "$($entry.NewName)$ext"

            if (-not (Test-Path $newPath)) {
                $journalItems += Invoke-JournaledRename -SourcePath $entry.Item.FullPath -NewName "$($entry.NewName)$ext" -OperationId $operationId
                Write-Log "Renamed: '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Success'
                $renamed++
            }
        }
        catch {
            Write-Log "Failed to rename: $($entry.Item.DisplayName) - $_" 'Error'
            $journalItems += New-JournalItem -ActionType 'Rename' -OriginalPath $entry.Item.FullPath -NewPath $newPath -Result 'Failed' -Message $_.Exception.Message
        }
    }

    if ($journalItems.Count -gt 0) {
        Add-JournalEntry -Type 'Rename' -Description "Rename $renamed item(s)" -Items $journalItems -OperationId $operationId
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
        $operations = @($toRename | ForEach-Object {
            $ext = [System.IO.Path]::GetExtension($_.Item.FullPath)
            New-PlanOperation -Action 'Rename' -SourcePath $_.Item.FullPath -NewName "$($_.NewName)$ext" -DisplayName $_.Item.DisplayName
        })
        Set-CurrentOperationPlan -Name "Clean names" -Operations $operations
        Write-Log "PREVIEW: Would clean $($toRename.Count) name(s):" 'Warning'
        foreach ($entry in $toRename) {
            Write-Log "  - '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Info'
        }
        return
    }
    
    $operationId = New-OperationId
    $journalItems = @()
    $renamed = 0
    foreach ($entry in $toRename) {
        if ($entry.Item.IsSystem -and -not $script:IsAdmin) {
            $journalItems += New-JournalItem -ActionType 'Rename' -OriginalPath $entry.Item.FullPath -Result 'Skipped' -Message 'Administrator privileges required'
            continue
        }
        if ($entry.Item.IsProtected) {
            $journalItems += New-JournalItem -ActionType 'Rename' -OriginalPath $entry.Item.FullPath -Result 'Skipped' -Message 'Protected Start Menu item'
            Write-Log "Skipped protected item: $($entry.Item.DisplayName)" 'Warning'
            continue
        }

        $newPath = $null
        try {
            $ext = [System.IO.Path]::GetExtension($entry.Item.FullPath)
            $newPath = Join-Path (Split-Path $entry.Item.FullPath -Parent) "$($entry.NewName)$ext"

            if (-not (Test-Path $newPath)) {
                $journalItems += Invoke-JournaledRename -SourcePath $entry.Item.FullPath -NewName "$($entry.NewName)$ext" -OperationId $operationId
                Write-Log "Cleaned: '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Success'
                $renamed++
            }
        }
        catch {
            Write-Log "Failed to clean: $($entry.Item.DisplayName) - $_" 'Error'
            $journalItems += New-JournalItem -ActionType 'Rename' -OriginalPath $entry.Item.FullPath -NewPath $newPath -Result 'Failed' -Message $_.Exception.Message
        }
    }

    if ($journalItems.Count -gt 0) {
        Add-JournalEntry -Type 'Rename' -Description "Clean $renamed name(s)" -Items $journalItems -OperationId $operationId
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
        $operations = @($toRename | ForEach-Object {
            $ext = [System.IO.Path]::GetExtension($_.Item.FullPath)
            New-PlanOperation -Action 'Rename' -SourcePath $_.Item.FullPath -NewName "$($_.NewName)$ext" -DisplayName $_.Item.DisplayName
        })
        Set-CurrentOperationPlan -Name "Find and replace names" -Operations $operations
        Write-Log "PREVIEW: Would rename $($toRename.Count) item(s):" 'Warning'
        foreach ($entry in $toRename) {
            Write-Log "  - '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Info'
        }
        return
    }
    
    $operationId = New-OperationId
    $journalItems = @()
    $renamed = 0
    foreach ($entry in $toRename) {
        if ($entry.Item.IsSystem -and -not $script:IsAdmin) {
            $journalItems += New-JournalItem -ActionType 'Rename' -OriginalPath $entry.Item.FullPath -Result 'Skipped' -Message 'Administrator privileges required'
            continue
        }
        if ($entry.Item.IsProtected) {
            $journalItems += New-JournalItem -ActionType 'Rename' -OriginalPath $entry.Item.FullPath -Result 'Skipped' -Message 'Protected Start Menu item'
            Write-Log "Skipped protected item: $($entry.Item.DisplayName)" 'Warning'
            continue
        }

        $newPath = $null
        try {
            $ext = [System.IO.Path]::GetExtension($entry.Item.FullPath)
            $newName = "$($entry.NewName)$ext"
            $newPath = Join-Path (Split-Path $entry.Item.FullPath -Parent) $newName
            $journalItems += Invoke-JournaledRename -SourcePath $entry.Item.FullPath -NewName $newName -OperationId $operationId
            Write-Log "Renamed: '$($entry.Item.DisplayName)' -> '$($entry.NewName)'" 'Success'
            $renamed++
        }
        catch {
            Write-Log "Failed: $($entry.Item.DisplayName) - $_" 'Error'
            $journalItems += New-JournalItem -ActionType 'Rename' -OriginalPath $entry.Item.FullPath -NewPath $newPath -Result 'Failed' -Message $_.Exception.Message
        }
    }

    if ($journalItems.Count -gt 0) {
        Add-JournalEntry -Type 'Rename' -Description "Find/replace rename $renamed item(s)" -Items $journalItems -OperationId $operationId
    }

    Write-Log "Renamed $renamed items" 'Info'
    Refresh-Items
}

function Get-NormalizedPath {
    param([string]$Path)
    return ([System.IO.Path]::GetFullPath($Path)).TrimEnd('\', '/')
}

function Test-ApprovedRestoreTarget {
    param([string]$Path)

    $candidate = Get-NormalizedPath $Path
    $allowedTargets = @(Get-ApprovedMutationRoots)

    return $allowedTargets -contains $candidate
}

function Copy-DirectoryContents {
    param(
        [string]$SourcePath,
        [string]$DestinationPath
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
        throw "Source folder does not exist: $SourcePath"
    }

    if (-not (Test-Path -LiteralPath $DestinationPath -PathType Container)) {
        [System.IO.Directory]::CreateDirectory($DestinationPath) | Out-Null
    }

    $children = @(Get-ChildItem -LiteralPath $SourcePath -Force -ErrorAction Stop)
    $copied = 0
    foreach ($child in $children) {
        if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
        Copy-Item -LiteralPath $child.FullName -Destination $DestinationPath -Recurse -Force -ErrorAction Stop
        $copied++
    }

    return $copied
}

function Clear-DirectoryContents {
    param([string]$Path)

    if (-not (Test-ApprovedRestoreTarget $Path)) {
        throw "Refusing to clear unapproved restore target: $Path"
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        [System.IO.Directory]::CreateDirectory($Path) | Out-Null
        return
    }

    $children = @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop)
    foreach ($child in $children) {
        $operation = Invoke-GuardedFileOperation -Action 'Delete' -SourcePath $child.FullName
        if ($operation.Result -ne 'Success') {
            throw $operation.Message
        }
    }
}

function New-RestorePlan {
    param(
        [string]$ScopeName,
        [string]$BackupPath,
        [string]$TargetPath,
        [string]$StagingRoot,
        [string]$RollbackRoot
    )

    if (-not (Test-Path -LiteralPath $BackupPath -PathType Container)) {
        throw "$ScopeName backup folder does not exist: $BackupPath"
    }

    if (-not (Test-ApprovedRestoreTarget $TargetPath)) {
        throw "$ScopeName target is not an approved Start Menu path: $TargetPath"
    }

    $stagedSource = Join-Path $StagingRoot $ScopeName
    $rollbackPath = Join-Path $RollbackRoot $ScopeName

    Copy-Item -LiteralPath $BackupPath -Destination $stagedSource -Recurse -Force -ErrorAction Stop
    if (-not (Test-Path -LiteralPath $stagedSource -PathType Container)) {
        throw "$ScopeName backup could not be staged for validation."
    }

    if (Test-Path -LiteralPath $TargetPath -PathType Container) {
        Copy-Item -LiteralPath $TargetPath -Destination $rollbackPath -Recurse -Force -ErrorAction Stop
    }
    else {
        [System.IO.Directory]::CreateDirectory($rollbackPath) | Out-Null
    }

    if (-not (Test-Path -LiteralPath $rollbackPath -PathType Container)) {
        throw "$ScopeName rollback snapshot could not be created."
    }

    Get-ChildItem -LiteralPath $stagedSource -Force -ErrorAction Stop | Out-Null
    Get-ChildItem -LiteralPath $rollbackPath -Force -ErrorAction Stop | Out-Null

    return [PSCustomObject]@{
        ScopeName    = $ScopeName
        BackupPath   = $BackupPath
        StagedSource = $stagedSource
        RollbackPath = $rollbackPath
        TargetPath   = $TargetPath
    }
}

function Invoke-RestoreRollback {
    param($Plan)

    Clear-DirectoryContents -Path $Plan.TargetPath
    Copy-DirectoryContents -SourcePath $Plan.RollbackPath -DestinationPath $Plan.TargetPath | Out-Null
}

function Restore-DirectoryFromPlan {
    param($Plan)

    try {
        Clear-DirectoryContents -Path $Plan.TargetPath
        $restoredCount = Copy-DirectoryContents -SourcePath $Plan.StagedSource -DestinationPath $Plan.TargetPath
        return [PSCustomObject]@{
            ScopeName          = $Plan.ScopeName
            Success            = $true
            RollbackSucceeded  = $null
            Message            = "$($Plan.ScopeName) Start Menu restored ($restoredCount root item(s))."
        }
    }
    catch {
        $restoreError = $_.Exception.Message
        Write-Log "$($Plan.ScopeName) restore failed; attempting rollback: $restoreError" 'Error'

        try {
            Invoke-RestoreRollback -Plan $Plan
            return [PSCustomObject]@{
                ScopeName          = $Plan.ScopeName
                Success            = $false
                RollbackSucceeded  = $true
                Message            = "$($Plan.ScopeName) restore failed and was rolled back: $restoreError"
            }
        }
        catch {
            return [PSCustomObject]@{
                ScopeName          = $Plan.ScopeName
                Success            = $false
                RollbackSucceeded  = $false
                Message            = "$($Plan.ScopeName) restore failed, then rollback failed: $restoreError; rollback error: $($_.Exception.Message)"
            }
        }
    }
}

function Restore-Backup {
    if (-not (Test-Path -LiteralPath $Config.BackupRoot)) {
        [System.Windows.MessageBox]::Show("No backups found.", "Info", "OK", "Information")
        return
    }

    $backups = Get-ChildItem -LiteralPath $Config.BackupRoot -Directory |
        Where-Object { $_.Name -match '^\d{8}_\d{6}$' } |
        Sort-Object Name -Descending

    if ($backups.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No backups found.", "Info", "OK", "Information")
        return
    }

    $result = [System.Windows.MessageBox]::Show(
        "Restore from latest backup?`n`n$($backups[0].Name)`n`nThis will replace current Start Menu contents!",
        "Confirm Restore", "YesNo", "Warning"
    )

    if ($result -ne 'Yes') { return }

    $latestBackup = $backups[0]
    $operationId = New-OperationId
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $rollbackRoot = Join-Path $Config.BackupRoot "_restore_rollback\$timestamp"
    $stagingRoot = Join-Path $env:TEMP "StartMenuOrganizer_Restore_$([System.Guid]::NewGuid().ToString('N'))"

    try {
        [System.IO.Directory]::CreateDirectory($rollbackRoot) | Out-Null
        [System.IO.Directory]::CreateDirectory($stagingRoot) | Out-Null

        $plans = @()
        $skipped = @()
        $planFailures = @()
        $scopes = @(Get-StartMenuScopes)

        foreach ($scope in $scopes) {
            $scopeBackup = Join-Path $latestBackup.FullName $scope.BackupName
            if (-not (Test-Path -LiteralPath $scopeBackup)) {
                $skipped += "$($scope.BackupName) backup folder was not present in $($latestBackup.Name)."
                continue
            }
            if ($scope.RequiresAdmin -and -not $script:IsAdmin) {
                $skipped += "$($scope.ScopeName) Start Menu skipped; run as Administrator to restore it."
                Write-Log "Skipped $($scope.ScopeName) Start Menu restore because the app is not elevated." 'Warning'
                continue
            }

            try {
                $plans += New-RestorePlan -ScopeName $scope.ScopeName -BackupPath $scopeBackup -TargetPath $scope.Path -StagingRoot $stagingRoot -RollbackRoot $rollbackRoot
                Write-Log "Prepared $($scope.ScopeName) Start Menu restore with rollback snapshot." 'Info'
            }
            catch {
                $planFailures += "$($scope.ScopeName): $($_.Exception.Message)"
                Write-Log "$($scope.ScopeName) restore preparation failed: $($_.Exception.Message)" 'Error'
            }
        }

        $legacyBackupNames = @('User', 'System')
        foreach ($legacyBackupName in $legacyBackupNames) {
            if ($scopes.BackupName -contains $legacyBackupName) { continue }
            $legacyBackup = Join-Path $latestBackup.FullName $legacyBackupName
            if (Test-Path -LiteralPath $legacyBackup) {
                $skipped += "$legacyBackupName backup folder is present but is not selected in the current scope."
            }
        }

        if ($plans.Count -eq 0) {
            $messages = @($planFailures + $skipped)
            throw "No restore targets were prepared. $($messages -join ' ')"
        }

        $results = @()
        foreach ($plan in $plans) {
            $restoreResult = Restore-DirectoryFromPlan -Plan $plan
            $results += $restoreResult
            if ($restoreResult.Success) {
                Write-Log $restoreResult.Message 'Success'
            }
            else {
                Write-Log $restoreResult.Message 'Error'
            }
        }

        $journalItems = @()
        foreach ($plan in $plans) {
            $scopeResult = $results | Where-Object { $_.ScopeName -eq $plan.ScopeName } | Select-Object -First 1
            $journalItems += New-JournalItem -ActionType 'Restore' -OriginalPath $plan.TargetPath -NewPath $plan.TargetPath -BackupPath $plan.RollbackPath -Result $(if ($scopeResult.Success) { 'Success' } else { 'Failed' }) -Message $scopeResult.Message
        }
        Add-JournalEntry -Type 'Restore' -Description "Restore backup $($latestBackup.Name)" -Items $journalItems -OperationId $operationId

        Write-Log "Pre-restore rollback snapshot preserved: $rollbackRoot" 'Info'
        Refresh-Items

        $failedResults = @($results | Where-Object { -not $_.Success })
        if ($failedResults.Count -gt 0 -or $planFailures.Count -gt 0 -or $skipped.Count -gt 0) {
            $details = @($failedResults.Message + $planFailures + $skipped)
            [System.Windows.MessageBox]::Show("Restore completed with warnings.`n`n$($details -join "`n")`n`nRollback snapshot:`n$rollbackRoot", "Restore Warning", "OK", "Warning")
        }
        else {
            Write-Log "Restore complete." 'Success'
            [System.Windows.MessageBox]::Show("Restore complete.`n`nRollback snapshot:`n$rollbackRoot", "Success", "OK", "Information")
        }
    }
    catch {
        Write-Log "Restore failed: $_" 'Error'
        [System.Windows.MessageBox]::Show("Restore failed: $_", "Error", "OK", "Error")
    }
    finally {
        if (Test-Path -LiteralPath $stagingRoot) {
            Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-ConfigurationSnapshot {
    $categorySnapshot = [ordered]@{}
    foreach ($category in $script:Categories.Keys) {
        $categorySnapshot[$category] = @($script:Categories[$category])
    }

    return [PSCustomObject]@{
        SchemaVersion = $Config.SettingsSchema
        SavedAt = (Get-Date).ToString('o')
        ScopeIndex = if ($cmbScope) { $cmbScope.SelectedIndex } else { 2 }
        ProfileRoot = if ($txtProfileRoot) { $txtProfileRoot.Text.Trim() } else { $Config.ProfileRoot }
        JunkPatterns = @($script:JunkPatterns)
        ProtectedFolders = @($script:ProtectedFolders)
        Categories = $categorySnapshot
    }
}

function Set-ObservableCollection {
    param(
        $Collection,
        [array]$Values
    )

    $Collection.Clear()
    foreach ($value in $Values) {
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            $Collection.Add([string]$value)
        }
    }
}

function Apply-ConfigurationSnapshot {
    param($Snapshot)

    if (-not $Snapshot) {
        throw "Configuration is empty."
    }
    if ($Snapshot.SchemaVersion -ne $Config.SettingsSchema) {
        throw "Unsupported configuration schema: $($Snapshot.SchemaVersion)"
    }
    if (-not $Snapshot.JunkPatterns -or -not $Snapshot.Categories) {
        throw "Configuration is missing required fields."
    }

    Set-ObservableCollection -Collection $script:JunkPatterns -Values @($Snapshot.JunkPatterns)
    if ($Snapshot.ProtectedFolders) {
        Set-ObservableCollection -Collection $script:ProtectedFolders -Values @($Snapshot.ProtectedFolders)
    }

    $newCategories = [ordered]@{}
    foreach ($property in $Snapshot.Categories.PSObject.Properties) {
        $newCategories[$property.Name] = @($property.Value)
    }
    if ($newCategories.Count -eq 0) {
        throw "Configuration must contain at least one category."
    }
    $script:Categories = $newCategories

    $previousLoadingState = $script:IsLoadingConfiguration
    $script:IsLoadingConfiguration = $true
    try {
        if ($cmbScope -and $Snapshot.PSObject.Properties.Name -contains 'ScopeIndex') {
            $scopeIndex = [int]$Snapshot.ScopeIndex
            if ($scopeIndex -ge 0 -and $scopeIndex -le 4) {
                $cmbScope.SelectedIndex = $scopeIndex
            }
        }
        if ($Snapshot.PSObject.Properties.Name -contains 'ProfileRoot') {
            $Config.ProfileRoot = [string]$Snapshot.ProfileRoot
            if ($txtProfileRoot) {
                $txtProfileRoot.Text = $Config.ProfileRoot
            }
        }
    }
    finally {
        $script:IsLoadingConfiguration = $previousLoadingState
    }
}

function Save-ApplicationConfiguration {
    try {
        $json = Get-ConfigurationSnapshot | ConvertTo-Json -Depth 8
        Write-AtomicFile -Path $Config.ConfigFile -Content $json
        Write-Log "Settings saved." 'Info'
    }
    catch {
        Write-Log "Failed to save settings: $($_.Exception.Message)" 'Error'
    }
}

function Load-ApplicationConfiguration {
    $raw = Read-WithBackupFallback -Path $Config.ConfigFile
    if (-not $raw) { return }

    try {
        $snapshot = $raw | ConvertFrom-Json
        Apply-ConfigurationSnapshot -Snapshot $snapshot
        Write-Log "Settings loaded." 'Success'
    }
    catch {
        Write-Log "Settings could not be parsed after recovery attempt." 'Warning'
    }
}

function Export-Configuration {
    $saveDialog = [System.Windows.Forms.SaveFileDialog]::new()
    $saveDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    $saveDialog.FileName = "StartMenuOrganizerConfig.json"

    if ($saveDialog.ShowDialog() -eq 'OK') {
        $json = Get-ConfigurationSnapshot | ConvertTo-Json -Depth 8
        [System.IO.File]::WriteAllText($saveDialog.FileName, $json, [System.Text.UTF8Encoding]::new($false))
        Write-Log "Configuration exported: $($saveDialog.FileName)" 'Success'
    }
}

function Import-Configuration {
    $openDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $openDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
    
    if ($openDialog.ShowDialog() -eq 'OK') {
        try {
            $config = Get-Content -LiteralPath $openDialog.FileName -Raw | ConvertFrom-Json
            Apply-ConfigurationSnapshot -Snapshot $config
            Save-ApplicationConfiguration
            Refresh-CategoryUI
            Refresh-JunkPatternsUI
            Refresh-CategoryPatternsUI
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

    Save-ApplicationConfiguration
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
$btnCancelWork.Add_Click({ Stop-BackgroundWorker })
$cmbScope.Add_SelectionChanged({
    if ($script:IsLoadingConfiguration) { return }
    Save-ApplicationConfiguration
    Refresh-Items
})
$txtProfileRoot.Add_LostFocus({
    if ($script:IsLoadingConfiguration) { return }
    $Config.ProfileRoot = $txtProfileRoot.Text.Trim()
    Save-ApplicationConfiguration
    if ($cmbScope.SelectedIndex -eq 3) {
        Refresh-Items
    }
})
$btnBrowseProfileRoot.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = "Select a local or offline Windows profile folder"
    $dialog.ShowNewFolderButton = $false
    if (-not [string]::IsNullOrWhiteSpace($txtProfileRoot.Text) -and (Test-Path -LiteralPath $txtProfileRoot.Text -PathType Container)) {
        $dialog.SelectedPath = $txtProfileRoot.Text
    }

    if ($dialog.ShowDialog() -eq 'OK') {
        $txtProfileRoot.Text = $dialog.SelectedPath
        $Config.ProfileRoot = $dialog.SelectedPath
        Save-ApplicationConfiguration
        if ($cmbScope.SelectedIndex -eq 3) {
            Refresh-Items
        }
    }
})
$btnUseDefaultProfileRoot.Add_Click({
    $txtProfileRoot.Text = Get-DefaultUserProfileRoot
    $Config.ProfileRoot = $txtProfileRoot.Text
    Save-ApplicationConfiguration
    if ($cmbScope.SelectedIndex -eq 3) {
        Refresh-Items
    }
})
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
$btnExportPlan.Add_Click({ Export-CurrentOperationPlan })
$btnImportPlan.Add_Click({ Import-OperationPlan })
$btnExecutePlan.Add_Click({ Invoke-CurrentOperationPlan })
$btnClearPlan.Add_Click({ Clear-OperationPlan })

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
        Save-ApplicationConfiguration
        Refresh-Items
    }
})

$btnRemoveJunkPattern.Add_Click({
    $selected = $lstJunkPatterns.SelectedItem
    if ($selected) {
        $script:JunkPatterns.Remove($selected)
        Refresh-JunkPatternsUI
        Write-Log "Removed junk pattern: $selected" 'Success'
        Save-ApplicationConfiguration
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
            Save-ApplicationConfiguration
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
        Save-ApplicationConfiguration
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
                $operationId = New-OperationId
                $journalItem = Invoke-JournaledRename -SourcePath $selected.FullPath -NewName "$newName$ext" -OperationId $operationId
                Add-JournalEntry -Type 'Rename' -Description "Rename $($selected.DisplayName)" -Items @($journalItem) -OperationId $operationId
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
    param($source, $e)
    $null = $source

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

Initialize-UiLocalizationAndAccessibility

# Set admin status
if ($script:IsAdmin) {
    $txtAdminStatus.Text = Get-UiString -Key 'Status.AdminFullAccess'
    $txtAdminStatus.Foreground = [System.Windows.Media.Brushes]::LightGreen
}
else {
    $txtAdminStatus.Text = Get-UiString -Key 'Status.StandardUser'
    $txtAdminStatus.Foreground = [System.Windows.Media.Brushes]::Orange
}

# Ensure config directory exists
$configDir = Split-Path $Config.ConfigFile -Parent
if (-not (Test-Path $configDir)) {
    New-Item -Path $configDir -ItemType Directory -Force | Out-Null
}
Initialize-FileLogging | Out-Null
Register-UnhandledExceptionHandlers
Load-ApplicationConfiguration
Load-UndoJournal

# Populate UI
Refresh-CategoryUI
Refresh-JunkPatternsUI
Refresh-CategoryPatternsUI

# Initial scan
Refresh-Items
Write-Log "Start Menu Organizer v$($Config.Version) initialized" 'Success'
Write-Log "File log: $script:LogFilePath" 'Info'
Write-Log "Keyboard shortcuts: Del=Delete, Ctrl+A=Select All, Ctrl+Z=Undo, Ctrl+F=Search, F5=Refresh" 'Info'

# Show window
try {
    $Window.ShowDialog() | Out-Null
}
catch {
    $crashPath = Write-CrashLog -Exception $_ -Context 'Fatal startup/runtime exception'
    $message = "Start Menu Organizer encountered an unexpected error."
    if ($crashPath) {
        $message = "$message`n`nCrash log:`n$crashPath"
    }
    [System.Windows.MessageBox]::Show($message, "Start Menu Organizer Error", "OK", "Error") | Out-Null
}
finally {
    if ($script:WScriptShell) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:WScriptShell) | Out-Null
    }
}
