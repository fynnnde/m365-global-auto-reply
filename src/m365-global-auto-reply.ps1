param([switch]$Mock)

$ErrorActionPreference = 'Stop'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072 } catch { }
try { $Host.UI.RawUI.WindowTitle = 'm365-global-auto-reply' } catch { }
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
if ($env:M365AR_MOCK) { $Mock = $true }

$base = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent ([Environment]::GetCommandLineArgs()[0]) }
$builtInDir = @("$base\templates", "$base\..\templates") | Where-Object { Test-Path $_ } | Select-Object -First 1
$dataDir = Join-Path $env:LOCALAPPDATA 'm365-global-auto-reply'
$userDir = Join-Path $dataDir 'templates'
$logDir = Join-Path $dataDir 'logs'
$logFile = Join-Path $logDir ('{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
New-Item -ItemType Directory -Force $userDir, $logDir | Out-Null
$utf8Bom = New-Object Text.UTF8Encoding $true

# EMBEDDED-ASSETS
if (-not $embeddedAssets) { $embeddedAssets = @{} }

# The built exe carries its icon and font inside, they are unpacked once so no extra files have to ship with it
$assetDir = Join-Path $dataDir 'assets'
foreach ($name in $embeddedAssets.Keys) {
    $target = Join-Path $assetDir $name
    if (-not (Test-Path $target)) {
        New-Item -ItemType Directory -Force $assetDir | Out-Null
        [IO.File]::WriteAllBytes($target, [Convert]::FromBase64String($embeddedAssets[$name]))
    }
}

# Stand-alone use: if no template folder ships next to the script, keep a default template in the user folder
$defaultTemplate = '<!doctype html><html><head><meta charset="utf-8"></head><body style="font-family: Arial, sans-serif; font-size: 12px; color: #333;"><p><b>Template</b></p><p>This is a template. Replace this text with your own auto-reply message in the Templates tab.</p><p>Best regards,<br>Your team</p></body></html>'
if (-not $builtInDir -and -not (Get-ChildItem $userDir -Filter *.html -ErrorAction SilentlyContinue)) {
    [IO.File]::WriteAllText((Join-Path $userDir 'default.html'), $defaultTemplate, $utf8Bom)
}

Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

public class Mailbox : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler PropertyChanged;
    bool selected, active;
    string state = "", period = "", snippet = "";

    public string Name { get; set; }
    public string Upn { get; set; }
    public string Html { get; set; }
    public string StartText { get; set; }
    public string EndText { get; set; }
    public bool Selected { get { return selected; } set { selected = value; Raise("Selected"); } }
    public bool Active { get { return active; } set { active = value; Raise("Active"); Raise("ToggleLabel"); } }
    public static string EnableText = "Enable", DisableText = "Disable";
    public string ToggleLabel { get { return active ? DisableText : EnableText; } }
    public string State { get { return state; } set { state = value; Raise("State"); } }
    public string Period { get { return period; } set { period = value; Raise("Period"); } }
    public string Snippet { get { return snippet; } set { snippet = value; Raise("Snippet"); } }

    void Raise(string n)
    {
        if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(n));
    }
}

public class LogLine
{
    public string Text { get; set; }
    public string Kind { get; set; }
}

public static class Dwm
{
    [DllImport("dwmapi.dll")]
    static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

    public static void Dark(IntPtr hwnd)
    {
        int v = 1;
        DwmSetWindowAttribute(hwnd, 20, ref v, 4);
    }
}
'@

# ---------------------------------------------------------------- xaml
$styles = @'
<Style x:Key="Muted" TargetType="TextBlock">
  <Setter Property="Foreground" Value="{DynamicResource MutedBrush}"/>
</Style>
<Style x:Key="H" TargetType="TextBlock">
  <Setter Property="Foreground" Value="{DynamicResource MutedBrush}"/>
  <Setter Property="FontSize" Value="11"/>
  <Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="Margin" Value="0,0,0,10"/>
</Style>
<Style x:Key="Card" TargetType="Border">
  <Setter Property="Background" Value="{DynamicResource CardBrush}"/>
  <Setter Property="BorderBrush" Value="{DynamicResource LineBrush}"/>
  <Setter Property="BorderThickness" Value="1"/>
  <Setter Property="CornerRadius" Value="10"/>
  <Setter Property="Padding" Value="16"/>
</Style>

<Style TargetType="Button">
  <Setter Property="Background" Value="{DynamicResource CardAltBrush}"/>
  <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
  <Setter Property="BorderBrush" Value="{DynamicResource LineBrush}"/>
  <Setter Property="BorderThickness" Value="1"/>
  <Setter Property="Padding" Value="14,7"/>
  <Setter Property="Cursor" Value="Hand"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="Button">
        <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="2" Padding="{TemplateBinding Padding}">
          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Opacity" Value="0.82"/></Trigger>
          <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.4"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>
<Style x:Key="Primary" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Background" Value="{DynamicResource AccentBrush}"/>
  <Setter Property="Foreground" Value="{DynamicResource AccentTextBrush}"/>
  <Setter Property="BorderThickness" Value="0"/>
  <Setter Property="FontWeight" Value="SemiBold"/>
</Style>
<Style x:Key="MsButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
  <Setter Property="Background" Value="White"/>
  <Setter Property="Foreground" Value="#3b3b3b"/>
  <Setter Property="BorderBrush" Value="#8c8c8c"/>
  <Setter Property="Padding" Value="22,11"/>
  <Setter Property="FontSize" Value="14"/>
  <Setter Property="FontWeight" Value="SemiBold"/>
</Style>

<Style TargetType="TextBox">
  <Setter Property="Background" Value="{DynamicResource CardAltBrush}"/>
  <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
  <Setter Property="BorderBrush" Value="{DynamicResource LineBrush}"/>
  <Setter Property="CaretBrush" Value="{DynamicResource TextBrush}"/>
  <Setter Property="Padding" Value="9,7"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="TextBox">
        <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                BorderThickness="1" CornerRadius="2">
          <ScrollViewer x:Name="PART_ContentHost" Padding="{TemplateBinding Padding}"/>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{DynamicResource AccentBrush}"/></Trigger>
          <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.45"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>

<Style TargetType="RadioButton">
  <Setter Property="Margin" Value="0,3"/>
  <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
  <Setter Property="Cursor" Value="Hand"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="RadioButton">
        <StackPanel x:Name="root" Orientation="Horizontal" Background="Transparent">
          <Grid Width="16" Height="16" VerticalAlignment="Center">
            <Ellipse x:Name="ring" Stroke="{DynamicResource MutedBrush}" StrokeThickness="1.5"/>
            <Ellipse x:Name="dot" Width="8" Height="8" Fill="{DynamicResource AccentBrush}" Visibility="Collapsed"/>
          </Grid>
          <ContentPresenter Margin="9,0,0,0" VerticalAlignment="Center"/>
        </StackPanel>
        <ControlTemplate.Triggers>
          <Trigger Property="IsChecked" Value="True">
            <Setter TargetName="dot" Property="Visibility" Value="Visible"/>
            <Setter TargetName="ring" Property="Stroke" Value="{DynamicResource AccentBrush}"/>
          </Trigger>
          <Trigger Property="IsEnabled" Value="False"><Setter TargetName="root" Property="Opacity" Value="0.45"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>

<Style x:Key="Tab" TargetType="RadioButton">
  <Setter Property="Foreground" Value="{DynamicResource MutedBrush}"/>
  <Setter Property="FontWeight" Value="SemiBold"/>
  <Setter Property="FontSize" Value="14"/>
  <Setter Property="Cursor" Value="Hand"/>
  <Setter Property="Margin" Value="0,0,26,0"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="RadioButton">
        <StackPanel Background="Transparent">
          <ContentPresenter Margin="0,2,0,8"/>
          <Border x:Name="line" Height="2" Background="Transparent"/>
        </StackPanel>
        <ControlTemplate.Triggers>
          <Trigger Property="IsChecked" Value="True">
            <Setter TargetName="line" Property="Background" Value="{DynamicResource AccentBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
          </Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>

<DataTemplate x:Key="FlagDe">
  <Canvas Width="22" Height="15">
    <Rectangle Width="22" Height="5" Fill="#000000"/>
    <Rectangle Width="22" Height="5" Canvas.Top="5" Fill="#dd0000"/>
    <Rectangle Width="22" Height="5" Canvas.Top="10" Fill="#ffce00"/>
  </Canvas>
</DataTemplate>
<DataTemplate x:Key="FlagEn">
  <Canvas Width="22" Height="15" ClipToBounds="True">
    <Rectangle Width="22" Height="15" Fill="#012169"/>
    <Line X1="0" Y1="0" X2="22" Y2="15" Stroke="White" StrokeThickness="3"/>
    <Line X1="22" Y1="0" X2="0" Y2="15" Stroke="White" StrokeThickness="3"/>
    <Line X1="0" Y1="0" X2="22" Y2="15" Stroke="#c8102e" StrokeThickness="1"/>
    <Line X1="22" Y1="0" X2="0" Y2="15" Stroke="#c8102e" StrokeThickness="1"/>
    <Rectangle Width="22" Height="5" Canvas.Top="5" Fill="White"/>
    <Rectangle Width="5" Height="15" Canvas.Left="8.5" Fill="White"/>
    <Rectangle Width="22" Height="3" Canvas.Top="6" Fill="#c8102e"/>
    <Rectangle Width="3" Height="15" Canvas.Left="9.5" Fill="#c8102e"/>
  </Canvas>
</DataTemplate>
<Style x:Key="Plain" TargetType="Button">
  <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
  <Setter Property="Cursor" Value="Hand"/>
  <Setter Property="HorizontalContentAlignment" Value="Left"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="Button">
        <Border x:Name="b" Background="Transparent" Padding="{TemplateBinding Padding}">
          <ContentPresenter HorizontalAlignment="Left" VerticalAlignment="Center"/>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="{DynamicResource HoverBrush}"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>

<Style TargetType="CheckBox">
  <Setter Property="Margin" Value="0,3"/>
  <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
  <Setter Property="Cursor" Value="Hand"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="CheckBox">
        <StackPanel x:Name="root" Orientation="Horizontal" Background="Transparent">
          <Border x:Name="box" Width="16" Height="16" CornerRadius="1" BorderThickness="1.5"
                  BorderBrush="{DynamicResource MutedBrush}" VerticalAlignment="Center">
            <Path x:Name="tick" Data="M0,4.5 L3.6,8 L10,0.8" Stretch="Uniform" Margin="2.5" HorizontalAlignment="Center" VerticalAlignment="Center"
                  Stroke="{DynamicResource AccentTextBrush}" StrokeThickness="2" StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Visibility="Collapsed"/>
          </Border>
          <ContentPresenter Margin="9,0,0,0" VerticalAlignment="Center"/>
        </StackPanel>
        <ControlTemplate.Triggers>
          <Trigger Property="IsChecked" Value="True">
            <Setter TargetName="tick" Property="Visibility" Value="Visible"/>
            <Setter TargetName="box" Property="Background" Value="{DynamicResource AccentBrush}"/>
            <Setter TargetName="box" Property="BorderBrush" Value="{DynamicResource AccentBrush}"/>
          </Trigger>
          <Trigger Property="IsEnabled" Value="False"><Setter TargetName="root" Property="Opacity" Value="0.45"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>

<Style TargetType="ListBox">
  <Setter Property="Background" Value="Transparent"/>
  <Setter Property="BorderThickness" Value="0"/>
  <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
  <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="ListBox">
        <ScrollViewer Focusable="False" CanContentScroll="True"><ItemsPresenter/></ScrollViewer>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>
<Style TargetType="ListBoxItem">
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="ListBoxItem">
        <Border x:Name="bd" CornerRadius="2" Padding="8,6" Background="Transparent">
          <ContentPresenter/>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource HoverBrush}"/></Trigger>
          <Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="{DynamicResource SelBrush}"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>

<Style x:Key="NoBtn" TargetType="RepeatButton">
  <Setter Property="Template">
    <Setter.Value><ControlTemplate TargetType="RepeatButton"><Border Background="Transparent"/></ControlTemplate></Setter.Value>
  </Setter>
</Style>
<Style x:Key="ThumbStyle" TargetType="Thumb">
  <Setter Property="Template">
    <Setter.Value><ControlTemplate TargetType="Thumb"><Border CornerRadius="1" Margin="1,0" Background="{DynamicResource ThumbBrush}"/></ControlTemplate></Setter.Value>
  </Setter>
</Style>
<Style TargetType="ScrollBar">
  <Setter Property="Width" Value="9"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="ScrollBar">
        <Grid Background="Transparent">
          <Track x:Name="PART_Track" IsDirectionReversed="True">
            <Track.DecreaseRepeatButton><RepeatButton Command="ScrollBar.PageUpCommand" Style="{StaticResource NoBtn}"/></Track.DecreaseRepeatButton>
            <Track.IncreaseRepeatButton><RepeatButton Command="ScrollBar.PageDownCommand" Style="{StaticResource NoBtn}"/></Track.IncreaseRepeatButton>
            <Track.Thumb><Thumb Style="{StaticResource ThumbStyle}"/></Track.Thumb>
          </Track>
        </Grid>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>

<Style TargetType="ProgressBar">
  <Setter Property="Height" Value="6"/>
  <Setter Property="Template">
    <Setter.Value>
      <ControlTemplate TargetType="ProgressBar">
        <Grid>
          <Border CornerRadius="0" Background="{DynamicResource LineBrush}"/>
          <Border x:Name="PART_Track">
            <Border x:Name="PART_Indicator" HorizontalAlignment="Left" CornerRadius="0" Background="{DynamicResource AccentBrush}"/>
          </Border>
        </Grid>
      </ControlTemplate>
    </Setter.Value>
  </Setter>
</Style>
'@

$mainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="m365-global-auto-reply" Width="1280" Height="800" WindowState="Maximized" MinWidth="1000" MinHeight="640"
        FontFamily="Segoe UI" FontSize="13"
        Background="{DynamicResource BgBrush}" Foreground="{DynamicResource TextBrush}"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
  <Window.Resources><!--STYLES--></Window.Resources>
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="60"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <Border Background="{DynamicResource CardBrush}" BorderBrush="{DynamicResource LineBrush}" BorderThickness="0,0,0,1">
      <Grid Margin="24,0">
        <TextBlock Text="One message for every mailbox. Set it once, and nobody is left on read."
                     VerticalAlignment="Center" FontFamily="Segoe UI Semilight" FontSize="19" Foreground="White"/>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
          <StackPanel x:Name="userBox" Orientation="Horizontal" Visibility="Collapsed" Margin="0,0,18,0">
            <TextBlock x:Name="connText" Style="{StaticResource Muted}" VerticalAlignment="Center" Margin="0,0,14,0"/>
            <Button x:Name="btnDisconnect" Padding="12,5"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf2f5;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="Sign out" VerticalAlignment="Center"/></StackPanel></Button>
          </StackPanel>
          <Button x:Name="btnLang" Padding="10,5">
            <StackPanel Orientation="Horizontal">
              <ContentControl x:Name="flagCur" Content="x" ContentTemplate="{StaticResource FlagEn}" VerticalAlignment="Center"/>
              <TextBlock x:Name="langText" Text="English" Visibility="Collapsed"/>
              <Path Data="M0,0 L4,4 L8,0" Stroke="{DynamicResource MutedBrush}" StrokeThickness="1.5" VerticalAlignment="Center" Margin="10,0,0,0"/>
            </StackPanel>
          </Button>
          <Popup x:Name="popLang" PlacementTarget="{Binding ElementName=btnLang}" Placement="Bottom" StaysOpen="False" HorizontalOffset="0" VerticalOffset="4">
            <Border Background="{DynamicResource CardAltBrush}" BorderBrush="{DynamicResource LineBrush}" BorderThickness="1" Padding="4" MinWidth="150">
              <StackPanel>
                <Button x:Name="langEn" Style="{StaticResource Plain}" Padding="8,6">
                  <StackPanel Orientation="Horizontal"><ContentControl Content="x" ContentTemplate="{StaticResource FlagEn}" VerticalAlignment="Center"/><TextBlock Text="English" Margin="10,0,0,0"/></StackPanel>
                </Button>
                <Button x:Name="langDe" Style="{StaticResource Plain}" Padding="8,6">
                  <StackPanel Orientation="Horizontal"><ContentControl Content="x" ContentTemplate="{StaticResource FlagDe}" VerticalAlignment="Center"/><TextBlock Text="Deutsch" Margin="10,0,0,0"/></StackPanel>
                </Button>
              </StackPanel>
            </Border>
          </Popup>
        </StackPanel>
      </Grid>
    </Border>

    <Grid x:Name="loginPane" Grid.Row="1">
      <StackPanel HorizontalAlignment="Center" VerticalAlignment="Center" Width="380">
        <TextBlock Text="Sign in to Microsoft 365" FontSize="22" FontWeight="SemiBold" HorizontalAlignment="Center"/>
        <TextBlock Text="Connect to Exchange Online to see and manage the automatic replies of your mailboxes. An Exchange administrator account is required."
                   Style="{StaticResource Muted}" TextWrapping="Wrap" TextAlignment="Center" Margin="0,10,0,26"/>
        <Button x:Name="btnConnect" Style="{StaticResource MsButton}" HorizontalAlignment="Center" AutomationProperties.Name="Sign in with Microsoft">
          <StackPanel Orientation="Horizontal">
            <Canvas Width="21" Height="21" Margin="0,0,12,0">
              <Rectangle Width="10" Height="10" Fill="#f25022"/>
              <Rectangle Width="10" Height="10" Fill="#7fba00" Canvas.Left="11"/>
              <Rectangle Width="10" Height="10" Fill="#00a4ef" Canvas.Top="11"/>
              <Rectangle Width="10" Height="10" Fill="#ffb900" Canvas.Left="11" Canvas.Top="11"/>
            </Canvas>
            <TextBlock Text="Sign in with Microsoft" VerticalAlignment="Center"/>
          </StackPanel>
        </Button>
        <TextBlock x:Name="loginStatus" Style="{StaticResource Muted}" TextWrapping="Wrap" TextAlignment="Center" Margin="0,16,0,0"/>
      </StackPanel>
    </Grid>

    <Grid x:Name="mainPane" Grid.Row="1" Margin="20" Visibility="Collapsed">
      <Grid.RowDefinitions>
        <RowDefinition Height="*"/>
        <RowDefinition Height="16"/>
        <RowDefinition Height="180"/>
      </Grid.RowDefinitions>

      <Border Style="{StaticResource Card}">
        <Grid>
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <Grid Margin="0,0,0,10">
            <StackPanel Orientation="Horizontal">
              <RadioButton x:Name="tabBoxes" Content="Mailboxes" Style="{StaticResource Tab}" IsChecked="True"/>
              <RadioButton x:Name="tabEditor" Content="Templates" Style="{StaticResource Tab}"/>
            </StackPanel>
            <Button x:Name="btnNew" Style="{StaticResource Primary}" HorizontalAlignment="Right" Padding="16,8"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf067;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="New auto-reply" VerticalAlignment="Center"/></StackPanel></Button>
          </Grid>

          <Grid x:Name="paneBoxes" Grid.Row="1">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid Margin="0,0,0,10">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="txtFilter" ToolTip="Filter by name, address or message text"/>
              <CheckBox x:Name="chkActiveOnly" Grid.Column="1" Content="Active only" Margin="16,0,0,0"/>
              <StackPanel Grid.Column="2" Orientation="Horizontal" Margin="16,0,0,0">
                <Button x:Name="btnSelAll" Content="Check shown" Padding="10,6"/>
                <Button x:Name="btnSelNone" Content="Uncheck all" Padding="10,6" Margin="6,0,0,0"/>
                <Button x:Name="btnDisableChecked" Padding="10,6" Margin="6,0,0,0"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf011;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="Disable checked" VerticalAlignment="Center"/></StackPanel></Button>
                <Button x:Name="btnRefresh" Padding="10,6" Margin="6,0,0,0"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf2f1;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="Refresh" VerticalAlignment="Center"/></StackPanel></Button>
              </StackPanel>
            </Grid>
            <Grid Grid.Row="1" Margin="8,0,25,4">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="26"/><ColumnDefinition Width="*" MinWidth="200"/><ColumnDefinition Width="1.4*" MinWidth="160"/><ColumnDefinition Width="104"/><ColumnDefinition Width="160"/><ColumnDefinition Width="210"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="1" Text="MAILBOX" Style="{StaticResource H}" Margin="0"/>
              <TextBlock Grid.Column="2" Text="CURRENT MESSAGE" Style="{StaticResource H}" Margin="0"/>
              <TextBlock Grid.Column="3" Text="STATE" Style="{StaticResource H}" Margin="0"/>
              <TextBlock Grid.Column="4" Text="PERIOD" Style="{StaticResource H}" Margin="0"/>
            </Grid>
            <ListBox x:Name="lstBoxes" Grid.Row="2" ScrollViewer.VerticalScrollBarVisibility="Visible" VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling">
              <ListBox.ItemTemplate>
                <DataTemplate>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="26"/><ColumnDefinition Width="*" MinWidth="200"/><ColumnDefinition Width="1.4*" MinWidth="160"/><ColumnDefinition Width="104"/><ColumnDefinition Width="160"/><ColumnDefinition Width="210"/>
                    </Grid.ColumnDefinitions>
                    <CheckBox IsChecked="{Binding Selected, Mode=TwoWay}" VerticalAlignment="Center" Margin="0"/>
                    <StackPanel Grid.Column="1" Margin="0,0,10,0">
                      <TextBlock Text="{Binding Name}" TextTrimming="CharacterEllipsis"/>
                      <TextBlock Text="{Binding Upn}" Style="{StaticResource Muted}" FontSize="11" TextTrimming="CharacterEllipsis"/>
                    </StackPanel>
                    <TextBlock Grid.Column="2" Text="{Binding Snippet}" Style="{StaticResource Muted}" TextTrimming="CharacterEllipsis" VerticalAlignment="Center" Margin="0,0,12,0"/>
                    <StackPanel Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Center">
                      <Ellipse x:Name="pill" Width="8" Height="8" Fill="{DynamicResource MutedBrush}" Margin="0,1,7,0"/>
                      <TextBlock x:Name="stateText" Text="{Binding State}" Style="{StaticResource Muted}"/>
                    </StackPanel>
                    <TextBlock Grid.Column="4" Text="{Binding Period}" Style="{StaticResource Muted}" FontSize="11" VerticalAlignment="Center"/>
                    <StackPanel Grid.Column="5" Orientation="Horizontal" HorizontalAlignment="Right">
                      <Button x:Name="rowEdit" Padding="10,3" FontSize="12"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf304;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="Edit" VerticalAlignment="Center"/></StackPanel></Button>
                      <Button x:Name="rowToggle" Padding="10,3" FontSize="12" Margin="6,0,0,0"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf011;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="{Binding ToggleLabel}" VerticalAlignment="Center"/></StackPanel></Button>
                    </StackPanel>
                  </Grid>
                  <DataTemplate.Triggers>
                    <DataTrigger Binding="{Binding State}" Value="On">
                      <Setter TargetName="pill" Property="Fill" Value="{DynamicResource OkBrush}"/>
                      <Setter TargetName="stateText" Property="Foreground" Value="{DynamicResource TextBrush}"/>
                      <DataTrigger.EnterActions>
                        <BeginStoryboard x:Name="blink">
                          <Storyboard RepeatBehavior="Forever" AutoReverse="True" Timeline.DesiredFrameRate="10">
                            <DoubleAnimation Storyboard.TargetName="pill" Storyboard.TargetProperty="Opacity" From="1" To="0.2" Duration="0:0:0.9"/>
                          </Storyboard>
                        </BeginStoryboard>
                      </DataTrigger.EnterActions>
                      <DataTrigger.ExitActions>
                        <StopStoryboard BeginStoryboardName="blink"/>
                      </DataTrigger.ExitActions>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding State}" Value="Off">
                      <Setter TargetName="pill" Property="Fill" Value="{DynamicResource ErrBrush}"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding Active}" Value="True">
                      <Setter TargetName="rowToggle" Property="Background" Value="{DynamicResource ErrBrush}"/>
                      <Setter TargetName="rowToggle" Property="BorderBrush" Value="{DynamicResource ErrBrush}"/>
                      <Setter TargetName="rowToggle" Property="Foreground" Value="#2a0c0e"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding Active}" Value="False">
                      <Setter TargetName="rowToggle" Property="Background" Value="{DynamicResource OkBrush}"/>
                      <Setter TargetName="rowToggle" Property="BorderBrush" Value="{DynamicResource OkBrush}"/>
                      <Setter TargetName="rowToggle" Property="Foreground" Value="#12281b"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding State}" Value="Scheduled">
                      <Setter TargetName="pill" Property="Fill" Value="{DynamicResource WarnBrush}"/>
                      <Setter TargetName="stateText" Property="Foreground" Value="{DynamicResource TextBrush}"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding State}" Value="Error">
                      <Setter TargetName="pill" Property="Fill" Value="{DynamicResource ErrBrush}"/>
                      <Setter TargetName="stateText" Property="Foreground" Value="{DynamicResource ErrBrush}"/>
                    </DataTrigger>
                  </DataTemplate.Triggers>
                </DataTemplate>
              </ListBox.ItemTemplate>
            </ListBox>
            <TextBlock x:Name="lblSummary" Grid.Row="3" Style="{StaticResource Muted}" FontSize="12" Margin="0,10,0,0"/>
          </Grid>

          <Grid x:Name="paneEditor" Grid.Row="1" Visibility="Collapsed">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="230"/><ColumnDefinition Width="16"/><ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <TextBlock Text="ALL TEMPLATES" Style="{StaticResource H}" VerticalAlignment="Center" Margin="0"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                  <Button x:Name="btnTplNew" ToolTip="New" Padding="10,6"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" VerticalAlignment="Center" Text="&#xf067;"/></Button>
                  <Button x:Name="btnTplOpen" ToolTip="Open file..." Padding="10,6" Margin="4,0,0,0"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" VerticalAlignment="Center" Text="&#xf07c;"/></Button>
                  <Button x:Name="btnTplDelete" ToolTip="Delete" Padding="10,6" Margin="4,0,0,0"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" VerticalAlignment="Center" Text="&#xf1f8;"/></Button>
                </StackPanel>
              </Grid>
              <ListBox x:Name="lstEdit" Grid.Row="1" DisplayMemberPath="Name"/>
            </Grid>
            <Grid Grid.Column="2">
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
              <Grid Margin="0,0,0,10">
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <TextBox x:Name="txtTplName" ToolTip="Template name"/>
                <CheckBox x:Name="chkSource" Grid.Column="1" Content="HTML source" Margin="16,0,0,0"/>
                <Button x:Name="btnTplSave" Grid.Column="2" Style="{StaticResource Primary}" Margin="16,0,0,0"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf0c7;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="Save template" VerticalAlignment="Center"/></StackPanel></Button>
              </Grid>
              <WrapPanel x:Name="editToolbar" Grid.Row="1" Margin="0,0,0,6"/>
              <Grid Grid.Row="2">
                <Border BorderBrush="{DynamicResource LineBrush}" BorderThickness="1" Background="White">
                  <WebBrowser x:Name="editBrowser"/>
                </Border>
                <TextBox x:Name="txtSource" Visibility="Collapsed" AcceptsReturn="True" AcceptsTab="True" TextWrapping="Wrap"
                         VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="13" VerticalContentAlignment="Top"/>
              </Grid>
            </Grid>
          </Grid>
        </Grid>
      </Border>

      <Border Grid.Row="2" Style="{StaticResource Card}">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <Grid>
            <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
            <TextBlock Text="ACTIVITY" Style="{StaticResource H}"/>
            <TextBlock x:Name="lblProgress" Grid.Column="1" Style="{StaticResource Muted}" FontSize="11" Margin="12,0,0,10" HorizontalAlignment="Left"/>
            <StackPanel Grid.Column="2" Orientation="Horizontal" Margin="0,-6,0,0">
              <Button x:Name="btnCancel" Padding="10,3" Visibility="Collapsed" Margin="0,0,6,0"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf00d;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="Cancel job" VerticalAlignment="Center"/></StackPanel></Button>
              <Button x:Name="btnCopyLog" Padding="10,3"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf0c5;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="Copy" VerticalAlignment="Center"/></StackPanel></Button>
              <Button x:Name="btnOpenLog" Padding="10,3" Margin="6,0,0,0"><StackPanel Orientation="Horizontal"><TextBlock FontFamily="{DynamicResource FA}" TextOptions.TextFormattingMode="Ideal" TextOptions.TextRenderingMode="Grayscale" Text="&#xf07c;" VerticalAlignment="Center" Margin="0,0,8,0"/><TextBlock Text="Open log folder" VerticalAlignment="Center"/></StackPanel></Button>
            </StackPanel>
          </Grid>
          <ProgressBar x:Name="bar" Grid.Row="1" Margin="0,2,0,10" Minimum="0" Maximum="100"/>
          <ListBox x:Name="lstLog" Grid.Row="2">
            <ListBox.ItemContainerStyle>
              <Style TargetType="ListBoxItem">
                <Setter Property="Template">
                  <Setter.Value><ControlTemplate TargetType="ListBoxItem"><ContentPresenter/></ControlTemplate></Setter.Value>
                </Setter>
              </Style>
            </ListBox.ItemContainerStyle>
            <ListBox.ItemTemplate>
              <DataTemplate>
                <TextBlock x:Name="t" Text="{Binding Text}" FontFamily="Consolas" FontSize="12" Foreground="{DynamicResource MutedBrush}" TextWrapping="Wrap"/>
                <DataTemplate.Triggers>
                  <DataTrigger Binding="{Binding Kind}" Value="ok"><Setter TargetName="t" Property="Foreground" Value="{DynamicResource OkBrush}"/></DataTrigger>
                  <DataTrigger Binding="{Binding Kind}" Value="err"><Setter TargetName="t" Property="Foreground" Value="{DynamicResource ErrBrush}"/></DataTrigger>
                  <DataTrigger Binding="{Binding Kind}" Value="warn"><Setter TargetName="t" Property="Foreground" Value="{DynamicResource WarnBrush}"/></DataTrigger>
                </DataTemplate.Triggers>
              </DataTemplate>
            </ListBox.ItemTemplate>
          </ListBox>
        </Grid>
      </Border>
    </Grid>
  </Grid>
</Window>
'@

$replyXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Auto-reply" Width="1260" Height="720" MinWidth="1100" MinHeight="660" ShowInTaskbar="False"
        WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="13"
        Background="{DynamicResource BgBrush}" Foreground="{DynamicResource TextBrush}" UseLayoutRounding="True">
  <Window.Resources><!--STYLES--></Window.Resources>
  <Grid Margin="24">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="330"/><ColumnDefinition Width="24"/><ColumnDefinition Width="330"/><ColumnDefinition Width="24"/><ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>

    <StackPanel Grid.ColumnSpan="5" Margin="0,0,0,18">
      <TextBlock x:Name="dlgTitle" FontSize="20" FontWeight="SemiBold"/>
      <TextBlock x:Name="dlgSub" Style="{StaticResource Muted}" Margin="0,2,0,0" TextTrimming="CharacterEllipsis"/>
    </StackPanel>

    <Grid Grid.Row="1">
      <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
      <TextBlock Text="MESSAGE" Style="{StaticResource H}"/>
      <Border Grid.Row="1" BorderBrush="{DynamicResource LineBrush}" BorderThickness="1">
        <ListBox x:Name="lstTplD" DisplayMemberPath="Name" Margin="4"/>
      </Border>
      <StackPanel Grid.Row="2" Margin="0,22,0,0">
        <TextBlock Text="OPTIONS" Style="{StaticResource H}"/>
        <CheckBox x:Name="chkOverwriteD" Content="Overwrite existing auto-replies"/>
        <TextBlock Text="Off: mailboxes that already have an active auto-reply are skipped." Style="{StaticResource Muted}" FontSize="11" Margin="25,0,0,6" TextWrapping="Wrap"/>
      </StackPanel>
    </Grid>

    <StackPanel Grid.Row="1" Grid.Column="2">
      <TextBlock Text="PERIOD" Style="{StaticResource H}"/>
      <RadioButton x:Name="rbPermD" Content="Until I turn it off" IsChecked="True"/>
      <RadioButton x:Name="rbUntilD" Content="From now until an end date"/>
      <RadioButton x:Name="rbRangeD" Content="Between a start and an end date"/>
      <StackPanel Margin="0,12,0,0">
        <TextBlock Text="Start (dd.MM.yyyy HH:mm)" Style="{StaticResource Muted}" FontSize="11" Margin="0,0,0,4"/>
        <TextBox x:Name="txtStartD"/>
      </StackPanel>
      <StackPanel Margin="0,10,0,0">
        <TextBlock Text="End (dd.MM.yyyy HH:mm)" Style="{StaticResource Muted}" FontSize="11" Margin="0,0,0,4"/>
        <TextBox x:Name="txtEndD"/>
      </StackPanel>
      <TextBlock x:Name="txtHintD" Foreground="{DynamicResource ErrBrush}" TextWrapping="Wrap" Margin="0,8,0,0" FontSize="12"/>

      <StackPanel x:Name="targetBox" Margin="0,22,0,0">
        <TextBlock Text="APPLY TO" Style="{StaticResource H}"/>
        <RadioButton x:Name="rbSelD" IsChecked="True"/>
        <RadioButton x:Name="rbAllD"/>
      </StackPanel>
    </StackPanel>

    <Grid Grid.Row="1" Grid.Column="4">
      <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
      <TextBlock Text="PREVIEW" Style="{StaticResource H}"/>
      <Border Grid.Row="1" BorderBrush="{DynamicResource LineBrush}" BorderThickness="1" Background="White">
        <WebBrowser x:Name="prevD"/>
      </Border>
    </Grid>

    <StackPanel Grid.Row="2" Grid.ColumnSpan="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,20,0,0">
      <Button x:Name="btnCancelD" Content="Cancel" Margin="0,0,10,0" Padding="20,8"/>
      <Button x:Name="btnOkD" Content="Apply" Style="{StaticResource Primary}" Padding="26,8"/>
    </StackPanel>
  </Grid>
</Window>
'@

$confirmXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Confirm" Width="480" SizeToContent="Height" ResizeMode="NoResize" ShowInTaskbar="False"
        WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" FontSize="13"
        Background="{DynamicResource BgBrush}" Foreground="{DynamicResource TextBrush}">
  <Window.Resources><!--STYLES--></Window.Resources>
  <StackPanel Margin="24">
    <TextBlock x:Name="msg" TextWrapping="Wrap" LineHeight="20"/>
    <StackPanel x:Name="guard" Margin="0,16,0,0" Visibility="Collapsed">
      <TextBlock x:Name="guardText" Style="{StaticResource Muted}" Margin="0,0,0,6"/>
      <TextBox x:Name="txtGuard"/>
    </StackPanel>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,22,0,0">
      <Button x:Name="no" Content="Cancel" Margin="0,0,8,0"/>
      <Button x:Name="yes" Content="Apply" Style="{StaticResource Primary}"/>
    </StackPanel>
  </StackPanel>
</Window>
'@

$viewXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Message preview" Width="760" Height="640" ShowInTaskbar="False" WindowStartupLocation="CenterOwner"
        FontFamily="Segoe UI" FontSize="13" Background="{DynamicResource BgBrush}" Foreground="{DynamicResource TextBrush}">
  <Window.Resources><!--STYLES--></Window.Resources>
  <Grid Margin="18">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
    <TextBlock x:Name="viewTitle" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,12"/>
    <Border Grid.Row="1" BorderBrush="{DynamicResource LineBrush}" BorderThickness="1" CornerRadius="2" Background="White">
      <WebBrowser x:Name="viewBrowser"/>
    </Border>
  </Grid>
</Window>
'@

$palette = @{
    Bg = '#141414'; Card = '#1e1e1e'; CardAlt = '#282828'; Line = '#333333'; Text = '#f2f2f2'; Muted = '#b3b3b3'
    Hover = '#2b2b2b'; Sel = '#223a2c'; Thumb = '#4a4a4a'
    Accent = '#8fd3a9'; AccentText = '#12281b'; Ok = '#7fd9a0'; Warn = '#e0b84a'; Err = '#ef7b7f'
}
$brushes = New-Object Windows.Media.BrushConverter

$de = @{
    'One message for every mailbox. Set it once, and nobody is left on read.' = 'Eine Nachricht für jedes Postfach. Einmal einrichten, und niemand bleibt unbeantwortet.'
    'Sign out' = 'Abmelden'
    'Sign in to Microsoft 365' = 'Bei Microsoft 365 anmelden'
    'Connect to Exchange Online to see and manage the automatic replies of your mailboxes. An Exchange administrator account is required.' = 'Verbinden Sie sich mit Exchange Online, um die automatischen Antworten Ihrer Postfächer zu sehen und zu verwalten. Ein Exchange-Administratorkonto ist erforderlich.'
    'Sign in with Microsoft' = 'Mit Microsoft anmelden'
    'Mailboxes' = 'Postfächer'
    'Templates' = 'Vorlagen'
    '+  New auto-reply' = '+  Neue automatische Antwort'
    'Filter by name, address or message text' = 'Nach Name, Adresse oder Nachrichtentext filtern'
    'Active only' = 'Nur aktive'
    'Check shown' = 'Angezeigte markieren'
    'Uncheck all' = 'Alle abwählen'
    'Disable checked' = 'Markierte deaktivieren'
    'Refresh' = 'Aktualisieren'
    'MAILBOX' = 'POSTFACH'
    'CURRENT MESSAGE' = 'AKTUELLE NACHRICHT'
    'STATE' = 'STATUS'
    'PERIOD' = 'ZEITRAUM'
    'Edit' = 'Bearbeiten'
    'ALL TEMPLATES' = 'ALLE VORLAGEN'
    'New' = 'Neu'
    'Delete' = 'Löschen'
    'Template name' = 'Name der Vorlage'
    'Save template' = 'Vorlage speichern'
    'Bold' = 'Fett'
    'Italic' = 'Kursiv'
    'Heading' = 'Überschrift'
    'List' = 'Liste'
    'Link' = 'Link'
    'Image' = 'Bild'
    'Divider' = 'Trennlinie'
    'ACTIVITY' = 'AKTIVITÄT'
    'Cancel job' = 'Abbrechen'
    'Copy' = 'Kopieren'
    'Open log folder' = 'Protokollordner öffnen'
    'Auto-reply' = 'Automatische Antwort'
    'MESSAGE' = 'NACHRICHT'
    'Template' = 'Vorlage'
    'Write a message' = 'Eigene Nachricht schreiben'
    'Messages are written in Markdown.' = 'Nachrichten werden in Markdown geschrieben.'
    'Until I turn it off' = 'Bis ich sie ausschalte'
    'From now until an end date' = 'Ab jetzt bis zu einem Enddatum'
    'Between a start and an end date' = 'Zwischen Start- und Enddatum'
    'Start (dd.MM.yyyy HH:mm)' = 'Start (TT.MM.JJJJ HH:mm)'
    'End (dd.MM.yyyy HH:mm)' = 'Ende (TT.MM.JJJJ HH:mm)'
    'OPTIONS' = 'OPTIONEN'
    'Overwrite existing auto-replies' = 'Bestehende automatische Antworten überschreiben'
    'Off: mailboxes that already have an active auto-reply are skipped.' = 'Aus: Postfächer mit aktiver automatischer Antwort werden übersprungen.'
    'APPLY TO' = 'ANWENDEN AUF'
    'PREVIEW' = 'VORSCHAU'
    'Cancel' = 'Abbrechen'
    'Apply' = 'Anwenden'
    'Confirm' = 'Bestätigen'
    'Message preview' = 'Nachrichtenvorschau'
    'Start must look like 24.12.2026 08:00' = 'Start muss so aussehen: 24.12.2026 08:00'
    'End must look like 27.12.2026 08:00' = 'Ende muss so aussehen: 27.12.2026 08:00'
    'End must be in the future and after the start' = 'Das Ende muss in der Zukunft und nach dem Start liegen'
    'Type {0} to confirm' = 'Zur Bestätigung {0} eingeben'
    'Insert image' = 'Bild einfügen'
    'Image URL (https://...). Leave empty to pick a file from disk instead.' = 'Bild-URL (https://...). Leer lassen, um stattdessen eine Datei von der Festplatte zu wählen.'
    'Current message of {0}' = 'Aktuelle Nachricht von {0}'
    'New auto-reply' = 'Neue automatische Antwort'
    'Choose the message, the period and the options.' = 'Wählen Sie Nachricht, Zeitraum und Optionen.'
    'Checked mailboxes ({0})' = 'Markierte Postfächer ({0})'
    'All mailboxes ({0})' = 'Alle Postfächer ({0})'
    'Edit auto-reply' = 'Automatische Antwort bearbeiten'
    'Enable auto-reply' = 'Automatische Antwort aktivieren'
    'Disable the auto-reply for {0}?' = 'Automatische Antwort für {0} deaktivieren?'
    'Disable the auto-reply for {0} mailboxes?' = 'Automatische Antwort für {0} Postfächer deaktivieren?'
    '{0} mailboxes will be changed.' = '{0} Postfächer werden geändert.'
    'until turned off' = 'bis zur Deaktivierung'
    '{0:dd.MM.yyyy HH:mm} to {1:dd.MM.yyyy HH:mm}' = '{0:dd.MM.yyyy HH:mm} bis {1:dd.MM.yyyy HH:mm}'
    'Set auto-reply, {0}' = 'Automatische Antwort setzen, {0}'
    'Set auto-reply for {0}, {1}' = 'Automatische Antwort für {0} setzen, {1}'
    '{0} mailboxes  |  {1} with an active auto-reply  |  {2} checked' = '{0} Postfächer  |  {1} mit aktiver Antwort  |  {2} markiert'
    'Connected as {0}' = 'Angemeldet als {0}'
    'Waiting for the Microsoft sign-in window...' = 'Warte auf das Microsoft-Anmeldefenster...'
    "Delete the template '{0}'?" = "Vorlage '{0}' löschen?"
    "A template named '{0}' already exists. Replace it?" = "Eine Vorlage namens '{0}' existiert bereits. Ersetzen?"
    'A job is still running. Cancel it and exit?' = 'Ein Vorgang läuft noch. Abbrechen und beenden?'
    ' (built-in)' = ' (mitgeliefert)'
    'Enable' = 'Aktivieren'
    'Disable' = 'Deaktivieren'
    'Open file...' = 'Datei öffnen...'
    'HTML source' = 'HTML-Quelltext'
    'Insert link' = 'Link einfügen'
    'Link address (https://...)' = 'Link-Adresse (https://...)'
    'Color as hex code, for example #1a73e8' = 'Farbe als Hex-Code, zum Beispiel #1a73e8'
    'Custom...' = 'Eigene...'
    'Font' = 'Schrift'
    'Size' = 'Größe'
    'Underline' = 'Unterstrichen'
    'Strikethrough' = 'Durchgestrichen'
    'Text color' = 'Textfarbe'
    'Highlight' = 'Hervorheben'
    'Left' = 'Links'
    'Align left' = 'Linksbündig'
    'Center' = 'Mitte'
    'Align right' = 'Rechtsbündig'
    'Right' = 'Rechts'
    'Bullets' = 'Liste'
    'Bulleted list' = 'Aufzählung'
    'Numbers' = 'Nummern'
    'Numbered list' = 'Nummerierung'
    'Indent' = 'Einrücken'
    'Outdent' = 'Ausrücken'
    'Unlink' = 'Link weg'
    'Remove link' = 'Link entfernen'
    'Line' = 'Linie'
    'Horizontal line' = 'Horizontale Linie'
    'Clear' = 'Leeren'
    'Clear formatting' = 'Formatierung entfernen'
    'Undo' = 'Zurück'
    'Redo' = 'Vor'
    'The Exchange Online PowerShell module is not installed. Install it now for your user account?' = 'Das PowerShell-Modul fÃ¼r Exchange Online ist nicht installiert. Jetzt fÃ¼r Ihr Benutzerkonto installieren?'
    'The Exchange Online module is required.' = 'Das Exchange-Online-Modul wird benÃ¶tigt.'
    'Image size' = 'BildgrÃ¶ÃŸe'
    'Float left' = 'Links umflieÃŸen'
    'Float right' = 'Rechts umflieÃŸen'
    'Direct link to the image, for example https://i.imgur.com/example.png' = 'Direkter Link zum Bild, zum Beispiel https://i.imgur.com/beispiel.png'
    'Upload the image to a site like imgur first. The link has to end with .png, .jpg, .jpeg or .gif.' = 'Lade das Bild zuerst bei einem Dienst wie imgur hoch. Der Link muss auf .png, .jpg, .jpeg oder .gif enden.'
    'Image links must be https and end with .png, .jpg, .jpeg or .gif' = 'Bild-Links mÃ¼ssen mit https beginnen und auf .png, .jpg, .jpeg oder .gif enden'
    'Align' = 'Ausrichtung'
    'Justify' = 'Blocksatz'
    'Lists' = 'Listen'
    'Paragraph' = 'Absatz'
    'Heading 1' = 'Überschrift 1'
    'Heading 2' = 'Überschrift 2'
    'Heading 3' = 'Überschrift 3'
}
$script:lang = 'en'
$strIndex = New-Object Collections.Hashtable

function T([string]$s) {
    if ($script:lang -eq 'de' -and $de.ContainsKey($s)) { $de[$s] } else { $s }
}

function Set-Strings($w) {
    foreach ($e in $strIndex.GetEnumerator()) { $w.Resources[$e.Value] = T $e.Key }
}

# Swaps literal texts in the xaml for resources so the language can change at runtime
function Register-Strings([string]$xaml) {
    [regex]::Replace($xaml, '(?<![\w.])(Text|Content|Title|ToolTip)="([^"{<&][^"]*)"', [Text.RegularExpressions.MatchEvaluator] {
            param($m)
            $value = $m.Groups[2].Value
            if ($value -eq 'm365-global-auto-reply' -or $value -notmatch '[A-Za-z]{2}') { return $m.Value }
            if (-not $strIndex.ContainsKey($value)) { $strIndex[$value] = 's' + $strIndex.Count }
            '{0}="{{DynamicResource {1}}}"' -f $m.Groups[1].Value, $strIndex[$value]
        })
}

$fontFile = @("$base\assets\fa-solid-900.ttf", "$base\..\assets\fa-solid-900.ttf", "$assetDir\fa-solid-900.ttf") | Where-Object { Test-Path $_ } | Select-Object -First 1
$iconFile = @("$base\assets\app.ico", "$base\..\assets\app.ico", "$assetDir\app.ico") | Where-Object { Test-Path $_ } | Select-Object -First 1
$faFont = [Windows.Media.FontFamily]::new($(if ($fontFile) { (New-Object Uri ((Split-Path (Resolve-Path $fontFile).Path -Parent) + '\')).AbsoluteUri + '#Font Awesome 6 Free Solid' } else { 'Segoe UI' }))
# Builds a window from xaml and applies colors, texts, font and icon
function New-Window($xaml) {
    $w = [Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader ([xml](Register-Strings $xaml.Replace('<!--STYLES-->', $styles)))))
    foreach ($k in $palette.Keys) { $w.Resources[$k + 'Brush'] = $brushes.ConvertFromString($palette[$k]) }
    $w.Resources['FA'] = $faFont
    if ($iconFile) { try { $w.Icon = [Windows.Media.Imaging.BitmapFrame]::Create((New-Object Uri (Resolve-Path $iconFile).Path)) } catch { } }
    Set-Strings $w
    $w.Add_SourceInitialized({ [Dwm]::Dark((New-Object Windows.Interop.WindowInteropHelper $w).Handle) }.GetNewClosure())
    $w
}

$win = New-Window $mainXaml
foreach ($n in 'btnLang', 'popLang', 'flagCur', 'langText', 'langEn', 'langDe', 'userBox', 'connText', 'btnDisconnect', 'loginPane', 'btnConnect', 'loginStatus', 'mainPane', 'tabBoxes', 'tabEditor', 'btnNew',
    'paneBoxes', 'txtFilter', 'chkActiveOnly', 'btnSelAll', 'btnSelNone', 'btnDisableChecked', 'btnRefresh', 'lstBoxes', 'lblSummary',
    'paneEditor', 'lstEdit', 'btnTplNew', 'btnTplDelete', 'txtTplName', 'btnTplSave', 'btnTplOpen', 'chkSource', 'txtSource', 'editToolbar', 'editBrowser', 'lblProgress', 'btnCancel', 'btnCopyLog', 'btnOpenLog', 'bar', 'lstLog') {
    Set-Variable -Name $n -Value $win.FindName($n)
}

# ---------------------------------------------------------------- state
$boxes = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$logItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$byUpn = @{}
$view = [Windows.Data.CollectionViewSource]::GetDefaultView($boxes)
$lstBoxes.ItemsSource = $view
$lstLog.ItemsSource = $logItems
$script:connected = $false
$script:job = $null
$script:afterJob = $null
$script:editingPath = $null

$sync = [hashtable]::Synchronized(@{ Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'; Cancel = $false; Mock = [bool]$Mock })
$rs = [runspacefactory]::CreateRunspace()
$rs.ApartmentState = 'STA'
$rs.ThreadOptions = 'ReuseThread'
$rs.Open()
$rs.SessionStateProxy.SetVariable('sync', $sync)

# Writes to the activity list and to the log file
function Add-Log($text, $kind = 'info') {
    $line = '[{0:HH:mm:ss}] {1}' -f (Get-Date), $text
    $logItems.Add([LogLine]@{ Text = $line; Kind = $kind })
    $lstLog.ScrollIntoView($logItems[$logItems.Count - 1])
    try { Add-Content -Path $logFile -Value $line -Encoding UTF8 } catch { }
}

# ---------------------------------------------------------------- worker scripts
# Everything in the worker blocks runs in the background runspace, so the window never freezes
$workerPrelude = {
    $ErrorActionPreference = 'Stop'
    $ProgressPreference = 'SilentlyContinue'
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072 } catch { }
    $ConfirmPreference = 'None'
    $q = $sync.Queue

    # Exchange throttles bulk calls, retry a few times before giving up
    function Invoke-Retry([scriptblock]$Action) {
        for ($try = 0; ; $try++) {
            try { return (& $Action) }
            catch {
                if ($try -ge 3 -or $_.Exception.Message -notmatch 'throttl|server busy|429') { throw }
                Start-Sleep -Seconds (2, 5, 10)[$try]
            }
        }
    }

    function ConvertTo-Local($t) {
        if ($t -is [datetime] -and $t.Kind -eq 'Utc') { return $t.ToLocalTime() }
        $t
    }

    # Reads the auto-reply of one mailbox and shapes it for the list
    function Get-AutoReplyInfo($upn) {
        $c = Invoke-Retry { Get-MailboxAutoReplyConfiguration -Identity $upn }
        $raw = [string]$c.AutoReplyState
        $from = ConvertTo-Local $c.StartTime
        $to = ConvertTo-Local $c.EndTime
        $active = $raw -eq 'Enabled' -or ($raw -eq 'Scheduled' -and $to -gt (Get-Date))
        $label, $period = switch ($raw) {
            'Enabled'   { 'On', 'always' }
            'Scheduled' { $(if ($active) { 'Scheduled' } else { 'Expired' }), ('{0:dd.MM.yy HH:mm} - {1:dd.MM.yy HH:mm}' -f $from, $to) }
            default     { 'Off', '' }
        }
        $html = [string]$c.InternalMessage
        $text = $html -replace '(?s)<(style|head)[^>]*>.*?</\1>', ' ' -replace '<[^>]+>', ' '
        $text = ([Net.WebUtility]::HtmlDecode($text) -replace '\s+', ' ').Trim()
        if ($text.Length -gt 120) { $text = $text.Substring(0, 120) }
        @{
            Type = 'status'; Upn = $upn; State = $label; Period = $period; Snippet = $text; Html = $html; Active = $active
            StartText = $(if ($raw -eq 'Scheduled') { '{0:dd.MM.yyyy HH:mm}' -f $from } else { '' })
            EndText = $(if ($raw -eq 'Scheduled') { '{0:dd.MM.yyyy HH:mm}' -f $to } else { '' })
        }
    }
}

# Installs the Exchange Online module for the current user, the user agreed to this before the job started
$workerInstall = {
    try {
        $q.Enqueue(@{ Type = 'info'; Text = 'Installing the ExchangeOnlineManagement module, this can take a minute' })
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) { Install-PackageProvider -Name NuGet -Scope CurrentUser -Force | Out-Null }
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
        $q.Enqueue(@{ Type = 'ok'; Text = 'Module installed' })
    } catch {
        $q.Enqueue(@{ Type = 'err'; Text = "Could not install the module: $($_.Exception.Message)" })
    }
}

$workerConnect = {
    try {
        if (-not $sync.Mock) { Import-Module ExchangeOnlineManagement -ErrorAction Stop }
    } catch {
        $q.Enqueue(@{ Type = 'err'; Text = 'ExchangeOnlineManagement module not found. Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser' })
        return
    }
    try {
        try { Connect-ExchangeOnline -ShowBanner:$false -DisableWAM }
        catch {
            $first = $_.Exception.Message
            $q.Enqueue(@{ Type = 'info'; Text = 'Browser sign-in failed, trying the Windows account broker' })
            try { Connect-ExchangeOnline -ShowBanner:$false } catch { throw $first }
        }
        $ci = Get-ConnectionInformation | Select-Object -First 1
        $q.Enqueue(@{ Type = 'connected'; Text = [string]$ci.UserPrincipalName })
    } catch {
        $q.Enqueue(@{ Type = 'err'; Text = "Connect failed: $($_.Exception.Message)" })
    }
}

$workerLoad = {
    try {
        $q.Enqueue(@{ Type = 'info'; Text = 'Loading mailboxes' })
        $list = New-Object Collections.Generic.List[object]
        Invoke-Retry {
            $list.Clear()
            $fetch = if (Get-Command Get-Mailbox -ErrorAction SilentlyContinue) {
                { Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited }
            } else {
                { Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited }
            }
            & $fetch | ForEach-Object {
                if ($sync.Cancel) { throw 'Cancelled' }
                $list.Add(@{ Name = [string]$_.DisplayName; Upn = [string]$_.UserPrincipalName })
                if ($list.Count % 50 -eq 0) { $q.Enqueue(@{ Type = 'info'; Text = "Loading mailboxes ($($list.Count))" }) }
            }
        } | Out-Null
        $q.Enqueue(@{ Type = 'mailboxes'; Data = @($list | Sort-Object { $_.Name }) })
    } catch {
        $q.Enqueue(@{ Type = 'err'; Text = "Could not load mailboxes: $($_.Exception.Message)" })
    }
}

$workerStatus = {
    $done = 0; $failed = 0
    foreach ($upn in $upns) {
        if ($sync.Cancel) { $q.Enqueue(@{ Type = 'warn'; Text = 'Status refresh cancelled' }); break }
        try { $q.Enqueue((Get-AutoReplyInfo $upn)) }
        catch {
            $failed++
            $q.Enqueue(@{ Type = 'err'; Text = "$upn : $($_.Exception.Message)" })
            $q.Enqueue(@{ Type = 'status'; Upn = $upn; State = 'Error'; Period = ''; Snippet = $_.Exception.Message; Html = ''; Active = $false; StartText = ''; EndText = '' })
        }
        $done++
        $q.Enqueue(@{ Type = 'progress'; Done = $done; Total = @($upns).Count; Failed = $failed; Label = 'Reading status' })
    }
    $q.Enqueue(@{ Type = 'info'; Text = "Status read for $done mailboxes, $failed failed" })
}

# Sets or disables the auto-reply per mailbox, one failing mailbox never stops the run
$workerApply = {
    $ok = 0; $skipped = 0; $failed = 0; $done = 0
    foreach ($upn in $upns) {
        if ($sync.Cancel) { $q.Enqueue(@{ Type = 'warn'; Text = 'Cancelled, remaining mailboxes untouched' }); break }
        $done++
        try {
            $cur = Get-AutoReplyInfo $upn
            if ($state -ne 'Disabled' -and $cur.Active -and -not $overwrite) {
                $skipped++
                $q.Enqueue(@{ Type = 'warn'; Text = "$upn : skipped, already has an auto-reply (enable overwrite to replace it)" })
            }
            elseif ($state -eq 'Disabled' -and -not $cur.Active) {
                $skipped++
                $q.Enqueue(@{ Type = 'info'; Text = "$upn : already off" })
            }
            else {
                $p = @{ Identity = $upn; AutoReplyState = $state }
                if ($state -ne 'Disabled') {
                    $p.InternalMessage = $html; $p.ExternalMessage = $html; $p.ExternalAudience = 'All'
                    if ($state -eq 'Scheduled') { $p.StartTime = $start; $p.EndTime = $end }
                }
                Invoke-Retry { Set-MailboxAutoReplyConfiguration @p } | Out-Null
                $ok++
                $q.Enqueue(@{ Type = 'ok'; Text = "$upn : $state" })
                try { $q.Enqueue((Get-AutoReplyInfo $upn)) } catch { }
            }
        }
        catch {
            $failed++
            $q.Enqueue(@{ Type = 'err'; Text = "$upn : $($_.Exception.Message)" })
        }
        $q.Enqueue(@{ Type = 'progress'; Done = $done; Total = @($upns).Count; Failed = $failed; Label = 'Applying' })
    }
    $q.Enqueue(@{ Type = 'info'; Text = "Finished: $ok changed, $skipped skipped, $failed failed" })
}

$mockInit = {
    $global:mockDb = @{}
    $global:mockUsers = 1..60 | ForEach-Object { [pscustomobject]@{ DisplayName = ('Test User {0:00}' -f $_); UserPrincipalName = ('user{0:00}@example.test' -f $_) } }
    foreach ($u in $global:mockUsers[0..5]) {
        $global:mockDb[$u.UserPrincipalName] = @{ State = 'Enabled'; Html = '<p>Thank you for your message. We are currently out of the office.</p>'; Start = $null; End = $null }
    }
    function global:Connect-ExchangeOnline { param([switch]$ShowBanner, [switch]$DisableWAM) Start-Sleep -Milliseconds 600 }
    function global:Disconnect-ExchangeOnline { param($Confirm) }
    function global:Get-ConnectionInformation { [pscustomobject]@{ UserPrincipalName = 'admin@example.test' } }
    function global:Get-Mailbox { param($RecipientTypeDetails, $ResultSize) Start-Sleep -Milliseconds 400; $global:mockUsers }
    function global:Get-EXOMailbox { param($RecipientTypeDetails, $ResultSize) Start-Sleep -Milliseconds 400; $global:mockUsers }
    function global:Get-MailboxAutoReplyConfiguration {
        param($Identity)
        Start-Sleep -Milliseconds 60
        $e = $global:mockDb[$Identity]
        if (-not $e) { $e = @{ State = 'Disabled'; Html = ''; Start = $null; End = $null } }
        [pscustomobject]@{ AutoReplyState = $e.State; InternalMessage = $e.Html; StartTime = $e.Start; EndTime = $e.End }
    }
    function global:Set-MailboxAutoReplyConfiguration {
        param($Identity, $AutoReplyState, $InternalMessage, $ExternalMessage, $ExternalAudience, $StartTime, $EndTime)
        Start-Sleep -Milliseconds (Get-Random -Minimum 40 -Maximum 220)
        if ($Identity -like 'user17@*') { throw 'Mock failure: mailbox is locked' }
        $e = $global:mockDb[$Identity]
        if (-not $e) { $e = @{ Html = '' }; $global:mockDb[$Identity] = $e }
        $e.State = $AutoReplyState; $e.Start = $StartTime; $e.End = $EndTime
        if ($PSBoundParameters.ContainsKey('InternalMessage')) { $e.Html = $InternalMessage }
    }
}

# ---------------------------------------------------------------- jobs
# Only one background job runs at a time
function Invoke-Worker([scriptblock[]]$Blocks, [hashtable]$Vars = @{}) {
    if ($script:job) { return }
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    foreach ($k in $Vars.Keys) { $rs.SessionStateProxy.SetVariable($k, $Vars[$k]) }
    $text = (@($workerPrelude) + $Blocks | ForEach-Object { $_.ToString() }) -join "`r`n"
    $null = $ps.AddScript($text)
    $sync.Cancel = $false
    $script:lastActivity = Get-Date; $script:warned = $false
    $bar.Value = 0
    $script:job = @{ Ps = $ps; Handle = $ps.BeginInvoke() }
    Update-State
}

# Called by the timer once the background job has finished
function Complete-Job {
    try { $script:job.Ps.EndInvoke($script:job.Handle) } catch { Add-Log $_.Exception.Message 'err' }
    foreach ($e in $script:job.Ps.Streams.Error) { Add-Log $e.ToString() 'err' }
    $script:job.Ps.Dispose()
    $script:job = $null
    $lblProgress.Text = ''
    $bar.Value = 0
    Update-State
    if ($script:afterJob) {
        $next = $script:afterJob
        $script:afterJob = $null
        & $next
    }
}

# Handles the messages the worker left in the queue
function Receive-Message($m) {
    $script:lastActivity = Get-Date; $script:warned = $false
    switch ($m.Type) {
        'info' { Add-Log $m.Text }
        'ok' { Add-Log $m.Text 'ok' }
        'warn' { Add-Log $m.Text 'warn' }
        'err' {
            Add-Log $m.Text 'err'
            if (-not $script:connected) {
                $loginStatus.Text = $m.Text
                $loginStatus.Foreground = $win.Resources['ErrBrush']
            }
        }
        'progress' {
            $bar.Value = [int](100 * $m.Done / [math]::Max(1, $m.Total))
            $lblProgress.Text = '{0}  {1}/{2}  {3} failed' -f $m.Label, $m.Done, $m.Total, $m.Failed
        }
        'connected' {
            $script:connected = $true
            $connText.Text = (T 'Connected as {0}') -f $m.Text
            Add-Log ('Connected as ' + $m.Text) 'ok'
            $script:afterJob = { Invoke-Worker @($workerLoad); $script:afterJob = { Start-StatusRefresh } }
        }
        'mailboxes' {
            $boxes.Clear(); $byUpn.Clear()
            foreach ($d in $m.Data) {
                $item = [Mailbox]@{ Name = $d.Name; Upn = $d.Upn; Html = ''; StartText = ''; EndText = '' }
                $boxes.Add($item); $byUpn[$d.Upn] = $item
            }
            Add-Log ('{0} mailboxes loaded' -f $boxes.Count) 'ok'
        }
        'status' {
            $item = $byUpn[$m.Upn]
            if ($item) {
                $item.State = $m.State; $item.Period = $m.Period; $item.Snippet = $m.Snippet
                $item.Html = $m.Html; $item.StartText = $m.StartText; $item.EndText = $m.EndText; $item.Active = $m.Active
            }
        }
    }
}

function Start-StatusRefresh {
    if ($boxes.Count -eq 0) { return }
    Invoke-Worker @($workerStatus) @{ upns = @($boxes | ForEach-Object { $_.Upn }) }
}

# ---------------------------------------------------------------- markdown
function ConvertTo-SafeUrl([string]$url) {
    if ($url -match '^(https?://|mailto:|data:image/(png|jpeg|gif);base64,)') { $url } else { '#' }
}

function Format-Inline([string]$text) {
    $s = [Net.WebUtility]::HtmlEncode($text)
    $s = [regex]::Replace($s, '!\[([^\]]*)\]\(([^)\s]+)\)', [Text.RegularExpressions.MatchEvaluator] {
            param($m) '<img src="{0}" alt="{1}" style="max-width:100%;"/>' -f (ConvertTo-SafeUrl $m.Groups[2].Value), $m.Groups[1].Value })
    $s = [regex]::Replace($s, '\[([^\]]+)\]\(([^)\s]+)\)', [Text.RegularExpressions.MatchEvaluator] {
            param($m) '<a href="{0}" target="_blank">{1}</a>' -f (ConvertTo-SafeUrl $m.Groups[2].Value), $m.Groups[1].Value })
    $s = $s -replace '\*\*(.+?)\*\*', '<strong>$1</strong>'
    $s -replace '(?<![*\w])\*(?!\s)(.+?)(?<!\s)\*(?![*\w])', '<em>$1</em>'
}

function ConvertFrom-Markdown([string]$md) {
    $out = New-Object Collections.Generic.List[string]
    $para = New-Object Collections.Generic.List[string]
    $list = ''
    $endPara = {
        if ($para.Count) {
            $out.Add('<p style="margin:0 0 12px 0;">' + (($para | ForEach-Object { Format-Inline $_ }) -join '<br/>') + '</p>')
            $para.Clear()
        }
    }
    $endList = { if ($list) { $out.Add("</$list>"); $list = '' } }

    foreach ($raw in @($md -split '\r?\n') + '') {
        $line = $raw.TrimEnd()
        if ($line -match '^(#{1,3})\s+(.*)$') {
            $level = $Matches[1].Length; $content = $Matches[2]
            . $endPara; . $endList
            $out.Add(('<h{0} style="font-size:{1}px; margin:0 0 10px 0;">{2}</h{0}>' -f $level, (0, 20, 16, 14)[$level], (Format-Inline $content)))
        }
        elseif ($line -match '^(-{3,}|\*{3,})$') { . $endPara; . $endList; $out.Add('<hr/>') }
        elseif ($line -match '^\s*[-*+]\s+(.*)$') {
            $content = $Matches[1]
            . $endPara
            if ($list -ne 'ul') { . $endList; $out.Add('<ul>'); $list = 'ul' }
            $out.Add('<li>' + (Format-Inline $content) + '</li>')
        }
        elseif ($line -match '^\s*\d+[.)]\s+(.*)$') {
            $content = $Matches[1]
            . $endPara
            if ($list -ne 'ol') { . $endList; $out.Add('<ol>'); $list = 'ol' }
            $out.Add('<li>' + (Format-Inline $content) + '</li>')
        }
        elseif (-not $line) { . $endPara; . $endList }
        else { . $endList; $para.Add($line) }
    }
    '<!doctype html><html><head><meta charset="utf-8"/></head><body style="font-family: Arial, sans-serif; font-size: 12px; color: #333;">' +
    ($out -join "`r`n") + '</body></html>'
}

function Add-Compat([string]$html) {
    $meta = '<meta http-equiv="X-UA-Compatible" content="IE=edge">'
    if ($html -match '(?i)<head[^>]*>') { ([regex]'(?i)(<head[^>]*>)').Replace($html, '$1' + $meta, 1) } else { $meta + $html }
}

# Strips scripts, frames, forms and event handlers from template html
function Clear-Html([string]$html) {
    $h = [regex]::Replace($html, '(?is)<(script|iframe|object|embed|applet|form|link|base)\b.*?</\1\s*>', '')
    $h = [regex]::Replace($h, '(?is)<(script|iframe|object|embed|applet|form|link|base|input|button)\b[^>]*>', '')
    $h = [regex]::Replace($h, '(?is)<meta[^>]*http-equiv\s*=\s*["'']?refresh[^>]*>', '')
    $h = [regex]::Replace($h, '(?is)\s+on\w+\s*=\s*("[^"]*"|''[^'']*''|[^\s>]+)', '')
    $h = [regex]::Replace($h, '(?is)(href|src|action|background)\s*=\s*(["''])\s*(javascript|vbscript):[^"'']*\2', '$1=$2#$2')
    $h = [regex]::Replace($h, '(?is)src\s*=\s*(["''])\s*data:(?!image/(png|jpe?g|gif);base64,)[^"'']*\1', 'src=$1$1')
    $h -replace '(?i)expression\s*\(', ''
}

function Lock-Browser($browser) {
    $browser.Add_Navigating({ param($s, $e) if ($e.Uri -and $e.Uri.Scheme -ne 'about') { $e.Cancel = $true } })
    # script error dialogs of the embedded page are of no use to anyone here, keep the control silent
    $browser.Add_LoadCompleted({
            param($sender, $e)
            try {
                $field = [Windows.Controls.WebBrowser].GetField('_axIWebBrowser2', [Reflection.BindingFlags]'NonPublic,Instance')
                $ax = $field.GetValue($sender)
                [void]$ax.GetType().InvokeMember('Silent', [Reflection.BindingFlags]::SetProperty, $null, $ax, @($true))
            } catch { }
        })
}

function Show-Html($browser, $html) {
    $doc = if ($html) { Add-Compat (Clear-Html $html) } else { '<html><body></body></html>' }
    try { $browser.NavigateToString($doc) } catch { }
}

function Get-Templates {
    $result = New-Object Collections.Generic.List[object]
    foreach ($src in @(@($builtInDir, $true), @($userDir, $false))) {
        if (-not $src[0]) { continue }
        foreach ($f in Get-ChildItem $src[0] -Filter *.html | Sort-Object Name) {
            $suffix = if ($src[1]) { T ' (built-in)' } else { '' }
            $result.Add([pscustomobject]@{ Name = $f.BaseName + $suffix; Title = $f.BaseName; Path = $f.FullName; BuiltIn = $src[1] })
        }
    }
    # a saved template with the same name replaces the built-in one, deleting it brings the built-in back
    $custom = @($result | Where-Object { -not $_.BuiltIn } | ForEach-Object { $_.Title })
    $result | Where-Object { -not ($_.BuiltIn -and $custom -contains $_.Title) }
}

# ---------------------------------------------------------------- dialogs
# Reads and validates the period fields, returns State/Start/End or an Error text
function Get-Schedule($rbPerm, $rbRange, $txtStart, $txtEnd) {
    if ($rbPerm.IsChecked) { return @{ State = 'Enabled' } }
    $fmt = 'dd.MM.yyyy HH:mm'
    $ci = [Globalization.CultureInfo]::InvariantCulture
    $from = Get-Date
    $to = [datetime]::MinValue
    if ($rbRange.IsChecked -and -not [datetime]::TryParseExact($txtStart.Text.Trim(), $fmt, $ci, 'None', [ref]$from)) {
        return @{ Error = T 'Start must look like 24.12.2026 08:00' }
    }
    if (-not [datetime]::TryParseExact($txtEnd.Text.Trim(), $fmt, $ci, 'None', [ref]$to)) {
        return @{ Error = T 'End must look like 27.12.2026 08:00' }
    }
    if ($to -le $from -or $to -le (Get-Date)) { return @{ Error = T 'End must be in the future and after the start' } }
    @{ State = 'Scheduled'; Start = $from; End = $to }
}

# Plain confirmation, with a typed number as extra guard for bulk changes
function Confirm-Action($message, $count) {
    $dlg = New-Window $confirmXaml
    $dlg.Owner = $win
    $dlg.FindName('msg').Text = $message
    $yes = $dlg.FindName('yes'); $guard = $dlg.FindName('txtGuard')
    if ($count -ge 10) {
        $dlg.FindName('guard').Visibility = 'Visible'
        $dlg.FindName('guardText').Text = (T 'Type {0} to confirm') -f $count
        $yes.IsEnabled = $false
        $guard.Add_TextChanged({ $yes.IsEnabled = ($guard.Text.Trim() -eq "$count") }.GetNewClosure())
    }
    $yes.Add_Click({ $dlg.DialogResult = $true }.GetNewClosure())
    $dlg.FindName('no').Add_Click({ $dlg.DialogResult = $false }.GetNewClosure())
    [bool]$dlg.ShowDialog()
}

function Read-Text($title, $prompt) {
    $dlg = New-Window $confirmXaml
    $dlg.Owner = $win
    $dlg.Title = $title
    $dlg.FindName('msg').Text = $prompt
    $dlg.FindName('guard').Visibility = 'Visible'
    $dlg.FindName('guardText').Visibility = 'Collapsed'
    $dlg.FindName('yes').Content = 'OK'
    $dlg.FindName('yes').Add_Click({ $dlg.DialogResult = $true }.GetNewClosure())
    $dlg.FindName('no').Add_Click({ $dlg.DialogResult = $false }.GetNewClosure())
    if ($dlg.ShowDialog()) { $dlg.FindName('txtGuard').Text }
}

function Show-Message($title, $html) {
    $v = New-Window $viewXaml
    $v.Owner = $win
    $v.FindName('viewTitle').Text = $title
    $browser = $v.FindName('viewBrowser')
    Lock-Browser $browser
    $v.Add_ContentRendered({ Show-Html $browser $html }.GetNewClosure())
    [void]$v.ShowDialog()
    try { $browser.Dispose() } catch { }
}

function Show-ReplyDialog($title, $subtitle, $nChecked, $nAll, $prefill) {
    $d = New-Window $replyXaml
    $d.Owner = $win
    $c = @{}
    foreach ($n in 'dlgTitle', 'dlgSub', 'lstTplD', 'rbPermD', 'rbUntilD', 'rbRangeD', 'txtStartD', 'txtEndD',
        'txtHintD', 'chkOverwriteD', 'targetBox', 'rbSelD', 'rbAllD', 'btnOkD', 'btnCancelD', 'prevD') { $c[$n] = $d.FindName($n) }
    Lock-Browser $c.prevD

    $c.dlgTitle.Text = $title
    $c.dlgSub.Text = $subtitle
    $c.lstTplD.ItemsSource = @(Get-Templates)
    $c.txtEndD.Text = (Get-Date).AddDays(3).Date.AddHours(8).ToString('dd.MM.yyyy HH:mm')
    $c.txtStartD.Text = (Get-Date).AddHours(1).ToString('dd.MM.yyyy HH:00')

    if ($prefill -and $prefill.State -in 'Scheduled', 'Expired') {
        $c.rbRangeD.IsChecked = $true
        if ($prefill.Start) { $c.txtStartD.Text = $prefill.Start }
        if ($prefill.End) { $c.txtEndD.Text = $prefill.End }
    }

    if ($null -eq $nChecked) {
        $c.targetBox.Visibility = 'Collapsed'
        $c.chkOverwriteD.IsChecked = $true
        $c.chkOverwriteD.IsEnabled = $false
    }
    else {
        $c.rbSelD.Content = (T 'Checked mailboxes ({0})') -f $nChecked
        $c.rbAllD.Content = (T 'All mailboxes ({0})') -f $nAll
        $c.rbSelD.IsEnabled = $nChecked -gt 0
        if ($nChecked -eq 0) { $c.rbAllD.IsChecked = $true }
    }

    $getHtml = {
        $t = $c.lstTplD.SelectedItem
        if ($t) { [IO.File]::ReadAllText($t.Path, [Text.Encoding]::UTF8) }
    }.GetNewClosure()

    $renderTimer = New-Object Windows.Threading.DispatcherTimer
    $renderTimer.Interval = [TimeSpan]::FromMilliseconds(250)
    $renderTimer.Add_Tick({ $renderTimer.Stop(); Show-Html $c.prevD (& $getHtml) }.GetNewClosure())

    $update = {
        $sch = Get-Schedule $c.rbPermD $c.rbRangeD $c.txtStartD $c.txtEndD
        $c.txtStartD.IsEnabled = [bool]$c.rbRangeD.IsChecked
        $c.txtEndD.IsEnabled = -not $c.rbPermD.IsChecked
        $c.txtHintD.Text = if ($sch.Error) { $sch.Error } else { '' }
        $c.btnOkD.IsEnabled = $null -ne $c.lstTplD.SelectedItem -and -not $sch.Error
    }.GetNewClosure()

    foreach ($r in $c.rbPermD, $c.rbUntilD, $c.rbRangeD) { $r.Add_Checked($update) }
    foreach ($t in $c.txtStartD, $c.txtEndD) { $t.Add_TextChanged($update) }
    $c.lstTplD.Add_SelectionChanged({ & $update; $renderTimer.Stop(); $renderTimer.Start() }.GetNewClosure())
    $d.Add_ContentRendered({ $renderTimer.Stop(); $renderTimer.Start() }.GetNewClosure())
    if ($c.lstTplD.Items.Count) { $c.lstTplD.SelectedIndex = 0 }

    $c.btnOkD.Add_Click({ $d.DialogResult = $true }.GetNewClosure())
    $c.btnCancelD.Add_Click({ $d.DialogResult = $false }.GetNewClosure())
    & $update

    $ok = $d.ShowDialog()
    $renderTimer.Stop()
    $result = $null
    if ($ok) {
        $result = @{
            Html      = Clear-Html (& $getHtml)
            Schedule  = Get-Schedule $c.rbPermD $c.rbRangeD $c.txtStartD $c.txtEndD
            Overwrite = [bool]$c.chkOverwriteD.IsChecked
            All       = [bool]$c.rbAllD.IsChecked
        }
    }
    try { $c.prevD.Dispose() } catch { }
    $result
}

# ---------------------------------------------------------------- actions
function Get-Checked { @($boxes | Where-Object { $_.Selected }) }

function Start-Apply($upns, $state, $html, $schedule, $overwrite, $summary) {
    if ($script:job) { Add-Log 'A job is already running' 'warn'; return }
    $upns = @($upns)
    if (-not $upns.Count) { Add-Log 'No mailboxes selected' 'warn'; return }
    if ($upns.Count -ge 10) {
        if (-not (Confirm-Action ($summary + "`n`n" + ((T '{0} mailboxes will be changed.') -f $upns.Count)) $upns.Count)) { return }
    }
    Add-Log ('{0} ({1} mailbox(es))' -f $summary, $upns.Count)
    Invoke-Worker @($workerApply) @{
        upns = $upns; state = $state; html = $html; start = $schedule.Start; end = $schedule.End
        overwrite = $overwrite
    }
}

function Get-PeriodText($schedule) {
    if ($schedule.State -eq 'Enabled') { T 'until turned off' }
    else { (T '{0:dd.MM.yyyy HH:mm} to {1:dd.MM.yyyy HH:mm}') -f $schedule.Start, $schedule.End }
}

function New-ReplyCore {
    $checked = Get-Checked
    $r = Show-ReplyDialog (T 'New auto-reply') (T 'Choose the message, the period and the options.') $checked.Count $boxes.Count $null
    if (-not $r) { return }
    $targets = if ($r.All) { $boxes } else { $checked }
    Start-Apply ($targets | ForEach-Object { $_.Upn }) $r.Schedule.State $r.Html $r.Schedule $r.Overwrite ((T 'Set auto-reply, {0}') -f (Get-PeriodText $r.Schedule))
}

function Edit-ReplyCore($item) {
    if ($script:job) { Add-Log 'Wait for the running job to finish' 'warn'; return }
    $prefill = @{ State = $item.State; Start = $item.StartText; End = $item.EndText }
    $title = if ($item.Active) { T 'Edit auto-reply' } else { T 'Enable auto-reply' }
    $r = Show-ReplyDialog $title ($item.Name + '  |  ' + $item.Upn) $null $null $prefill
    if (-not $r) { return }
    Start-Apply @($item.Upn) $r.Schedule.State $r.Html $r.Schedule $true ((T 'Set auto-reply for {0}, {1}') -f $item.Upn, (Get-PeriodText $r.Schedule))
}

function Disable-ReplyCore($items) {
    if ($script:job) { Add-Log 'Wait for the running job to finish' 'warn'; return }
    $items = @($items)
    if (-not $items.Count) { return }
    $what = if ($items.Count -eq 1) { (T 'Disable the auto-reply for {0}?') -f $items[0].Name } else { (T 'Disable the auto-reply for {0} mailboxes?') -f $items.Count }
    if (-not (Confirm-Action $what $items.Count)) { return }
    Invoke-Worker @($workerApply) @{
        upns = @($items | ForEach-Object { $_.Upn }); state = 'Disabled'; html = $null; start = $null; end = $null; overwrite = $true
    }
}

function New-Reply { Invoke-Dialog { New-ReplyCore } }
function Edit-Reply($item) { Invoke-Dialog { Edit-ReplyCore $item }.GetNewClosure() }
function Disable-Reply($items) { Invoke-Dialog { Disable-ReplyCore $items }.GetNewClosure() }

# Enables or disables controls depending on the connection and running jobs
function Update-State {
    $busy = [bool]$script:job
    $mainPane.Visibility = if ($script:connected) { 'Visible' } else { 'Collapsed' }
    $userBox.Visibility = $mainPane.Visibility
    $loginPane.Visibility = if ($script:connected) { 'Collapsed' } else { 'Visible' }
    $btnConnect.IsEnabled = -not $script:connected -and -not $busy
    $btnDisconnect.IsEnabled = $script:connected -and -not $busy
    $btnNew.IsEnabled = $script:connected -and -not $busy
    $btnRefresh.IsEnabled = $script:connected -and -not $busy
    $checked = @(Get-Checked).Count
    $btnDisableChecked.IsEnabled = $script:connected -and -not $busy -and $checked -gt 0
    $btnCancel.Visibility = if ($busy) { 'Visible' } else { 'Collapsed' }
    $active = @($boxes | Where-Object { $_.Active }).Count
    $lblSummary.Text = (T '{0} mailboxes  |  {1} with an active auto-reply  |  {2} checked') -f $boxes.Count, $active, $checked
}

# ---------------------------------------------------------------- template editor
$blankDoc = '<!doctype html><html><head><meta charset="utf-8"></head><body style="font-family: Arial, sans-serif; font-size: 12px; color: #333;"><p>Write your message here.</p></body></html>'
$script:pendingHtml = $null
$script:editorReady = $false
Lock-Browser $editBrowser

$navTimer = New-Object Windows.Threading.DispatcherTimer
$navTimer.Interval = [TimeSpan]::FromMilliseconds(250)
$navTimer.Add_Tick({
        $navTimer.Stop()
        if ($paneEditor.Visibility -eq 'Visible' -and $null -ne $script:pendingHtml) {
            $html = $script:pendingHtml
            $script:pendingHtml = $null
            Show-Html $editBrowser $html
        }
    })

# The page only accepts typing once editing is switched on. That can fail when the page is not fully built yet,
# so it is tried right after loading and once more shortly afterwards.
function Enable-Editing([bool]$fresh) {
    try {
        $d = $editBrowser.Document
        $d.designMode = 'On'
        if (([string]$d.designMode) -ne 'On' -and $d.body) { $d.body.contentEditable = 'true' }
        if ($fresh) { Reset-History; Push-History }
    } catch { }
}

$enableTimer = New-Object Windows.Threading.DispatcherTimer
$enableTimer.Interval = [TimeSpan]::FromMilliseconds(400)
$enableTimer.Add_Tick({ $enableTimer.Stop(); Enable-Editing $false })
$editBrowser.Add_LoadCompleted({ Enable-Editing $true; $enableTimer.Stop(); $enableTimer.Start() })

function Request-EditorNav($html) {
    $script:pendingHtml = $html
    $navTimer.Stop(); $navTimer.Start()
}

function Update-EditList($selectPath) {
    $items = @(Get-Templates)
    $lstEdit.ItemsSource = $items
    $match = $items | Where-Object { $_.Path -eq $selectPath } | Select-Object -First 1
    if ($match) { $lstEdit.SelectedItem = $match }
}

function Open-Template($t) {
    $script:editingPath = if ($t -and -not $t.BuiltIn) { $t.Path } else { $null }
    $txtTplName.Text = if ($t) { $t.Title } else { '' }
    $chkSource.IsChecked = $false
    Request-EditorNav $(if ($t) { [IO.File]::ReadAllText($t.Path, [Text.Encoding]::UTF8) } else { $blankDoc })
}

# Reads the edited page, falls back to the body or to the last snapshot when the page refuses to answer
function Get-PageHtml {
    $d = $editBrowser.Document
    try { $html = [string]$d.documentElement.outerHTML; if ($html) { return $html } } catch { }
    $shell = '<html><head></head><body style="font-family: Arial, sans-serif; font-size: 12px; color: #333;">{0}</body></html>'
    try { $body = [string]$d.body.innerHTML; if ($body) { return $shell -f $body } } catch { }
    try { Push-History } catch { }
    if ($script:historyPos -ge 0) { return $shell -f $script:history[$script:historyPos] }
    throw 'The editor page is not ready yet, click into the text once and try again'
}
# Cleans the edited page up for mail clients (image widths, charset) and returns the full document
function Get-EditorHtml([bool]$fromDom) {
    if (-not $fromDom) { return Clear-Html $txtSource.Text }
    $d = $editBrowser.Document
    $embedded = 0
    $images = try { @($d.getElementsByTagName('img')) } catch { @() }
    foreach ($img in $images) {
        try {
            $oldW = [int]$img.width
            $oldH = [int]$img.height
            $w = [int]([math]::Round($oldW / 10.0) * 10)
            if ($w -gt 600) { $w = 600 }
            if ($w -gt 0 -and $oldW -gt 0) {
                $img.setAttribute('width', "$w")
                if ($oldH -gt 0) { $img.setAttribute('height', [string][int][math]::Round($oldH * $w / $oldW)) }
            }
            $img.setAttribute('border', '0')
            $img.style.display = 'block'
            if (-not $img.getAttribute('alt')) { $img.setAttribute('alt', 'image') }
            if ([string]$img.getAttribute('src') -like 'data:*') { $embedded++ }
        } catch { }
    }
    if ($embedded) { Add-Log ("{0} embedded image(s): Outlook will not show them, use hosted https images" -f $embedded) 'warn' }    $html = Clear-Html (Get-PageHtml)
    $html = $html -replace '(?i)<meta[^>]*X-UA-Compatible[^>]*>', ''
    if ($html -notmatch '(?i)<meta[^>]*charset') { $html = ([regex]'(?i)(<head[^>]*>)').Replace($html, '$1<meta charset="utf-8">', 1) }
    "<!doctype html>`r`n" + $html
}

# Undo and redo keep their own snapshots of the page, changes made through markup do not end up in the browser's undo stack
$script:history = New-Object Collections.Generic.List[string]
$script:historyPos = -1

function Reset-History {
    $script:history.Clear()
    $script:historyPos = -1
}

function Push-History {
    $html = try { [string]$editBrowser.Document.body.innerHTML } catch { return }
    if ($script:historyPos -ge 0 -and $script:history[$script:historyPos] -eq $html) { return }
    if ($script:historyPos -lt $script:history.Count - 1) {
        $script:history.RemoveRange($script:historyPos + 1, $script:history.Count - $script:historyPos - 1)
    }
    $script:history.Add($html)
    if ($script:history.Count -gt 50) { $script:history.RemoveAt(0) }
    $script:historyPos = $script:history.Count - 1
}

function Restore-History([int]$position) {
    $script:historyPos = $position
    $script:savedRange = $null
    try { $editBrowser.Document.body.innerHTML = $script:history[$position] }
    catch { Add-Log ('Undo failed: ' + $_.Exception.Message) 'warn' }
}

function Undo-Edit { Push-History; if ($script:historyPos -gt 0) { Restore-History ($script:historyPos - 1) } }
function Redo-Edit { if ($script:historyPos -lt $script:history.Count - 1) { Restore-History ($script:historyPos + 1) } }
# Fast double clicks must not start a second handler while the first one is still busy
$script:uiBusy = $false
$script:dialogOpen = $false

function Invoke-Guarded([scriptblock]$action) {
    if ($script:uiBusy) { return }
    $script:uiBusy = $true
    try { & $action }
    catch { Add-Log ('Action failed: ' + $_.Exception.Message) 'warn' }
    finally { $script:uiBusy = $false }
}

function Invoke-Dialog([scriptblock]$action) {
    if ($script:dialogOpen) { return }
    $script:dialogOpen = $true
    try { & $action }
    catch { Add-Log ('Action failed: ' + $_.Exception.Message) 'warn' }
    finally { $script:dialogOpen = $false }
}
# Runs after the popup has closed completely, running page commands inside the popup click can freeze the window
function Invoke-Later([scriptblock]$action) {
    [void]$win.Dispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::ApplicationIdle, [Action]$action)
}

# The selection is stored when a dropdown opens, the popup takes the focus and the page loses its selection
function Save-Selection {
    try {
        $range = $editBrowser.Document.selection.createRange()
        if ($range) { $script:savedRange = $range }
    } catch { }
}

function Get-Range {
    if ($script:savedRange) { return $script:savedRange }
    try { $editBrowser.Document.selection.createRange() } catch { $null }
}

# Nearest block element (paragraph, heading, list item ...) around the selection, $null when the text sits directly in the body
function Get-BlockElement($range) {
    $el = try { $range.parentElement() } catch { $null }
    while ($el) {
        if ($el.tagName -eq 'BODY') { return $null }
        if ($el.tagName -in 'P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'TD', 'BLOCKQUOTE') { return $el }
        $el = $el.parentElement
    }
}

function Set-Align($command) {
    $value = @{ justifyLeft = 'left'; justifyCenter = 'center'; justifyRight = 'right'; justifyFull = 'justify' }[$command]
    $range = Get-Range
    $el = Get-BlockElement $range
    if ($el) { $el.setAttribute('align', $value); return }
    if ($range -and $range.text) { $range.pasteHTML("<div align=`"$value`">" + $range.htmlText + '</div>'); $script:savedRange = $null }
    else { Add-Log 'Click into a paragraph or select text first' 'warn' }
}

# Moves a paragraph in steps of 40 px, which is also the grid for horizontal positioning
function Set-Indent([int]$step) {
    $range = Get-Range
    $el = Get-BlockElement $range
    if ($el) {
        $current = [int](([string]$el.style.marginLeft) -replace '[^\d]', '0')
        $el.style.marginLeft = ([math]::Max(0, $current + $step)).ToString() + 'px'
        return
    }
    if ($range -and $range.text -and $step -gt 0) { $range.pasteHTML('<div style="margin-left:40px">' + $range.htmlText + '</div>'); $script:savedRange = $null }
    else { Add-Log 'Click into a paragraph or select text first' 'warn' }
}

function Set-List($command) {
    $tag = if ($command -eq 'insertOrderedList') { 'ol' } else { 'ul' }
    $range = Get-Range
    $lines = @(([string]$range.text) -split '\r?\n' | Where-Object { $_.Trim() })
    if (-not $lines) { $lines = @('List item') }
    $items = ($lines | ForEach-Object { '<li>' + [Net.WebUtility]::HtmlEncode($_.Trim()) + '</li>' }) -join ''
    Insert-Html "<$tag>$items</$tag>"
}

function Remove-Link {
    $el = try { (Get-Range).parentElement() } catch { $null }
    while ($el -and $el.tagName -ne 'A') { $el = $el.parentElement }
    if ($el) { $el.outerHTML = $el.innerHTML } else { Add-Log 'Click into a link first' 'warn' }
}

function Clear-Format {
    $range = Get-Range
    if (-not $range -or -not $range.text) { Add-Log 'Select some text first' 'warn'; return }
    $range.pasteHTML([Net.WebUtility]::HtmlEncode([string]$range.text) -replace '\r?\n', '<br>')
    $script:savedRange = $null
}

# Images are resized in presets that fit the 600 px mail width, native drag handles snap to 10 px when the template is saved
function Set-ImageSize($value) {
    $image = try { (Get-Range).item(0) } catch { $null }
    if (-not $image -or $image.tagName -ne 'IMG') { Add-Log 'Click an image first' 'warn'; return }
    switch -Regex ($value) {
        '^\d+$' { $image.removeAttribute('height'); $image.setAttribute('width', [string][int](600 * [int]$value / 100)) }
        '^float-(left|right)$' { $image.setAttribute('align', $Matches[1]) }
        '^center$' { $image.removeAttribute('align'); $image.parentElement.setAttribute('align', 'center') }
    }
}
# Puts open/close around the text inside every selected block, a wrapper around the block itself would add an empty line
function Format-Html($html, $open, $close) {
    $script:fmtOpen = $open
    $script:fmtClose = $close
    $blocks = '(?is)(<(p|div|h[1-6]|li)\b[^>]*>)(.*?)(</\2>)'
    if ([regex]::IsMatch($html, $blocks)) {
        [regex]::Replace($html, $blocks, [Text.RegularExpressions.MatchEvaluator] { param($m) $m.Groups[1].Value + $script:fmtOpen + $m.Groups[3].Value + $script:fmtClose + $m.Groups[4].Value })
    }
    else { $open + $html + $close }
}

# Dropdown commands wrap the selected html directly, execCommand froze the window after a popup was used
function Set-RangeFormat($command, $value) {
    $open, $close = switch ($command) {
        'fontName'   { "<font face=`"$value`">", '</font>' }
        'fontSize'   { "<font size=`"$value`">", '</font>' }
        'foreColor'  { "<font color=`"$value`">", '</font>' }
        'backColor'  { "<span style=`"background-color:$value`">", '</span>' }
        'formatBlock' { $tag = $value.Trim('<', '>'); "<$tag>", "</$tag>" }
    }
    $range = Get-Range
    if (-not $range.text) { Add-Log 'Select some text first' 'warn'; return }
    $range.pasteHTML((Format-Html ([string]$range.htmlText) $open $close))
    $script:savedRange = $null
}
# Bold, italic, underline and strikethrough toggle by wrapping or unwrapping the selected html
function Set-Toggle($pattern, $tag) {
    try {
        $range = $editBrowser.Document.selection.createRange()
        if (-not $range.text) { Add-Log 'Select some text first' 'warn'; return }
        $html = [string]$range.htmlText
        $m = [regex]::Match($html, "(?is)^\s*<($pattern)\b[^>]*>(.*)</\1>\s*$")
        if ($m.Success) { $range.pasteHTML($m.Groups[2].Value) } else { $range.pasteHTML((Format-Html $html "<$tag>" "</$tag>")) }
    }
    catch { Add-Log ('Format failed: ' + $_.Exception.Message) 'warn' }
}
# Paragraph and heading styles replace the tag of the block around the cursor, nesting a heading inside a paragraph breaks the page
function Set-Block($value) {
    $tag = $value.Trim('<', '>')
    $style = @{ h1 = ' style="font-size:20px; margin:0 0 10px 0;"'; h2 = ' style="font-size:16px; margin:0 0 10px 0;"'; h3 = ' style="font-size:14px; margin:0 0 10px 0;"'; p = ' style="margin:0 0 12px 0;"' }[$tag]
    $range = Get-Range
    $script:savedRange = $null
    try {
        $el = Get-BlockElement $range
        if ($el -and $el.tagName -in 'LI', 'TD') { $el.innerHTML = "<$tag$style>" + $el.innerHTML + "</$tag>"; return }
        if ($el) { $el.outerHTML = "<$tag$style>" + $el.innerHTML + "</$tag>"; return }
        if ($range -and $range.text) {
            $html = [string]$range.htmlText
            $blocks = '(?is)<(p|div|h[1-6])\b[^>]*>(.*?)</\1>'
            if ([regex]::IsMatch($html, $blocks)) { $html = [regex]::Replace($html, $blocks, "<$tag$style>`$2</$tag>") }
            else { $html = "<$tag$style>$html</$tag>" }
            $range.pasteHTML($html)
        }
        else { Add-Log 'Click into a paragraph or select text first' 'warn' }
    }
    catch { Add-Log ('Format failed: ' + $_.Exception.Message) 'warn' }
}
# Closes the dropdown, then applies the chosen value once the window is idle again
function Complete-Pick($data) {
    Invoke-Guarded { Complete-PickCore $data }.GetNewClosure()
}

function Complete-PickCore($data) {
    $data.Pop.IsOpen = $false
    $script:pick = $data
    $p = $script:pick
    $command = if ($p.Command) { $p.Command } else { $p.Name }
    try {
        if ($command -eq 'formatBlock') { Set-Block $p.Value }
        elseif ($command -in 'fontName', 'fontSize', 'foreColor', 'backColor') { Set-RangeFormat $command $p.Value }
        elseif ($p.Pick) { & $p.Pick $p.Value }
        else { Invoke-Edit $command $p.Value }
    }
    catch { Add-Log ('Format failed: ' + $_.Exception.Message) 'warn' }
}

function Invoke-Edit($command, $value = $null) {
    if ($chkSource.IsChecked) { return }
    try {
        $res = $editBrowser.Document.execCommand($command, $false, $value)
    }
    catch { Add-Log ('Editor command failed: ' + $_.Exception.Message) 'warn' }
}

function New-Icon([int]$code, [double]$size = 13) {
    $i = New-Object Windows.Controls.TextBlock
    $i.FontFamily = $win.Resources['FA']
    [Windows.Media.TextOptions]::SetTextFormattingMode($i, 'Ideal')
    [Windows.Media.TextOptions]::SetTextRenderingMode($i, 'Grayscale')
    $i.Text = [string][char]$code
    $i.FontSize = $size
    $i.VerticalAlignment = 'Center'
    $i
}

function New-IconContent([int]$code, $text, [bool]$caret = $false) {
    $p = New-Object Windows.Controls.StackPanel
    $p.Orientation = 'Horizontal'
    [void]$p.Children.Add((New-Icon $code))
    if ($text) {
        $t = New-Object Windows.Controls.TextBlock
        $t.Text = $text; $t.Margin = '8,0,0,0'; $t.VerticalAlignment = 'Center'
        [void]$p.Children.Add($t)
    }
    if ($caret) {
        $c = New-Icon 0xf078 8
        $c.Margin = '9,1,0,0'
        [void]$p.Children.Add($c)
    }
    $p
}

function New-Separator {
    $s = New-Object Windows.Controls.Border
    $s.Width = 1; $s.Height = 22; $s.Margin = '4,0,8,4'
    $s.Background = $win.Resources['LineBrush']
    [void]$editToolbar.Children.Add($s)
}

function New-ToolButton($content, $tip, [scriptblock]$action) {
    $b = New-Object Windows.Controls.Button
    $b.Content = $content
    $b.ToolTip = $tip
    $b.Focusable = $false
    $b.Padding = '9,5'
    $b.MinWidth = 32
    $b.Margin = '0,0,4,4'
    $b.Tag = $action
    $b.Add_Click({ param($sender, $e) Invoke-Guarded $sender.Tag })
    [void]$editToolbar.Children.Add($b)
    $b
}

function New-Popup($button, $content) {
    $pop = New-Object Windows.Controls.Primitives.Popup
    $pop.PlacementTarget = $button
    $pop.Placement = 'Bottom'
    $pop.StaysOpen = $false
    $border = New-Object Windows.Controls.Border
    $border.Background = $win.Resources['CardAltBrush']
    $border.BorderBrush = $win.Resources['LineBrush']
    $border.BorderThickness = 1
    $border.Padding = 4
    $border.Child = $content
    $pop.Child = $border
    [void]$editToolbar.Children.Add($pop)
    $button.Add_Click({ if (-not $pop.IsOpen) { Save-Selection }; $pop.IsOpen = -not $pop.IsOpen }.GetNewClosure())
    $pop
}

function New-MenuButton($content, $tip, $items, $commandName, [scriptblock]$onPick) {
    $b = New-ToolButton $content $tip {}
    $panel = New-Object Windows.Controls.StackPanel
    $pop = New-Popup $b $panel
    foreach ($item in $items) {
        $entry = New-Object Windows.Controls.Button
        $entry.Style = $win.Resources['Plain']
        $entry.Padding = '10,6'
        $entry.MinWidth = 160
        $entry.Focusable = $false
        $entry.Content = if ($item.Icon) { New-IconContent $item.Icon $item.Text } else { $item.Text }
        if ($item.Font) { $entry.FontFamily = New-Object Windows.Media.FontFamily $item.Font }
        $value = $item.Value
        $entry.Tag = @{ Pop = $pop; Pick = $onPick; Value = $value; Name = $commandName }
        $entry.Add_Click({ param($sender, $e) Complete-Pick $sender.Tag })
        [void]$panel.Children.Add($entry)
    }
}

function New-ColorButton([int]$icon, $tip, $command) {
    $b = New-ToolButton (New-IconContent $icon $null $true) $tip {}
    $grid = New-Object Windows.Controls.Primitives.UniformGrid
    $grid.Columns = 8
    $pop = New-Popup $b $grid
    $colors = '#000000', '#404040', '#808080', '#c0c0c0', '#ffffff', '#c00000', '#ed7d31', '#ffc000', '#70ad47', '#2e9e8f', '#0070c0', '#002060', '#7030a0', '#e83e8c', '#7f4f24', '#8fd3a9'
    foreach ($color in $colors) {
        $s = New-Object Windows.Controls.Button
        $s.Width = 24; $s.Height = 24; $s.Padding = 0; $s.Margin = 2; $s.Focusable = $false
        $s.Background = $brushes.ConvertFromString($color)
        $s.ToolTip = $color
        $hex = $color
        $s.Tag = @{ Pop = $pop; Command = $command; Value = $hex }
        $s.Add_Click({ param($sender, $e) Complete-Pick $sender.Tag })
        [void]$grid.Children.Add($s)
    }
    $custom = New-Object Windows.Controls.Button
    $custom.Content = T 'Custom...'
    $custom.Focusable = $false
    $custom.Padding = '6,3'
    $custom.Margin = 2
    $custom.Add_Click({
            $pop.IsOpen = $false
            $hex = Read-Text (T 'Custom...') (T 'Color as hex code, for example #1a73e8')
            if ($hex -match '^#[0-9a-fA-F]{6}$') { Invoke-Edit $command $hex }
        }.GetNewClosure())
    [void]$grid.Children.Add($custom)
}

function Build-Toolbar {
    $fonts = 'Arial', 'Calibri', 'Cambria', 'Georgia', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Times New Roman', 'Courier New', 'Segoe UI'
    New-MenuButton (New-IconContent 0xf031 (T 'Font') $true) (T 'Font') ($fonts | ForEach-Object { @{ Text = $_; Value = $_; Font = $_ } }) 'fontName' { param($v) Invoke-Edit 'fontName' $v }
    $sizes = @(@('8 pt', '1'), @('10 pt', '2'), @('12 pt', '3'), @('14 pt', '4'), @('18 pt', '5'), @('24 pt', '6'), @('36 pt', '7'))
    New-MenuButton (New-IconContent 0xf034 (T 'Size') $true) (T 'Size') ($sizes | ForEach-Object { @{ Text = $_[0]; Value = $_[1] } }) 'fontSize' { param($v) Invoke-Edit 'fontSize' $v }
    $blocks = @(@('Paragraph', '<p>'), @('Heading 1', '<h1>'), @('Heading 2', '<h2>'), @('Heading 3', '<h3>'))
    New-MenuButton (New-IconContent 0xf1dc (T 'Heading') $true) (T 'Heading') ($blocks | ForEach-Object { @{ Text = T $_[0]; Value = $_[1] } }) 'formatBlock' { param($v) Invoke-Edit 'formatBlock' $v }
    New-Separator

    [void](New-ToolButton (New-Icon 0xf032) (T 'Bold') { Set-Toggle 'b|strong' 'b' })
    [void](New-ToolButton (New-Icon 0xf033) (T 'Italic') { Set-Toggle 'i|em' 'i' })
    [void](New-ToolButton (New-Icon 0xf0cd) (T 'Underline') { Set-Toggle 'u' 'u' })
    [void](New-ToolButton (New-Icon 0xf0cc) (T 'Strikethrough') { Set-Toggle 'strike|s|del' 'strike' })
    New-ColorButton 0xf53f (T 'Text color') 'foreColor'
    New-ColorButton 0xf591 (T 'Highlight') 'backColor'
    New-Separator

    $align = @(@(0xf036, 'Left', 'justifyLeft'), @(0xf037, 'Center', 'justifyCenter'), @(0xf038, 'Right', 'justifyRight'), @(0xf039, 'Justify', 'justifyFull'))
    New-MenuButton (New-IconContent 0xf036 $null $true) (T 'Align') ($align | ForEach-Object { @{ Icon = $_[0]; Text = T $_[1]; Value = $_[2] } }) 'align' { param($v) Set-Align $v }
    $lists = @(@(0xf0ca, 'Bulleted list', 'insertUnorderedList'), @(0xf0cb, 'Numbered list', 'insertOrderedList'))
    New-MenuButton (New-IconContent 0xf0ca $null $true) (T 'Lists') ($lists | ForEach-Object { @{ Icon = $_[0]; Text = T $_[1]; Value = $_[2] } }) 'lists' { param($v) Set-List $v }
    [void](New-ToolButton (New-Icon 0xf03b) (T 'Outdent') { Set-Indent -40 })
    [void](New-ToolButton (New-Icon 0xf03c) (T 'Indent') { Set-Indent 40 })
    $sizeItems = @(@('25 %', '25'), @('50 %', '50'), @('75 %', '75'), @('100 %', '100'), @((T 'Float left'), 'float-left'), @((T 'Float right'), 'float-right'), @((T 'Center'), 'center'))
    New-MenuButton (New-IconContent 0xf424 $null $true) (T 'Image size') ($sizeItems | ForEach-Object { @{ Text = $_[0]; Value = $_[1] } }) 'imgsize' { param($v) Set-ImageSize $v }
    New-Separator

    [void](New-ToolButton (New-Icon 0xf0c1) (T 'Insert link') { Add-EditorLink })
    [void](New-ToolButton (New-Icon 0xf127) (T 'Remove link') { Remove-Link })
    [void](New-ToolButton (New-Icon 0xf03e) (T 'Insert image') { Add-ImageFromUrl })
    [void](New-ToolButton (New-Icon 0xf068) (T 'Horizontal line') { Insert-Html '<hr>' })
    New-Separator

    [void](New-ToolButton (New-Icon 0xf12d) (T 'Clear formatting') { Clear-Format })
    [void](New-ToolButton (New-Icon 0xf0e2) (T 'Undo') { Undo-Edit })
    [void](New-ToolButton (New-Icon 0xf01e) (T 'Redo') { Redo-Edit })
}

function Add-EditorLink {
    Save-Selection
    $url = Read-Text (T 'Insert link') (T 'Link address (https://...)')
    if (-not $url -or $url.Trim() -notmatch '^(https?://|mailto:)') {
        if ($url) { Add-Log 'Links must start with http://, https:// or mailto:' 'warn' }
        return
    }
    $safe = [Net.WebUtility]::HtmlEncode($url.Trim())
    $range = Get-Range
    $inner = try { if ($range.text) { [string]$range.htmlText } else { $safe } } catch { $safe }
    Insert-Html "<a href=`"$safe`">$inner</a>"
}
# Puts html at the stored cursor position, or at the end of the page when there is none
function Insert-Html($html) {
    $range = Get-Range
    $script:savedRange = $null
    try {
        if ($range) { $range.pasteHTML($html); return }
    } catch { }
    try { $editBrowser.Document.body.insertAdjacentHTML('beforeend', $html) }
    catch { Add-Log ('Could not insert: ' + $_.Exception.Message) 'warn' }
}

# Mail clients only reliably show hosted images, so only direct links to a png, jpg or gif are accepted
function Add-ImageFromUrl {
    Save-Selection
    $hint = (T 'Direct link to the image, for example https://i.imgur.com/example.png') + "`n" + (T 'Upload the image to a site like imgur first. The link has to end with .png, .jpg, .jpeg or .gif.')
    $url = Read-Text (T 'Insert image') $hint
    if (-not $url) { return }
    $url = $url.Trim()
    if ($url -notmatch '^https://[^\s"''<>]+\.(png|jpe?g|gif)(\?[^\s"''<>]*)?$') {
        Add-Log (T 'Image links must be https and end with .png, .jpg, .jpeg or .gif') 'warn'
        return
    }
    Insert-Html ('<img src="{0}" alt="" border="0">' -f [Net.WebUtility]::HtmlEncode($url))
}
Build-Toolbar
$editToolbar.Add_PreviewMouseLeftButtonDown({
        $open = @($editToolbar.Children | Where-Object { $_ -is [Windows.Controls.Primitives.Popup] -and $_.IsOpen })
        if (-not $open) { Push-History; Save-Selection }
    })
Update-EditList $null

# ---------------------------------------------------------------- events
$btnConnect.Add_Click({
        $needInstall = -not $Mock -and -not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)
        if ($needInstall -and -not (Confirm-Action (T 'The Exchange Online PowerShell module is not installed. Install it now for your user account?') 0)) {
            $loginStatus.Text = T 'The Exchange Online module is required.'
            return
        }
        Add-Log 'Connecting to Exchange Online'
        $loginStatus.Text = T 'Waiting for the Microsoft sign-in window...'
        $loginStatus.Foreground = $win.Resources['MutedBrush']
        $script:afterJob = { if ($script:connected) { Invoke-Worker @($workerLoad); $script:afterJob = { Start-StatusRefresh } } }
        Invoke-Worker $(if ($needInstall) { @($workerInstall, $workerConnect) } else { @($workerConnect) })
    })

function Disconnect-Safe {
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript('Disconnect-ExchangeOnline -Confirm:$false')
    try { $ps.Invoke() | Out-Null } finally { $ps.Dispose() }
}

$btnDisconnect.Add_Click({
        try { Disconnect-Safe } catch { }
        $script:connected = $false
        $boxes.Clear(); $byUpn.Clear()
        $loginStatus.Text = ''
        Add-Log 'Signed out'
        Update-State
    })

$btnNew.Add_Click({ New-Reply })
$btnRefresh.Add_Click({ Start-StatusRefresh })
$btnCancel.Add_Click({ $sync.Cancel = $true })
$btnDisableChecked.Add_Click({ Disable-Reply (Get-Checked) })

$lstBoxes.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler] {
        param($s, $e)
        $src = $e.OriginalSource
        if ($src -is [Windows.Controls.Button]) {
            $item = $src.DataContext
            if ($src.Name -eq 'rowEdit') { Edit-Reply $item }
            elseif ($src.Name -eq 'rowToggle') { if ($item.Active) { Disable-Reply @($item) } else { Edit-Reply $item } }
        }
        Update-State
    })

$lstBoxes.Add_MouseDoubleClick({
        $m = $lstBoxes.SelectedItem
        if ($m -and $m.Html) { Show-Message ((T 'Current message of {0}') -f $m.Name) $m.Html }
    })

$view.Filter = [Predicate[object]] {
    param($o)
    if ($chkActiveOnly.IsChecked -and -not $o.Active) { return $false }
    $f = $txtFilter.Text.Trim()
    -not $f -or $o.Name -like "*$f*" -or $o.Upn -like "*$f*" -or $o.Snippet -like "*$f*"
}
# Filtering a long list on every key stroke is wasteful, wait until typing pauses
$filterTimer = New-Object Windows.Threading.DispatcherTimer
$filterTimer.Interval = [TimeSpan]::FromMilliseconds(250)
$filterTimer.Add_Tick({ $filterTimer.Stop(); $view.Refresh() })
$txtFilter.Add_TextChanged({ $filterTimer.Stop(); $filterTimer.Start() })
$chkActiveOnly.Add_Click({ $view.Refresh() })
$btnSelAll.Add_Click({ foreach ($o in $view) { $o.Selected = $true }; Update-State })
$btnSelNone.Add_Click({ foreach ($o in $boxes) { $o.Selected = $false }; Update-State })

$tabBoxes.Add_Checked({ $paneBoxes.Visibility = 'Visible'; $paneEditor.Visibility = 'Collapsed' })
$tabEditor.Add_Checked({
        $paneBoxes.Visibility = 'Collapsed'; $paneEditor.Visibility = 'Visible'
        if (-not $script:editorReady -and $null -eq $script:pendingHtml) { $script:pendingHtml = $blankDoc }
        $script:editorReady = $true
        $navTimer.Stop(); $navTimer.Start()
    })

$lstEdit.Add_SelectionChanged({ if ($lstEdit.SelectedItem) { Open-Template $lstEdit.SelectedItem } })
$btnTplNew.Add_Click({ $lstEdit.SelectedItem = $null; Open-Template $null })

$btnTplOpen.Add_Click({
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Filter = 'Templates (*.html;*.htm;*.md)|*.html;*.htm;*.md'
        if (-not $dlg.ShowDialog()) { return }
        try {
            $text = [IO.File]::ReadAllText($dlg.FileName, [Text.Encoding]::UTF8)
            if ($dlg.FileName -like '*.md') { $text = ConvertFrom-Markdown $text }
            $lstEdit.SelectedItem = $null
            $script:editingPath = $null
            $txtTplName.Text = [IO.Path]::GetFileNameWithoutExtension($dlg.FileName)
            $chkSource.IsChecked = $false
            Request-EditorNav (Clear-Html $text)
            Add-Log ('Loaded {0}, review it and click Save' -f [IO.Path]::GetFileName($dlg.FileName))
        } catch { Add-Log "Could not open file: $($_.Exception.Message)" 'err' }
    })

$btnTplDelete.Add_Click({
        $t = $lstEdit.SelectedItem
        if (-not $t) { return }
        if ($t.BuiltIn) { Add-Log 'Built-in templates cannot be deleted' 'warn'; return }
        if ([Windows.MessageBox]::Show(((T 'Delete the template ''{0}''?') -f $t.Title), 'm365-global-auto-reply', 'YesNo', 'Question') -ne 'Yes') { return }
        try { [IO.File]::Delete($t.Path); Update-EditList $null; Open-Template $null; Add-Log "Template deleted: $($t.Title)" }
        catch { Add-Log "Could not delete template: $($_.Exception.Message)" 'err' }
    })

$btnTplSave.Add_Click({
        $safe = ($txtTplName.Text -replace '[\\/:*?"<>|]', '').Trim().TrimEnd('.')
        if (-not $safe) { Add-Log 'Give the template a name first' 'warn'; return }
        $path = Join-Path $userDir "$safe.html"
        if ((Test-Path -LiteralPath $path) -and $path -ne $script:editingPath) {
            if ([Windows.MessageBox]::Show(((T 'A template named ''{0}'' already exists. Replace it?') -f $safe), 'm365-global-auto-reply', 'YesNo', 'Question') -ne 'Yes') { return }
        }
        try {
            $html = Get-EditorHtml (-not $chkSource.IsChecked)
            [IO.File]::WriteAllText($path, $html, $utf8Bom)
            Update-EditList $path
            Add-Log "Template saved: $safe" 'ok'
        } catch { Add-Log "Could not save template: $($_.Exception.Message)" 'err' }
    })

$chkSource.Add_Checked({
        try { $txtSource.Text = Get-EditorHtml $true } catch { $txtSource.Text = ''; Add-Log ('Could not read the page: ' + $_.Exception.Message) 'warn' }
        $txtSource.Visibility = 'Visible'
        $editBrowser.Visibility = 'Collapsed'
    })
$chkSource.Add_Unchecked({
        $txtSource.Visibility = 'Collapsed'
        $editBrowser.Visibility = 'Visible'
        if ($txtSource.Text) { Request-EditorNav (Clear-Html $txtSource.Text) }
    })

$btnCopyLog.Add_Click({ try { [Windows.Clipboard]::SetText((($logItems | ForEach-Object { $_.Text }) -join "`r`n")) } catch { } })
$btnOpenLog.Add_Click({ Start-Process explorer.exe $logDir })

$lstBoxes.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler] { Update-State })

# Polls the worker queue
$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(120)
$timer.Add_Tick({
        $finished = $script:job -and $script:job.Handle.IsCompleted
        if ($script:job -and -not $script:warned -and ((Get-Date) - $script:lastActivity).TotalSeconds -gt 45) {
            Add-Log 'Still waiting for Exchange Online. This can take a while, use Cancel job to stop.' 'warn'
            $script:warned = $true
        }
        $m = $null
        $changed = $false
        $handled = 0
        # at most 150 messages per tick so a burst of status updates cannot freeze the window
        while ($handled -lt 150 -and $sync.Queue.TryDequeue([ref]$m)) { Receive-Message $m; $changed = $true; $handled++ }
        if ($finished) { Complete-Job }
        elseif ($changed) { Update-State }
    })

$win.Add_Closing({
        param($s, $e)
        if ($script:job) {
            $answer = [Windows.MessageBox]::Show((T 'A job is still running. Cancel it and exit?'), 'm365-global-auto-reply', 'YesNo', 'Question')
            if ($answer -ne 'Yes') { $e.Cancel = $true; return }
            $sync.Cancel = $true
            $null = $script:job.Handle.AsyncWaitHandle.WaitOne(15000)
        }
        if ($script:connected) { try { Disconnect-Safe } catch { } }
        $timer.Stop(); $navTimer.Stop()
        try { $editBrowser.Dispose() } catch { }
        $rs.Dispose()
    })

$win.Dispatcher.Add_UnhandledException({
        param($s, $e)
        try {
            Add-Log ('Unexpected error: ' + $e.Exception.Message) 'err'
            $rec = $e.Exception.ErrorRecord
            if ($rec) { Add-Log ($rec.ScriptStackTrace -replace '\r?\n', ' <- ') 'err' }
        } catch { }
        $e.Handled = $true
    })

if ($Mock) {
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript($mockInit.ToString())
    $ps.Invoke() | Out-Null
    $ps.Dispose()
}

function Set-Language($code) {
    $script:lang = $code
    Set-Strings $win
    $flagCur.ContentTemplate = if ($code -eq 'de') { $win.Resources['FlagDe'] } else { $win.Resources['FlagEn'] }
    $langText.Text = if ($code -eq 'de') { 'Deutsch' } else { 'English' }
    $editToolbar.Children.Clear()
    Build-Toolbar
    [Mailbox]::EnableText = T 'Enable'
    [Mailbox]::DisableText = T 'Disable'
    foreach ($o in $boxes) { $o.Active = $o.Active }
    Update-State
    try { Set-Content -LiteralPath (Join-Path $dataDir 'language.txt') -Value $code } catch { }
}

$btnLang.Add_Click({ $popLang.IsOpen = -not $popLang.IsOpen })
$langEn.Add_Click({ $popLang.IsOpen = $false; Set-Language 'en' })
$langDe.Add_Click({ $popLang.IsOpen = $false; Set-Language 'de' })

$saved = try { (Get-Content -LiteralPath (Join-Path $dataDir 'language.txt') -ErrorAction Stop | Select-Object -First 1).Trim() } catch { '' }
Set-Language $(if ($saved -eq 'de') { 'de' } else { 'en' })
Update-State
$timer.Start()

try { [void]$win.ShowDialog() }
catch { [Windows.MessageBox]::Show($_.Exception.Message, 'm365-global-auto-reply', 'OK', 'Error') | Out-Null }
