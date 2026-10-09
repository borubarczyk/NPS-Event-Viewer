<#
.SYNOPSIS
    NPS Event Viewer - GUI (WPF) do przeglądania i filtrowania zdarzeń autoryzacji NPS.

.DESCRIPTION
    Zakładka 1 - Dziennik Security:
      6272 udzielono dostępu, 6273 odmowa, 6274 odrzucono żądanie, 6276 kwarantanna,
      6277 dostęp warunkowy, 6278 pełny dostęp, 6279 konto zablokowane, 6280 konto odblokowane.
    Zakładka 2 - Pliki logów RADIUS (.log) z rejestrowania NPS (Accounting -> Log File):
      format DTS-compliant (XML, zalecany) oraz IAS (best effort). Lokalizacja do wyboru:
      folder, pojedyncze pliki albo ścieżka UNC (np. \\NPS01\c$\Windows\System32\LogFiles).
      Odpowiedzi Access-Accept / Access-Reject są uzupełniane danymi z poprzedzającego
      Access-Request (User-Name, MAC, port) od tego samego klienta RADIUS.
    Zakładka 3 - Zdarzenia systemowe NPS (dziennik System, opcjonalnie Application), czyli wszystko
      co nie jest audytem autoryzacji: niezgodny shared secret / Message-Authenticator (ID 18),
      brak Message-Authenticator (14), nieznany klient RADIUS (13), problemy z kontrolerem domeny itd.
    Filtrowanie po wyniku, switchu (klient RADIUS), zasadzie sieciowej, typie uwierzytelniania,
    użytkowniku / MAC (dowolny format), komputerze / serwerze i pełnotekstowo.
    Wczytywanie działa w tle, okno się nie blokuje.

.NOTES
    Wymagania: uruchom jako administrator (albo konto w grupie Event Log Readers).
    Serwer zdalny: reguła firewalla "Remote Event Log Management" na serwerze NPS.
    Brak zdarzeń: sprawdź audyt -> auditpol /get /subcategory:"Network Policy Server"
    Pliki .log: domyślnie %SystemRoot%\System32\LogFiles\IN*.log (sprawdź w konsoli NPS:
    Accounting -> Change Log File Properties). Plik jest otwierany bez blokowania - NPS może pisać dalej.
    Zapisz plik jako UTF-8 z BOM (polskie znaki w Windows PowerShell 5.1).
#>

#Requires -Version 5.1

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $exe = (Get-Process -Id $PID).Path
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# --- XAML ----------------------------------------------------------------------------------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="NPS Event Viewer" Height="900" Width="1560" MinHeight="600" MinWidth="960"
        WindowStartupLocation="CenterScreen" Background="#15161C"
        FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <!-- Paleta -->
    <SolidColorBrush x:Key="Card"   Color="#1E2029"/>
    <SolidColorBrush x:Key="Field"  Color="#272A36"/>
    <SolidColorBrush x:Key="Hover"  Color="#2B2E3B"/>
    <SolidColorBrush x:Key="Line"   Color="#343847"/>
    <SolidColorBrush x:Key="Accent" Color="#4F8CFF"/>
    <SolidColorBrush x:Key="Text"   Color="#E8EAF0"/>
    <SolidColorBrush x:Key="Muted"  Color="#8A90A2"/>

    <!-- Skala typografii: tytuł 20, karta 14, tekst 13, drobny / monospace 12, etykieta 11.
         Wysokość pól, list i przycisków: 28 (małe przyciski 24). Odstępy: 4 / 8 / 12 / 16. -->
    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
    </Style>
    <Style x:Key="PageTitle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="FontSize" Value="20"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="PageSubtitle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Margin" Value="0,2,0,0"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="Caption" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="CardTitle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="Counts" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>

    <!-- Karty i układ -->
    <Style x:Key="CardBox" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource Card}"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="14,12"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <!-- Karta z polami w WrapPanel: dolny odstęp pól (8) + padding (4) = 12, tyle co u góry -->
    <Style x:Key="FormCard" TargetType="Border" BasedOn="{StaticResource CardBox}">
      <Setter Property="Padding" Value="14,12,14,4"/>
    </Style>
    <Style x:Key="PanelCard" TargetType="Border" BasedOn="{StaticResource CardBox}">
      <Setter Property="Margin" Value="0"/>
    </Style>
    <Style x:Key="CardHeader" TargetType="DockPanel">
      <Setter Property="Margin" Value="0,0,0,8"/>
      <Setter Property="MinHeight" Value="24"/>
      <Setter Property="LastChildFill" Value="True"/>
    </Style>
    <!-- Pole formularza: etykieta + kontrolka; wszystkie pola mają tę samą wysokość -->
    <Style x:Key="FieldBox" TargetType="StackPanel">
      <Setter Property="Margin" Value="0,0,12,8"/>
      <Setter Property="VerticalAlignment" Value="Bottom"/>
    </Style>
    <!-- Wiersz pól wyboru / przycisków o wysokości kontrolki (28) - wyrównany z polami tekstowymi -->
    <Style x:Key="InlineRow" TargetType="StackPanel">
      <Setter Property="Orientation" Value="Horizontal"/>
      <Setter Property="Height" Value="28"/>
    </Style>
    <Style x:Key="Splitter" TargetType="GridSplitter">
      <Setter Property="Width" Value="12"/>
      <Setter Property="HorizontalAlignment" Value="Stretch"/>
      <Setter Property="VerticalAlignment" Value="Stretch"/>
      <Setter Property="ResizeBehavior" Value="PreviousAndNext"/>
      <Setter Property="ResizeDirection" Value="Columns"/>
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Cursor" Value="SizeWE"/>
      <Setter Property="ToolTip" Value="Przeciągnij, aby zmienić szerokość panelu"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="GridSplitter">
            <Grid Background="Transparent">
              <Rectangle x:Name="grip" Width="2" Height="40" RadiusX="1" RadiusY="1" Fill="{StaticResource Line}"
                         HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="grip" Property="Fill" Value="{StaticResource Accent}"/>
                <Setter TargetName="grip" Property="Height" Value="64"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Przyciski -->
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Background" Value="#2C3040"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="Padding" Value="12,0"/>
      <Setter Property="MinHeight" Value="28"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Opacity" Value="0.85"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="b" Property="BorderBrush" Value="{StaticResource Accent}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="b" Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnPrimary" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Background" Value="{StaticResource Accent}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Accent}"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="16,0"/>
    </Style>
    <Style x:Key="BtnSmall" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Padding" Value="10,0"/>
      <Setter Property="MinHeight" Value="24"/>
      <Setter Property="FontSize" Value="12"/>
    </Style>

    <!-- Pola tekstowe: jedna linia = 28 px; pole wielowierszowe (szczegóły) ustawia własny Padding -->
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{StaticResource Field}"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="CaretBrush" Value="{StaticResource Text}"/>
      <Setter Property="SelectionBrush" Value="{StaticResource Accent}"/>
      <Setter Property="Padding" Value="6,0"/>
      <Setter Property="MinHeight" Value="28"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="6" SnapsToDevicePixels="True">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"
                            VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="BorderBrush" Value="#4A5068"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="b" Property="BorderBrush" Value="{StaticResource Accent}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="b" Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="DetailsBox" TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="AcceptsReturn" Value="True"/>
      <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
      <Setter Property="VerticalContentAlignment" Value="Stretch"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Padding" Value="8,6"/>
    </Style>

    <Style TargetType="ComboBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="MinHeight" Value="28"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton Focusable="False" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="bd" Background="#272A36" BorderBrush="#343847" BorderThickness="1" CornerRadius="6" SnapsToDevicePixels="True">
                      <Path HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0"
                            Data="M 0 0 L 4 4 L 8 0" Stroke="#8A90A2" StrokeThickness="1.6"/>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="bd" Property="BorderBrush" Value="#4A5068"/>
                      </Trigger>
                      <Trigger Property="IsChecked" Value="True">
                        <Setter TargetName="bd" Property="BorderBrush" Value="#4F8CFF"/>
                      </Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter IsHitTestVisible="False" Margin="8,0,26,0" VerticalAlignment="Center"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
              <Popup IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom" AllowsTransparency="True"
                     Focusable="False" PopupAnimation="Slide">
                <Border Background="#272A36" BorderBrush="#343847" BorderThickness="1" CornerRadius="6" Margin="0,2,0,0" Padding="0,4"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}" MaxHeight="360">
                  <ScrollViewer>
                    <StackPanel IsItemsHost="True"/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Padding" Value="8,5"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="bd" Background="Transparent" Padding="{TemplateBinding Padding}">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#2F3A55"/>
              </Trigger>
              <Trigger Property="IsHighlighted" Value="True">
                <Setter TargetName="bd" Property="Background" Value="#3A4A70"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Pole wyboru w ciemnym motywie (domyślne jest jasne i odstaje od reszty) -->
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,16,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Grid Background="Transparent">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <Border x:Name="box" Width="16" Height="16" CornerRadius="4" BorderThickness="1" VerticalAlignment="Center"
                      Background="{StaticResource Field}" BorderBrush="#4A5068" SnapsToDevicePixels="True">
                <Path x:Name="mark" Data="M 3 8 L 6.5 11.5 L 13 4.5" Stroke="White" StrokeThickness="2"
                      StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round" Visibility="Collapsed"/>
              </Border>
              <ContentPresenter Grid.Column="1" Margin="8,0,0,0" VerticalAlignment="Center" RecognizesAccessKey="True"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource Accent}"/>
              </Trigger>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource Accent}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="box" Property="Background" Value="{StaticResource Accent}"/>
                <Setter TargetName="box" Property="BorderBrush" Value="{StaticResource Accent}"/>
                <Setter TargetName="mark" Property="Visibility" Value="Visible"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="TabControl">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabControl">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border BorderBrush="#343847" BorderThickness="0,0,0,1" Margin="0,0,0,12">
                <TabPanel IsItemsHost="True"/>
              </Border>
              <ContentPresenter Grid.Row="1" ContentSource="SelectedContent"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Background="Transparent" BorderBrush="Transparent" BorderThickness="0,0,0,2"
                    Padding="14,8" Margin="0,0,4,-1">
              <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Foreground" Value="#E8EAF0"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#4F8CFF"/>
                <Setter Property="Foreground" Value="#E8EAF0"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="{StaticResource Field}"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="RowBackground" Value="#272A36"/>
      <Setter Property="AlternatingRowBackground" Value="#2B2E3B"/>
      <Setter Property="GridLinesVisibility" Value="None"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="RowHeight" Value="28"/>
      <Setter Property="SelectionMode" Value="Single"/>
      <Setter Property="EnableRowVirtualization" Value="True"/>
    </Style>
    <!-- Nagłówek i komórka mają ten sam lewy odstęp (6), więc tekst kolumn jest wyrównany z nagłówkiem -->
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#1E2029"/>
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Padding" Value="6,6"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Padding" Value="6,0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#33405E"/>
          <Setter Property="Foreground" Value="White"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="ResultCell" TargetType="TextBlock">
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Foreground" Value="#FBBF24"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding Level}" Value="OK">
          <Setter Property="Foreground" Value="#4ADE80"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Level}" Value="Error">
          <Setter Property="Foreground" Value="#F87171"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Level}" Value="Info">
          <Setter Property="Foreground" Value="#8A90A2"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>

    <Style TargetType="ProgressBar">
      <Setter Property="Height" Value="6"/>
      <Setter Property="Foreground" Value="{StaticResource Accent}"/>
      <Setter Property="Background" Value="{StaticResource Field}"/>
      <Setter Property="BorderThickness" Value="0"/>
    </Style>

    <!-- Pasek przewijania w ciemnym motywie: cienki, bez strzałek -->
    <Style x:Key="ScrollThumb" TargetType="Thumb">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Thumb">
            <Border x:Name="t" Background="#4A5068" CornerRadius="4"/>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="t" Property="Background" Value="#5E6584"/>
              </Trigger>
              <Trigger Property="IsDragging" Value="True">
                <Setter TargetName="t" Property="Background" Value="{StaticResource Accent}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="ScrollPage" TargetType="RepeatButton">
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="IsTabStop" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RepeatButton">
            <Border Background="Transparent"/>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ScrollBar">
      <Setter Property="Background" Value="Transparent"/>
      <Setter Property="Width" Value="10"/>
      <Setter Property="MinWidth" Value="10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Border Background="{TemplateBinding Background}" Padding="2">
              <Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.DecreaseRepeatButton>
                  <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageUpCommand"/>
                </Track.DecreaseRepeatButton>
                <Track.Thumb>
                  <Thumb Style="{StaticResource ScrollThumb}" MinHeight="24"/>
                </Track.Thumb>
                <Track.IncreaseRepeatButton>
                  <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageDownCommand"/>
                </Track.IncreaseRepeatButton>
              </Track>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="Orientation" Value="Horizontal">
          <Setter Property="Width" Value="Auto"/>
          <Setter Property="MinWidth" Value="0"/>
          <Setter Property="Height" Value="10"/>
          <Setter Property="MinHeight" Value="10"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="ScrollBar">
                <Border Background="{TemplateBinding Background}" Padding="2">
                  <Track x:Name="PART_Track" IsDirectionReversed="False">
                    <Track.DecreaseRepeatButton>
                      <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageLeftCommand"/>
                    </Track.DecreaseRepeatButton>
                    <Track.Thumb>
                      <Thumb Style="{StaticResource ScrollThumb}" MinWidth="24"/>
                    </Track.Thumb>
                    <Track.IncreaseRepeatButton>
                      <RepeatButton Style="{StaticResource ScrollPage}" Command="ScrollBar.PageRightCommand"/>
                    </Track.IncreaseRepeatButton>
                  </Track>
                </Border>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Trigger>
      </Style.Triggers>
    </Style>
  </Window.Resources>

  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- NAGŁÓWEK -->
    <DockPanel Grid.Row="0" Margin="0,0,0,12">
      <Border DockPanel.Dock="Right" Background="{StaticResource Card}" CornerRadius="14" Padding="12,0" Height="28" VerticalAlignment="Center">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse Name="dotState" Width="8" Height="8" Fill="#8A90A2" VerticalAlignment="Center" Margin="0,0,8,0"/>
          <TextBlock Name="lblState" Text="Brak danych" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
        </StackPanel>
      </Border>
      <Button Name="btnHistory" DockPanel.Dock="Right" Content="Historia urządzenia / użytkownika..." Style="{StaticResource Btn}" Margin="0,0,12,0"
              ToolTip="Oś czasu wszystkich zdarzeń jednego MAC, użytkownika albo komputera (też: dwuklik na wierszu tabeli)"/>
      <StackPanel Margin="0,0,16,0" VerticalAlignment="Center">
        <TextBlock Text="NPS Event Viewer" Style="{StaticResource PageTitle}"/>
        <TextBlock Name="lblSubtitle" Text="Zdarzenia autoryzacji Network Policy Server - dziennik Security (6272-6280), pliki logów RADIUS (.log) i zdarzenia systemowe"
                   Style="{StaticResource PageSubtitle}"/>
      </StackPanel>
    </DockPanel>

    <TabControl Name="tabMain" Grid.Row="1">

      <!-- ========================= ZAKŁADKA 1: DZIENNIK SECURITY ========================= -->
      <TabItem Header="Dziennik Security">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <!-- ZAPYTANIE -->
          <Border Grid.Row="0" Style="{StaticResource FormCard}">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="190">
                <TextBlock Text="SERWER NPS (puste = lokalny)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtServer"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="170">
                <TextBlock Text="ZAKRES CZASU" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbRange" SelectedIndex="1">
                  <ComboBoxItem Content="Ostatnia godzina"/>
                  <ComboBoxItem Content="Ostatnie 24 godziny"/>
                  <ComboBoxItem Content="Ostatnie 7 dni"/>
                  <ComboBoxItem Content="Ostatnie 30 dni"/>
                  <ComboBoxItem Content="Własny zakres"/>
                </ComboBox>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="OD (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtFrom" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="DO (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtTo" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="100">
                <TextBlock Text="MAKS. ZDARZEŃ" Style="{StaticResource Caption}"/>
                <TextBox Name="txtMax" Text="5000" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <StackPanel Style="{StaticResource InlineRow}">
                  <CheckBox Name="chkAuto" Content="Auto-odświeżanie co 60 s" Margin="0"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <StackPanel Style="{StaticResource InlineRow}">
                  <Button Name="btnLoad" Content="Wczytaj zdarzenia" Style="{StaticResource BtnPrimary}"/>
                </StackPanel>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- FILTRY -->
          <Border Grid.Row="1" Style="{StaticResource FormCard}">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="120">
                <TextBlock Text="WYNIK" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbResult" SelectedIndex="0">
                  <ComboBoxItem Content="Wszystkie"/>
                  <ComboBoxItem Content="Udzielono"/>
                  <ComboBoxItem Content="Odmowa"/>
                  <ComboBoxItem Content="Inne"/>
                </ComboBox>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="170">
                <TextBlock Text="SWITCH (KLIENT RADIUS)" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbSwitch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="180">
                <TextBlock Text="ZASADA SIECIOWA" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbPolicy"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="130">
                <TextBlock Text="UWIERZYTELNIANIE" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbAuth"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="160">
                <TextBlock Text="UŻYTKOWNIK / MAC" Style="{StaticResource Caption}"/>
                <TextBox Name="txtUser" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="130">
                <TextBlock Text="KOMPUTER" Style="{StaticResource Caption}"/>
                <TextBox Name="txtComputer"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="170">
                <TextBlock Text="SZUKAJ WSZĘDZIE" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSearch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <StackPanel Style="{StaticResource InlineRow}">
                  <Button Name="btnResetFilters" Content="Wyczyść filtry" Style="{StaticResource Btn}"/>
                </StackPanel>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- TABELA + SZCZEGÓŁY (szerokość panelu szczegółów: przeciągnij separator) -->
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="5*" MinWidth="360"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="2*" MinWidth="300"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Style="{StaticResource PanelCard}">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Style="{StaticResource CardHeader}">
                  <Button Name="btnExport" DockPanel.Dock="Right" Content="Eksport CSV" Style="{StaticResource BtnSmall}" Margin="12,0,0,0"
                          ToolTip="Zapisz zdarzenia widoczne w tabeli (po filtrach) do pliku CSV"/>
                  <TextBlock Text="Zdarzenia" Style="{StaticResource CardTitle}" DockPanel.Dock="Left" Margin="0,0,12,0"/>
                  <TextBlock Name="lblCounts" Text="" Style="{StaticResource Counts}" HorizontalAlignment="Right"/>
                </DockPanel>
                <DataGrid Name="dgEvents" Grid.Row="1">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Czas" Binding="{Binding TimeStr}" Width="155" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Wynik" Binding="{Binding Result}" Width="95" ElementStyle="{StaticResource ResultCell}"/>
                    <DataGridTextColumn Header="Użytkownik" Binding="{Binding User}" Width="150"/>
                    <DataGridTextColumn Header="MAC klienta" Binding="{Binding CallingStation}" Width="145" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Komputer" Binding="{Binding Machine}" Width="150"/>
                    <DataGridTextColumn Header="Switch" Binding="{Binding Client}" Width="160"/>
                    <DataGridTextColumn Header="Port" Binding="{Binding NasPort}" Width="55"/>
                    <DataGridTextColumn Header="Zasada sieciowa" Binding="{Binding Policy}" Width="180"/>
                    <DataGridTextColumn Header="Uwierz." Binding="{Binding AuthType}" Width="75"/>
                    <DataGridTextColumn Header="Kod" Binding="{Binding ReasonCode}" Width="50"/>
                    <DataGridTextColumn Header="Przyczyna" Binding="{Binding Reason}" Width="*" MinWidth="160"/>
                  </DataGrid.Columns>
                </DataGrid>
              </Grid>
            </Border>

            <GridSplitter Grid.Column="1" Style="{StaticResource Splitter}"/>

            <Border Grid.Column="2" Style="{StaticResource PanelCard}">
              <Grid Name="sideEv">
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*" MinHeight="90"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Style="{StaticResource CardHeader}">
                  <Button Name="btnCopy" DockPanel.Dock="Right" Content="Kopiuj" Style="{StaticResource BtnSmall}"/>
                  <TextBlock Text="Szczegóły" Style="{StaticResource CardTitle}"/>
                </DockPanel>
                <TextBox Name="txtDetails" Grid.Row="1" Style="{StaticResource DetailsBox}" Text="Zaznacz zdarzenie na liście."/>
                <TextBlock Grid.Row="2" Text="Statystyki widoku" Style="{StaticResource CardTitle}" Margin="0,12,0,8"/>
                <Border Name="bdStats" Grid.Row="3" Background="{StaticResource Field}" CornerRadius="6" Padding="8,6" MaxHeight="240">
                  <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBlock Name="txtStats" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap" Text="-"/>
                  </ScrollViewer>
                </Border>
              </Grid>
            </Border>
          </Grid>
        </Grid>
      </TabItem>

      <!-- ========================= ZAKŁADKA 2: PLIKI .LOG ========================= -->
      <TabItem Header="Pliki logów RADIUS (.log)">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <!-- ŹRÓDŁO -->
          <Border Grid.Row="0" Style="{StaticResource FormCard}">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="380">
                <TextBlock Text="LOKALIZACJA (folder, plik lub UNC; kilka ścieżek rozdziel ;)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogPath" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <StackPanel Style="{StaticResource InlineRow}">
                  <Button Name="btnLogFolder" Content="Folder..." Style="{StaticResource Btn}" Margin="0,0,8,0"/>
                  <Button Name="btnLogFiles" Content="Pliki..." Style="{StaticResource Btn}" Margin="0,0,8,0"/>
                  <Button Name="btnLogDefault" Content="Domyślna" Style="{StaticResource Btn}"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="100">
                <TextBlock Text="MASKA PLIKÓW" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogMask" Text="IN*.log" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="170">
                <TextBlock Text="ZAKRES CZASU" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbLogRange" SelectedIndex="1">
                  <ComboBoxItem Content="Ostatnia godzina"/>
                  <ComboBoxItem Content="Ostatnie 24 godziny"/>
                  <ComboBoxItem Content="Ostatnie 7 dni"/>
                  <ComboBoxItem Content="Ostatnie 30 dni"/>
                  <ComboBoxItem Content="Wszystko"/>
                  <ComboBoxItem Content="Własny zakres"/>
                </ComboBox>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="OD (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogFrom" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="DO (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogTo" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="100">
                <TextBlock Text="MAKS. WPISÓW" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogMax" Text="20000" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <StackPanel Style="{StaticResource InlineRow}">
                  <Button Name="btnLogLoad" Content="Wczytaj logi" Style="{StaticResource BtnPrimary}"/>
                </StackPanel>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- FILTRY -->
          <Border Grid.Row="1" Style="{StaticResource FormCard}">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="130">
                <TextBlock Text="WYNIK" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbLResult" SelectedIndex="0">
                  <ComboBoxItem Content="Wszystkie"/>
                  <ComboBoxItem Content="Udzielono"/>
                  <ComboBoxItem Content="Odmowa"/>
                  <ComboBoxItem Content="Inne"/>
                </ComboBox>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="190">
                <TextBlock Text="SWITCH (KLIENT RADIUS)" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbLSwitch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="210">
                <TextBlock Text="ZASADA SIECIOWA" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbLPolicy"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="UWIERZYTELNIANIE" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbLAuth"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="170">
                <TextBlock Text="UŻYTKOWNIK / MAC" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLUser" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="150">
                <TextBlock Text="SERWER NPS" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLComputer"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="190">
                <TextBlock Text="SZUKAJ WSZĘDZIE" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLSearch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <TextBlock Text="POKAŻ TAKŻE" Style="{StaticResource Caption}"/>
                <StackPanel Style="{StaticResource InlineRow}">
                  <CheckBox Name="chkLogReq" Content="Access-Request / Challenge"/>
                  <CheckBox Name="chkLogAcct" Content="Accounting" Margin="0"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <StackPanel Style="{StaticResource InlineRow}">
                  <Button Name="btnLResetFilters" Content="Wyczyść filtry" Style="{StaticResource Btn}"/>
                </StackPanel>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- TABELA + SZCZEGÓŁY -->
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="5*" MinWidth="360"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="2*" MinWidth="300"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Style="{StaticResource PanelCard}">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Style="{StaticResource CardHeader}">
                  <Button Name="btnLExport" DockPanel.Dock="Right" Content="Eksport CSV" Style="{StaticResource BtnSmall}" Margin="12,0,0,0"
                          ToolTip="Zapisz wpisy widoczne w tabeli (po filtrach) do pliku CSV"/>
                  <TextBlock Text="Wpisy z logów" Style="{StaticResource CardTitle}" DockPanel.Dock="Left" Margin="0,0,12,0"/>
                  <TextBlock Name="lblLCounts" Text="" Style="{StaticResource Counts}" HorizontalAlignment="Right"/>
                </DockPanel>
                <DataGrid Name="dgLog" Grid.Row="1">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Czas" Binding="{Binding TimeStr}" Width="155" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Wynik" Binding="{Binding Result}" Width="95" ElementStyle="{StaticResource ResultCell}"/>
                    <DataGridTextColumn Header="Użytkownik" Binding="{Binding User}" Width="150"/>
                    <DataGridTextColumn Header="MAC klienta" Binding="{Binding CallingStation}" Width="145" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Switch" Binding="{Binding Client}" Width="160"/>
                    <DataGridTextColumn Header="Port" Binding="{Binding NasPort}" Width="55"/>
                    <DataGridTextColumn Header="Zasada sieciowa" Binding="{Binding Policy}" Width="180"/>
                    <DataGridTextColumn Header="Uwierz." Binding="{Binding AuthType}" Width="80"/>
                    <DataGridTextColumn Header="Kod" Binding="{Binding ReasonCode}" Width="50"/>
                    <DataGridTextColumn Header="Przyczyna" Binding="{Binding Reason}" Width="*" MinWidth="160"/>
                    <DataGridTextColumn Header="Serwer NPS" Binding="{Binding Machine}" Width="110"/>
                  </DataGrid.Columns>
                </DataGrid>
              </Grid>
            </Border>

            <GridSplitter Grid.Column="1" Style="{StaticResource Splitter}"/>

            <Border Grid.Column="2" Style="{StaticResource PanelCard}">
              <Grid Name="sideLog">
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*" MinHeight="90"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Style="{StaticResource CardHeader}">
                  <Button Name="btnLCopy" DockPanel.Dock="Right" Content="Kopiuj" Style="{StaticResource BtnSmall}"/>
                  <TextBlock Text="Szczegóły" Style="{StaticResource CardTitle}"/>
                </DockPanel>
                <TextBox Name="txtLDetails" Grid.Row="1" Style="{StaticResource DetailsBox}" Text="Wskaż lokalizację logów i kliknij Wczytaj logi."/>
                <TextBlock Grid.Row="2" Text="Statystyki widoku" Style="{StaticResource CardTitle}" Margin="0,12,0,8"/>
                <Border Name="bdLStats" Grid.Row="3" Background="{StaticResource Field}" CornerRadius="6" Padding="8,6" MaxHeight="240">
                  <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBlock Name="txtLStats" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap" Text="-"/>
                  </ScrollViewer>
                </Border>
              </Grid>
            </Border>
          </Grid>
        </Grid>
      </TabItem>
      <!-- ========================= ZAKŁADKA 3: ZDARZENIA SYSTEMOWE NPS ========================= -->
      <TabItem Header="Zdarzenia systemowe NPS">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>

          <!-- ZAPYTANIE -->
          <Border Grid.Row="0" Style="{StaticResource FormCard}">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="190">
                <TextBlock Text="SERWER NPS (puste = lokalny)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSServer"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="250">
                <TextBlock Text="ŹRÓDŁA ZDARZEŃ (rozdziel ;)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSProviders" Text="NPS; IAS; Microsoft-Windows-NPS" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <TextBlock Text="DZIENNIKI" Style="{StaticResource Caption}"/>
                <StackPanel Style="{StaticResource InlineRow}">
                  <CheckBox Name="chkSSystem" Content="System" IsChecked="True"/>
                  <CheckBox Name="chkSApp" Content="Application" Margin="0"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="170">
                <TextBlock Text="ZAKRES CZASU" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbSRange" SelectedIndex="2">
                  <ComboBoxItem Content="Ostatnia godzina"/>
                  <ComboBoxItem Content="Ostatnie 24 godziny"/>
                  <ComboBoxItem Content="Ostatnie 7 dni"/>
                  <ComboBoxItem Content="Ostatnie 30 dni"/>
                  <ComboBoxItem Content="Wszystko"/>
                  <ComboBoxItem Content="Własny zakres"/>
                </ComboBox>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="OD (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSFrom" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="DO (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSTo" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="100">
                <TextBlock Text="MAKS. ZDARZEŃ" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSMax" Text="5000" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <StackPanel Style="{StaticResource InlineRow}">
                  <Button Name="btnSLoad" Content="Wczytaj zdarzenia" Style="{StaticResource BtnPrimary}"/>
                </StackPanel>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- FILTRY -->
          <Border Grid.Row="1" Style="{StaticResource FormCard}">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="170">
                <TextBlock Text="POZIOM" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbSLevel" SelectedIndex="0">
                  <ComboBoxItem Content="Wszystkie"/>
                  <ComboBoxItem Content="Krytyczne i błędy"/>
                  <ComboBoxItem Content="Ostrzeżenia"/>
                  <ComboBoxItem Content="Informacje"/>
                </ComboBox>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="130">
                <TextBlock Text="ID ZDARZENIA" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbSId"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="170">
                <TextBlock Text="ŹRÓDŁO" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbSSource"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="210">
                <TextBlock Text="KLIENT RADIUS (IP)" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbSClient"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="250">
                <TextBlock Text="SZUKAJ W TREŚCI" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSSearch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}">
                <StackPanel Style="{StaticResource InlineRow}">
                  <Button Name="btnSResetFilters" Content="Wyczyść filtry" Style="{StaticResource Btn}"/>
                </StackPanel>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- TABELA + SZCZEGÓŁY -->
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="5*" MinWidth="360"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="2*" MinWidth="300"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Style="{StaticResource PanelCard}">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Style="{StaticResource CardHeader}">
                  <Button Name="btnSExport" DockPanel.Dock="Right" Content="Eksport CSV" Style="{StaticResource BtnSmall}" Margin="12,0,0,0"
                          ToolTip="Zapisz zdarzenia widoczne w tabeli (po filtrach) do pliku CSV"/>
                  <TextBlock Text="Zdarzenia systemowe" Style="{StaticResource CardTitle}" DockPanel.Dock="Left" Margin="0,0,12,0"/>
                  <TextBlock Name="lblSCounts" Text="" Style="{StaticResource Counts}" HorizontalAlignment="Right"/>
                </DockPanel>
                <DataGrid Name="dgSys" Grid.Row="1">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Czas" Binding="{Binding TimeStr}" Width="155" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Poziom" Binding="{Binding Result}" Width="95" ElementStyle="{StaticResource ResultCell}"/>
                    <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="55"/>
                    <DataGridTextColumn Header="Źródło" Binding="{Binding Source}" Width="85"/>
                    <DataGridTextColumn Header="Dziennik" Binding="{Binding LogName}" Width="85"/>
                    <DataGridTextColumn Header="Klient RADIUS" Binding="{Binding ClientIp}" Width="125" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Switch" Binding="{Binding ClientName}" Width="130"/>
                    <DataGridTextColumn Header="Wiadomość" Binding="{Binding Short}" Width="*" MinWidth="200"/>
                  </DataGrid.Columns>
                </DataGrid>
              </Grid>
            </Border>

            <GridSplitter Grid.Column="1" Style="{StaticResource Splitter}"/>

            <Border Grid.Column="2" Style="{StaticResource PanelCard}">
              <Grid Name="sideSys">
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*" MinHeight="90"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Style="{StaticResource CardHeader}">
                  <Button Name="btnSCopy" DockPanel.Dock="Right" Content="Kopiuj" Style="{StaticResource BtnSmall}"/>
                  <TextBlock Text="Szczegóły" Style="{StaticResource CardTitle}"/>
                </DockPanel>
                <TextBox Name="txtSDetails" Grid.Row="1" Style="{StaticResource DetailsBox}" Text="Kliknij Wczytaj zdarzenia."/>
                <TextBlock Grid.Row="2" Text="Statystyki widoku" Style="{StaticResource CardTitle}" Margin="0,12,0,8"/>
                <Border Name="bdSStats" Grid.Row="3" Background="{StaticResource Field}" CornerRadius="6" Padding="8,6" MaxHeight="240">
                  <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBlock Name="txtSStats" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap" Text="-"/>
                  </ScrollViewer>
                </Border>
              </Grid>
            </Border>
          </Grid>
        </Grid>
      </TabItem>
    </TabControl>

    <!-- PASEK STATUSU -->
    <Grid Grid.Row="2" Margin="0,12,0,0" MinHeight="20">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="240"/>
      </Grid.ColumnDefinitions>
      <TextBlock Name="lblStatus" Text="Gotowy" Foreground="{StaticResource Muted}" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Margin="0,0,12,0"/>
      <ProgressBar Name="pb" Grid.Column="1" IsIndeterminate="False" VerticalAlignment="Center" Visibility="Hidden"/>
    </Grid>
  </Grid>
</Window>
'@

# --- Okno ----------------------------------------------------------------------------------
$reader = New-Object System.Xml.XmlNodeReader $xaml
$win    = [Windows.Markup.XamlReader]::Load($reader)
$ui     = @{}
$xaml.SelectNodes('//*[@Name]') | ForEach-Object { $ui[$_.Name] = $win.FindName($_.Name) }

# Okno nie może być większe niż ekran (np. laptop 1366x768 albo skalowanie 150%) - inaczej
# panel szczegółów i pasek statusu lądują poza ekranem.
function Set-WindowFit($Window) {
    $wa = [System.Windows.SystemParameters]::WorkArea
    if ($Window.Width -gt $wa.Width - 20) { $Window.Width = [Math]::Max(600, $wa.Width - 20) }
    if ($Window.Height -gt $wa.Height - 20) { $Window.Height = [Math]::Max(400, $wa.Height - 20) }
    if ($Window.MinWidth -gt $Window.Width) { $Window.MinWidth = $Window.Width }
    if ($Window.MinHeight -gt $Window.Height) { $Window.MinHeight = $Window.Height }
}
Set-WindowFit $win

# Statystyki widoku zajmują najwyżej ~40% wysokości panelu, a w niskim panelu są chowane -
# miejsce zostaje dla szczegółów zaznaczonego zdarzenia. W niskim oknie znika też podtytuł.
foreach ($side in 'sideEv', 'sideLog', 'sideSys') {
    $ui[$side].Add_SizeChanged({
            $vis = $(if ($this.ActualHeight -lt 300) { 'Collapsed' } else { 'Visible' })
            foreach ($c in $this.Children) {
                if ([System.Windows.Controls.Grid]::GetRow($c) -ge 2) { $c.Visibility = $vis }
                if ($c -is [System.Windows.Controls.Border]) { $c.MaxHeight = [Math]::Max(60, [Math]::Min(240, [Math]::Floor($this.ActualHeight * 0.4))) }
            }
        })
}
$win.Add_SizeChanged({ $ui.lblSubtitle.Visibility = $(if ($this.ActualHeight -lt 760) { 'Collapsed' } else { 'Visible' }) })

# --- Stan ----------------------------------------------------------------------------------
$script:Colors   = @{ OK = '#4ADE80'; Warn = '#FBBF24'; Error = '#F87171'; Info = '#8A90A2' }
$script:AllLabel = '(wszystkie)'
$script:DefaultLogPath = Join-Path $env:SystemRoot 'System32\LogFiles'

# Kontekst każdej zakładki: dane + kontrolki. Kontrolki dostają Tag = nazwa kontekstu,
# dzięki temu jedna procedura obsługi zdarzenia działa dla obu zakładek.
function New-Ctx([string]$Name, [hashtable]$Map) {
    $c = @{
        Name = $Name; Suspend = $false; Busy = $false; PS = $null; Handle = $null; LoadStarted = $null; Timer = $null
        All  = New-Object System.Collections.Generic.List[object]
        View = New-Object System.Collections.Generic.List[object]
    }
    foreach ($k in $Map.Keys) {
        $c[$k] = $ui[$Map[$k]]
        if ($c[$k] -is [System.Windows.FrameworkElement]) { $c[$k].Tag = $Name }
    }
    return $c
}

$script:Ctx = @{
    Ev  = New-Ctx 'Ev' @{
        Result = 'cbResult'; Switch = 'cbSwitch'; Policy = 'cbPolicy'; Auth = 'cbAuth'
        User = 'txtUser'; Computer = 'txtComputer'; Search = 'txtSearch'
        Grid = 'dgEvents'; Counts = 'lblCounts'; Stats = 'txtStats'; Details = 'txtDetails'
        Load = 'btnLoad'; Reset = 'btnResetFilters'; Export = 'btnExport'; Copy = 'btnCopy'
    }
    Log = New-Ctx 'Log' @{
        Result = 'cbLResult'; Switch = 'cbLSwitch'; Policy = 'cbLPolicy'; Auth = 'cbLAuth'
        User = 'txtLUser'; Computer = 'txtLComputer'; Search = 'txtLSearch'
        Grid = 'dgLog'; Counts = 'lblLCounts'; Stats = 'txtLStats'; Details = 'txtLDetails'
        Load = 'btnLogLoad'; Reset = 'btnLResetFilters'; Export = 'btnLExport'; Copy = 'btnLCopy'
        ShowReq = 'chkLogReq'; ShowAcct = 'chkLogAcct'
    }
    Sys = New-Ctx 'Sys' @{
        Result = 'cbSLevel'; Id = 'cbSId'; Source = 'cbSSource'; Client = 'cbSClient'; Search = 'txtSSearch'
        Grid = 'dgSys'; Counts = 'lblSCounts'; Stats = 'txtSStats'; Details = 'txtSDetails'
        Load = 'btnSLoad'; Reset = 'btnSResetFilters'; Export = 'btnSExport'; Copy = 'btnSCopy'
    }
}

# Podpowiedzi do zdarzeń systemowych NPS (dziennik System)
$script:SysHints = @{
    '13' = 'Żądanie od adresu, który nie jest skonfigurowany jako klient RADIUS w NPS. Dodaj klienta (switch/AP) albo sprawdź, z jakiego IP faktycznie wysyła (interfejs zarządzający, VLAN, NAT).'
    '14' = 'Access-Request bez atrybutu Message-Authenticator, a klient RADIUS go wymaga (ochrona przed BlastRADIUS). Włącz wysyłanie Message-Authenticator na urządzeniu / zaktualizuj firmware albo zdejmij wymaganie w ustawieniach klienta RADIUS w NPS.'
    '18' = 'Nieprawidłowy Message-Authenticator - prawie zawsze niezgodny wspólny klucz tajny (shared secret) między urządzeniem a klientem RADIUS w NPS. Wpisz ten sam sekret po obu stronach (uwaga na spacje / znaki specjalne) i sprawdź, czy pod tym IP nie ma innego urządzenia.'
    '4402' = 'Brak dostępnego kontrolera domeny dla wskazanej domeny - NPS nie może uwierzytelniać kont z AD. Sprawdź łączność z DC, DNS i czy serwer NPS jest w grupie RAS and IAS Servers.'
}

# Najczęstsze kody przyczyn NPS - krótkie podpowiedzi
$script:ReasonHints = @{
    '0'   = 'Sukces.'
    '8'   = 'Konto nie istnieje - sprawdź nazwę / format MAC.'
    '16'  = 'Złe poświadczenia - login lub hasło (przy MAB: hasło konta musi być = MAC).'
    '22'  = 'Nie można przetworzyć typu EAP - niezgodna metoda między klientem a zasadą.'
    '23'  = 'Błąd podczas EAP - najczęściej problem z certyfikatem serwera lub klienta.'
    '34'  = 'Konto wyłączone.'
    '36'  = 'Konto zablokowane.'
    '48'  = 'Żądanie nie pasuje do żadnej zasady sieciowej - sprawdź warunki (grupa, typ uwierzytelniania, typ portu).'
    '49'  = 'Żądanie nie pasuje do żadnej zasady żądań połączeń.'
    '65'  = 'Odmowa przez uprawnienia dostępu (zakładka Telefonowanie) albo trafienie w domyślną zasadę odmawiającą - sprawdź, czy właściwa zasada w ogóle pasuje.'
    '66'  = 'Metoda uwierzytelnienia nie jest dozwolona w ograniczeniach zasady sieciowej.'
    '265' = 'Łańcuch certyfikatu wystawiony przez niezaufany urząd.'
}

# --- Pomocnicze ----------------------------------------------------------------------------
function Get-Brush([string]$Hex) { (New-Object System.Windows.Media.BrushConverter).ConvertFromString($Hex) }

function Set-Status([string]$Text, [string]$Level = 'Info') {
    $ui.lblStatus.Text = $Text
    $ui.lblStatus.Foreground = Get-Brush $script:Colors[$Level]
}

function Set-State([string]$Text, [string]$Level) {
    $ui.lblState.Text = $Text
    $ui.dotState.Fill = Get-Brush $script:Colors[$Level]
}

function Show-Msg([string]$Text, [string]$Icon = 'Information') {
    [void][System.Windows.MessageBox]::Show($win, $Text, 'NPS Event Viewer', 'OK', $Icon)
}

function Get-HexMac([string]$Value) {
    if ($Value -match '^[0-9A-Fa-f:\.\-\s]+$') {
        $h = ($Value -replace '[^0-9A-Fa-f]', '').ToUpper()
        if ($h.Length -ge 4) { return $h }
    }
    return ''
}

function Test-Contains([string]$Haystack, [string]$Needle) {
    if (-not $Needle) { return $true }
    if (-not $Haystack) { return $false }
    return $Haystack.IndexOf($Needle, [StringComparison]::OrdinalIgnoreCase) -ge 0
}

# Ostatnia pozycja kombi = "Własny zakres"; pozycja "Wszystko" = bez ograniczenia czasu
function Get-TimeRange($Combo, $TxtFrom, $TxtTo) {
    $now = Get-Date
    $idx = $Combo.SelectedIndex
    if ($idx -eq $Combo.Items.Count - 1) {
        $fmt = 'yyyy-MM-dd HH:mm'
        $inv = [Globalization.CultureInfo]::InvariantCulture
        $s = [datetime]::ParseExact($TxtFrom.Text.Trim(), $fmt, $inv)
        $e = [datetime]::ParseExact($TxtTo.Text.Trim(), $fmt, $inv)
        if ($e -le $s) { throw 'Data "DO" musi być późniejsza niż "OD".' }
        return @($s, $e)
    }
    switch ($idx) {
        0 { return @($now.AddHours(-1), $now) }
        1 { return @($now.AddHours(-24), $now) }
        2 { return @($now.AddDays(-7), $now) }
        3 { return @($now.AddDays(-30), $now) }
    }
    return @([datetime]::MinValue, [datetime]::MaxValue)
}

function Format-Range($Range) {
    if ($Range[0] -eq [datetime]::MinValue) { return 'cały zakres' }
    return "$($Range[0].ToString('yyyy-MM-dd HH:mm')) - $($Range[1].ToString('yyyy-MM-dd HH:mm'))"
}

function Update-CustomRange($Combo, $TxtFrom, $TxtTo) {
    $custom = ($Combo.SelectedIndex -eq $Combo.Items.Count - 1)
    $TxtFrom.IsEnabled = $custom
    $TxtTo.IsEnabled   = $custom
    if ($custom -and -not $TxtFrom.Text) {
        $TxtFrom.Text = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd HH:mm')
        $TxtTo.Text   = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    }
}

function Get-MaxValue($TextBox, [int]$Default) {
    $max = 0
    if (-not [int]::TryParse($TextBox.Text.Trim(), [ref]$max) -or $max -lt 0) { $max = $Default; $TextBox.Text = "$Default" }
    return $max
}

# --- Wczytywanie w tle: dziennik Security --------------------------------------------------
$script:EvLoader = {
    # $DataValues (opcjonalne, okno historii): tylko zdarzenia, w których któreś pole EventData ma
    # dokładnie jedną z tych wartości (np. MAC w kilku zapisach) - filtr wykonuje sam dziennik.
    param($Computer, $Start, $End, $Max, [string[]]$DataValues)

    $ids = 6272, 6273, 6274, 6276, 6277, 6278, 6279, 6280
    $p = @{
        FilterHashtable = @{ LogName = 'Security'; Id = $ids; StartTime = $Start; EndTime = $End }
        ErrorAction     = 'Stop'
    }
    if ($Max -gt 0) { $p.MaxEvents = $Max }
    if ($Computer)  { $p.ComputerName = $Computer }
    if ($DataValues) {
        # Dla historii zapytanie XML zamiast FilterHashtable: FilterHashtable w Windows PowerShell 5.1
        # po cichu zwraca "brak zdarzeń", gdy zapytanie okaże się za złożone (limit ok. 20 wyrażeń
        # XPath), a XML zgłasza błąd. Zakres ID (2 wyrażenia zamiast 8) zostawia miejsce na warianty MAC.
        # Porównanie Data='...' w dzienniku rozróżnia wielkość liter - stąd warianty małymi i dużymi.
        # Warianty dzielimy na kilka <Select> (suma wyników), żeby każdy miał mniej niż 20 wyrażeń.
        # Czas zawsze w formacie niezależnym od ustawień regionalnych (inaczej np. kalendarz buddyjski
        # albo kropka jako separator godziny dają zapytanie bez wyników).
        $fmt = 'yyyy-MM-dd\THH:mm:ss.fff\Z'
        $inv = [Globalization.CultureInfo]::InvariantCulture
        $time = "TimeCreated[@SystemTime&gt;='$($Start.ToUniversalTime().ToString($fmt, $inv))' and @SystemTime&lt;='$($End.ToUniversalTime().ToString($fmt, $inv))']"
        $vals = @($DataValues | ForEach-Object { "Data='$($_ -replace "[^0-9A-Za-z\-:\.]", '')'" })
        $sel = ''
        for ($i = 0; $i -lt $vals.Count; $i += 8) {
            $part = $vals[$i..([Math]::Min($i + 7, $vals.Count - 1))] -join ' or '
            $sel += "<Select Path='Security'>*[System[(EventID&gt;=6272 and EventID&lt;=6280 and EventID!=6275) and $time] and EventData[$part]]</Select>"
        }
        $p.Remove('FilterHashtable')
        $p.FilterXml = [xml]"<QueryList><Query Id='0' Path='Security'>$sel</Query></QueryList>"
    }

    try { $events = @(Get-WinEvent @p) }
    catch {
        if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { return }
        throw
    }

    $names = @{
        6272 = 'Udzielono'; 6273 = 'Odmowa'; 6274 = 'Odrzucono'; 6276 = 'Kwarantanna'
        6277 = 'Warunkowo'; 6278 = 'Pełny dostęp'; 6279 = 'Zablokowano'; 6280 = 'Odblokowano'
    }

    foreach ($ev in $events) {
        $x = [xml]$ev.ToXml()
        $d = @{}
        foreach ($n in $x.Event.EventData.Data) {
            $t = [string]$n.InnerText
            $d[[string]$n.Name] = $(if ($t) { $t.Trim() } else { '-' })
        }

        $level = if ($ev.Id -in 6272, 6278, 6280) { 'OK' } elseif ($ev.Id -in 6273, 6279) { 'Error' } else { 'Warn' }

        $machine = $d['FullyQualifiedSubjectMachineName']
        if (-not $machine -or $machine -eq '-') { $machine = $d['SubjectMachineName'] }

        $user    = [string]$d['SubjectUserName']
        $calling = [string]$d['CallingStationID']
        # Przy VPN Calling-Station-Id to adres IP klienta - to nie MAC (192.168.100.200 też ma 12 cyfr)
        $macHex  = $(if ($calling -match '^\s*\d{1,3}(\.\d{1,3}){3}\s*$' -or $calling.Contains('::')) { '' } else { ($calling -replace '[^0-9A-Fa-f]', '').ToUpper() })

        $o = [pscustomobject]@{
            Time           = $ev.TimeCreated
            TimeStr        = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            Id             = $ev.Id
            Result         = $names[$ev.Id]
            Level          = $level
            User           = $user
            Domain         = $d['SubjectDomainName']
            UserFQ         = $d['FullyQualifiedSubjectUserName']
            Machine        = $machine
            CallingStation = $calling
            CalledStation  = $d['CalledStationID']
            NasIp          = $d['NASIPv4Address']
            NasId          = $d['NASIdentifier']
            NasPortType    = $d['NASPortType']
            NasPort        = $d['NASPort']
            Client         = $d['ClientName']
            ClientIp       = $d['ClientIPAddress']
            CRP            = $d['ProxyPolicyName']
            Policy         = $d['NetworkPolicyName']
            AuthProvider   = $d['AuthenticationProvider']
            AuthServer     = $d['AuthenticationServer']
            AuthType       = $d['AuthenticationType']
            EapType        = $d['EAPType']
            ReasonCode     = $d['ReasonCode']
            Reason         = $d['Reason']
            Server         = $ev.MachineName
            RecordId       = $ev.RecordId
            MacHex         = $macHex
            UserHex        = ($user -replace '[^0-9A-Fa-f]', '').ToUpper()
            SearchText     = ''
        }
        $o.SearchText = (($o.PSObject.Properties | Where-Object { $_.Name -ne 'SearchText' -and $_.Value }) |
                ForEach-Object { [string]$_.Value }) -join ' | '
        $o
    }
}

# --- Wczytywanie w tle: pliki .log NPS (DTS/XML + IAS) -------------------------------------
$script:LogLoader = {
    # $MatchMode/$MatchValue (opcjonalne, okno historii): 'Mac' + 12 cyfr hex, 'User' / 'Computer'
    # + nazwa po normalizacji - zostają tylko wpisy tego urządzenia / użytkownika / komputera.
    # $SkipReq: bez Access-Request / Challenge (okno historii, gdy ich nie pokazuje) - nie zajmują limitu.
    param([string[]]$Paths, [string]$Mask, [datetime]$Start, [datetime]$End, [int]$Max,
          [string]$MatchMode = '', [string]$MatchValue = '', [bool]$MatchPartial = $false, [bool]$SkipReq = $false)

    # --- słowniki ---
    $reasonNames = @{
        '0' = 'Sukces'; '1' = 'Błąd wewnętrzny'; '2' = 'Odmowa dostępu'; '3' = 'Nieprawidłowe żądanie'
        '4' = 'Wykaz globalny niedostępny'; '5' = 'Domena niedostępna'; '6' = 'Serwer niedostępny'
        '7' = 'Brak takiej domeny'; '8' = 'Brak takiego użytkownika'; '16' = 'Błąd uwierzytelnienia (złe poświadczenia)'
        '17' = 'Błąd zmiany hasła'; '18' = 'Nieobsługiwany typ uwierzytelniania'; '22' = 'Nie można przetworzyć typu EAP'
        '23' = 'Nieoczekiwany błąd podczas EAP'; '32' = 'Tylko użytkownicy lokalni'; '33' = 'Wymagana zmiana hasła'
        '34' = 'Konto wyłączone'; '35' = 'Konto wygasło'; '36' = 'Konto zablokowane'; '37' = 'Niedozwolone godziny logowania'
        '38' = 'Ograniczenie konta'; '48' = 'Brak pasującej zasady sieciowej'; '49' = 'Brak pasującej zasady żądań połączeń'
        '64' = 'Blokada dostępu zdalnego'; '65' = 'Odmowa - uprawnienia dostępu / zasada odmawiająca'
        '66' = 'Niedozwolona metoda uwierzytelniania'; '67' = 'Niedozwolony Calling-Station-Id'
        '68' = 'Niedozwolone godziny dostępu'; '69' = 'Niedozwolony Called-Station-Id'; '70' = 'Niedozwolony typ portu'
        '265' = 'Certyfikat z niezaufanego urzędu'
    }
    $authNames = @{ '1' = 'PAP'; '2' = 'CHAP'; '3' = 'MS-CHAP'; '4' = 'MS-CHAPv2'; '5' = 'EAP'; '7' = 'Brak'; '8' = 'Custom'; '9' = 'MS-CHAP CPW'; '10' = 'MS-CHAPv2 CPW'; '11' = 'PEAP' }
    $portTypes = @{ '0' = 'Async'; '1' = 'Sync'; '2' = 'ISDN Sync'; '5' = 'Virtual'; '15' = 'Ethernet'; '17' = 'Cable'; '18' = 'Wireless-Other'; '19' = 'Wireless-802.11' }
    $termCauses = @{
        '1' = 'User-Request (wylogowanie / roaming)'; '2' = 'Lost-Carrier (odłączony kabel / poza zasięgiem)'; '3' = 'Lost-Service'
        '4' = 'Idle-Timeout (bezczynność)'; '5' = 'Session-Timeout (koniec czasu sesji)'; '6' = 'Admin-Reset'; '7' = 'Admin-Reboot'
        '8' = 'Port-Error'; '9' = 'NAS-Error'; '10' = 'NAS-Request'; '11' = 'NAS-Reboot'; '16' = 'Callback'
        '19' = 'Supplicant-Restart'; '20' = 'Reauthentication-Failure (nieudana reautoryzacja)'; '21' = 'Port-Reinit'; '22' = 'Port-Disabled'
    }
    $providers = @{ '0' = 'Brak'; '1' = 'Windows'; '2' = 'Zdalny serwer RADIUS' }
    $acctTypes = @{ '1' = 'Start'; '2' = 'Stop'; '3' = 'Interim'; '7' = 'Acct-On'; '8' = 'Acct-Off' }

    # Format IAS: numery atrybutów (best effort)
    $iasMap = @{
        '1' = 'User-Name'; '4' = 'NAS-IP-Address'; '5' = 'NAS-Port'; '6' = 'Service-Type'; '30' = 'Called-Station-Id'
        '31' = 'Calling-Station-Id'; '32' = 'NAS-Identifier'; '40' = 'Acct-Status-Type'; '44' = 'Acct-Session-Id'
        '61' = 'NAS-Port-Type'; '4108' = 'Client-IP-Address'; '4116' = 'Client-Vendor'; '4127' = 'Authentication-Type'
        '4128' = 'Client-Friendly-Name'; '4129' = 'SAM-Account-Name'; '4130' = 'Fully-Qualifed-User-Name'
        '4132' = 'EAP-Friendly-Name'; '4136' = 'Packet-Type'; '4142' = 'Reason-Code'; '4149' = 'NP-Policy-Name'
        '4154' = 'Proxy-Policy-Name'; '4155' = 'Provider-Type'
        '8' = 'Framed-IP-Address'; '25' = 'Class'; '41' = 'Acct-Delay-Time'; '46' = 'Acct-Session-Time'
        '49' = 'Acct-Terminate-Cause'; '87' = 'NAS-Port-Id'
    }
    $mergeKeys = 'User-Name', 'Calling-Station-Id', 'Called-Station-Id', 'NAS-Port', 'NAS-Port-Type', 'NAS-Identifier', 'NAS-IP-Address', 'NAS-Port-Id'

    $rxAttr = New-Object System.Text.RegularExpressions.Regex '<([A-Za-z0-9\-]+)(?:\s[^>]*)?>([^<]*)</\1>', 'Compiled'
    $rxTs   = New-Object System.Text.RegularExpressions.Regex '<Timestamp[^>]*>([^<]+)</Timestamp>', 'Compiled'
    $inv    = [Globalization.CultureInfo]::InvariantCulture
    $fmts   = [string[]]@('MM/dd/yyyy HH:mm:ss.fff', 'MM/dd/yyyy HH:mm:ss', 'M/d/yyyy H:mm:ss.fff', 'M/d/yyyy H:mm:ss', 'yyyy-MM-dd HH:mm:ss.fff', 'yyyy-MM-dd HH:mm:ss')

    function ConvertTo-Time([string]$s) {
        $t = [datetime]::MinValue
        $s = $s.Trim()
        if ([datetime]::TryParseExact($s, $fmts, $inv, [Globalization.DateTimeStyles]::None, [ref]$t)) { return $t }
        if ([datetime]::TryParse($s, $inv, [Globalization.DateTimeStyles]::None, [ref]$t)) { return $t }
        return $null
    }

    # Normalizacja jak Get-NormUser / Get-NormComputer w oknie historii (ten runspace ich nie widzi).
    function Get-IdNorm([string]$v, [bool]$AsComputer) {
        $v = $v.Trim().ToLowerInvariant()
        if ($AsComputer -and $v.StartsWith('host/')) { $v = $v.Substring(5) }
        $i = $v.LastIndexOf('\'); if ($i -ge 0) { $v = $v.Substring($i + 1) }
        if ($AsComputer) {
            $v = $v.TrimEnd('$')
            $j = $v.IndexOf('.'); if ($j -gt 0) { $v = $v.Substring(0, $j) }
        }
        elseif (-not $v.StartsWith('host/') -and $v -match '^([^@]+)@') { $v = $Matches[1] }
        return $v
    }
    function Test-IdMatch($d) {
        if ($MatchMode -eq 'Mac') {
            foreach ($k in 'Calling-Station-Id', 'User-Name') {
                $v = [string]$d[$k]
                if (-not $v) { continue }
                if ($k -eq 'User-Name' -and $v -notmatch '^[0-9A-Fa-f:\.\-\s]+$') { continue }
                if ($v -match '^\s*\d{1,3}(\.\d{1,3}){3}\s*$' -or $v.Contains('::')) { continue }   # VPN: adres IP, nie MAC
                $h = ($v -replace '[^0-9A-Fa-f]', '').ToUpper()
                if ($k -eq 'User-Name' -and $h.Length -ne 12) { continue }
                if ($MatchPartial) { if ($h -and $h.Contains($MatchValue)) { return $true } }
                elseif ($h -eq $MatchValue) { return $true }
            }
            return $false
        }
        foreach ($k in 'User-Name', 'SAM-Account-Name', 'Fully-Qualifed-User-Name', 'Fully-Qualified-User-Name') {
            $v = [string]$d[$k]
            if (-not $v) { continue }
            $comp = ($MatchMode -eq 'Computer')
            if ($comp -and $v -notmatch '^host/|\$$') { continue }
            $n = Get-IdNorm $v $comp
            if (-not $n) { continue }
            if ($MatchPartial) { if ($n.Contains($MatchValue)) { return $true } }
            elseif ($n -eq $MatchValue) { return $true }
        }
        return $false
    }

    # --- pliki ---
    $files = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    $seen  = @{}
    foreach ($raw in $Paths) {
        $p = [Environment]::ExpandEnvironmentVariables($raw.Trim().Trim('"'))
        if (-not $p) { continue }
        $found = @()
        if (Test-Path -LiteralPath $p -PathType Container) { $found = @(Get-ChildItem -LiteralPath $p -Filter $Mask -File -ErrorAction Stop) }
        elseif (Test-Path -LiteralPath $p -PathType Leaf) { $found = @(Get-Item -LiteralPath $p -ErrorAction Stop) }
        elseif ($p -match '[\*\?]') { $found = @(Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue) }
        else { throw "Nie znaleziono lokalizacji: $p" }
        foreach ($f in $found) { if (-not $seen.ContainsKey($f.FullName)) { $seen[$f.FullName] = 1; $files.Add($f) } }
    }
    if ($files.Count -eq 0) { throw "Brak plików pasujących do maski '$Mask' we wskazanej lokalizacji." }

    # pomijamy pliki, do których nic nie dopisano od początku zakresu
    $files = @($files | Where-Object { $_.LastWriteTime -ge $Start } | Sort-Object LastWriteTime)
    $preStart = $(if ($Start -gt [datetime]::MinValue.AddMinutes(5)) { $Start.AddMinutes(-2) } else { $Start })

    $queue   = New-Object 'System.Collections.Generic.Queue[object]'
    $lastReq = @{}
    # Żądanie i odpowiedź (Accept / Reject / Challenge) mają ten sam atrybut Class - to pewny klucz.
    # "Ostatnie żądanie od tego samego klienta RADIUS" zostaje jako zapas dla wpisów bez Class:
    # przy WLC / switchu uwierzytelniającym wielu klientów naraz przypisywało MAC innego urządzenia.
    $reqByClass = @{}
    $stat    = @{ Lines = 0; Parsed = 0; Skipped = 0; Unsupported = 0; Dropped = 0 }
    $errors  = New-Object System.Collections.Generic.List[string]

    foreach ($fi in $files) {
        $sr = $null
        try {
            $fs = New-Object System.IO.FileStream($fi.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]'ReadWrite, Delete')
            $sr = New-Object System.IO.StreamReader($fs, [Text.Encoding]::UTF8, $true)
            $lineNo = 0
            $buf = ''
            while ($null -ne ($line = $sr.ReadLine())) {
                $lineNo++
                $stat.Lines++
                if (-not $line.Trim()) { continue }

                $chunks = New-Object System.Collections.Generic.List[string]
                if ($buf -or $line.TrimStart().StartsWith('<')) {
                    # --- DTS / XML: zdarzenie może (teoretycznie) być rozbite na kilka linii
                    $buf += $line
                    while (($i = $buf.IndexOf('</Event>')) -ge 0) {
                        $chunks.Add($buf.Substring(0, $i + 8))
                        $buf = $buf.Substring($i + 8)
                    }
                    if ($buf.Length -gt 200000) { $buf = ''; $stat.Skipped++ }
                    foreach ($ch in $chunks) {
                        $m = $rxTs.Match($ch)
                        if (-not $m.Success) { $stat.Skipped++; continue }
                        $time = ConvertTo-Time $m.Groups[1].Value
                        if ($null -eq $time) { $stat.Skipped++; continue }
                        if ($time -lt $preStart -or $time -gt $End) { continue }

                        $d = @{}
                        foreach ($a in $rxAttr.Matches($ch)) {
                            $v = [System.Net.WebUtility]::HtmlDecode($a.Groups[2].Value).Trim()
                            if ($v) { $d[$a.Groups[1].Value] = $v }
                        }
                        $d['#Time'] = $time
                        $d['#Line'] = $lineNo
                        $d['#File'] = $fi
                        $d['#Raw']  = $ch
                        $stat.Parsed++
                        $pt  = [string]$d['Packet-Type']
                        $key = $d['Client-IP-Address']; if (-not $key) { $key = $d['Client-Friendly-Name'] }
                        $cls = [string]$d['Class']
                        if ($pt -eq '1') {
                            if ($key) { $lastReq[$key] = $d }
                            if ($cls) { if ($reqByClass.Count -gt 50000) { $reqByClass.Clear() }; $reqByClass[$cls] = $d }
                        }
                        elseif ($pt -eq '2' -or $pt -eq '3' -or $pt -eq '11') {
                            $rq = $null
                            if ($cls -and $reqByClass.ContainsKey($cls)) { $rq = $reqByClass[$cls]; $reqByClass.Remove($cls) }
                            elseif ($key -and $lastReq.ContainsKey($key)) { $rq = $lastReq[$key] }
                            if ($rq) {
                                foreach ($k in $mergeKeys) { if (-not $d[$k] -and $rq[$k]) { $d[$k] = $rq[$k]; $d['#Merged'] = $true } }
                            }
                        }
                        if ($time -lt $Start) { continue }
                        if ($SkipReq -and ($pt -eq '11' -or ($pt -eq '1' -and -not ($d['Reason-Code'] -and $d['Reason-Code'] -ne '0')))) { continue }
                        if ($MatchMode -and -not (Test-IdMatch $d)) { continue }
                        $queue.Enqueue($d)
                        if ($Max -gt 0 -and $queue.Count -gt $Max) { [void]$queue.Dequeue(); $stat.Dropped++ }
                    }
                    continue
                }

                if ($line.StartsWith('"')) { $stat.Unsupported++; continue }   # format "zgodny z bazą danych" (ODBC)

                # --- IAS: NAS-IP, User, Data, Czas, Usługa, Komputer, [nr atrybutu, wartość]...
                $f = $line.Split(',')
                if ($f.Count -lt 8) { $stat.Unsupported++; continue }
                $time = ConvertTo-Time "$($f[2]) $($f[3])"
                if ($null -eq $time) { $stat.Unsupported++; continue }
                if ($time -lt $preStart -or $time -gt $End) { continue }
                $d = @{ 'NAS-IP-Address' = $f[0].Trim(); 'User-Name' = $f[1].Trim(); 'Computer-Name' = $f[5].Trim() }
                for ($i = 6; $i + 1 -lt $f.Count; $i += 2) {
                    $id = $f[$i].Trim(); $v = $f[$i + 1].Trim().Trim('"')
                    if (-not $v) { continue }
                    $n = $iasMap[$id]; if (-not $n) { $n = "Attr-$id" }
                    $d[$n] = $v
                }
                $d['#Time'] = $time; $d['#Line'] = $lineNo; $d['#File'] = $fi; $d['#Raw'] = $line
                $stat.Parsed++
                $pt  = [string]$d['Packet-Type']
                $key = $d['Client-IP-Address']; if (-not $key) { $key = $d['Client-Friendly-Name'] }
                $cls = [string]$d['Class']
                if ($pt -eq '1') {
                    if ($key) { $lastReq[$key] = $d }
                    if ($cls) { if ($reqByClass.Count -gt 50000) { $reqByClass.Clear() }; $reqByClass[$cls] = $d }
                }
                elseif ($pt -eq '2' -or $pt -eq '3' -or $pt -eq '11') {
                    $rq = $null
                    if ($cls -and $reqByClass.ContainsKey($cls)) { $rq = $reqByClass[$cls]; $reqByClass.Remove($cls) }
                    elseif ($key -and $lastReq.ContainsKey($key)) { $rq = $lastReq[$key] }
                    if ($rq) {
                        foreach ($k in $mergeKeys) { if (-not $d[$k] -and $rq[$k]) { $d[$k] = $rq[$k]; $d['#Merged'] = $true } }
                    }
                }
                if ($time -lt $Start) { continue }
                if ($SkipReq -and ($pt -eq '11' -or ($pt -eq '1' -and -not ($d['Reason-Code'] -and $d['Reason-Code'] -ne '0')))) { continue }
                if ($MatchMode -and -not (Test-IdMatch $d)) { continue }
                $queue.Enqueue($d)
                if ($Max -gt 0 -and $queue.Count -gt $Max) { [void]$queue.Dequeue(); $stat.Dropped++ }
            }
        }
        catch { $errors.Add("$($fi.Name): $($_.Exception.Message)") }
        finally { if ($sr) { $sr.Dispose() } }
    }

    # --- budowa obiektów (najnowsze na górze) ---
    $arr = $queue.ToArray()
    [array]::Reverse($arr)
    foreach ($d in $arr) {
        $pt = [string]$d['Packet-Type']
        $rc = [string]$d['Reason-Code']
        switch ($pt) {
            '1' {
                if ($rc -and $rc -ne '0') { $res = 'Odrzucono'; $lvl = 'Warn'; $kind = 'Resp' }
                else { $res = 'Żądanie'; $lvl = 'Info'; $kind = 'Req' }
            }
            '2'  { $res = 'Udzielono'; $lvl = 'OK'; $kind = 'Resp' }
            '3'  { $res = 'Odmowa'; $lvl = 'Error'; $kind = 'Resp' }
            '11' { $res = 'Challenge'; $lvl = 'Info'; $kind = 'Req' }
            '4'  {
                $as = [string]$d['Acct-Status-Type']
                $res = 'Acct ' + $(if ($acctTypes[$as]) { $acctTypes[$as] } else { $as }); $lvl = 'Info'; $kind = 'Acct'
            }
            '5'  { $res = 'Acct-odp.'; $lvl = 'Info'; $kind = 'Acct' }
            default { $res = $(if ($pt) { "Pakiet $pt" } else { '?' }); $lvl = 'Warn'; $kind = 'Resp' }
        }

        $user = [string]$d['User-Name']
        $sam  = [string]$d['SAM-Account-Name']
        $dom  = ''
        if ($sam -match '^(.+?)\\') { $dom = $Matches[1] }
        $fq = $d['Fully-Qualifed-User-Name']; if (-not $fq) { $fq = $d['Fully-Qualified-User-Name'] }; if (-not $fq) { $fq = $sam }

        $at = [string]$d['Authentication-Type'];  if ($authNames[$at]) { $at = $authNames[$at] }
        $npt = [string]$d['NAS-Port-Type'];       if ($portTypes[$npt]) { $npt = $portTypes[$npt] }
        $prov = [string]$d['Provider-Type'];      if ($providers[$prov]) { $prov = $providers[$prov] }
        $reason = ''
        if ($rc) { $reason = $(if ($reasonNames[$rc]) { $reasonNames[$rc] } else { "Kod $rc" }) }
        $calling = [string]$d['Calling-Station-Id']
        $macHex  = $(if ($calling -match '^\s*\d{1,3}(\.\d{1,3}){3}\s*$' -or $calling.Contains('::')) { '' } else { ($calling -replace '[^0-9A-Fa-f]', '').ToUpper() })
        $fi = $d['#File']

        $o = [pscustomobject]@{
            Time           = $d['#Time']
            TimeStr        = $d['#Time'].ToString('yyyy-MM-dd HH:mm:ss')
            Id             = $pt
            Result         = $res
            Level          = $lvl
            Kind           = $kind
            User           = $user
            Domain         = $dom
            Sam            = $sam
            UserFQ         = $fq
            Machine        = $d['Computer-Name']
            CallingStation = $calling
            CalledStation  = $d['Called-Station-Id']
            NasIp          = $d['NAS-IP-Address']
            NasId          = $d['NAS-Identifier']
            NasPortType    = $npt
            NasPort        = $d['NAS-Port']
            Client         = $d['Client-Friendly-Name']
            ClientIp       = $d['Client-IP-Address']
            CRP            = $d['Proxy-Policy-Name']
            Policy         = $d['NP-Policy-Name']
            AuthProvider   = $prov
            AuthServer     = ''
            AuthType       = $at
            EapType        = $d['EAP-Friendly-Name']
            ReasonCode     = $rc
            Reason         = $reason
            Server         = $d['Computer-Name']
            RecordId       = $d['#Line']
            File           = $fi.Name
            FilePath       = $fi.FullName
            Merged         = [bool]$d['#Merged']
            FramedIp       = $d['Framed-IP-Address']
            SessionTime    = $d['Acct-Session-Time']
            NasPortId      = $d['NAS-Port-Id']
            AcctTerminate  = $(if ($d['Acct-Terminate-Cause']) { $tc = [string]$d['Acct-Terminate-Cause']; if ($termCauses[$tc]) { $termCauses[$tc] } else { "kod $tc" } } else { '' })
            MacHex         = $macHex
            UserHex        = ($user -replace '[^0-9A-Fa-f]', '').ToUpper()
            Attrs          = $d
            SearchText     = ''
        }
        $o.SearchText = (($o.PSObject.Properties | Where-Object { $_.Name -notin 'SearchText', 'Attrs', 'Time' -and $_.Value }) |
                ForEach-Object { [string]$_.Value }) -join ' | '
        $o
    }

    [pscustomobject]@{
        IsSummary   = $true
        Files       = $files.Count
        FileNames   = (@($files | ForEach-Object Name) -join ', ')
        Lines       = $stat.Lines
        Parsed      = $stat.Parsed
        Skipped     = $stat.Skipped
        Unsupported = $stat.Unsupported
        Dropped     = $stat.Dropped
        Errors      = @($errors)
    }
}

# --- Wczytywanie w tle: zdarzenia systemowe NPS (System / Application) --------------------
$script:SysLoader = {
    param($Computer, [datetime]$Start, [datetime]$End, [int]$Max, [string[]]$Logs, [string[]]$Providers)

    $prov = ($Providers | ForEach-Object { "@Name='$($_ -replace "'", '')'" }) -join ' or '
    $cond = "Provider[$prov]"
    if ($Start -gt [datetime]::MinValue) {
        $fmt = 'yyyy-MM-dd\THH:mm:ss.fff\Z'
        $inv = [Globalization.CultureInfo]::InvariantCulture
        $cond += " and TimeCreated[@SystemTime>='$($Start.ToUniversalTime().ToString($fmt, $inv))' and @SystemTime<='$($End.ToUniversalTime().ToString($fmt, $inv))']"
    }
    $xp = "*[System[$cond]]"

    $levels = @{ 1 = 'Krytyczny'; 2 = 'Błąd'; 3 = 'Ostrzeżenie'; 4 = 'Informacja'; 0 = 'Informacja'; 5 = 'Szczegółowy' }
    $rxIp   = [regex]'\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b'
    $all    = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[string]

    foreach ($log in $Logs) {
        $p = @{ LogName = $log; FilterXPath = $xp; ErrorAction = 'Stop' }
        if ($Max -gt 0) { $p.MaxEvents = $Max }
        if ($Computer)  { $p.ComputerName = $Computer }
        try { foreach ($ev in @(Get-WinEvent @p)) { $all.Add($ev) } }
        catch {
            if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { continue }
            $errors.Add("${log}: $($_.Exception.Message)")
        }
    }

    $sorted = @($all | Sort-Object TimeCreated -Descending)
    if ($Max -gt 0 -and $sorted.Count -gt $Max) { $sorted = $sorted[0..($Max - 1)] }

    foreach ($ev in $sorted) {
        $props = @($ev.Properties | ForEach-Object { [string]$_.Value })
        $msg = $null
        try { $msg = $ev.FormatDescription() } catch { }
        if (-not $msg) { $msg = $ev.Message }
        if (-not $msg) { $msg = "(brak opisu - dane: $($props -join ' | '))" }
        $short = (($msg -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -First 1)
        $ipM = $rxIp.Match($msg)
        if (-not $ipM.Success) { $ipM = $rxIp.Match(($props -join ' ')) }
        $lvlNo = [int]$ev.Level
        $lvl = $(if ($lvlNo -in 1, 2) { 'Error' } elseif ($lvlNo -eq 3) { 'Warn' } else { 'Info' })

        $o = [pscustomobject]@{
            Time       = $ev.TimeCreated
            TimeStr    = $ev.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss')
            Id         = [string]$ev.Id
            Result     = $(if ($levels.ContainsKey($lvlNo)) { $levels[$lvlNo] } else { "Poziom $lvlNo" })
            Level      = $lvl
            Source     = $ev.ProviderName
            LogName    = $ev.LogName
            ClientIp   = $(if ($ipM.Success) { $ipM.Value } else { '' })
            ClientName = ''
            Short      = $short
            Message    = $msg
            Props      = $props
            Server     = $ev.MachineName
            RecordId   = $ev.RecordId
            SearchText = ''
        }
        $o.SearchText = "$($o.Id) | $($o.Result) | $($o.Source) | $($o.ClientIp) | $msg | $($props -join ' | ')"
        $o
    }

    if ($errors.Count) {
        [pscustomobject]@{ IsSummary = $true; Errors = @($errors) }
    }
}

# --- Obsługa wczytywania -------------------------------------------------------------------
function Set-Busy($C, [bool]$Busy) {
    $C.Busy = $Busy
    $C.Load.IsEnabled = -not $Busy
    $any = @($script:Ctx.Values | Where-Object { $_.Busy }).Count -gt 0
    $ui.pb.Visibility      = $(if ($any) { 'Visible' } else { 'Hidden' })
    $ui.pb.IsIndeterminate = $any
}

function Start-EvLoad {
    $C = $script:Ctx.Ev
    if ($C.Busy) { return }
    try { $range = Get-TimeRange $ui.cbRange $ui.txtFrom $ui.txtTo }
    catch { Show-Msg "Nieprawidłowy zakres czasu.`n$($_.Exception.Message)" 'Warning'; return }

    $max    = Get-MaxValue $ui.txtMax 5000
    $server = $ui.txtServer.Text.Trim()

    Set-Busy $C $true
    $target = $(if ($server) { $server } else { $env:COMPUTERNAME })
    Set-Status "Wczytywanie zdarzeń z $target ($(Format-Range $range))..."
    Set-State 'Wczytywanie...' 'Warn'

    $C.LoadStarted = Get-Date
    $C.PS = [powershell]::Create()
    [void]$C.PS.AddScript($script:EvLoader).AddArgument($server).AddArgument($range[0]).AddArgument($range[1]).AddArgument($max)
    $C.Handle = $C.PS.BeginInvoke()
    $C.Timer.Start()
}

function Start-LogLoad {
    $C = $script:Ctx.Log
    if ($C.Busy) { return }
    try { $range = Get-TimeRange $ui.cbLogRange $ui.txtLogFrom $ui.txtLogTo }
    catch { Show-Msg "Nieprawidłowy zakres czasu.`n$($_.Exception.Message)" 'Warning'; return }

    $paths = @($ui.txtLogPath.Text -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not $paths.Count) { Show-Msg 'Podaj lokalizację plików .log (folder, plik lub ścieżka UNC).' 'Warning'; return }
    $mask = $ui.txtLogMask.Text.Trim()
    if (-not $mask) { $mask = 'IN*.log'; $ui.txtLogMask.Text = $mask }
    $max = Get-MaxValue $ui.txtLogMax 20000

    Set-Busy $C $true
    Set-Status "Wczytywanie logów z $($paths -join '; ') ($(Format-Range $range))..."
    Set-State 'Wczytywanie logów...' 'Warn'

    $C.LoadStarted = Get-Date
    $C.PS = [powershell]::Create()
    [void]$C.PS.AddScript($script:LogLoader).AddArgument([string[]]$paths).AddArgument($mask).AddArgument($range[0]).AddArgument($range[1]).AddArgument($max)
    $C.Handle = $C.PS.BeginInvoke()
    $C.Timer.Start()
}

function Start-SysLoad {
    $C = $script:Ctx.Sys
    if ($C.Busy) { return }
    try { $range = Get-TimeRange $ui.cbSRange $ui.txtSFrom $ui.txtSTo }
    catch { Show-Msg "Nieprawidłowy zakres czasu.`n$($_.Exception.Message)" 'Warning'; return }

    $logs = @()
    if ($ui.chkSSystem.IsChecked) { $logs += 'System' }
    if ($ui.chkSApp.IsChecked)    { $logs += 'Application' }
    if (-not $logs.Count) { Show-Msg 'Zaznacz co najmniej jeden dziennik.' 'Warning'; return }
    $prov = @($ui.txtSProviders.Text -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not $prov.Count) { $prov = @('NPS', 'IAS', 'Microsoft-Windows-NPS'); $ui.txtSProviders.Text = $prov -join '; ' }
    $max    = Get-MaxValue $ui.txtSMax 5000
    $server = $ui.txtSServer.Text.Trim()

    Set-Busy $C $true
    $target = $(if ($server) { $server } else { $env:COMPUTERNAME })
    Set-Status "Wczytywanie zdarzeń systemowych NPS z $target ($(Format-Range $range))..."
    Set-State 'Wczytywanie...' 'Warn'

    $C.LoadStarted = Get-Date
    $C.PS = [powershell]::Create()
    [void]$C.PS.AddScript($script:SysLoader).AddArgument($server).AddArgument($range[0]).AddArgument($range[1]).AddArgument($max).AddArgument([string[]]$logs).AddArgument([string[]]$prov)
    $C.Handle = $C.PS.BeginInvoke()
    $C.Timer.Start()
}

function Complete-Load($C) {
    if (-not $C.Handle -or -not $C.Handle.IsCompleted) { return }
    $C.Timer.Stop()
    try {
        $result  = $C.PS.EndInvoke($C.Handle)
        $list    = New-Object System.Collections.Generic.List[object]
        $summary = $null
        foreach ($r in $result) {
            if ($r.PSObject.Properties['IsSummary']) { $summary = $r } else { $list.Add($r) }
        }
        $C.All = $list
        if ($C.Name -eq 'Sys') { Resolve-SysClientNames }

        Update-FilterCombos $C
        Invoke-Filter $C

        $secs = [Math]::Round(((Get-Date) - $C.LoadStarted).TotalSeconds, 1)
        if ($C.Name -eq 'Ev') {
            $srv = $(if ($ui.txtServer.Text.Trim()) { $ui.txtServer.Text.Trim() } else { $env:COMPUTERNAME })
            if ($list.Count -eq 0) {
                Set-Status "Brak zdarzeń NPS w wybranym zakresie na $srv. Jeśli to niespodziewane: auditpol /get /subcategory:`"Network Policy Server`"" 'Warn'
                Set-State "$srv - brak zdarzeń" 'Warn'
            }
            else {
                $max  = [int]$ui.txtMax.Text
                $note = $(if ($max -gt 0 -and $list.Count -ge $max) { ' (osiągnięto limit - zwiększ MAKS. ZDARZEŃ lub zawęź zakres)' } else { '' })
                Set-Status "Wczytano $($list.Count) zdarzeń z $srv w $secs s. Ostatnie odświeżenie: $((Get-Date).ToString('HH:mm:ss'))$note" 'OK'
                Set-State "$srv - $($list.Count) zdarzeń" 'OK'
            }
        }
        elseif ($C.Name -eq 'Sys') {
            $srv  = $(if ($ui.txtSServer.Text.Trim()) { $ui.txtSServer.Text.Trim() } else { $env:COMPUTERNAME })
            $errs = $(if ($summary -and $summary.Errors.Count) { " | BŁĘDY: $($summary.Errors -join '; ')" } else { '' })
            if ($list.Count -eq 0) {
                Set-Status "Brak zdarzeń NPS w dziennikach systemowych na $srv w wybranym zakresie$errs" $(if ($errs) { 'Error' } else { 'Warn' })
                Set-State "$srv - brak zdarzeń systemowych" 'Warn'
            }
            else {
                $err = @($list | Where-Object Level -eq 'Error').Count
                Set-Status "Wczytano $($list.Count) zdarzeń systemowych NPS z $srv w $secs s (błędów: $err)$errs" $(if ($errs) { 'Warn' } else { 'OK' })
                Set-State "$srv - $($list.Count) zdarzeń systemowych" $(if ($err) { 'Warn' } else { 'OK' })
            }
        }
        else {
            $info = ''
            if ($summary) {
                $info = " | plików: $($summary.Files), linii: $($summary.Lines)"
                if ($summary.Skipped) { $info += ", pominięto: $($summary.Skipped)" }
                if ($summary.Errors.Count) { $info += " | BŁĘDY: $($summary.Errors -join '; ')" }
            }
            if ($summary -and $summary.Files -eq 0) {
                Set-Status "Żaden plik nie był modyfikowany w wybranym zakresie czasu - wybierz dłuższy zakres lub 'Wszystko'." 'Warn'
                Set-State 'Logi - brak plików' 'Warn'
            }
            elseif ($summary -and $summary.Parsed -eq 0 -and $summary.Unsupported -gt 0) {
                Set-Status "Nie rozpoznano formatu ($($summary.Unsupported) linii). Obsługiwane: DTS-compliant (XML) i IAS. Zmień w NPS: Accounting -> Log File Properties.$info" 'Error'
                Set-State 'Logi - nieznany format' 'Error'
            }
            elseif ($list.Count -eq 0) {
                Set-Status "Brak wpisów w wybranym zakresie czasu$info" 'Warn'
                Set-State 'Logi - brak wpisów' 'Warn'
            }
            else {
                $max  = [int]$ui.txtLogMax.Text
                $note = $(if ($max -gt 0 -and $list.Count -ge $max) { ' (osiągnięto limit - pokazano najnowsze)' } else { '' })
                $lvl  = $(if ($summary -and $summary.Errors.Count) { 'Warn' } else { 'OK' })
                Set-Status "Wczytano $($list.Count) wpisów w $secs s$note$info" $lvl
                Set-State "Logi - $($list.Count) wpisów" 'OK'
            }
        }
    }
    catch {
        $msg = $_.Exception.Message
        if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
        if ($msg -match 'unauthorized|Odmowa dostępu|Access is denied|nieautoryzowan') {
            $msg += $(if ($C.Name -eq 'Ev') { ' - uruchom jako administrator lub dodaj konto do grupy Event Log Readers.' } else { ' - brak uprawnień do folderu z logami (przy UNC: prawa do udziału c$ lub osobny udział).' })
        }
        Set-Status "Błąd wczytywania: $msg" 'Error'
        Set-State 'Błąd' 'Error'
    }
    finally {
        if ($C.PS) { $C.PS.Dispose(); $C.PS = $null }
        $C.Handle = $null
        Set-Busy $C $false
    }
}

foreach ($n in 'Ev', 'Log', 'Sys') {
    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(250)
    $t.Tag = $n
    $t.Add_Tick({ Complete-Load $script:Ctx[$this.Tag] })
    $script:Ctx[$n].Timer = $t
}

$script:AutoTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:AutoTimer.Interval = [TimeSpan]::FromSeconds(60)
$script:AutoTimer.Add_Tick({ if ($ui.chkAuto.IsChecked -and -not $script:Ctx.Ev.Busy) { Start-EvLoad } })

# --- Filtrowanie ---------------------------------------------------------------------------
function Set-ComboItems($Combo, $Values) {
    $current = $Combo.SelectedItem
    $Combo.Items.Clear()
    [void]$Combo.Items.Add($script:AllLabel)
    foreach ($v in $Values) { [void]$Combo.Items.Add($v) }
    if ($current -and $Combo.Items.Contains($current)) { $Combo.SelectedItem = $current } else { $Combo.SelectedIndex = 0 }
}

function Update-FilterCombos($C) {
    if ($C.Name -eq 'Sys') { Update-SysCombos $C; return }
    $C.Suspend = $true
    try {
        $all  = $C.All
        $pick = { param($prop) $all | ForEach-Object { $_.$prop } | Where-Object { $_ -and $_ -ne '-' } | Sort-Object -Unique }
        Set-ComboItems $C.Switch (& $pick 'Client')
        Set-ComboItems $C.Policy (& $pick 'Policy')
        Set-ComboItems $C.Auth   (& $pick 'AuthType')
    }
    finally { $C.Suspend = $false }
}

function Get-ComboFilter($Combo) {
    $v = [string]$Combo.SelectedItem
    if (-not $v -or $v -eq $script:AllLabel) { return $null }
    return $v
}

function Invoke-Filter($C) {
    if ($C.Suspend) { return }
    if ($C.Name -eq 'Sys') { Invoke-SysFilter $C; return }

    $res     = $C.Result.SelectedIndex
    $sw      = Get-ComboFilter $C.Switch
    $pol     = Get-ComboFilter $C.Policy
    $auth    = Get-ComboFilter $C.Auth
    $user    = $C.User.Text.Trim()
    $userHex = Get-HexMac $user
    $comp    = $C.Computer.Text.Trim()
    $any     = $C.Search.Text.Trim()
    $showReq  = $(if ($C.ShowReq)  { [bool]$C.ShowReq.IsChecked }  else { $true })
    $showAcct = $(if ($C.ShowAcct) { [bool]$C.ShowAcct.IsChecked } else { $true })

    $list = New-Object System.Collections.Generic.List[object]
    foreach ($e in $C.All) {
        if (-not $showReq  -and $e.Kind -eq 'Req')  { continue }
        if (-not $showAcct -and $e.Kind -eq 'Acct') { continue }
        if ($res -eq 1 -and $e.Level -ne 'OK')    { continue }
        if ($res -eq 2 -and $e.Level -ne 'Error') { continue }
        if ($res -eq 3 -and $e.Level -in 'OK', 'Error') { continue }
        if ($sw   -and $e.Client   -ne $sw)   { continue }
        if ($pol  -and $e.Policy   -ne $pol)  { continue }
        if ($auth -and $e.AuthType -ne $auth) { continue }

        if ($user) {
            $hit = (Test-Contains $e.User $user) -or (Test-Contains $e.UserFQ $user) -or (Test-Contains $e.CallingStation $user)
            if (-not $hit -and $userHex) { $hit = (Test-Contains $e.MacHex $userHex) -or (Test-Contains $e.UserHex $userHex) }
            if (-not $hit) { continue }
        }
        if ($comp -and -not (Test-Contains $e.Machine $comp)) { continue }
        if ($any  -and -not (Test-Contains $e.SearchText $any)) { continue }

        $list.Add($e)
    }

    $C.View = $list
    $C.Grid.ItemsSource = $list
    Update-Stats $C
}

function Update-Stats($C) {
    $v = $C.View
    $ok   = @($v | Where-Object Level -eq 'OK').Count
    $err  = @($v | Where-Object Level -eq 'Error').Count
    $oth  = $v.Count - $ok - $err
    $C.Counts.Text = "Widok: $($v.Count) / $($C.All.Count)    Udzielono: $ok    Odmowa: $err    Inne: $oth"

    if ($v.Count -eq 0) { $C.Stats.Text = 'Brak wpisów w widoku.'; return }

    $nm = { param($s) if ($s) { $s } else { '(brak)' } }
    $sb = New-Object System.Text.StringBuilder
    $denied = @($v | Where-Object Level -eq 'Error')

    [void]$sb.AppendLine('Najczęstsze przyczyny odmów:')
    if ($denied.Count) {
        foreach ($g in ($denied | Group-Object ReasonCode | Sort-Object Count -Descending | Select-Object -First 5)) {
            $txt = [string]$g.Group[0].Reason
            if ($txt.Length -gt 60) { $txt = $txt.Substring(0, 60) + '...' }
            [void]$sb.AppendLine(("  {0,4} x  [{1}] {2}" -f $g.Count, $g.Name, $txt))
        }

        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Odmowy wg switcha:')
        foreach ($g in ($denied | Group-Object Client | Sort-Object Count -Descending | Select-Object -First 5)) {
            [void]$sb.AppendLine(("  {0,4} x  {1}" -f $g.Count, (& $nm $g.Name)))
        }

        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Najczęściej odrzucani (użytkownik / MAC):')
        foreach ($g in ($denied | Group-Object User | Sort-Object Count -Descending | Select-Object -First 5)) {
            [void]$sb.AppendLine(("  {0,4} x  {1}" -f $g.Count, (& $nm $g.Name)))
        }
    }
    else { [void]$sb.AppendLine('  brak odmów') }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Użycie zasad sieciowych:')
    foreach ($g in ($v | Where-Object { $_.Kind -ne 'Req' -and $_.Kind -ne 'Acct' } | Group-Object Policy | Sort-Object Count -Descending | Select-Object -First 6)) {
        [void]$sb.AppendLine(("  {0,4} x  {1}" -f $g.Count, (& $nm $g.Name)))
    }
    $C.Stats.Text = $sb.ToString().TrimEnd()
}

# --- Zdarzenia systemowe: filtr, statystyki, nazwy klientów --------------------------------
# Nazwa switcha dla IP z komunikatu - z danych wczytanych w pozostałych zakładkach
function Resolve-SysClientNames {
    $map = @{}
    foreach ($cx in $script:Ctx.Ev, $script:Ctx.Log) {
        foreach ($e in $cx.All) {
            if ($e.ClientIp -and $e.ClientIp -ne '-' -and $e.Client -and $e.Client -ne '-' -and -not $map.ContainsKey($e.ClientIp)) { $map[$e.ClientIp] = $e.Client }
        }
    }
    foreach ($e in $script:Ctx.Sys.All) {
        if ($e.ClientIp -and $map.ContainsKey($e.ClientIp)) { $e.ClientName = $map[$e.ClientIp] }
    }
}

function Update-SysCombos($C) {
    $C.Suspend = $true
    try {
        $all = $C.All
        Set-ComboItems $C.Id     ($all | ForEach-Object { $_.Id } | Sort-Object { [int]$_ } -Unique)
        Set-ComboItems $C.Source ($all | ForEach-Object { $_.Source } | Where-Object { $_ } | Sort-Object -Unique)
        Set-ComboItems $C.Client ($all | ForEach-Object { if ($_.ClientIp) { $(if ($_.ClientName) { "$($_.ClientIp) ($($_.ClientName))" } else { $_.ClientIp }) } } | Sort-Object -Unique)
    }
    finally { $C.Suspend = $false }
}

function Invoke-SysFilter($C) {
    $lv  = $C.Result.SelectedIndex
    $id  = Get-ComboFilter $C.Id
    $src = Get-ComboFilter $C.Source
    $cl  = Get-ComboFilter $C.Client
    if ($cl) { $cl = ($cl -split ' ')[0] }
    $any = $C.Search.Text.Trim()

    $list = New-Object System.Collections.Generic.List[object]
    foreach ($e in $C.All) {
        if ($lv -eq 1 -and $e.Level -ne 'Error') { continue }
        if ($lv -eq 2 -and $e.Level -ne 'Warn')  { continue }
        if ($lv -eq 3 -and $e.Level -ne 'Info')  { continue }
        if ($id  -and $e.Id -ne $id)        { continue }
        if ($src -and $e.Source -ne $src)   { continue }
        if ($cl  -and $e.ClientIp -ne $cl)  { continue }
        if ($any -and -not (Test-Contains $e.SearchText $any)) { continue }
        $list.Add($e)
    }
    $C.View = $list
    $C.Grid.ItemsSource = $list

    $v = $list
    $err  = @($v | Where-Object Level -eq 'Error').Count
    $warn = @($v | Where-Object Level -eq 'Warn').Count
    $C.Counts.Text = "Widok: $($v.Count) / $($C.All.Count)    Błędy: $err    Ostrzeżenia: $warn    Informacje: $($v.Count - $err - $warn)"
    if ($v.Count -eq 0) { $C.Stats.Text = 'Brak zdarzeń w widoku.'; return }

    $now = Get-Date
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('Wg ID zdarzenia:')
    [void]$sb.AppendLine(('  {0,-6} {1,6} {2,6} {3,6}  {4}' -f 'ID', '1 h', '24 h', 'widok', 'opis'))
    foreach ($g in ($v | Group-Object Id | Sort-Object Count -Descending | Select-Object -First 10)) {
        $h1  = @($g.Group | Where-Object { $_.Time -ge $now.AddHours(-1) }).Count
        $h24 = @($g.Group | Where-Object { $_.Time -ge $now.AddHours(-24) }).Count
        $txt = [string]$g.Group[0].Short
        if ($txt.Length -gt 45) { $txt = $txt.Substring(0, 45) + '...' }
        [void]$sb.AppendLine(('  {0,-6} {1,6} {2,6} {3,6}  {4}' -f $g.Name, $h1, $h24, $g.Count, $txt))
    }

    $withIp = @($v | Where-Object { $_.ClientIp })
    if ($withIp.Count) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Wg klienta RADIUS:')
        foreach ($g in ($withIp | Group-Object ClientIp | Sort-Object Count -Descending | Select-Object -First 8)) {
            $name = $g.Group[0].ClientName
            $ids  = ($g.Group | ForEach-Object Id | Sort-Object -Unique) -join ','
            [void]$sb.AppendLine(('  {0,5} x  {1}{2}  [ID {3}]' -f $g.Count, $g.Name, $(if ($name) { " ($name)" } else { '' }), $ids))
        }
    }
    $C.Stats.Text = $sb.ToString().TrimEnd()
}

function Reset-Filters($C) {
    $C.Suspend = $true
    try {
        if ($C.Name -eq 'Sys') {
            foreach ($cb in $C.Result, $C.Id, $C.Source, $C.Client) { $cb.SelectedIndex = 0 }
            $C.Search.Clear()
            return
        }
        $C.Result.SelectedIndex = 0
        $C.Switch.SelectedIndex = 0
        $C.Policy.SelectedIndex = 0
        $C.Auth.SelectedIndex   = 0
        $C.User.Clear(); $C.Computer.Clear(); $C.Search.Clear()
    }
    finally { $C.Suspend = $false }
    Invoke-Filter $C
}

# --- Szczegóły -----------------------------------------------------------------------------
function Show-Details($C, $e) {
    if (-not $e) {
        $C.Details.Text = $(if ($C.Name -eq 'Ev') { 'Zaznacz zdarzenie na liście.' } else { 'Zaznacz wpis na liście.' })
        return
    }
    $C.Details.Text = Get-DetailsText $C.Name $e
}

# Tekst szczegółów zdarzenia ($Name: 'Ev' = dziennik Security, 'Log' = plik .log, 'Sys' = systemowe).
# Używany w zakładkach i w oknie historii.
function Get-DetailsText([string]$Name, $e) {
    if ($Name -eq 'Sys') {
        $sb = New-Object System.Text.StringBuilder
        $fields = [ordered]@{
            'Czas'          = $e.TimeStr
            'Poziom'        = $e.Result
            'ID zdarzenia'  = $e.Id
            'Źródło'        = $e.Source
            'Dziennik'      = $e.LogName
            'Klient RADIUS' = $(if ($e.ClientName) { "$($e.ClientIp) ($($e.ClientName))" } else { $e.ClientIp })
            'Serwer NPS'    = $e.Server
            'RecordId'      = $e.RecordId
        }
        foreach ($k in $fields.Keys) { [void]$sb.AppendLine(("{0,-15}: {1}" -f $k, $fields[$k])) }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('TREŚĆ:')
        [void]$sb.AppendLine($e.Message.Trim())
        $hint = $script:SysHints[[string]$e.Id]
        if ($hint) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('PODPOWIEDŹ:')
            [void]$sb.AppendLine($hint)
        }
        if ($e.Props.Count) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('DANE ZDARZENIA:')
            for ($i = 0; $i -lt $e.Props.Count; $i++) { [void]$sb.AppendLine(("  [{0}] {1}" -f $i, $e.Props[$i])) }
        }
        return $sb.ToString().TrimEnd()
    }

    if ($Name -eq 'Ev') {
        $fields = [ordered]@{
            'Czas'                    = $e.TimeStr
            'Wynik'                   = "$($e.Result) (ID $($e.Id))"
            'Użytkownik'              = "$($e.Domain)\$($e.User)"
            'Nazwa FQ'                = $e.UserFQ
            'Komputer'                = $e.Machine
            'MAC klienta (Calling)'   = $e.CallingStation
            'Called-Station-ID'       = $e.CalledStation
            'Switch (klient RADIUS)'  = "$($e.Client) [$($e.ClientIp)]"
            'NAS IPv4'                = $e.NasIp
            'NAS Identifier'          = $e.NasId
            'Typ / port NAS'          = "$($e.NasPortType) / $($e.NasPort)"
            'Zasada żądań połączeń'   = $e.CRP
            'Zasada sieciowa'         = $e.Policy
            'Typ uwierzytelniania'    = $e.AuthType
            'Typ EAP'                 = $e.EapType
            'Dostawca'                = $e.AuthProvider
            'Serwer uwierzytelniania' = $e.AuthServer
            'Kod przyczyny'           = $e.ReasonCode
            'Przyczyna'               = $e.Reason
            'Serwer NPS'              = $e.Server
            'RecordId'                = $e.RecordId
        }
    }
    else {
        $pktNames = @{ '1' = 'Access-Request'; '2' = 'Access-Accept'; '3' = 'Access-Reject'; '4' = 'Accounting-Request'; '5' = 'Accounting-Response'; '11' = 'Access-Challenge' }
        $pkt = $pktNames[[string]$e.Id]; if (-not $pkt) { $pkt = "typ $($e.Id)" }
        $fields = [ordered]@{
            'Czas'                    = $e.TimeStr
            'Wynik'                   = "$($e.Result) ($pkt)"
            'User-Name'               = $e.User
            'SAM-Account-Name'        = $e.Sam
            'Nazwa FQ'                = $e.UserFQ
            'MAC klienta (Calling)'   = $e.CallingStation
            'Called-Station-ID'       = $e.CalledStation
            'Switch (klient RADIUS)'  = "$($e.Client) [$($e.ClientIp)]"
            'NAS IPv4'                = $e.NasIp
            'NAS Identifier'          = $e.NasId
            'Typ / port NAS'          = "$($e.NasPortType) / $($e.NasPort)"
            'Zasada żądań połączeń'   = $e.CRP
            'Zasada sieciowa'         = $e.Policy
            'Typ uwierzytelniania'    = $e.AuthType
            'Typ EAP'                 = $e.EapType
            'Dostawca'                = $e.AuthProvider
            'Kod przyczyny'           = $e.ReasonCode
            'Przyczyna'               = $e.Reason
            'Serwer NPS'              = $e.Server
            'Plik'                    = $e.FilePath
            'Linia'                   = $e.RecordId
        }
    }

    $sb = New-Object System.Text.StringBuilder
    foreach ($k in $fields.Keys) { [void]$sb.AppendLine(("{0,-24}: {1}" -f $k, $fields[$k])) }

    $hint = $script:ReasonHints[[string]$e.ReasonCode]
    if ($hint -and $e.Level -ne 'OK') {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('PODPOWIEDŹ:')
        [void]$sb.AppendLine($hint)
    }
    if ($e.Level -eq 'Error' -and $e.Policy -and $e.Policy -ne '-' -and $e.ReasonCode -eq '65') {
        [void]$sb.AppendLine("Obsłużyła zasada: '$($e.Policy)' - jeśli to domyślna zasada odmawiająca, żądanie nie pasowało do zasady docelowej.")
    }

    if ($Name -eq 'Log') {
        if ($e.Merged) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('UWAGA: User-Name / MAC / port uzupełniono z poprzedniego Access-Request od tego samego klienta RADIUS.')
        }
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('WSZYSTKIE ATRYBUTY Z LOGU:')
        foreach ($k in ($e.Attrs.Keys | Where-Object { -not $_.StartsWith('#') } | Sort-Object)) {
            [void]$sb.AppendLine(("  {0,-30} {1}" -f $k, $e.Attrs[$k]))
        }
    }
    return $sb.ToString().TrimEnd()
}

# --- Eksport -------------------------------------------------------------------------------
function Invoke-Export($C) {
    if ($C.View.Count -eq 0) { Show-Msg 'Brak wpisów w widoku.'; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv'
    if ($C.Name -eq 'Ev') {
        $dlg.FileName = "NPS_zdarzenia_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        $cols = 'TimeStr', 'Id', 'Result', 'Domain', 'User', 'UserFQ', 'Machine', 'CallingStation', 'CalledStation',
        'Client', 'ClientIp', 'NasIp', 'NasId', 'NasPortType', 'NasPort', 'CRP', 'Policy', 'AuthType', 'EapType',
        'AuthProvider', 'AuthServer', 'ReasonCode', 'Reason', 'Server', 'RecordId'
    }
    elseif ($C.Name -eq 'Sys') {
        $dlg.FileName = "NPS_system_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        $cols = 'TimeStr', 'Result', 'Id', 'Source', 'LogName', 'ClientIp', 'ClientName', 'Server', 'RecordId',
        @{ n = 'Message'; e = { ($_.Message -replace '\s*\r?\n\s*', ' ').Trim() } }
    }
    else {
        $dlg.FileName = "NPS_logi_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        $cols = 'TimeStr', 'Id', 'Result', 'User', 'Sam', 'UserFQ', 'CallingStation', 'CalledStation',
        'Client', 'ClientIp', 'NasIp', 'NasId', 'NasPortType', 'NasPort', 'CRP', 'Policy', 'AuthType', 'EapType',
        'AuthProvider', 'ReasonCode', 'Reason', 'Server', 'File', 'RecordId', 'Merged'
    }
    if ($dlg.ShowDialog($win)) {
        $C.View | Select-Object -Property $cols | Export-Csv -Path $dlg.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
        Set-Status "Zapisano $($C.View.Count) wpisów: $($dlg.FileName)" 'OK'
    }
}

# --- Menu kontekstowe tabel ----------------------------------------------------------------
function Add-MenuItem($Menu, [string]$Name, [string]$Header, [scriptblock]$Action) {
    $mi = New-Object System.Windows.Controls.MenuItem
    $mi.Header = $Header
    $mi.Tag    = $Name
    $mi.Add_Click($Action)
    [void]$Menu.Items.Add($mi)
}

function New-GridMenu([string]$Name) {
    $cm = New-Object System.Windows.Controls.ContextMenu
    Add-MenuItem $cm $Name 'Filtruj: ten użytkownik / MAC' {
        $cx = $script:Ctx[$this.Tag]; $e = $cx.Grid.SelectedItem; if (-not $e) { return }
        $cx.User.Text = $(if ($e.MacHex.Length -eq 12) { $e.MacHex } else { $e.User })
    }
    Add-MenuItem $cm $Name 'Filtruj: ten switch' {
        $cx = $script:Ctx[$this.Tag]; $e = $cx.Grid.SelectedItem; if (-not $e) { return }
        if ($cx.Switch.Items.Contains($e.Client)) { $cx.Switch.SelectedItem = $e.Client }
    }
    Add-MenuItem $cm $Name 'Filtruj: ta zasada sieciowa' {
        $cx = $script:Ctx[$this.Tag]; $e = $cx.Grid.SelectedItem; if (-not $e) { return }
        if ($cx.Policy.Items.Contains($e.Policy)) { $cx.Policy.SelectedItem = $e.Policy }
    }
    Add-MenuItem $cm $Name $(if ($Name -eq 'Ev') { 'Filtruj: ten komputer' } else { 'Filtruj: ten serwer NPS' }) {
        $cx = $script:Ctx[$this.Tag]; $e = $cx.Grid.SelectedItem; if (-not $e -or -not $e.Machine -or $e.Machine -eq '-') { return }
        $cx.Computer.Text = $e.Machine
    }
    [void]$cm.Items.Add((New-Object System.Windows.Controls.Separator))
    Add-MenuItem $cm $Name 'Historia tego urządzenia (MAC)...' {
        Open-HistoryFromEvent $script:Ctx[$this.Tag].Grid.SelectedItem 'Mac'
    }
    Add-MenuItem $cm $Name 'Historia tego użytkownika...' {
        Open-HistoryFromEvent $script:Ctx[$this.Tag].Grid.SelectedItem 'User'
    }
    Add-MenuItem $cm $Name 'Historia tego komputera...' {
        Open-HistoryFromEvent $script:Ctx[$this.Tag].Grid.SelectedItem 'Computer'
    }
    [void]$cm.Items.Add((New-Object System.Windows.Controls.Separator))
    Add-MenuItem $cm $Name 'Kopiuj szczegóły' {
        $cx = $script:Ctx[$this.Tag]; if ($cx.Grid.SelectedItem -and $cx.Details.Text) { [System.Windows.Clipboard]::SetText($cx.Details.Text) }
    }
    Add-MenuItem $cm $Name 'Kopiuj MAC' {
        $e = $script:Ctx[$this.Tag].Grid.SelectedItem; if ($e -and $e.CallingStation) { [System.Windows.Clipboard]::SetText($e.CallingStation) }
    }
    if ($Name -eq 'Log') {
        Add-MenuItem $cm $Name 'Kopiuj surowy wpis z logu' {
            $e = $script:Ctx[$this.Tag].Grid.SelectedItem; if ($e -and $e.Attrs['#Raw']) { [System.Windows.Clipboard]::SetText([string]$e.Attrs['#Raw']) }
        }
    }
    return $cm
}

function New-SysMenu {
    $cm = New-Object System.Windows.Controls.ContextMenu
    Add-MenuItem $cm 'Sys' 'Filtruj: to ID zdarzenia' {
        $cx = $script:Ctx.Sys; $e = $cx.Grid.SelectedItem; if (-not $e) { return }
        if ($cx.Id.Items.Contains($e.Id)) { $cx.Id.SelectedItem = $e.Id }
    }
    Add-MenuItem $cm 'Sys' 'Filtruj: ten klient RADIUS' {
        $cx = $script:Ctx.Sys; $e = $cx.Grid.SelectedItem; if (-not $e -or -not $e.ClientIp) { return }
        foreach ($it in $cx.Client.Items) { if (([string]$it -split ' ')[0] -eq $e.ClientIp) { $cx.Client.SelectedItem = $it; break } }
    }
    Add-MenuItem $cm 'Sys' 'Pokaż autoryzacje tego klienta (dziennik Security)' {
        $e = $script:Ctx.Sys.Grid.SelectedItem; if (-not $e -or -not $e.ClientIp) { return }
        $ui.txtSearch.Text = $e.ClientIp
        $ui.tabMain.SelectedIndex = 0
    }
    [void]$cm.Items.Add((New-Object System.Windows.Controls.Separator))
    Add-MenuItem $cm 'Sys' 'Kopiuj szczegóły' {
        $cx = $script:Ctx.Sys; if ($cx.Grid.SelectedItem -and $cx.Details.Text) { [System.Windows.Clipboard]::SetText($cx.Details.Text) }
    }
    Add-MenuItem $cm 'Sys' 'Kopiuj IP klienta' {
        $e = $script:Ctx.Sys.Grid.SelectedItem; if ($e -and $e.ClientIp) { [System.Windows.Clipboard]::SetText($e.ClientIp) }
    }
    return $cm
}

# --- Historia (oś czasu) urządzenia / użytkownika / komputera ------------------------------
# Osobne, niemodalne okno: wszystkie zdarzenia jednego MAC, użytkownika albo komputera ułożone
# w czasie (jak "łańcuch zdarzeń" w antywirusach) - z zaznaczonymi zmianami sieci (LAN <-> Wi-Fi,
# inny switch / SSID), przerwami i odmowami. Dane: z zakładek (od ręki) albo wczytane w tle dla
# wybranego zakresu czasu z dziennika Security i/lub plików .log.

$script:HistXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Historia" Height="860" Width="1320" MinHeight="560" MinWidth="900"
        WindowStartupLocation="CenterOwner" Background="#15161C"
        FontFamily="Segoe UI" FontSize="13">
  <Grid Margin="16">
    <Grid.Resources>
      <Style x:Key="TimelineItem" TargetType="ListBoxItem">
        <Setter Property="Padding" Value="0"/>
        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
        <Setter Property="Template">
          <Setter.Value>
            <ControlTemplate TargetType="ListBoxItem">
              <Border x:Name="bd" Background="Transparent">
                <ContentPresenter/>
              </Border>
              <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                  <Setter TargetName="bd" Property="Background" Value="#2B2E3B"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                  <Setter TargetName="bd" Property="Background" Value="#33405E"/>
                </Trigger>
              </ControlTemplate.Triggers>
            </ControlTemplate>
          </Setter.Value>
        </Setter>
      </Style>
    </Grid.Resources>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- NAGŁÓWEK -->
    <DockPanel Grid.Row="0" Margin="0,0,0,12">
      <Border DockPanel.Dock="Right" Background="{StaticResource Card}" CornerRadius="14" Padding="12,0" Height="28" VerticalAlignment="Center">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Ellipse Name="hDot" Width="8" Height="8" Fill="#8A90A2" VerticalAlignment="Center" Margin="0,0,8,0"/>
          <TextBlock Name="hState" Text="Brak danych" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
        </StackPanel>
      </Border>
      <StackPanel Margin="0,0,16,0" VerticalAlignment="Center">
        <TextBlock Name="hTitle" Text="Historia urządzenia / użytkownika" Style="{StaticResource PageTitle}"/>
        <TextBlock Name="hSubtitle" Style="{StaticResource PageSubtitle}"
                   Text="Wszystkie zdarzenia jednego MAC, użytkownika albo komputera w kolejności czasu - ze zmianami sieci, przerwami i odmowami"/>
      </StackPanel>
    </DockPanel>

    <!-- ZAPYTANIE -->
    <Border Grid.Row="1" Style="{StaticResource FormCard}">
      <WrapPanel>
        <StackPanel Style="{StaticResource FieldBox}" Width="150">
          <TextBlock Text="SZUKAJ WG" Style="{StaticResource Caption}"/>
          <ComboBox Name="hMode" SelectedIndex="0">
            <ComboBoxItem Content="MAC (urządzenie)"/>
            <ComboBoxItem Content="Użytkownik"/>
            <ComboBoxItem Content="Komputer"/>
          </ComboBox>
        </StackPanel>
        <StackPanel Style="{StaticResource FieldBox}" Width="200">
          <TextBlock Text="WARTOŚĆ (MAC w dowolnym formacie)" Style="{StaticResource Caption}"/>
          <TextBox Name="hValue" FontFamily="Consolas"/>
        </StackPanel>
        <StackPanel Style="{StaticResource FieldBox}" Width="200">
          <TextBlock Text="ZAKRES CZASU" Style="{StaticResource Caption}"/>
          <ComboBox Name="hRange" SelectedIndex="0">
            <ComboBoxItem Content="Dane wczytane w zakładkach"/>
            <ComboBoxItem Content="Ostatnia godzina"/>
            <ComboBoxItem Content="Ostatnie 24 godziny"/>
            <ComboBoxItem Content="Ostatnie 7 dni"/>
            <ComboBoxItem Content="Ostatnie 30 dni"/>
            <ComboBoxItem Content="Ostatnie 90 dni"/>
            <ComboBoxItem Content="Własny zakres"/>
          </ComboBox>
        </StackPanel>
        <StackPanel Style="{StaticResource FieldBox}" Width="140">
          <TextBlock Text="OD (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
          <TextBox Name="hFrom" FontFamily="Consolas" IsEnabled="False"/>
        </StackPanel>
        <StackPanel Style="{StaticResource FieldBox}" Width="140">
          <TextBlock Text="DO (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
          <TextBox Name="hTo" FontFamily="Consolas" IsEnabled="False"/>
        </StackPanel>
        <StackPanel Style="{StaticResource FieldBox}" Width="100">
          <TextBlock Text="MAKS. ZDARZEŃ" Style="{StaticResource Caption}"/>
          <TextBox Name="hMax" Text="20000" FontFamily="Consolas"/>
        </StackPanel>
        <StackPanel Style="{StaticResource FieldBox}">
          <TextBlock Text="ŹRÓDŁA (przy wczytywaniu)" Style="{StaticResource Caption}"/>
          <StackPanel Style="{StaticResource InlineRow}">
            <CheckBox Name="hSrcEv" Content="Dziennik Security" IsChecked="True"/>
            <CheckBox Name="hSrcLog" Content="Pliki .log" Margin="0"/>
          </StackPanel>
        </StackPanel>
        <StackPanel Style="{StaticResource FieldBox}">
          <StackPanel Style="{StaticResource InlineRow}">
            <CheckBox Name="hPartial" Content="Dopasowanie częściowe" Margin="0" ToolTip="Fragment MAC / nazwy zamiast dokładnej wartości (wolniejsze - bez filtra po stronie serwera)"/>
          </StackPanel>
        </StackPanel>
        <StackPanel Style="{StaticResource FieldBox}">
          <StackPanel Style="{StaticResource InlineRow}">
            <Button Name="hLoad" Content="Pokaż historię" Style="{StaticResource BtnPrimary}"/>
          </StackPanel>
        </StackPanel>
      </WrapPanel>
    </Border>

    <!-- PODSUMOWANIE + PASEK AKTYWNOŚCI + POWIĄZANE -->
    <Border Grid.Row="2" Style="{StaticResource CardBox}" Padding="14,12,14,6">
      <ScrollViewer Name="hSummary" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
      <StackPanel>
        <WrapPanel Name="hChips"/>
        <DockPanel Margin="0,2,0,0">
          <TextBlock Name="hStripInfo" DockPanel.Dock="Right" Foreground="{StaticResource Muted}" FontSize="11" Margin="12,0,0,0" VerticalAlignment="Center"
                     Text="Kliknij słupek, aby zawęzić widok"/>
          <TextBlock Text="AKTYWNOŚĆ W CZASIE" Style="{StaticResource Caption}" Margin="0"/>
        </DockPanel>
        <Border Name="hStripHost" Background="{StaticResource Field}" CornerRadius="6" Height="48" Margin="0,4,0,2" ClipToBounds="True">
          <Canvas Name="hStrip" Background="Transparent" Cursor="Hand"/>
        </Border>
        <DockPanel Margin="0,0,0,8">
          <TextBlock Name="hStripTo" DockPanel.Dock="Right" Foreground="{StaticResource Muted}" FontSize="11" FontFamily="Consolas"/>
          <TextBlock Name="hStripFrom" Foreground="{StaticResource Muted}" FontSize="11" FontFamily="Consolas"/>
        </DockPanel>
        <StackPanel Name="hRelated"/>
      </StackPanel>
      </ScrollViewer>
    </Border>

    <!-- OŚ CZASU + SZCZEGÓŁY (szerokość panelu szczegółów: przeciągnij separator) -->
    <Grid Grid.Row="3">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="2*" MinWidth="420"/>
        <ColumnDefinition Width="12"/>
        <ColumnDefinition Width="*" MinWidth="280"/>
      </Grid.ColumnDefinitions>

      <Border Grid.Column="0" Style="{StaticResource PanelCard}">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <DockPanel Grid.Row="0" Style="{StaticResource CardHeader}">
            <TextBlock Text="Oś czasu" Style="{StaticResource CardTitle}" DockPanel.Dock="Left" Margin="0,0,12,0"/>
            <TextBlock Name="hCounts" Style="{StaticResource Counts}" HorizontalAlignment="Right"/>
          </DockPanel>
          <WrapPanel Grid.Row="1" Margin="0,0,0,4">
            <ComboBox Name="hResult" Width="130" SelectedIndex="0" Margin="0,0,16,8" VerticalAlignment="Center">
              <ComboBoxItem Content="Wszystkie"/>
              <ComboBoxItem Content="Udzielono"/>
              <ComboBoxItem Content="Odmowa"/>
              <ComboBoxItem Content="Inne"/>
            </ComboBox>
            <CheckBox Name="hGroup" Content="Grupuj powtórzenia" IsChecked="True" Margin="0,0,16,8"/>
            <CheckBox Name="hMarkers" Content="Zmiany sieci i przerwy" IsChecked="True" Margin="0,0,16,8"/>
            <CheckBox Name="hAcct" Content="Accounting (sesje, IP)" IsChecked="True" Margin="0,0,16,8" ToolTip="Start / stop sesji z plików .log (adres IP, czas sesji)"/>
            <CheckBox Name="hReq" Content="Access-Request / Challenge" Margin="0,0,16,8" ToolTip="Pośrednie pakiety z plików .log - zwykle tylko szum"/>
            <CheckBox Name="hNewest" Content="Najnowsze na górze" Margin="0,0,16,8"/>
            <StackPanel Orientation="Horizontal" Height="28" Margin="0,0,0,8" VerticalAlignment="Center">
              <TextBlock Name="hZoomText" Foreground="{StaticResource Muted}" VerticalAlignment="Center" Text="Widok: cały wczytany zakres"/>
              <Button Name="hZoomClear" Content="Cały zakres" Style="{StaticResource BtnSmall}" Margin="8,0,0,0" Visibility="Collapsed"/>
            </StackPanel>
          </WrapPanel>
          <ListBox Name="hList" Grid.Row="2" Background="{StaticResource Field}" BorderThickness="0"
                   ItemContainerStyle="{StaticResource TimelineItem}" HorizontalContentAlignment="Stretch"
                   ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                   VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling" VirtualizingPanel.ScrollUnit="Pixel">
            <ListBox.ItemTemplate>
              <DataTemplate>
                <Grid>
                  <!-- nagłówek dnia -->
                  <Border Visibility="{Binding DayVis}" Padding="12,12,12,6" BorderBrush="#343847" BorderThickness="0,0,0,1">
                    <DockPanel>
                      <TextBlock Text="{Binding Detail}" DockPanel.Dock="Right" Foreground="#8A90A2" FontSize="12" VerticalAlignment="Bottom" Margin="12,0,0,0"/>
                      <TextBlock Text="{Binding Title}" FontWeight="SemiBold" FontSize="14" Foreground="#E8EAF0" TextTrimming="CharacterEllipsis"/>
                    </DockPanel>
                  </Border>
                  <!-- zmiana sieci / przerwa -->
                  <Grid Visibility="{Binding MarkerVis}">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="130"/>
                      <ColumnDefinition Width="26"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Rectangle Grid.Column="1" Width="2" Fill="#343847" HorizontalAlignment="Center"/>
                    <Border Grid.Column="2" Background="{Binding MediumBg}" CornerRadius="6" Padding="8,4" Margin="4,4,12,4" HorizontalAlignment="Left">
                      <TextBlock Text="{Binding Title}" Foreground="{Binding TitleColor}" FontStyle="Italic" TextWrapping="Wrap"/>
                    </Border>
                  </Grid>
                  <!-- zdarzenie -->
                  <Grid Visibility="{Binding EventVis}">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="130"/>
                      <ColumnDefinition Width="26"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Margin="8,6,4,6" HorizontalAlignment="Right">
                      <TextBlock Text="{Binding TimeStr}" FontFamily="Consolas" FontSize="12" Foreground="#C9CDD8" HorizontalAlignment="Right"/>
                      <TextBlock Text="{Binding SubTime}" FontSize="11" Foreground="#8A90A2" HorizontalAlignment="Right" TextTrimming="CharacterEllipsis"/>
                    </StackPanel>
                    <Rectangle Grid.Column="1" Width="2" Fill="#343847" HorizontalAlignment="Center"/>
                    <Ellipse Grid.Column="1" Width="12" Height="12" Fill="{Binding Color}" Stroke="#272A36" StrokeThickness="2"
                             HorizontalAlignment="Center" VerticalAlignment="Top" Margin="0,9,0,0"/>
                    <StackPanel Grid.Column="2" Margin="4,6,12,6">
                      <WrapPanel>
                        <TextBlock Text="{Binding Title}" FontWeight="SemiBold" Foreground="{Binding TitleColor}" Margin="0,0,8,0"/>
                        <Border Visibility="{Binding MediumVis}" Background="{Binding MediumBg}" CornerRadius="9" Padding="8,1" Margin="0,0,8,0" VerticalAlignment="Center">
                          <TextBlock Text="{Binding Medium}" FontSize="11" FontWeight="SemiBold" Foreground="#E8EAF0"/>
                        </Border>
                        <TextBlock Text="{Binding Location}" Foreground="#C9CDD8"/>
                      </WrapPanel>
                      <TextBlock Text="{Binding Detail}" Foreground="#8A90A2" FontSize="12" TextWrapping="Wrap" Margin="0,2,0,0"/>
                    </StackPanel>
                  </Grid>
                </Grid>
              </DataTemplate>
            </ListBox.ItemTemplate>
          </ListBox>
        </Grid>
      </Border>

      <GridSplitter Grid.Column="1" Style="{StaticResource Splitter}"/>

      <Border Grid.Column="2" Style="{StaticResource PanelCard}">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <DockPanel Grid.Row="0" Style="{StaticResource CardHeader}">
            <Button Name="hCopy" DockPanel.Dock="Right" Content="Kopiuj" Style="{StaticResource BtnSmall}"/>
            <TextBlock Text="Szczegóły" Style="{StaticResource CardTitle}"/>
          </DockPanel>
          <TextBox Name="hDetails" Grid.Row="1" Style="{StaticResource DetailsBox}"
                   Text="Wpisz MAC, użytkownika albo komputer i kliknij Pokaż historię."/>
        </Grid>
      </Border>
    </Grid>

    <!-- PASEK STATUSU -->
    <Grid Grid.Row="4" Margin="0,12,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="200"/>
      </Grid.ColumnDefinitions>
      <TextBlock Name="hStatus" Text="Gotowy" Foreground="{StaticResource Muted}" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
      <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="12,0,12,0">
        <Button Name="hCopyAll" Content="Kopiuj oś czasu" Style="{StaticResource Btn}" Margin="0,0,8,0"/>
        <Button Name="hExport" Content="Eksport CSV" Style="{StaticResource Btn}"/>
      </StackPanel>
      <ProgressBar Name="hPb" Grid.Column="2" VerticalAlignment="Center" Visibility="Hidden"/>
    </Grid>
  </Grid>
</Window>
'@

$script:HistWins       = @{}
$script:HistSeq        = 0
$script:HistGapMinutes = 60
$script:PlCulture      = [Globalization.CultureInfo]::GetCultureInfo('pl-PL')
$script:MediumBg       = @{ 'Wi-Fi' = '#2E4A7D'; 'LAN' = '#24584A'; 'VPN' = '#55437A' }

# --- Silnik osi czasu (C#) -----------------------------------------------------------------
# Dopasowanie, elementy i wiersze osi czasu, podsumowanie, filtry widoku i pasek aktywności dla
# dziesiątek tysięcy zdarzeń. W Windows PowerShell 5.1 te same pętle w PowerShellu trwały sekundy
# (okna zamarzały przy każdym przełączeniu widoku), w C# - milisekundy. Kompilacja (Add-Type) raz
# na sesję, w tle zaraz po starcie skryptu. Funkcje PowerShell poniżej to tylko nakładki.
$script:HistEngineSource = @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Management.Automation;
using System.Text;
using System.Text.RegularExpressions;

namespace NpsHistory
{
    // Element osi czasu: jedno zdarzenie z dziennika Security albo z pliku .log (E = obiekt zdarzenia)
    public class Item
    {
        public DateTime Time { get; set; }
        public string Src { get; set; }
        public bool Both { get; set; }
        public string Kind { get; set; }
        public string Level { get; set; }
        public object E { get; set; }
        public string Medium { get; set; }
        public string PortKey { get; set; }
        public string LocText { get; set; }
        public string Ssid { get; set; }
        public string Client { get; set; }
        public string UserN { get; set; }
        public string UserDisp { get; set; }
        public string Computer { get; set; }
        public string MacHex { get; set; }
        public string Title { get; set; }
        public List<string> Parts { get; set; }
        public string Sig { get; set; }
    }

    // Wiersz osi czasu (ListBox): dzień, zdarzenie (z grupą powtórzeń) albo znacznik (zmiana sieci, przerwa, inny użytkownik)
    public class Row
    {
        public Row(string kind, DateTime time)
        {
            Kind = kind; Time = time; First = time; Last = time;
            TimeStr = ""; SubTime = ""; Color = "#8A90A2"; Title = ""; TitleColor = "#E8EAF0";
            Medium = ""; MediumBg = "Transparent"; MediumVis = "Collapsed"; Location = ""; Detail = ""; Sig = ""; Info = "";
            DayVis = kind == "Day" ? "Visible" : "Collapsed";
            EventVis = kind == "Event" ? "Visible" : "Collapsed";
            MarkerVis = (kind == "Change" || kind == "Gap" || kind == "Who") ? "Visible" : "Collapsed";
            Items = new List<Item>();
        }
        public string Kind { get; set; }
        public DateTime Time { get; set; }
        public string TimeStr { get; set; }
        public string SubTime { get; set; }
        public string Color { get; set; }
        public string Title { get; set; }
        public string TitleColor { get; set; }
        public string Medium { get; set; }
        public string MediumBg { get; set; }
        public string MediumVis { get; set; }
        public string Location { get; set; }
        public string Detail { get; set; }
        public string DayVis { get; set; }
        public string EventVis { get; set; }
        public string MarkerVis { get; set; }
        public List<Item> Items { get; set; }
        public string Sig { get; set; }
        public DateTime First { get; set; }
        public DateTime Last { get; set; }
        public string Info { get; set; }
    }

    // Ostatnie znane miejsce w sieci (porównanie tylko pól znanych po obu stronach)
    internal class LocState
    {
        public Item It, Prev;
        public string Medium = "", Ssid = "", Port = "";
    }

    public static class Engine
    {
        const RegexOptions Ci = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;
        static readonly Regex RxHexChars = new Regex(@"^[0-9A-Fa-f:\.\-\s]+$", Ci);
        static readonly Regex RxNonHex = new Regex(@"[^0-9A-Fa-f]");
        static readonly Regex RxSsid = new Regex(@"^(?:[0-9A-Fa-f]{2}[-:.]?){5}[0-9A-Fa-f]{2}[:;](.+)$", Ci);
        // Typ portu w dzienniku Security jest w języku systemu ("Wireless - IEEE 802.11", "Drahtlos - IEEE 802.11",
        // "Virtual"...), w plikach .log - numer (19/18 Wi-Fi, 15 LAN, 5 VPN)
        static readonly Regex RxWifi = new Regex(@"802\.11|wireless|wi-?fi|bezprzew|drahtlos|sans fil|inal[aá]mbr|^1[89]$", Ci);
        static readonly Regex RxLan = new Regex(@"ethernet|^15$", Ci);
        static readonly Regex RxVpn = new Regex(@"virtu|wirtu|vpn|^5$", Ci);
        static readonly Regex RxIpv4 = new Regex(@"^\d{1,3}(\.\d{1,3}){3}$");
        static readonly Regex RxUpn = new Regex(@"^([^@]+)@", Ci);
        static readonly Regex RxDigits = new Regex(@"^\d+$");
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // --- dostęp do właściwości obiektów PowerShell ---
        static PSObject P(object o) { return o == null ? null : PSObject.AsPSObject(o); }

        static object Raw(PSObject p, string name)
        {
            if (p == null) return null;
            PSPropertyInfo pi = p.Properties[name];
            if (pi == null) return null;
            object v;
            try { v = pi.Value; } catch { return null; }
            PSObject pv = v as PSObject;
            return pv != null ? pv.BaseObject : v;
        }

        static string Str(PSObject p, string name)
        {
            object v = Raw(p, name);
            if (v == null) return "";
            return Convert.ToString(v, Inv) ?? "";
        }

        static string QS(IDictionary q, string key)
        {
            object v = q[key];
            PSObject pv = v as PSObject;
            if (pv != null) v = pv.BaseObject;
            return v == null ? "" : (Convert.ToString(v, Inv) ?? "");
        }

        static bool Eq(string a, string b) { return string.Equals(a ?? "", b ?? "", StringComparison.OrdinalIgnoreCase); }

        static Item AsItem(object o)
        {
            PSObject p = o as PSObject;
            if (p != null) o = p.BaseObject;
            return o as Item;
        }

        static string Lookup(IDictionary map, string key)
        {
            if (map == null || key == null) return null;
            object v = map[key];
            PSObject pv = v as PSObject;
            if (pv != null) v = pv.BaseObject;
            return v == null ? null : Convert.ToString(v, Inv);
        }

        // --- normalizacja identyfikatorów ---
        public static string HexMac(string v)
        {
            if (string.IsNullOrEmpty(v) || !RxHexChars.IsMatch(v)) return "";
            string h = RxNonHex.Replace(v, "").ToUpperInvariant();
            return h.Length >= 4 ? h : "";
        }

        // Użytkownik: bez domeny ("CONTOSO\jan", "jan@contoso.com" -> "jan"), małymi literami; host/... zostaje
        public static string NormUser(string v)
        {
            if (string.IsNullOrEmpty(v) || v == "-") return "";
            v = v.Trim().ToLowerInvariant();
            int i = v.LastIndexOf('\\');
            if (i >= 0) v = v.Substring(i + 1);
            if (!v.StartsWith("host/", StringComparison.Ordinal))
            {
                Match m = RxUpn.Match(v);
                if (m.Success) v = m.Groups[1].Value;
            }
            return v;
        }

        // Komputer: "host/PC01.contoso.com", "CONTOSO\PC01$", "PC01.contoso.com" -> "pc01"
        public static string NormComputer(string v)
        {
            if (string.IsNullOrEmpty(v) || v == "-") return "";
            v = v.Trim().ToLowerInvariant();
            if (v.StartsWith("host/", StringComparison.Ordinal)) v = v.Substring(5);
            int i = v.LastIndexOf('\\');
            if (i >= 0) v = v.Substring(i + 1);
            v = v.TrimEnd('$');
            int j = v.IndexOf('.');
            if (j > 0) v = v.Substring(0, j);
            return v;
        }

        public static bool IsComputerAccount(string v)
        {
            if (string.IsNullOrEmpty(v)) return false;
            string t = v.Trim();
            return t.StartsWith("host/", StringComparison.OrdinalIgnoreCase) || t.EndsWith("$", StringComparison.Ordinal);
        }

        public static string FormatMac(string hex)
        {
            if (hex == null || hex.Length != 12) return hex ?? "";
            StringBuilder sb = new StringBuilder(17);
            for (int i = 0; i < 12; i += 2) { if (i > 0) sb.Append('-'); sb.Append(hex, i, 2); }
            return sb.ToString();
        }

        public static string FormatSpan(TimeSpan s)
        {
            if (s.TotalDays >= 1) return string.Format("{0} d {1} godz.", (int)Math.Floor(s.TotalDays), s.Hours);
            if (s.TotalHours >= 1) return string.Format("{0} godz. {1} min", (int)Math.Floor(s.TotalHours), s.Minutes);
            if (s.TotalMinutes >= 1) return string.Format("{0} min", (int)Math.Floor(s.TotalMinutes));
            return string.Format("{0} s", (int)Math.Floor(s.TotalSeconds));
        }

        // --- pola zdarzenia ---
        static string UserMacP(PSObject p) { string h = HexMac(Str(p, "User")); return h.Length == 12 ? h : ""; }
        public static string UserMac(object e) { return UserMacP(P(e)); }

        // Tylko obiekty z plików .log mają nazwę pliku
        static string SrcP(PSObject p) { return Raw(p, "File") != null ? "Log" : "Ev"; }
        public static string EventSrc(object e) { return SrcP(P(e)); }

        static string UserDisplayP(PSObject p)
        {
            string u = Str(p, "User");
            if (u == "" || u == "-") return "";
            string d = Str(p, "Domain");
            if (d != "" && d != "-" && u.IndexOf('\\') < 0 && u.IndexOf('@') < 0 && SrcP(p) == "Ev") return d + "\\" + u;
            return u;
        }
        public static string EventUserDisplay(object e) { return UserDisplayP(P(e)); }

        // Komputer zdarzenia: nazwa maszyny z dziennika Security albo konto komputera (host/..., NAZWA$)
        // w nazwie użytkownika. W plikach .log "Computer-Name" to serwer NPS, nie klient.
        static string ComputerP(PSObject p)
        {
            if (SrcP(p) == "Ev")
            {
                string m = Str(p, "Machine");
                if (m != "" && m != "-") return NormComputer(m);
            }
            foreach (string f in new string[] { "User", "UserFQ" })
            {
                string u = Str(p, f);
                if (IsComputerAccount(u)) return NormComputer(u);
            }
            return "";
        }
        public static string EventComputer(object e) { return ComputerP(P(e)); }

        static string SsidP(PSObject p)
        {
            Match m = RxSsid.Match(Str(p, "CalledStation"));
            return m.Success ? m.Groups[1].Value.Trim() : "";
        }
        public static string EventSsid(object e) { return SsidP(P(e)); }

        static string MediumP(PSObject p)
        {
            string t = Str(p, "NasPortType");
            if (RxWifi.IsMatch(t)) return "Wi-Fi";
            if (RxLan.IsMatch(t)) return "LAN";
            if (RxVpn.IsMatch(t)) return "VPN";
            if (SsidP(p) != "") return "Wi-Fi";
            if (RxIpv4.IsMatch(Str(p, "CallingStation"))) return "VPN";
            return "";
        }
        public static string EventMedium(object e) { return MediumP(P(e)); }

        static string ClientNameP(PSObject p)
        {
            foreach (string f in new string[] { "Client", "ClientIp", "NasIp", "NasId" })
            {
                string v = Str(p, f);
                if (v != "" && v != "-") return v;
            }
            return "";
        }
        public static string EventClientName(object e) { return ClientNameP(P(e)); }

        // --- dopasowanie do zapytania ($Q z New-HistoryQuery: Mode, Hex, Norm, Partial) ---
        // Najpierw tanie odrzucenie po surowym tekście, potem dokładne porównanie po normalizacji.
        static bool MatchP(PSObject p, IDictionary q)
        {
            string mode = QS(q, "Mode");
            bool partial = LanguagePrimitives.IsTrue(q["Partial"]);
            if (mode == "Mac")
            {
                // MAC tylko z pełnych 12 cyfr: przy VPN Calling-Station-Id to adres IP (loadery nie liczą
                // z niego MacHex), a nazwa użytkownika "abc.def" też "wygląda" na szesnastkową.
                string hex = QS(q, "Hex");
                string mh = Str(p, "MacHex");
                bool inUser = Str(p, "UserHex").Contains(hex);
                if (!(inUser || mh.Contains(hex))) return false;
                string userHex = inUser ? UserMacP(p) : "";
                string macHex = mh.Length == 12 ? mh : "";
                if (partial) return (macHex != "" && macHex.Contains(hex)) || (userHex != "" && userHex.Contains(hex));
                return Eq(macHex, hex) || Eq(userHex, hex);
            }
            string norm = QS(q, "Norm");
            if (norm == "") return false;
            if (mode == "User")
            {
                foreach (string f in new string[] { "User", "UserFQ", "Sam" })
                {
                    string s = Str(p, f);
                    if (s == "" || s.IndexOf(norm, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    string n = NormUser(s);
                    if (n == "") continue;
                    if (partial ? n.Contains(norm) : Eq(n, norm)) return true;
                }
                return false;
            }
            if (mode == "Computer")
            {
                // Nazwa maszyny tylko z dziennika Security, poza tym konta komputerów w nazwie użytkownika
                bool isEv = SrcP(p) == "Ev";
                string[] fields = new string[] { "Machine", "User", "UserFQ", "Sam" };
                for (int i = 0; i < fields.Length; i++)
                {
                    string s = Str(p, fields[i]);
                    if (s == "" || s.IndexOf(norm, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    if (i == 0) { if (!isEv || s == "-") continue; }
                    else if (!IsComputerAccount(s)) continue;
                    string n = NormComputer(s);
                    if (n == "") continue;
                    if (partial ? n.Contains(norm) : Eq(n, norm)) return true;
                }
                return false;
            }
            return false;
        }
        public static bool Match(object e, IDictionary q) { return e != null && MatchP(P(e), q); }

        // --- element osi czasu ---
        static Item NewItemP(PSObject p, string src, IDictionary q)
        {
            string mode = QS(q, "Mode");
            string medium = MediumP(p);
            string client = ClientNameP(p);
            string ssid = SsidP(p);
            // Port do porównań: NAS-Port (numer) jest w obu źródłach. NAS-Port-Id (nazwa, np.
            // GigabitEthernet1/0/12) jest tylko w plikach .log - służy wyłącznie do opisu.
            string portKey = Str(p, "NasPort"); if (portKey == "-") portKey = "";
            string portTxt = portKey;
            string portId = Str(p, "NasPortId"); if (portId != "") portTxt = portId;
            string locText;
            if (medium == "Wi-Fi")
            {
                portKey = "";
                locText = (ssid != "" && client != "") ? "SSID " + ssid + " · " + client : (ssid != "" ? "SSID " + ssid : client);
            }
            else if (medium == "VPN")
            {
                // NAS-Port przy VPN to numer tunelu / sesji - zmienia się przy każdym połączeniu
                portKey = "";
                locText = client;
            }
            else
            {
                locText = (portTxt != "" && client != "") ? client + " · port " + portTxt : (portTxt != "" ? "port " + portTxt : client);
            }

            string kind = Str(p, "Kind"); if (kind == "") kind = "Resp";
            string level = Str(p, "Level");
            string result = Str(p, "Result");
            string rc = Str(p, "ReasonCode");
            string userDisp = UserDisplayP(p);
            string userN = NormUser(Str(p, "User"));
            string comp = ComputerP(p);
            string macHex = Str(p, "MacHex"); if (macHex.Length != 12) macHex = "";
            string calling = Str(p, "CallingStation");

            List<string> parts = new List<string>();
            if (mode != "User" && userDisp != "") parts.Add("użytkownik " + userDisp);
            if (mode != "Mac" && calling != "" && calling != "-") parts.Add(macHex != "" ? "MAC " + calling : "z adresu " + calling);
            if (mode != "Computer" && comp != "" && !IsComputerAccount(userDisp)) parts.Add("komputer " + comp.ToUpperInvariant());
            string policy = Str(p, "Policy");
            if (policy != "" && policy != "-") parts.Add("zasada: " + policy);
            string at = Str(p, "AuthType"); if (at == "-") at = "";
            string et = Str(p, "EapType"); if (et == "-") et = "";
            string auth = (at != "" && et != "" && !Eq(at, et)) ? at + " / " + et : (at != "" ? at : et);
            if (auth != "") parts.Add(auth);
            string ip = Str(p, "FramedIp");
            if (ip != "") parts.Add("IP " + ip);
            string st = Str(p, "SessionTime");
            if (st != "" && RxDigits.IsMatch(st)) parts.Add("czas sesji " + FormatSpan(TimeSpan.FromSeconds(double.Parse(st, Inv))));
            string term = Str(p, "AcctTerminate");
            if (term != "") parts.Add("koniec sesji: " + term);

            string title = result;
            if (!Eq(level, "OK") && rc != "" && rc != "0" && rc != "-")
            {
                string r = Str(p, "Reason");
                if (r.Length > 90) r = r.Substring(0, 90) + "...";
                title += (r != "" && r != "-") ? " - " + r + " (kod " + rc + ")" : " (kod " + rc + ")";
            }

            object t = Raw(p, "Time");
            Item it = new Item();
            it.Time = t is DateTime ? (DateTime)t : (t == null ? DateTime.MinValue : Convert.ToDateTime(t, Inv));
            it.Src = src; it.Both = false; it.Kind = kind; it.Level = level; it.E = p;
            it.Medium = medium; it.PortKey = portKey; it.LocText = locText; it.Ssid = ssid; it.Client = client;
            it.UserN = userN; it.UserDisp = userDisp; it.Computer = comp; it.MacHex = macHex; it.Title = title; it.Parts = parts;
            it.Sig = string.Join("|", new string[] { level, result, rc, client, medium, ssid, portKey, userN, macHex, kind });
            return it;
        }
        public static Item NewItem(object e, string src, IDictionary q) { return NewItemP(P(e), src, q); }

        // Dopasowanie i budowa elementów dla całego wyniku loadera / danych z zakładki (runspace w tle).
        // Obiekt z IsSummary to podsumowanie loadera plików .log.
        public static Hashtable BuildItems(IEnumerable source, IDictionary q, string src)
        {
            List<Item> items = new List<Item>();
            object summary = null;
            int raw = 0;
            if (source != null)
            {
                foreach (object o in source)
                {
                    if (o == null) continue;
                    PSObject p = PSObject.AsPSObject(o);
                    if (p.Properties["IsSummary"] != null) { summary = o; continue; }
                    raw++;
                    if (MatchP(p, q)) items.Add(NewItemP(p, src, q));
                }
            }
            Hashtable r = new Hashtable(StringComparer.OrdinalIgnoreCase);
            r["Items"] = items; r["Raw"] = raw; r["Summary"] = summary;
            return r;
        }

        // Łączy elementy z dziennika Security i z plików .log. Odpowiedź (Accept/Reject) z logu, która ma
        // odpowiednik w dzienniku Security (ten sam wynik i urządzenie, do 2 s różnicy), nie jest pokazywana
        // drugi raz - zdarzenie z Security dostaje oznaczenie "Security + log". Wynik rosnąco po czasie;
        // przy równym czasie późniejsza pozycja wejścia pierwsza (loadery podają najnowsze najpierw).
        public static List<Item> Merge(IEnumerable evItems, IEnumerable logItems)
        {
            List<Item> all = new List<Item>();
            Dictionary<string, List<Item>> index = new Dictionary<string, List<Item>>(StringComparer.OrdinalIgnoreCase);
            if (evItems != null)
            {
                foreach (object o in evItems)
                {
                    Item it = AsItem(o); if (it == null) continue;
                    all.Add(it);
                    string k = it.Level + "|" + (it.Time.Ticks / TimeSpan.TicksPerSecond).ToString(Inv);
                    List<Item> l;
                    if (!index.TryGetValue(k, out l)) { l = new List<Item>(); index[k] = l; }
                    l.Add(it);
                }
            }
            if (logItems != null)
            {
                foreach (object o in logItems)
                {
                    Item it = AsItem(o); if (it == null) continue;
                    if (Eq(it.Kind, "Resp") && index.Count > 0)
                    {
                        long sec = it.Time.Ticks / TimeSpan.TicksPerSecond;
                        Item dup = null;
                        for (long d = -2; d <= 2 && dup == null; d++)
                        {
                            List<Item> l;
                            if (!index.TryGetValue(it.Level + "|" + (sec + d).ToString(Inv), out l)) continue;
                            foreach (Item c in l)
                            {
                                if (c.Both) continue;
                                bool same = (c.MacHex != "" && it.MacHex != "") ? Eq(c.MacHex, it.MacHex) : (c.UserN != "" && Eq(c.UserN, it.UserN));
                                if (same) { dup = c; break; }
                            }
                        }
                        if (dup != null) { dup.Both = true; continue; }
                    }
                    all.Add(it);
                }
            }
            int n = all.Count;
            int[] idx = new int[n];
            for (int i = 0; i < n; i++) idx[i] = i;
            Array.Sort(idx, delegate (int a, int b)
            {
                int c = all[a].Time.CompareTo(all[b].Time);
                return c != 0 ? c : b.CompareTo(a);
            });
            List<Item> sorted = new List<Item>(n);
            foreach (int i in idx) sorted.Add(all[i]);
            return sorted;
        }

        // Czy element jest w innym miejscu sieci niż ostatnie znane? Porównujemy tylko pola znane po obu
        // stronach: wpis accounting bez typu portu albo zdarzenie bez numeru portu nie oznacza zmiany.
        static bool StepLocation(LocState s, Item it)
        {
            if (it.Client == "") return false;
            if (s.It == null) { s.It = it; s.Medium = it.Medium; s.Ssid = it.Ssid; s.Port = it.PortKey; return false; }
            bool chg = !Eq(s.It.Client, it.Client)
                || (s.Medium != "" && it.Medium != "" && !Eq(s.Medium, it.Medium))
                || (s.Ssid != "" && it.Ssid != "" && !Eq(s.Ssid, it.Ssid))
                || (s.Port != "" && it.PortKey != "" && !Eq(s.Port, it.PortKey));
            if (chg)
            {
                s.Prev = s.It; s.It = it; s.Medium = it.Medium; s.Ssid = it.Ssid; s.Port = it.PortKey;
                return true;
            }
            if (s.Medium == "" && it.Medium != "") { s.Medium = it.Medium; s.It = it; }
            if (s.Ssid == "") s.Ssid = it.Ssid;
            if (s.Port == "" && s.Medium != "Wi-Fi" && s.Medium != "VPN") s.Port = it.PortKey;
            return false;
        }

        static string FormatLocation(Item it)
        {
            string l = it.LocText != "" ? it.LocText : "(nieznane miejsce)";
            return it.Medium != "" ? it.Medium + " · " + l : l;
        }

        // Wiersze osi czasu z elementów posortowanych rosnąco po czasie: zdarzenia (z grupowaniem
        // powtórzeń), znaczniki zmian sieci / urządzenia / użytkownika, przerwy i nagłówki dni.
        public static Hashtable BuildRows(IEnumerable items, IDictionary q, bool group, bool markers, bool newest,
            int gapMinutes, IDictionary colors, IDictionary mediumBg, CultureInfo dayCulture)
        {
            string mode = QS(q, "Mode");
            List<Item> list = new List<Item>();
            if (items != null) foreach (object o in items) { Item it = AsItem(o); if (it != null) list.Add(it); }
            List<Row> seq = new List<Row>();
            TimeSpan gap = TimeSpan.FromMinutes(gapMinutes);
            Item prev = null;
            LocState loc = new LocState();
            string lastWho = null, lastWhoDisp = null;
            Row cur = null;
            int changes = 0;
            string whatTxt = mode == "User" ? "użytkownika" : (mode == "Computer" ? "komputera" : "urządzenia");
            const string fmt = "yyyy-MM-dd HH:mm:ss";

            foreach (Item it in list)
            {
                if (prev != null)
                {
                    TimeSpan dt = it.Time - prev.Time;
                    if (markers && dt > gap)
                    {
                        Row m = new Row("Gap", it.Time);
                        m.Title = "brak zdarzeń przez " + FormatSpan(dt);
                        m.TitleColor = "#8A90A2";
                        m.Info = "Od " + prev.Time.ToString(fmt) + " do " + it.Time.ToString(fmt) + " nie było zdarzeń tego " + whatTxt +
                            ". Urządzenie mogło być wyłączone, odłączone albo po prostu nie uwierzytelniało się ponownie (zależy od ustawień reauth na switchu / AP).";
                        seq.Add(m);
                        cur = null;
                    }
                }

                // Zmiana miejsca w sieci (LAN <-> Wi-Fi, inny switch / port / SSID)
                if (StepLocation(loc, it))
                {
                    changes++;
                    if (markers)
                    {
                        Item from = loc.Prev;
                        Row m = new Row("Change", it.Time);
                        string kindTxt = (from.Medium != "" && it.Medium != "" && !Eq(from.Medium, it.Medium)) ? "zmiana sieci " + from.Medium + " → " + it.Medium : "zmiana miejsca w sieci";
                        m.Title = kindTxt + ":  " + FormatLocation(from) + "  →  " + FormatLocation(it);
                        m.TitleColor = "#93C5FD";
                        m.MediumBg = "#1F2A40";
                        m.Info = "Poprzednie miejsce (" + from.Time.ToString(fmt) + "): " + FormatLocation(from) + "\nNastępne (" + it.Time.ToString(fmt) + "): " + FormatLocation(it);
                        seq.Add(m);
                        cur = null;
                    }
                }

                // Inny użytkownik na tym samym urządzeniu (tryb MAC / komputer) albo inne urządzenie tego samego
                // użytkownika (tryb użytkownik; MacHex jest pusty, gdy to nie MAC - np. VPN)
                string who = mode == "User" ? it.MacHex : it.UserN;
                if (who != "")
                {
                    if (markers && lastWho != null && !Eq(lastWho, who))
                    {
                        Row m = new Row("Who", it.Time);
                        m.Title = mode == "User" ? "inne urządzenie:  " + FormatMac(lastWho) + "  →  " + FormatMac(who) : "inny użytkownik:  " + lastWhoDisp + "  →  " + it.UserDisp;
                        m.TitleColor = "#FCD34D";
                        m.MediumBg = "#3A3320";
                        m.Info = m.Title;
                        seq.Add(m);
                        cur = null;
                    }
                    lastWho = who;
                    lastWhoDisp = it.UserDisp != "" ? it.UserDisp : who;
                }

                if (group && cur != null && Eq(cur.Sig, it.Sig) && cur.Last.Date == it.Time.Date && (it.Time - cur.Last) <= gap)
                {
                    cur.Items.Add(it);
                    cur.Last = it.Time;
                    prev = it;
                    continue;
                }

                Row r = new Row("Event", it.Time);
                r.Sig = it.Sig;
                r.Items.Add(it);
                r.Color = Lookup(colors, it.Level) ?? "#8A90A2";
                r.Title = it.Title;
                r.TitleColor = Eq(it.Level, "Error") ? "#F87171" : (Eq(it.Level, "Warn") ? "#FBBF24" : (Eq(it.Level, "Info") ? "#C9CDD8" : "#E8EAF0"));
                if (it.Medium != "") { r.Medium = it.Medium; r.MediumVis = "Visible"; r.MediumBg = Lookup(mediumBg, it.Medium) ?? "Transparent"; }
                r.Location = it.LocText;
                seq.Add(r);
                cur = r;
                prev = it;
            }

            // Uzupełnienie wierszy (czas, licznik powtórzeń, źródło)
            foreach (Row r in seq)
            {
                if (r.Kind != "Event") continue;
                int n = r.Items.Count;
                Item last = r.Items[n - 1];
                r.TimeStr = (newest ? r.Last : r.First).ToString("HH:mm:ss");
                if (n > 1) r.SubTime = "×" + n.ToString(Inv) + " · " + (newest ? "od " + r.First.ToString("HH:mm:ss") : "do " + r.Last.ToString("HH:mm:ss"));
                bool both = false, hasEv = false, hasLog = false;
                foreach (Item x in r.Items) { if (x.Both) both = true; if (x.Src == "Log") hasLog = true; else hasEv = true; }
                string srcTxt = (both || (hasEv && hasLog)) ? "Security + log" : (hasLog ? "log .log" : "Security");
                r.Detail = last.Parts.Count > 0 ? string.Join("  ·  ", last.Parts.ToArray()) + "  ·  " + srcTxt : srcTxt;
            }

            if (newest) seq.Reverse();

            // Nagłówki dni (z liczbą zdarzeń i odmów danego dnia)
            Dictionary<DateTime, int[]> dayStats = new Dictionary<DateTime, int[]>();
            foreach (Item it in list)
            {
                int[] st;
                if (!dayStats.TryGetValue(it.Time.Date, out st)) { st = new int[2]; dayStats[it.Time.Date] = st; }
                st[0]++;
                if (Eq(it.Level, "Error")) st[1]++;
            }
            List<Row> rows = new List<Row>(seq.Count + 16);
            DateTime day = DateTime.MinValue;
            bool first = true;
            foreach (Row r in seq)
            {
                if (first || day != r.Time.Date)
                {
                    first = false;
                    day = r.Time.Date;
                    Row hd = new Row("Day", day);
                    hd.Title = day.ToString("dddd, d MMMM yyyy", dayCulture ?? CultureInfo.CurrentCulture);
                    int[] st;
                    if (dayStats.TryGetValue(day, out st)) hd.Detail = "zdarzeń: " + st[0].ToString(Inv) + (st[1] > 0 ? "  ·  odmów: " + st[1].ToString(Inv) : "");
                    rows.Add(hd);
                }
                rows.Add(r);
            }
            Hashtable res = new Hashtable(StringComparer.OrdinalIgnoreCase);
            res["Rows"] = rows; res["Changes"] = changes;
            return res;
        }

        // Filtry widoku (wynik, żądania / accounting, zawężenie czasu): Base = po filtrach, View = Base w zawężeniu
        public static Hashtable Filter(IEnumerable items, int result, bool showReq, bool showAcct, object zoomFrom, object zoomTo)
        {
            PSObject zf = zoomFrom as PSObject; if (zf != null) zoomFrom = zf.BaseObject;
            PSObject zt = zoomTo as PSObject; if (zt != null) zoomTo = zt.BaseObject;
            bool zoom = zoomFrom is DateTime && zoomTo is DateTime;
            DateTime from = zoom ? (DateTime)zoomFrom : DateTime.MinValue, to = zoom ? (DateTime)zoomTo : DateTime.MaxValue;
            List<Item> bas = new List<Item>(), view = zoom ? new List<Item>() : bas;
            int ok = 0, er = 0;
            if (items != null)
            {
                foreach (object o in items)
                {
                    Item it = AsItem(o); if (it == null) continue;
                    if (!showReq && Eq(it.Kind, "Req")) continue;
                    if (!showAcct && Eq(it.Kind, "Acct")) continue;
                    bool isOk = Eq(it.Level, "OK"), isErr = Eq(it.Level, "Error");
                    if (result == 1 && !isOk) continue;
                    if (result == 2 && !isErr) continue;
                    if (result == 3 && (isOk || isErr)) continue;
                    bas.Add(it);
                    if (zoom) { if (it.Time >= from && it.Time < to) view.Add(it); else continue; }
                    if (isOk) ok++; else if (isErr) er++;
                }
            }
            Hashtable res = new Hashtable(StringComparer.OrdinalIgnoreCase);
            res["Base"] = bas; res["View"] = view; res["Ok"] = ok; res["Er"] = er;
            return res;
        }

        static void Count(Hashtable map, string key)
        {
            if (string.IsNullOrEmpty(key)) return;
            object v = map[key];
            map[key] = (v == null ? 0 : (int)v) + 1;
        }

        // Podsumowanie: pierwsze / ostatnie, udzielono / odmowa, zmiany miejsca, powiązani użytkownicy,
        // urządzenia, komputery, switche / AP, SSID, zasady i udział LAN / Wi-Fi / VPN
        public static Hashtable Summarize(IEnumerable items)
        {
            Hashtable users = new Hashtable(StringComparer.OrdinalIgnoreCase), macs = new Hashtable(StringComparer.OrdinalIgnoreCase),
                comps = new Hashtable(StringComparer.OrdinalIgnoreCase), clients = new Hashtable(StringComparer.OrdinalIgnoreCase),
                ssids = new Hashtable(StringComparer.OrdinalIgnoreCase), pols = new Hashtable(StringComparer.OrdinalIgnoreCase),
                media = new Hashtable(StringComparer.OrdinalIgnoreCase);
            int ok = 0, err = 0, changes = 0, count = 0, logOnly = 0;
            Item lastErr = null, firstIt = null, lastIt = null;
            LocState loc = new LocState();
            if (items != null)
            {
                foreach (object o in items)
                {
                    Item it = AsItem(o); if (it == null) continue;
                    count++;
                    if (firstIt == null) firstIt = it;
                    lastIt = it;
                    if (Eq(it.Level, "OK")) ok++; else if (Eq(it.Level, "Error")) { err++; lastErr = it; }
                    Count(users, it.UserDisp);
                    if (it.MacHex.Length == 12) Count(macs, it.MacHex);
                    Count(comps, it.Computer);
                    Count(clients, it.Client);
                    Count(ssids, it.Ssid);
                    string pol = Str(P(it.E), "Policy");
                    if (pol != "-") Count(pols, pol);
                    Count(media, it.Medium);
                    if (StepLocation(loc, it)) changes++;
                    if (it.Src == "Log" && Eq(it.Kind, "Resp")) logOnly++;   // odpowiedź tylko z pliku .log (bez pary w Security)
                }
            }
            Hashtable res = new Hashtable(StringComparer.OrdinalIgnoreCase);
            res["Count"] = count; res["Ok"] = ok; res["Err"] = err; res["Changes"] = changes; res["LastErr"] = lastErr; res["LogOnly"] = logOnly;
            res["First"] = firstIt == null ? (object)null : firstIt.Time; res["Last"] = lastIt == null ? (object)null : lastIt.Time;
            res["Users"] = users; res["Macs"] = macs; res["Comps"] = comps; res["Clients"] = clients;
            res["Ssids"] = ssids; res["Pols"] = pols; res["Media"] = media;
            return res;
        }

        // Liczniki paska aktywności w n równych odcinkach [t0, t1] (zdarzenie z końca zakresu - w ostatnim)
        public static Hashtable Buckets(IEnumerable items, DateTime t0, DateTime t1, int n)
        {
            int[] ok = new int[n], er = new int[n], ot = new int[n];
            double span = (double)(t1 - t0).Ticks / n;
            if (items != null && span > 0)
            {
                foreach (object o in items)
                {
                    Item it = AsItem(o); if (it == null) continue;
                    int i = (int)Math.Floor((it.Time.Ticks - t0.Ticks) / span);
                    if (i < 0) i = 0; else if (i >= n) i = n - 1;
                    if (Eq(it.Level, "OK")) ok[i]++; else if (Eq(it.Level, "Error")) er[i]++; else ot[i]++;
                }
            }
            Hashtable res = new Hashtable(StringComparer.OrdinalIgnoreCase);
            res["Ok"] = ok; res["Er"] = er; res["Ot"] = ot;
            return res;
        }
    }
}
'@

$script:HistEngine = $null
$script:HistEngineBuild = $null
if (-not ('NpsHistory.Engine' -as [type])) {
    $hb = [powershell]::Create()
    [void]$hb.AddScript('param($Source) Add-Type -TypeDefinition $Source -PassThru')
    [void]$hb.AddArgument($script:HistEngineSource)
    $script:HistEngineBuild = @{ PS = $hb; Handle = $hb.BeginInvoke() }
}

function Get-HistoryEngine {
    if ($script:HistEngine) { return $script:HistEngine }
    $t = $null
    $b = $script:HistEngineBuild
    if ($b) {
        # kompilacja w tle jeszcze trwa albo się skończyła - czekamy na nią zamiast kompilować drugi raz
        $script:HistEngineBuild = $null
        try { foreach ($x in $b.PS.EndInvoke($b.Handle)) { if ($x -and $x.FullName -eq 'NpsHistory.Engine') { $t = [type]$x } } } catch { }
        try { $b.PS.Dispose() } catch { }
    }
    if (-not $t) { $t = 'NpsHistory.Engine' -as [type] }
    if (-not $t) { $t = @(Add-Type -TypeDefinition $script:HistEngineSource -PassThru | Where-Object { $_.FullName -eq 'NpsHistory.Engine' })[0] }
    $script:HistEngine = $t
    return $t
}

# --- Normalizacja identyfikatorów i pola zdarzenia (nakładki na silnik) ----------------------
# Użytkownik: bez domeny ("CONTOSO\jan", "jan@contoso.com" -> "jan"), małymi literami;
# konta komputerów ("host/pc01.contoso.com") zostają w całości.
function Get-NormUser([string]$Value) { (Get-HistoryEngine)::NormUser($Value) }
# Komputer: "host/PC01.contoso.com", "CONTOSO\PC01$", "PC01.contoso.com" -> "pc01".
function Get-NormComputer([string]$Value) { (Get-HistoryEngine)::NormComputer($Value) }
function Test-ComputerAccount([string]$Value) { (Get-HistoryEngine)::IsComputerAccount($Value) }
# MAC zapisany jako nazwa użytkownika (MAB) - tylko gdy cała nazwa to 12 cyfr szesnastkowych.
function Get-UserMac($e) { (Get-HistoryEngine)::UserMac($e) }
function Format-Mac([string]$Hex) { (Get-HistoryEngine)::FormatMac($Hex) }
function Get-EventSrc($e) { (Get-HistoryEngine)::EventSrc($e) }
function Get-EventUserDisplay($e) { (Get-HistoryEngine)::EventUserDisplay($e) }
# Komputer zdarzenia: nazwa maszyny z dziennika Security albo konto komputera (host/..., NAZWA$)
# w nazwie użytkownika. W plikach .log "Computer-Name" to serwer NPS, nie klient.
function Get-EventComputer($e) { (Get-HistoryEngine)::EventComputer($e) }
function Get-EventSsid($e) { (Get-HistoryEngine)::EventSsid($e) }
# LAN / Wi-Fi / VPN z typu portu (tekst w języku systemu albo numer z pliku .log), SSID albo adresu IP
function Get-EventMedium($e) { (Get-HistoryEngine)::EventMedium($e) }
function Get-EventClientName($e) { (Get-HistoryEngine)::EventClientName($e) }
function Format-Span([TimeSpan]$Span) { (Get-HistoryEngine)::FormatSpan($Span) }

# --- Zapytanie i dopasowanie ---------------------------------------------------------------
function New-HistoryQuery([string]$Mode, [string]$Value, [bool]$Partial) {
    $v = ([string]$Value).Trim()
    $q = @{ Mode = $Mode; Value = $v; Partial = $Partial; Hex = ''; Norm = ''; Label = ''; Valid = $false }
    switch ($Mode) {
        'Mac' {
            $q.Hex = Get-HexMac $v
            if ($q.Hex.Length -ne 12) { $q.Partial = $true }
            $q.Valid = $q.Hex.Length -ge 4
            $q.Label = "MAC $(Format-Mac $q.Hex)"
        }
        'User' {
            $q.Norm = Get-NormUser $v
            $q.Valid = [bool]$q.Norm
            $q.Label = "użytkownik $v"
        }
        'Computer' {
            $q.Norm = Get-NormComputer $v
            $q.Valid = [bool]$q.Norm
            $q.Label = "komputer $($q.Norm.ToUpperInvariant())"
        }
    }
    return $q
}
function Test-HistoryMatch($e, $Q) { (Get-HistoryEngine)::Match($e, $Q) }

# Wartości do filtra po stronie serwera (zapytanie XML z EventData/Data = ...): dokładne
# zapisy MAC w formatach spotykanych u producentów. Dla użytkownika i komputera nie da się
# przewidzieć zapisu (wielkość liter, domena, host/), więc tam skanujemy zakres i filtrujemy tutaj.
function Get-HistoryDataValues($Q) {
    if ($Q.Mode -ne 'Mac' -or $Q.Partial -or $Q.Hex.Length -ne 12) { return $null }
    $h = $Q.Hex.ToUpperInvariant()
    $p = @(0..5 | ForEach-Object { $h.Substring($_ * 2, 2) })
    $forms = @(
        ($p -join '-'), ($p -join ':'), $h,               # Cisco / RFC 3580, Linux / UniFi, Aruba / HP
        ('{0}{1}.{2}{3}.{4}{5}' -f $p),                   # Cisco IOS (aabb.ccdd.eeff)
        ('{0}{1}{2}-{3}{4}{5}' -f $p),                    # HP ProCurve (aabbcc-ddeeff)
        ('{0}{1}-{2}{3}-{4}{5}' -f $p),                   # Huawei / H3C / Comware (aabb-ccdd-eeff)
        ($p -join '.')                                    # aa.bb.cc.dd.ee.ff
    )
    $all = New-Object System.Collections.Generic.List[string]
    foreach ($f in $forms) { $all.Add($f); $all.Add($f.ToLowerInvariant()) }
    return [string[]]@($all | Select-Object -Unique)
}
# --- Elementy i wiersze osi czasu ----------------------------------------------------------
# Element: zdarzenie z obu źródeł z policzonymi raz polami (miejsce w sieci, użytkownik, opis...).
function New-HistoryItem($e, [string]$Src, $Q) { (Get-HistoryEngine)::NewItem($e, $Src, $Q) }
# Scalenie Security + .log (odpowiedź widoczna w obu - raz, "Security + log"), rosnąco po czasie.
function Merge-HistoryItems($EvItems, $LogItems) { , (Get-HistoryEngine)::Merge($EvItems, $LogItems) }
# Wiersze: zdarzenia (z grupowaniem powtórzeń), znaczniki zmian sieci / urządzenia / użytkownika,
# przerwy dłuższe niż $script:HistGapMinutes i nagłówki dni. Zwraca @{ Rows; Changes }.
function Build-HistoryRows($Items, $Q, [hashtable]$Opt) {
    (Get-HistoryEngine)::BuildRows($Items, $Q, [bool]$Opt.Group, [bool]$Opt.Markers, [bool]$Opt.Newest, [int]$script:HistGapMinutes, $script:Colors, $script:MediumBg, $script:PlCulture)
}

# --- Okno historii -------------------------------------------------------------------------
function Get-HistoryRange($H) {
    $u = $H.Ui
    $now = Get-Date
    switch ($u.hRange.SelectedIndex) {
        1 { return @($now.AddHours(-1), $now) }
        2 { return @($now.AddHours(-24), $now) }
        3 { return @($now.AddDays(-7), $now) }
        4 { return @($now.AddDays(-30), $now) }
        5 { return @($now.AddDays(-90), $now) }
    }
    $fmt = 'yyyy-MM-dd HH:mm'
    $inv = [Globalization.CultureInfo]::InvariantCulture
    $s = [datetime]::ParseExact($u.hFrom.Text.Trim(), $fmt, $inv)
    $e = [datetime]::ParseExact($u.hTo.Text.Trim(), $fmt, $inv)
    if ($e -le $s) { throw 'Data "DO" musi być późniejsza niż "OD".' }
    return @($s, $e)
}

function Set-HistoryStatus($H, [string]$Text, [string]$Level = 'Info') {
    $H.Ui.hStatus.Text = $Text
    $H.Ui.hStatus.Foreground = Get-Brush $script:Colors[$Level]
}

function Set-HistoryState($H, [string]$Text, [string]$Level) {
    $H.Ui.hState.Text = $Text
    $H.Ui.hDot.Fill = Get-Brush $script:Colors[$Level]
}

function Show-HistoryWindow {
    param([string]$Mode = 'Mac', [string]$Value = '', [int]$RangeIndex = -1, [string]$From = '', [string]$To = '')

    [xml]$hx = $script:HistXaml
    # Te same style i kolory co w oknie głównym - kopiujemy jego Window.Resources do tego okna.
    $res = $xaml.DocumentElement.SelectSingleNode("*[local-name()='Window.Resources']")
    [void]$hx.DocumentElement.PrependChild($hx.ImportNode($res, $true))
    $hw = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $hx))

    $id = ++$script:HistSeq
    $H = @{
        Id = $id; Win = $hw; Ui = @{}; Q = $null; Suspend = $true
        Items = New-Object System.Collections.Generic.List[object]
        Base = New-Object System.Collections.Generic.List[object]
        ZoomFrom = $null; ZoomTo = $null; Buckets = @(); Busy = $false
        PS = @{}; Handles = @{}; LoadStarted = $null; Range = $null; Max = 0; Prefilter = $false; RangeText = ''
        TabMode = $false; ReqLoaded = $true; BaseVer = 0; StripCache = $null
    }
    $hx.SelectNodes('//*[@Name]') | ForEach-Object {
        $c = $hw.FindName($_.Name)
        $H.Ui[$_.Name] = $c
        if ($c -is [System.Windows.FrameworkElement]) { $c.Tag = $id }
    }
    $script:HistWins[$id] = $H
    $u = $H.Ui

    Set-WindowFit $hw
    try { $hw.Owner = $win } catch { }

    $u.hMode.SelectedIndex = [Math]::Max(0, @('Mac', 'User', 'Computer').IndexOf($Mode))
    $u.hValue.Text = $Value
    $hasTabs = ($script:Ctx.Ev.All.Count + $script:Ctx.Log.All.Count) -gt 0
    if ($RangeIndex -lt 0) { $RangeIndex = $(if ($hasTabs) { 0 } else { 3 }) }
    $u.hRange.SelectedIndex = $RangeIndex
    $u.hFrom.Text = $From; $u.hTo.Text = $To
    Update-CustomRange $u.hRange $u.hFrom $u.hTo
    $u.hSrcLog.IsChecked = ($script:Ctx.Log.All.Count -gt 0)
    $srv = $(if ($ui.txtServer.Text.Trim()) { $ui.txtServer.Text.Trim() } else { $env:COMPUTERNAME })
    $u.hSubtitle.Text = "Źródła przy wczytywaniu: dziennik Security na $srv  ·  pliki .log: $($ui.txtLogPath.Text.Trim()) ($($ui.txtLogMask.Text.Trim()))"

    $H.Timer = New-Object System.Windows.Threading.DispatcherTimer
    $H.Timer.Interval = [TimeSpan]::FromMilliseconds(300)
    $H.Timer.Tag = $id
    $H.Timer.Add_Tick({ Complete-HistoryLoad $script:HistWins[[int]$this.Tag] })

    $u.hLoad.Add_Click({ Start-HistoryLoad $script:HistWins[[int]$this.Tag] })
    $u.hValue.Add_KeyDown({ if ($_.Key -eq 'Return') { Start-HistoryLoad $script:HistWins[[int]$this.Tag] } })
    $u.hRange.Add_SelectionChanged({ $hh = $script:HistWins[[int]$this.Tag]; Update-CustomRange $hh.Ui.hRange $hh.Ui.hFrom $hh.Ui.hTo })
    $u.hResult.Add_SelectionChanged({ Update-HistoryView $script:HistWins[[int]$this.Tag] })
    foreach ($k in 'hGroup', 'hMarkers', 'hAcct', 'hReq', 'hNewest') {
        $u[$k].Add_Checked({ Update-HistoryView $script:HistWins[[int]$this.Tag] })
        $u[$k].Add_Unchecked({ Update-HistoryView $script:HistWins[[int]$this.Tag] })
    }
    # Access-Request / Challenge z plików .log są wczytywane tylko, gdy są pokazywane - po
    # zaznaczeniu trzeba je doczytać.
    $u.hReq.Add_Checked({
            $hh = $script:HistWins[[int]$this.Tag]
            if ($hh -and $hh.Q -and -not $hh.ReqLoaded -and -not $hh.Busy) { Start-HistoryLoad $hh }
        })
    $u.hList.Add_SelectionChanged({ $hh = $script:HistWins[[int]$this.Tag]; Show-HistoryDetails $hh $hh.Ui.hList.SelectedItem })
    $u.hList.ContextMenu = New-HistoryMenu $id
    $u.hZoomClear.Add_Click({ Set-HistoryZoom $script:HistWins[[int]$this.Tag] $null $null })
    $u.hStrip.Add_MouseLeftButtonUp({ Select-HistoryBucket $script:HistWins[[int]$this.Tag] $_.GetPosition($this).X })
    $u.hStripHost.Add_SizeChanged({ Update-HistoryStrip $script:HistWins[[int]$this.Tag] })
    # Podsumowanie i powiązane zajmują najwyżej ~30% wysokości okna (dalej przewijane) - na niskim
    # ekranie zostaje miejsce na oś czasu i szczegóły.
    $hw.Add_SizeChanged({ $hh = $script:HistWins[[int]$this.Tag]; if ($hh) { $hh.Ui.hSummary.MaxHeight = [Math]::Max(90, [Math]::Floor($this.ActualHeight * 0.3)) } })
    $u.hCopy.Add_Click({ $hh = $script:HistWins[[int]$this.Tag]; if ($hh.Ui.hDetails.Text) { [System.Windows.Clipboard]::SetText($hh.Ui.hDetails.Text) } })
    $u.hCopyAll.Add_Click({ $hh = $script:HistWins[[int]$this.Tag]; $t = Get-HistoryText $hh; if ($t) { [System.Windows.Clipboard]::SetText($t); Set-HistoryStatus $hh 'Skopiowano oś czasu do schowka.' 'OK' } })
    $u.hExport.Add_Click({ Export-History $script:HistWins[[int]$this.Tag] })
    $hw.Tag = $id
    $hw.Add_Closed({
            $hh = $script:HistWins[[int]$this.Tag]
            if ($hh) {
                $hh.Timer.Stop()
                Stop-HistoryWorkers $hh
                $script:HistWins.Remove([int]$this.Tag)
            }
        })

    $H.Suspend = $false
    $hw.Show()
    if ($Value.Trim()) { Start-HistoryLoad $H }
    else { [void]$u.hValue.Focus() }
    return $H
}

function Open-HistoryFromEvent($e, [string]$Mode, [int]$RangeIndex = -1, [string]$From = '', [string]$To = '') {
    if (-not $e) { return }
    switch ($Mode) {
        'Mac' {
            $hex = $(if (([string]$e.MacHex).Length -eq 12) { $e.MacHex } else { Get-UserMac $e })
            if (-not $hex -or $hex.Length -ne 12) { Show-Msg 'To zdarzenie nie zawiera adresu MAC klienta.' 'Information'; return }
            [void](Show-HistoryWindow -Mode 'Mac' -Value (Format-Mac $hex) -RangeIndex $RangeIndex -From $From -To $To)
        }
        'User' {
            $usr = Get-EventUserDisplay $e
            if (-not $usr) { Show-Msg 'To zdarzenie nie zawiera nazwy użytkownika.' 'Information'; return }
            [void](Show-HistoryWindow -Mode 'User' -Value $usr -RangeIndex $RangeIndex -From $From -To $To)
        }
        'Computer' {
            $comp = Get-EventComputer $e
            if (-not $comp) { Show-Msg 'To zdarzenie nie wskazuje komputera (ani nazwy maszyny, ani konta komputera host/...).' 'Information'; return }
            [void](Show-HistoryWindow -Mode 'Computer' -Value $comp.ToUpperInvariant() -RangeIndex $RangeIndex -From $From -To $To)
        }
    }
}

# Domyślnie historia urządzenia (MAC), a gdy zdarzenie nie ma MAC - historia użytkownika.
function Open-HistoryDefault($e) {
    if (-not $e) { return }
    $hex = $(if (([string]$e.MacHex).Length -eq 12) { $e.MacHex } else { Get-UserMac $e })
    if ($hex -and $hex.Length -eq 12) { Open-HistoryFromEvent $e 'Mac' } else { Open-HistoryFromEvent $e 'User' }
}

# --- Wczytywanie historii w tle ------------------------------------------------------------
# Loader (dziennik Security / pliki .log) albo dane z zakładek, a potem dopasowanie i budowa
# elementów osi czasu (silnik C#) - wszystko w runspace w tle, okna się nie zatrzymują.
# Zwraca jeden obiekt: elementy, liczbę zdarzeń ze źródła i podsumowanie loadera plików .log.
$script:HistWorker = {
    param($Engine, [string]$Loader, [object[]]$LoaderArgs, [object[]]$Objects, $Q, [string]$Src)
    # Każdy błąd przerywa wczytywanie (EndInvoke go zgłosi), zamiast po cichu dać pustą historię.
    # Bez $(...): wynik z jednym obiektem (jedno zdarzenie, samo podsumowanie loadera) nie może
    # przestać być tablicą.
    try {
        if ($Loader) { $source = @(& ([scriptblock]::Create($Loader)) @LoaderArgs) } else { $source = $Objects }
        $r = $Engine::BuildItems([object[]]$source, $Q, $Src)
    }
    catch { throw }
    [pscustomobject]@{ HistWorker = $true; Items = $r.Items; Raw = $r.Raw; Summary = $r.Summary }
}

function New-HistoryWorker($Loader, $LoaderArgs, $Objects, $Q, $Src) {
    $ps = [powershell]::Create()
    [void]$ps.AddScript([string]$script:HistWorker)
    [void]$ps.AddArgument((Get-HistoryEngine))
    [void]$ps.AddArgument([string]$Loader)
    [void]$ps.AddArgument($LoaderArgs)
    [void]$ps.AddArgument($Objects)
    [void]$ps.AddArgument($Q)
    [void]$ps.AddArgument([string]$Src)
    return $ps
}

# Zatrzymanie bez czekania: PowerShell.Stop() czeka, aż Get-WinEvent skończy bieżące skanowanie
# dziennika (przy filtrze MAC i dużym / zdalnym dzienniku - minuty), a w tym czasie wszystkie okna
# stoją. BeginStop wraca od razu; zatrzymane instancje sprząta zegar.
$script:HistStopping = New-Object System.Collections.Generic.List[object]
$script:HistReaper = New-Object System.Windows.Threading.DispatcherTimer
$script:HistReaper.Interval = [TimeSpan]::FromSeconds(2)
$script:HistReaper.Add_Tick({
        foreach ($ps in $script:HistStopping.ToArray()) {
            if ([string]$ps.InvocationStateInfo.State -in 'Stopped', 'Completed', 'Failed', 'NotStarted') {
                try { $ps.Dispose() } catch { }
                [void]$script:HistStopping.Remove($ps)
            }
        }
        if (-not $script:HistStopping.Count) { $script:HistReaper.Stop() }
    })

function Stop-HistoryWorkers($H) {
    foreach ($ps in @($H.PS.Values)) {
        try {
            if ([string]$ps.InvocationStateInfo.State -in 'Running', 'Stopping') {
                if ([string]$ps.InvocationStateInfo.State -eq 'Running') { [void]$ps.BeginStop($null, $null) }
                $script:HistStopping.Add($ps)
            }
            else { $ps.Dispose() }
        }
        catch { }
    }
    $H.PS = @{}; $H.Handles = @{}
    if ($script:HistStopping.Count) { $script:HistReaper.Start() }
}

function Start-HistoryLoad($H) {
    if (-not $H -or $H.Busy) { return }
    $u = $H.Ui
    $mode = @('Mac', 'User', 'Computer')[[Math]::Max(0, $u.hMode.SelectedIndex)]
    $q = New-HistoryQuery $mode $u.hValue.Text ([bool]$u.hPartial.IsChecked)
    if (-not $q.Valid) {
        $what = $(switch ($mode) { 'Mac' { 'adres MAC (co najmniej 4 cyfry szesnastkowe)' } 'User' { 'nazwę użytkownika' } default { 'nazwę komputera' } })
        Set-HistoryStatus $H "Wpisz $what." 'Warn'
        return
    }

    # Najpierw wszystkie sprawdzenia: nieudany start nie może zmienić tytułu ani zapytania osi
    # czasu, która jest na ekranie (kopiowanie / eksport podpisałyby ją cudzą nazwą).
    $tab = ($u.hRange.SelectedIndex -eq 0)
    $nEv = $script:Ctx.Ev.All.Count; $nLog = $script:Ctx.Log.All.Count
    if ($tab) {
        if (($nEv + $nLog) -eq 0) {
            Set-HistoryStatus $H 'W zakładkach nie ma jeszcze danych - wybierz zakres czasu (np. Ostatnie 7 dni) i kliknij Pokaż historię.' 'Warn'
            Set-HistoryState $H 'Brak danych' 'Warn'
            return
        }
    }
    else {
        try { $range = Get-HistoryRange $H }
        catch { Set-HistoryStatus $H "Nieprawidłowy zakres czasu: $($_.Exception.Message)" 'Error'; return }
        $useEv = [bool]$u.hSrcEv.IsChecked
        $useLog = [bool]$u.hSrcLog.IsChecked
        if (-not $useEv -and -not $useLog) { Set-HistoryStatus $H 'Zaznacz co najmniej jedno źródło (dziennik Security albo pliki .log).' 'Warn'; return }
        $paths = @($ui.txtLogPath.Text -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        if ($useLog -and -not $paths.Count) { Set-HistoryStatus $H 'Podaj lokalizację plików .log w zakładce Pliki logów RADIUS.' 'Warn'; return }
        $mask = $ui.txtLogMask.Text.Trim(); if (-not $mask) { $mask = 'IN*.log' }
    }

    $H.Q = $q
    $H.ZoomFrom = $null; $H.ZoomTo = $null
    $u.hTitle.Text = "Historia: $($q.Label)"
    $H.Win.Title = "Historia - $($q.Label)"
    $H.TabMode = $tab
    $H.PS = @{}; $H.Handles = @{}
    $H.Prefilter = $false
    if ($tab) {
        $H.Range = $null; $H.Max = 0
        $H.RangeText = 'dane wczytane w zakładkach'
        $H.ReqLoaded = $true
        # Kopia danych zakładki (ToArray, nie @(...): @() na liście z New-Object wywala binder
        # Windows PowerShell 5.1 - "Argument types do not match")
        if ($nEv) { $H.PS['Ev'] = New-HistoryWorker '' $null $script:Ctx.Ev.All.ToArray() $q 'Ev' }
        if ($nLog) { $H.PS['Log'] = New-HistoryWorker '' $null $script:Ctx.Log.All.ToArray() $q 'Log' }
        $msg = "Szukanie w danych zakładek: $($q.Label)..."
    }
    else {
        $H.Max = Get-MaxValue $u.hMax 20000
        $H.Range = $range
        $H.RangeText = Format-Range $range
        # Access-Request / Challenge (kilkanaście na jedno logowanie EAP) wczytujemy tylko, gdy są
        # pokazywane - inaczej zajmowałyby limit zdarzeń i starsza historia by przepadła.
        $skipReq = -not [bool]$u.hReq.IsChecked
        $H.ReqLoaded = -not ($useLog -and $skipReq)
        if ($useEv) {
            $dv = Get-HistoryDataValues $q
            $H.Prefilter = [bool]$dv
            $H.PS['Ev'] = New-HistoryWorker ([string]$script:EvLoader) @($ui.txtServer.Text.Trim(), $range[0], $range[1], $H.Max, $dv) $null $q 'Ev'
        }
        if ($useLog) {
            $mv = $(if ($q.Mode -eq 'Mac') { $q.Hex } else { $q.Norm })
            $H.PS['Log'] = New-HistoryWorker ([string]$script:LogLoader) @([string[]]$paths, $mask, $range[0], $range[1], $H.Max, $q.Mode, $mv, [bool]$q.Partial, $skipReq) $null $q 'Log'
        }
        $msg = "Wczytywanie historii: $($q.Label), $($H.RangeText)$(if ($H.Prefilter) { ' (filtr MAC po stronie serwera)' } else { ' - skanowanie zakresu, może potrwać' })..."
    }
    foreach ($k in @($H.PS.Keys)) { $H.Handles[$k] = $H.PS[$k].BeginInvoke() }
    $H.Busy = $true
    $H.LoadStarted = Get-Date
    $u.hLoad.IsEnabled = $false
    $u.hPb.Visibility = 'Visible'; $u.hPb.IsIndeterminate = $true
    Set-HistoryStatus $H $msg 'Warn'
    Set-HistoryState $H 'Wczytywanie...' 'Warn'
    $H.Timer.Start()
}

function Complete-HistoryLoad($H) {
    if (-not $H -or -not $H.Busy) { return }
    foreach ($hnd in $H.Handles.Values) { if (-not $hnd.IsCompleted) { return } }
    $H.Timer.Stop()
    $evItems = $null; $logItems = $null; $logSum = $null
    $notes = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]
    $rawEv = 0
    $hadEv = $H.PS.ContainsKey('Ev'); $hadLog = $H.PS.ContainsKey('Log')
    try {
        foreach ($src in @($H.PS.Keys)) {
            $ps = $H.PS[$src]
            try {
                $w = $null
                foreach ($r in $ps.EndInvoke($H.Handles[$src])) { if ($r -and $r.PSObject.Properties['HistWorker']) { $w = $r } }
                if (-not $w) { throw 'wczytywanie nie zwróciło wyniku' }
                if ($src -eq 'Ev') { $evItems = $w.Items; $rawEv = $w.Raw } else { $logItems = $w.Items; $logSum = $w.Summary }
            }
            catch {
                $msg = $_.Exception.Message
                if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
                $errors.Add($(if ($src -eq 'Ev') { "dziennik Security: $msg" } else { "pliki .log: $msg" }))
            }
            finally { try { $ps.Dispose() } catch { } }
        }
        $nEv = [int]$evItems.Count; $nLog = [int]$logItems.Count
        if ($logSum) {
            if ($logSum.PSObject.Properties['Errors'] -and $logSum.Errors.Count) { foreach ($x in $logSum.Errors) { $errors.Add("pliki .log: $x") } }
            if ($logSum.PSObject.Properties['Files'] -and $logSum.Files -eq 0) { $notes.Add('żaden plik .log nie był modyfikowany w tym zakresie') }
            if ($logSum.PSObject.Properties['Dropped'] -and $logSum.Dropped -gt 0) {
                $notes.Add("z plików .log pokazano $($H.Max) najnowszych pasujących wpisów, starszych ($($logSum.Dropped)) nie wczytano - zwiększ MAKS. ZDARZEŃ albo zawęź zakres")
            }
        }
        $H.Items = Merge-HistoryItems $evItems $logItems
        if (-not $H.TabMode) {
            if ($hadEv) {
                if ($H.Max -gt 0 -and $rawEv -ge $H.Max) {
                    $notes.Add($(if ($H.Prefilter) { "osiągnięto limit $($H.Max) zdarzeń z dziennika - pokazano najnowsze" } else { "przeszukano tylko $($H.Max) najnowszych zdarzeń z dziennika - zwiększ MAKS. ZDARZEŃ albo zawęź zakres" }))
                }
                if ($H.Prefilter) {
                    $st = (Get-HistoryEngine)::Summarize($H.Items)
                    $logOnly = [int]$st.LogOnly
                    if ($nEv -eq 0) { $notes.Add('w dzienniku szukano dokładnego zapisu MAC (AA-BB-CC-DD-EE-FF, aabbccddeeff, aabb.ccdd.eeff, aabb-ccdd-eeff...) - jeśli switch zapisuje go inaczej, zaznacz Dopasowanie częściowe') }
                    elseif ($logOnly -gt 0) { $notes.Add("$logOnly odpowiedzi (Accept / Reject) z plików .log nie ma w dzienniku Security - to inny serwer NPS albo switch zapisuje MAC inaczej (wtedy zaznacz Dopasowanie częściowe)") }
                }
            }
        }
        Update-HistorySummary $H
        Update-HistoryView $H
        if ($H.TabMode) {
            $txt = "Z danych zakładek: dziennik Security $nEv, pliki .log $nLog."
            if ($H.Items.Count -eq 0) { $txt += ' Nic nie znaleziono - wybierz dłuższy zakres czasu, żeby przeszukać dziennik / logi.' }
        }
        else {
            $secs = [Math]::Round(((Get-Date) - $H.LoadStarted).TotalSeconds, 1)
            $txt = "Wczytano w $secs s: dziennik Security $nEv$(if ($hadEv) { '' } else { ' (pominięty)' }), pliki .log $nLog$(if ($hadLog) { '' } else { ' (pominięte)' })."
        }
        if ($notes.Count) { $txt += ' Uwaga: ' + ($notes -join '; ') + '.' }
        if ($errors.Count) { $txt += ' BŁĘDY: ' + ($errors -join ' | ') }
        $lvl = $(if ($errors.Count) { 'Error' } elseif ($notes.Count -or $H.Items.Count -eq 0) { 'Warn' } else { 'OK' })
        Set-HistoryStatus $H $txt $lvl
        Set-HistoryState $H $(if ($errors.Count) { 'Błąd wczytywania' } else { "$($H.Items.Count) zdarzeń" }) $lvl
    }
    finally {
        $H.PS = @{}; $H.Handles = @{}
        $H.Busy = $false
        $H.Ui.hLoad.IsEnabled = $true
        $H.Ui.hPb.Visibility = 'Hidden'; $H.Ui.hPb.IsIndeterminate = $false
    }
    # "Access-Request / Challenge" zaznaczone w trakcie wczytywania - doczytaj je
    if (-not $H.ReqLoaded -and $H.Ui.hReq.IsChecked -and -not $errors.Count) { Start-HistoryLoad $H }
}

# Filtry widoku (wynik, żądania/accounting, zawężenie czasu) i przebudowa osi czasu.
function Update-HistoryView($H) {
    if (-not $H -or $H.Suspend) { return }
    $u = $H.Ui
    $f = (Get-HistoryEngine)::Filter($H.Items, [int]$u.hResult.SelectedIndex, [bool]$u.hReq.IsChecked, [bool]$u.hAcct.IsChecked, $H.ZoomFrom, $H.ZoomTo)
    $H.Base = $f.Base
    $H.BaseVer++
    $view = $f.View
    $H.View = $view
    if (-not $H.Q) { return }
    $b = Build-HistoryRows $view $H.Q @{ Group = [bool]$u.hGroup.IsChecked; Markers = [bool]$u.hMarkers.IsChecked; Newest = [bool]$u.hNewest.IsChecked }
    $H.Rows = $b.Rows
    $u.hList.ItemsSource = $b.Rows
    $u.hCounts.Text = "W widoku: $($view.Count) / $($H.Items.Count)    Udzielono: $($f.Ok)    Odmowa: $($f.Er)    Zmian miejsca w sieci: $($b.Changes)"
    if ($H.ZoomFrom) {
        $u.hZoomText.Text = "Widok zawężony: $($H.ZoomFrom.ToString('yyyy-MM-dd HH:mm')) - $($H.ZoomTo.ToString('yyyy-MM-dd HH:mm'))"
        $u.hZoomClear.Visibility = 'Visible'
    }
    else {
        $u.hZoomText.Text = "Widok: $($H.RangeText)"
        $u.hZoomClear.Visibility = 'Collapsed'
    }
    $u.hDetails.Text = $(if ($view.Count) { 'Zaznacz wiersz na osi czasu.' } else { 'Brak zdarzeń w widoku.' })
    Update-HistoryStrip $H
}

function Set-HistoryZoom($H, $From, $To) {
    if (-not $H) { return }
    $H.ZoomFrom = $From
    $H.ZoomTo = $To
    Update-HistoryView $H
}

function New-HistoryChip([string]$Text, [string]$Fg = '#E8EAF0', [string]$Bg = '#272A36', [string]$Tip = '') {
    $b = New-Object System.Windows.Controls.Border
    $b.Background = Get-Brush $Bg
    $b.CornerRadius = New-Object System.Windows.CornerRadius 10
    $b.Padding = New-Object System.Windows.Thickness 10, 0, 10, 0
    $b.Height = 24
    $b.Margin = New-Object System.Windows.Thickness 0, 0, 8, 8
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $Text
    $t.FontSize = 12
    $t.VerticalAlignment = 'Center'
    $t.Foreground = Get-Brush $Fg
    $b.Child = $t
    if ($Tip) { $b.ToolTip = $Tip }
    return $b
}

# Kafelki podsumowania i "powiązane": użytkownicy / urządzenia / komputery (klik = ich historia),
# switche / AP, SSID i zasady sieciowe.
function Update-HistorySummary($H) {
    $u = $H.Ui
    $u.hChips.Children.Clear()
    $u.hRelated.Children.Clear()
    $items = $H.Items
    $s = (Get-HistoryEngine)::Summarize($items)
    $H.Stats = $s
    if (-not $items.Count) {
        [void]$u.hChips.Children.Add((New-HistoryChip 'Brak zdarzeń dla wybranego zapytania w tym zakresie' '#FBBF24'))
        return
    }
    $first = $s.First; $last = $s.Last; $ok = $s.Ok; $err = $s.Err; $lastErr = $s.LastErr; $changes = $s.Changes
    $users = $s.Users; $macs = $s.Macs; $comps = $s.Comps; $clients = $s.Clients; $ssids = $s.Ssids; $pols = $s.Pols; $media = $s.Media
    $c = $u.hChips.Children
    [void]$c.Add((New-HistoryChip "Pierwsze: $($first.ToString('yyyy-MM-dd HH:mm:ss'))"))
    [void]$c.Add((New-HistoryChip "Ostatnie: $($last.ToString('yyyy-MM-dd HH:mm:ss'))"))
    [void]$c.Add((New-HistoryChip "Zdarzeń: $($items.Count)"))
    [void]$c.Add((New-HistoryChip "Udzielono: $ok" '#4ADE80' '#1F3A2C'))
    [void]$c.Add((New-HistoryChip "Odmowa: $err" $(if ($err) { '#F87171' } else { '#8A90A2' }) $(if ($err) { '#3F2226' } else { '#272A36' })))
    if ($lastErr) {
        $r = [string]$lastErr.E.Reason; if ($r.Length -gt 60) { $r = $r.Substring(0, 60) + '...' }
        [void]$c.Add((New-HistoryChip "Ostatnia odmowa: $($lastErr.Time.ToString('yyyy-MM-dd HH:mm')) (kod $($lastErr.E.ReasonCode))" '#F87171' '#3F2226' $r))
    }
    [void]$c.Add((New-HistoryChip "Zmian miejsca w sieci: $changes" '#93C5FD' '#1F2A40'))
    if ($media.Count) {
        $mt = ($media.Keys | Sort-Object { -$media[$_] } | ForEach-Object { "$_ $([Math]::Round(100 * $media[$_] / $items.Count))%" }) -join ' · '
        [void]$c.Add((New-HistoryChip "Sieć: $mt"))
    }

    $q = $H.Q
    $addRow = {
        param([string]$Caption, [hashtable]$Map, [string]$PivotMode, [string]$Skip)
        if (-not $Map.Count) { return }
        $wp = New-Object System.Windows.Controls.WrapPanel
        $wp.Margin = New-Object System.Windows.Thickness 0, 0, 0, 2
        $cap = New-Object System.Windows.Controls.TextBlock
        $cap.Text = $Caption
        $cap.Width = 150
        $cap.Style = $H.Win.FindResource('Caption')
        $cap.Margin = New-Object System.Windows.Thickness 0, 5, 8, 0
        [void]$wp.Children.Add($cap)
        $keys = @($Map.Keys | Sort-Object { -$Map[$_] })
        foreach ($k in ($keys | Select-Object -First 8)) {
            $label = $(if ($PivotMode -eq 'Mac') { Format-Mac $k } elseif ($PivotMode -eq 'Computer') { $k.ToUpperInvariant() } else { $k })
            if ($PivotMode -and $k -ne $Skip) {
                $btn = New-Object System.Windows.Controls.Button
                $btn.Content = "$label  ($($Map[$k]))"
                $btn.Style = $H.Win.FindResource('BtnSmall')
                $btn.Margin = New-Object System.Windows.Thickness 0, 0, 8, 8
                $btn.ToolTip = "Pokaż historię: $label"
                $btn.Tag = @{ Id = $H.Id; Mode = $PivotMode; Value = $label }
                $btn.Add_Click({ Open-HistoryPivot $this.Tag })
                [void]$wp.Children.Add($btn)
            }
            else { [void]$wp.Children.Add((New-HistoryChip "$label  ($($Map[$k]))")) }
        }
        if ($keys.Count -gt 8) { [void]$wp.Children.Add((New-HistoryChip "+$($keys.Count - 8) więcej" '#8A90A2')) }
        [void]$H.Ui.hRelated.Children.Add($wp)
    }
    $skipUser = $(if ($q.Mode -eq 'User') { ($users.Keys | Where-Object { (Get-NormUser $_) -eq $q.Norm } | Select-Object -First 1) } else { '' })
    $skipMac = $(if ($q.Mode -eq 'Mac') { $q.Hex } else { '' })
    $skipComp = $(if ($q.Mode -eq 'Computer') { $q.Norm } else { '' })
    & $addRow 'UŻYTKOWNICY' $users 'User' $skipUser
    & $addRow 'URZĄDZENIA (MAC)' $macs 'Mac' $skipMac
    & $addRow 'KOMPUTERY' $comps 'Computer' $skipComp
    & $addRow 'SWITCHE / AP' $clients '' ''
    & $addRow 'SSID' $ssids '' ''
    & $addRow 'ZASADY SIECIOWE' $pols '' ''
}

function Open-HistoryPivot($Tag) {
    $H = $script:HistWins[[int]$Tag.Id]
    $ri = -1; $from = ''; $to = ''
    if ($H) { $ri = $H.Ui.hRange.SelectedIndex; $from = $H.Ui.hFrom.Text; $to = $H.Ui.hTo.Text }
    [void](Show-HistoryWindow -Mode $Tag.Mode -Value $Tag.Value -RangeIndex $ri -From $from -To $to)
}

# Pasek aktywności: słupki (zielone = udzielono, żółte = inne, czerwone = odmowy) w równych
# odcinkach czasu całego wczytanego zakresu. Klik w słupek zawęża widok do tego odcinka.
function Update-HistoryStrip($H) {
    if (-not $H) { return }
    $u = $H.Ui
    $cv = $u.hStrip
    $cv.Children.Clear()
    $H.Buckets = @()
    $items = $H.Base
    $w = $u.hStripHost.ActualWidth - 8
    $ht = $u.hStripHost.ActualHeight - 8
    $u.hStripFrom.Text = ''; $u.hStripTo.Text = ''
    if (-not $items -or $items.Count -eq 0 -or $w -lt 40 -or $ht -lt 10) { return }

    $t0 = $items[0].Time; $t1 = $items[$items.Count - 1].Time
    if ($H.Range -and -not $H.TabMode) { $t0 = $H.Range[0]; $t1 = $H.Range[1] }
    if (($t1 - $t0).TotalMinutes -lt 1) { $t0 = $t0.AddMinutes(-30); $t1 = $t1.AddMinutes(30) }
    # Stała liczba odcinków: liczniki liczymy raz na dane widoku, a zmiana rozmiaru okna
    # tylko przerysowuje słupki (przy dziesiątkach tysięcy zdarzeń liczenie trwa).
    $n = 120
    $span = ($t1 - $t0).Ticks / $n
    $key = "$($H.BaseVer)|$($t0.Ticks)|$($t1.Ticks)"
    if (-not $H.StripCache -or $H.StripCache.Key -ne $key) {
        $bk = (Get-HistoryEngine)::Buckets($items, $t0, $t1, $n)
        $H.StripCache = @{ Key = $key; Ok = $bk.Ok; Er = $bk.Er; Ot = $bk.Ot }
    }
    $ok = $H.StripCache.Ok; $er = $H.StripCache.Er; $ot = $H.StripCache.Ot
    $mx = 1
    for ($i = 0; $i -lt $n; $i++) { $s = $ok[$i] + $er[$i] + $ot[$i]; if ($s -gt $mx) { $mx = $s } }
    $bw = $w / $n
    $buckets = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $n; $i++) {
        $from = $t0.AddTicks([int64]($span * $i))
        # Ostatni odcinek włącznie z końcem - inaczej klik w niego pomijałby najnowsze zdarzenie
        $to = $(if ($i -eq $n - 1) { $t1.AddTicks(1) } else { $t0.AddTicks([int64]($span * ($i + 1))) })
        $buckets.Add(@{ From = $from; To = $to })
        $tot = $ok[$i] + $er[$i] + $ot[$i]
        if (-not $tot) { continue }
        $y = $ht + 4
        foreach ($seg in @(@($ok[$i], '#4ADE80'), @($ot[$i], '#FBBF24'), @($er[$i], '#F87171'))) {
            if (-not $seg[0]) { continue }
            $sh = [Math]::Max(2, $ht * $seg[0] / $mx)
            $rc = New-Object System.Windows.Shapes.Rectangle
            $rc.Width = [Math]::Max(1, $bw - 2)
            $rc.Height = $sh
            $rc.Fill = Get-Brush $seg[1]
            $rc.RadiusX = 1; $rc.RadiusY = 1
            [System.Windows.Controls.Canvas]::SetLeft($rc, 4 + $bw * $i)
            [System.Windows.Controls.Canvas]::SetTop($rc, $y - $sh)
            $rc.ToolTip = "$($from.ToString('yyyy-MM-dd HH:mm')) - $($to.ToString('yyyy-MM-dd HH:mm'))`nudzielono: $($ok[$i])   odmowa: $($er[$i])   inne: $($ot[$i])"
            [void]$cv.Children.Add($rc)
            $y -= $sh
        }
    }
    if ($H.ZoomFrom) {
        $x0 = 4 + $w * (($H.ZoomFrom - $t0).Ticks / ($t1 - $t0).Ticks)
        $x1 = 4 + $w * (($H.ZoomTo - $t0).Ticks / ($t1 - $t0).Ticks)
        $x0 = [Math]::Max(0, [Math]::Min($w + 8, $x0)); $x1 = [Math]::Max($x0 + 2, [Math]::Min($w + 8, $x1))
        $ov = New-Object System.Windows.Shapes.Rectangle
        $ov.Width = $x1 - $x0; $ov.Height = $ht + 8
        $ov.Fill = Get-Brush '#334F8CFF'
        $ov.IsHitTestVisible = $false
        [System.Windows.Controls.Canvas]::SetLeft($ov, $x0)
        [System.Windows.Controls.Canvas]::SetTop($ov, 0)
        [void]$cv.Children.Add($ov)
    }
    $H.Buckets = $buckets
    $H.StripLeft = 4; $H.StripBw = $bw
    $u.hStripFrom.Text = $t0.ToString('yyyy-MM-dd HH:mm')
    $u.hStripTo.Text = $t1.ToString('yyyy-MM-dd HH:mm')
}

function Select-HistoryBucket($H, [double]$X) {
    if (-not $H -or -not $H.Buckets -or $H.Buckets.Count -eq 0) { return }
    $i = [int][Math]::Floor(($X - $H.StripLeft) / $H.StripBw)
    if ($i -lt 0 -or $i -ge $H.Buckets.Count) { return }
    $b = $H.Buckets[$i]
    Set-HistoryZoom $H $b.From $b.To
}

function Show-HistoryDetails($H, $row) {
    if (-not $H) { return }
    $u = $H.Ui
    if (-not $row) { return }
    switch ($row.Kind) {
        'Day' { $u.hDetails.Text = "$($row.Title)`n$($row.Detail)"; return }
        { $_ -in 'Gap', 'Change', 'Who' } { $u.hDetails.Text = "$($row.Title)`n`n$($row.Info)"; return }
    }
    $its = $row.Items
    $last = $its[$its.Count - 1]
    $sb = New-Object System.Text.StringBuilder
    if ($its.Count -gt 1) {
        [void]$sb.AppendLine("GRUPA: $($its.Count) takich samych zdarzeń")
        [void]$sb.AppendLine("od $($row.First.ToString('yyyy-MM-dd HH:mm:ss')) do $($row.Last.ToString('yyyy-MM-dd HH:mm:ss'))")
        [void]$sb.AppendLine('Czasy: ' + ((@($its | Select-Object -First 40 | ForEach-Object { $_.Time.ToString('HH:mm:ss') })) -join ', ') + $(if ($its.Count -gt 40) { ', ...' } else { '' }))
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('OSTATNIE ZDARZENIE Z GRUPY:')
    }
    if ($last.Both) { [void]$sb.AppendLine('(to samo zdarzenie jest też w pliku .log)') }
    [void]$sb.AppendLine((Get-DetailsText $last.Src $last.E))
    $u.hDetails.Text = $sb.ToString().TrimEnd()
}

function New-HistoryMenu([int]$Id) {
    $cm = New-Object System.Windows.Controls.ContextMenu
    $add = {
        param([string]$Header, [scriptblock]$Action)
        $mi = New-Object System.Windows.Controls.MenuItem
        $mi.Header = $Header
        $mi.Tag = $Id
        $mi.Add_Click($Action)
        [void]$cm.Items.Add($mi)
    }
    & $add 'Pokaż tylko ten dzień' {
        $hh = $script:HistWins[[int]$this.Tag]; $r = $hh.Ui.hList.SelectedItem; if (-not $r) { return }
        Set-HistoryZoom $hh $r.Time.Date $r.Time.Date.AddDays(1)
    }
    & $add 'Pokaż godzinę przed i po' {
        $hh = $script:HistWins[[int]$this.Tag]; $r = $hh.Ui.hList.SelectedItem; if (-not $r) { return }
        Set-HistoryZoom $hh $r.First.AddHours(-1) $r.Last.AddHours(1)
    }
    & $add 'Pokaż od tego momentu' {
        $hh = $script:HistWins[[int]$this.Tag]; $r = $hh.Ui.hList.SelectedItem; if (-not $r -or -not $hh.Items.Count) { return }
        Set-HistoryZoom $hh $r.First ($hh.Items[$hh.Items.Count - 1].Time.AddSeconds(1))
    }
    & $add 'Pokaż do tego momentu' {
        $hh = $script:HistWins[[int]$this.Tag]; $r = $hh.Ui.hList.SelectedItem; if (-not $r -or -not $hh.Items.Count) { return }
        Set-HistoryZoom $hh $hh.Items[0].Time ($r.Last.AddSeconds(1))
    }
    & $add 'Cały zakres' { Set-HistoryZoom $script:HistWins[[int]$this.Tag] $null $null }
    [void]$cm.Items.Add((New-Object System.Windows.Controls.Separator))
    & $add 'Historia urządzenia (MAC) z tego zdarzenia' {
        $hh = $script:HistWins[[int]$this.Tag]; $r = $hh.Ui.hList.SelectedItem; if (-not $r -or $r.Kind -ne 'Event') { return }
        Open-HistoryFromEvent $r.Items[0].E 'Mac' $hh.Ui.hRange.SelectedIndex $hh.Ui.hFrom.Text $hh.Ui.hTo.Text
    }
    & $add 'Historia użytkownika z tego zdarzenia' {
        $hh = $script:HistWins[[int]$this.Tag]; $r = $hh.Ui.hList.SelectedItem; if (-not $r -or $r.Kind -ne 'Event') { return }
        Open-HistoryFromEvent $r.Items[0].E 'User' $hh.Ui.hRange.SelectedIndex $hh.Ui.hFrom.Text $hh.Ui.hTo.Text
    }
    & $add 'Historia komputera z tego zdarzenia' {
        $hh = $script:HistWins[[int]$this.Tag]; $r = $hh.Ui.hList.SelectedItem; if (-not $r -or $r.Kind -ne 'Event') { return }
        Open-HistoryFromEvent $r.Items[0].E 'Computer' $hh.Ui.hRange.SelectedIndex $hh.Ui.hFrom.Text $hh.Ui.hTo.Text
    }
    [void]$cm.Items.Add((New-Object System.Windows.Controls.Separator))
    & $add 'Kopiuj szczegóły' {
        $hh = $script:HistWins[[int]$this.Tag]; if ($hh.Ui.hDetails.Text) { [System.Windows.Clipboard]::SetText($hh.Ui.hDetails.Text) }
    }
    & $add 'Kopiuj MAC' {
        $hh = $script:HistWins[[int]$this.Tag]; $r = $hh.Ui.hList.SelectedItem
        if ($r -and $r.Kind -eq 'Event' -and $r.Items[0].E.CallingStation) { [System.Windows.Clipboard]::SetText([string]$r.Items[0].E.CallingStation) }
    }
    return $cm
}

# Oś czasu jako tekst (do zgłoszenia / maila) - dokładnie to, co widać w oknie.
function Get-HistoryText($H) {
    if (-not $H.Rows -or $H.Rows.Count -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Historia: $($H.Q.Label)  ($($H.RangeText))")
    foreach ($r in $H.Rows) {
        switch ($r.Kind) {
            'Day' { [void]$sb.AppendLine(''); [void]$sb.AppendLine("== $($r.Title)  ($($r.Detail)) ==") }
            'Event' {
                $line = '{0,-8} {1,-14} {2}' -f $r.TimeStr, $r.SubTime, $r.Title
                if ($r.Medium -or $r.Location) { $line += "  [$((@($r.Medium, $r.Location) | Where-Object { $_ }) -join ' ')]" }
                [void]$sb.AppendLine($line)
                [void]$sb.AppendLine("                        $($r.Detail)")
            }
            default { [void]$sb.AppendLine("         --- $($r.Title)") }
        }
    }
    return $sb.ToString().TrimEnd()
}

function Export-History($H) {
    if (-not $H.View -or $H.View.Count -eq 0) { Set-HistoryStatus $H 'Brak zdarzeń w widoku.' 'Warn'; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv'
    $safe = ($H.Q.Label -replace '[^\w\-]+', '_').Trim('_')
    $dlg.FileName = "NPS_historia_${safe}_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if (-not $dlg.ShowDialog($H.Win)) { return }
    $rows = foreach ($it in $H.View) {
        $e = $it.E
        [pscustomobject]@{
            Czas = $it.Time.ToString('yyyy-MM-dd HH:mm:ss'); Zrodlo = $(if ($it.Both) { 'Security+log' } elseif ($it.Src -eq 'Log') { 'log' } else { 'Security' })
            Wynik = $e.Result; Kod = $e.ReasonCode; Przyczyna = $e.Reason; Uzytkownik = $it.UserDisp; MAC = $e.CallingStation
            Komputer = $it.Computer; Siec = $it.Medium; Switch = $it.Client; Port = $e.NasPort; SSID = $it.Ssid
            Zasada = $e.Policy; Uwierzytelnianie = $e.AuthType; EAP = $e.EapType
            IP = $(if ($e.PSObject.Properties['FramedIp']) { $e.FramedIp } else { '' })
        }
    }
    $rows | Export-Csv -Path $dlg.FileName -Delimiter ';' -NoTypeInformation -Encoding UTF8
    Set-HistoryStatus $H "Zapisano $($H.View.Count) zdarzeń: $($dlg.FileName)" 'OK'
}

# --- Zdarzenia -----------------------------------------------------------------------------
$S = $script:Ctx.Sys
foreach ($k in 'Result', 'Id', 'Source', 'Client') { $S[$k].Add_SelectionChanged({ Invoke-Filter $script:Ctx.Sys }) }
$S.Search.Add_TextChanged({ Invoke-Filter $script:Ctx.Sys })
$S.Grid.Add_SelectionChanged({ Show-Details $script:Ctx.Sys $script:Ctx.Sys.Grid.SelectedItem })
$S.Reset.Add_Click({ Reset-Filters $script:Ctx.Sys; Invoke-Filter $script:Ctx.Sys })
$S.Export.Add_Click({ Invoke-Export $script:Ctx.Sys })
$S.Copy.Add_Click({ $cx = $script:Ctx.Sys; if ($cx.Details.Text) { [System.Windows.Clipboard]::SetText($cx.Details.Text) } })
$S.Grid.ContextMenu = New-SysMenu
foreach ($k in 'Id', 'Source', 'Client') { Set-ComboItems $S[$k] @() }
$ui.btnSLoad.Add_Click({ Start-SysLoad })
$ui.cbSRange.Add_SelectionChanged({ Update-CustomRange $ui.cbSRange $ui.txtSFrom $ui.txtSTo })

foreach ($n in 'Ev', 'Log') {
    $C = $script:Ctx[$n]
    foreach ($k in 'Result', 'Switch', 'Policy', 'Auth') { $C[$k].Add_SelectionChanged({ Invoke-Filter $script:Ctx[$this.Tag] }) }
    foreach ($k in 'User', 'Computer', 'Search') { $C[$k].Add_TextChanged({ Invoke-Filter $script:Ctx[$this.Tag] }) }
    $C.Grid.Add_SelectionChanged({ $cx = $script:Ctx[$this.Tag]; Show-Details $cx $cx.Grid.SelectedItem })
    $C.Reset.Add_Click({ Reset-Filters $script:Ctx[$this.Tag] })
    $C.Export.Add_Click({ Invoke-Export $script:Ctx[$this.Tag] })
    $C.Copy.Add_Click({ $cx = $script:Ctx[$this.Tag]; if ($cx.Details.Text) { [System.Windows.Clipboard]::SetText($cx.Details.Text) } })
    $C.Grid.ContextMenu = New-GridMenu $n
    # Dwuklik na wierszu = historia tego urządzenia (MAC), a bez MAC - użytkownika
    $C.Grid.Add_MouseDoubleClick({
            $dep = $_.OriginalSource
            try {
                while ($dep -and $dep -isnot [System.Windows.Controls.DataGridRow]) {
                    $dep = $(if ($dep -is [System.Windows.Media.Visual]) { [System.Windows.Media.VisualTreeHelper]::GetParent($dep) } else { [System.Windows.LogicalTreeHelper]::GetParent($dep) })
                }
            }
            catch { $dep = $null }
            if ($dep) { Open-HistoryDefault $dep.Item }
        })
    Set-ComboItems $C.Switch @()
    Set-ComboItems $C.Policy @()
    Set-ComboItems $C.Auth   @()
}
foreach ($k in 'ShowReq', 'ShowAcct') {
    $script:Ctx.Log[$k].Add_Checked({ Invoke-Filter $script:Ctx[$this.Tag] })
    $script:Ctx.Log[$k].Add_Unchecked({ Invoke-Filter $script:Ctx[$this.Tag] })
}

$ui.btnLoad.Add_Click({ Start-EvLoad })
$ui.btnLogLoad.Add_Click({ Start-LogLoad })

# Historia: dla zaznaczonego wiersza bieżącej zakładki, a bez zaznaczenia - puste okno do wpisania
$ui.btnHistory.Add_Click({
        $cx = $(switch ($ui.tabMain.SelectedIndex) { 0 { $script:Ctx.Ev } 1 { $script:Ctx.Log } default { $null } })
        if ($cx -and $cx.Grid.SelectedItem) { Open-HistoryDefault $cx.Grid.SelectedItem }
        else { [void](Show-HistoryWindow) }
    })

$ui.cbRange.Add_SelectionChanged({ Update-CustomRange $ui.cbRange $ui.txtFrom $ui.txtTo })
$ui.cbLogRange.Add_SelectionChanged({ Update-CustomRange $ui.cbLogRange $ui.txtLogFrom $ui.txtLogTo })

$ui.chkAuto.Add_Checked({ $script:AutoTimer.Start(); Set-Status 'Auto-odświeżanie włączone (co 60 s).' })
$ui.chkAuto.Add_Unchecked({ $script:AutoTimer.Stop(); Set-Status 'Auto-odświeżanie wyłączone.' })

# Wybór lokalizacji logów
$ui.txtLogPath.Text = $script:DefaultLogPath

$ui.btnLogDefault.Add_Click({ $ui.txtLogPath.Text = $script:DefaultLogPath })

$ui.btnLogFolder.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Wskaż folder z logami NPS (np. C:\Windows\System32\LogFiles lub \\serwer\udział)'
        $cur = ($ui.txtLogPath.Text -split ';')[0].Trim()
        if ($cur -and (Test-Path -LiteralPath $cur -PathType Container)) { $dlg.SelectedPath = $cur }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $ui.txtLogPath.Text = $dlg.SelectedPath }
        $dlg.Dispose()
    })

$ui.btnLogFiles.Add_Click({
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Filter      = 'Logi NPS (*.log)|*.log|Wszystkie pliki (*.*)|*.*'
        $dlg.Multiselect = $true
        $dlg.Title       = 'Wybierz pliki logów NPS'
        $cur = ($ui.txtLogPath.Text -split ';')[0].Trim()
        if ($cur -and (Test-Path -LiteralPath $cur -PathType Container)) { $dlg.InitialDirectory = $cur }
        elseif ($cur -and (Test-Path -LiteralPath $cur -PathType Leaf)) { $dlg.InitialDirectory = Split-Path -Parent $cur }
        if ($dlg.ShowDialog($win)) { $ui.txtLogPath.Text = $dlg.FileNames -join '; ' }
    })

$ui.txtLogPath.Add_KeyDown({ if ($_.Key -eq 'Return') { Start-LogLoad } })

$win.Add_KeyDown({
        if ($_.Key -eq 'F5') {
            switch ($ui.tabMain.SelectedIndex) {
                1 { Start-LogLoad }
                2 { Start-SysLoad }
                default { Start-EvLoad }
            }
        }
    })

$win.Add_Closing({
        $script:AutoTimer.Stop()
        foreach ($cx in $script:Ctx.Values) {
            $cx.Timer.Stop()
            if ($cx.PS) { try { $cx.PS.Stop(); $cx.PS.Dispose() } catch { } }
        }
    })

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { Set-Status 'Uwaga: okno nie działa z uprawnieniami administratora - odczyt dziennika Security może się nie udać.' 'Warn' }

# $global:NpsViewerNoShow - testy (tests\Smoke-NpsViewer.ps1) wczytują skrypt bez pokazywania okna.
if (-not $global:NpsViewerNoShow) {
    $win.Add_ContentRendered({ Start-EvLoad })
    [void]$win.ShowDialog()
}
