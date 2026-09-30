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
        Title="NPS Event Viewer" Height="900" Width="1560" MinHeight="700" MinWidth="1200"
        WindowStartupLocation="CenterScreen" Background="#15161C"
        FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <SolidColorBrush x:Key="Card"   Color="#1E2029"/>
    <SolidColorBrush x:Key="Field"  Color="#272A36"/>
    <SolidColorBrush x:Key="Line"   Color="#343847"/>
    <SolidColorBrush x:Key="Accent" Color="#4F8CFF"/>
    <SolidColorBrush x:Key="Text"   Color="#E8EAF0"/>
    <SolidColorBrush x:Key="Muted"  Color="#8A90A2"/>

    <Style TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
    </Style>
    <Style x:Key="Caption" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,5"/>
    </Style>
    <Style x:Key="CardTitle" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="FontSize" Value="15"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="FieldBox" TargetType="StackPanel">
      <Setter Property="Margin" Value="0,0,12,10"/>
    </Style>

    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Background" Value="#2C3040"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="Padding" Value="14,7"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Opacity" Value="0.85"/>
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
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Background" Value="{StaticResource Field}"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="CaretBrush" Value="{StaticResource Text}"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="6">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
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

    <Style TargetType="ComboBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="MinHeight" Value="32"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton Focusable="False" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="bd" Background="#272A36" BorderBrush="#343847" BorderThickness="1" CornerRadius="6">
                      <Path HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,10,0"
                            Data="M 0 0 L 4 4 L 8 0" Stroke="#8A90A2" StrokeThickness="1.6"/>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True">
                        <Setter TargetName="bd" Property="BorderBrush" Value="#4F8CFF"/>
                      </Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter IsHitTestVisible="False" Margin="10,0,28,0" VerticalAlignment="Center"
                                Content="{TemplateBinding SelectionBoxItem}"
                                ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
              <Popup IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom" AllowsTransparency="True"
                     Focusable="False" PopupAnimation="Slide">
                <Border Background="#272A36" BorderBrush="#343847" BorderThickness="1" CornerRadius="6" Margin="0,2,0,0"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}" MaxHeight="360">
                  <ScrollViewer>
                    <StackPanel IsItemsHost="True"/>
                  </ScrollViewer>
                </Border>
              </Popup>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="Padding" Value="10,6"/>
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

    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
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
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Background="Transparent" BorderBrush="Transparent" BorderThickness="0,0,0,2"
                    Padding="16,8" Margin="0,0,6,-1">
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
      <Setter Property="RowBackground" Value="#272A36"/>
      <Setter Property="AlternatingRowBackground" Value="#2B2E3B"/>
      <Setter Property="GridLinesVisibility" Value="None"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="RowHeight" Value="27"/>
      <Setter Property="SelectionMode" Value="Single"/>
      <Setter Property="EnableRowVirtualization" Value="True"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#1E2029"/>
      <Setter Property="Foreground" Value="{StaticResource Muted}"/>
      <Setter Property="Padding" Value="8,6"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Foreground" Value="{StaticResource Text}"/>
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
  </Window.Resources>

  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- NAGŁÓWEK -->
    <DockPanel Grid.Row="0" Margin="0,0,0,10">
      <Border DockPanel.Dock="Right" Background="{StaticResource Card}" CornerRadius="14" Padding="12,6" VerticalAlignment="Center">
        <StackPanel Orientation="Horizontal">
          <Ellipse Name="dotState" Width="9" Height="9" Fill="#8A90A2" VerticalAlignment="Center" Margin="0,0,8,0"/>
          <TextBlock Name="lblState" Text="Brak danych" Foreground="{StaticResource Muted}"/>
        </StackPanel>
      </Border>
      <StackPanel>
        <TextBlock Text="NPS Event Viewer" FontSize="24" FontWeight="SemiBold"/>
        <TextBlock Text="Zdarzenia autoryzacji Network Policy Server - dziennik Security (6272-6280), pliki logów RADIUS (.log) i zdarzenia systemowe"
                   Foreground="{StaticResource Muted}" Margin="0,2,0,0"/>
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
          <Border Grid.Row="0" Background="{StaticResource Card}" CornerRadius="10" Padding="16,14,16,4" Margin="0,0,0,12">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="200">
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
              <StackPanel Style="{StaticResource FieldBox}" Width="150">
                <TextBlock Text="OD (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtFrom" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="150">
                <TextBlock Text="DO (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtTo" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="110">
                <TextBlock Text="MAKS. ZDARZEŃ" Style="{StaticResource Caption}"/>
                <TextBox Name="txtMax" Text="5000" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Margin="4,0,16,17">
                <CheckBox Name="chkAuto" Content="Auto-odświeżanie co 60 s"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button Name="btnLoad" Content="Wczytaj zdarzenia" Style="{StaticResource BtnPrimary}" Padding="20,8"/>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- FILTRY -->
          <Border Grid.Row="1" Background="{StaticResource Card}" CornerRadius="10" Padding="16,14,16,4" Margin="0,0,0,12">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="WYNIK" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbResult" SelectedIndex="0">
                  <ComboBoxItem Content="Wszystkie"/>
                  <ComboBoxItem Content="Udzielono"/>
                  <ComboBoxItem Content="Odmowa"/>
                  <ComboBoxItem Content="Inne"/>
                </ComboBox>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="210">
                <TextBlock Text="SWITCH (KLIENT RADIUS)" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbSwitch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="230">
                <TextBlock Text="ZASADA SIECIOWA" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbPolicy"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="UWIERZYTELNIANIE" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbAuth"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="190">
                <TextBlock Text="UŻYTKOWNIK / MAC" Style="{StaticResource Caption}"/>
                <TextBox Name="txtUser" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="170">
                <TextBlock Text="KOMPUTER" Style="{StaticResource Caption}"/>
                <TextBox Name="txtComputer"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="220">
                <TextBlock Text="SZUKAJ WSZĘDZIE" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSearch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button Name="btnResetFilters" Content="Wyczyść filtry" Style="{StaticResource Btn}" Margin="0,0,8,0"/>
                <Button Name="btnExport" Content="Eksport CSV" Style="{StaticResource Btn}"/>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- TABELA + SZCZEGÓŁY -->
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="14"/>
              <ColumnDefinition Width="430"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Background="{StaticResource Card}" CornerRadius="10" Padding="16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Margin="0,0,0,10">
                  <TextBlock Name="lblCounts" DockPanel.Dock="Right" Text="" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
                  <TextBlock Text="Zdarzenia" Style="{StaticResource CardTitle}"/>
                </DockPanel>
                <DataGrid Name="dgEvents" Grid.Row="1">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Czas" Binding="{Binding TimeStr}" Width="145" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Wynik" Binding="{Binding Result}" Width="95" ElementStyle="{StaticResource ResultCell}"/>
                    <DataGridTextColumn Header="Użytkownik" Binding="{Binding User}" Width="150"/>
                    <DataGridTextColumn Header="MAC klienta" Binding="{Binding CallingStation}" Width="140" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Komputer" Binding="{Binding Machine}" Width="150"/>
                    <DataGridTextColumn Header="Switch" Binding="{Binding Client}" Width="160"/>
                    <DataGridTextColumn Header="Port" Binding="{Binding NasPort}" Width="50"/>
                    <DataGridTextColumn Header="Zasada sieciowa" Binding="{Binding Policy}" Width="180"/>
                    <DataGridTextColumn Header="Uwierz." Binding="{Binding AuthType}" Width="70"/>
                    <DataGridTextColumn Header="Kod" Binding="{Binding ReasonCode}" Width="45"/>
                    <DataGridTextColumn Header="Przyczyna" Binding="{Binding Reason}" Width="*"/>
                  </DataGrid.Columns>
                </DataGrid>
              </Grid>
            </Border>

            <Border Grid.Column="2" Background="{StaticResource Card}" CornerRadius="10" Padding="16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Margin="0,0,0,10">
                  <Button Name="btnCopy" DockPanel.Dock="Right" Content="Kopiuj" Style="{StaticResource Btn}" Padding="12,4"/>
                  <TextBlock Text="Szczegóły" Style="{StaticResource CardTitle}" VerticalAlignment="Center"/>
                </DockPanel>
                <TextBox Name="txtDetails" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                         VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"
                         Text="Zaznacz zdarzenie na liście."/>
                <TextBlock Grid.Row="2" Text="Statystyki widoku" Style="{StaticResource CardTitle}" Margin="0,14,0,8"/>
                <Border Grid.Row="3" Background="{StaticResource Field}" CornerRadius="6" Padding="10" MaxHeight="260">
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
          <Border Grid.Row="0" Background="{StaticResource Card}" CornerRadius="10" Padding="16,14,16,4" Margin="0,0,0,12">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="420">
                <TextBlock Text="LOKALIZACJA (folder, plik lub UNC; kilka ścieżek rozdziel ;)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogPath" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button Name="btnLogFolder" Content="Folder..." Style="{StaticResource Btn}" Margin="0,0,8,0"/>
                <Button Name="btnLogFiles" Content="Pliki..." Style="{StaticResource Btn}" Margin="0,0,8,0"/>
                <Button Name="btnLogDefault" Content="Domyślna" Style="{StaticResource Btn}"/>
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
              <StackPanel Style="{StaticResource FieldBox}" Width="150">
                <TextBlock Text="OD (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogFrom" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="150">
                <TextBlock Text="DO (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogTo" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="110">
                <TextBlock Text="MAKS. WPISÓW" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLogMax" Text="20000" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button Name="btnLogLoad" Content="Wczytaj logi" Style="{StaticResource BtnPrimary}" Padding="20,8"/>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- FILTRY -->
          <Border Grid.Row="1" Background="{StaticResource Card}" CornerRadius="10" Padding="16,14,16,4" Margin="0,0,0,12">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="WYNIK" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbLResult" SelectedIndex="0">
                  <ComboBoxItem Content="Wszystkie"/>
                  <ComboBoxItem Content="Udzielono"/>
                  <ComboBoxItem Content="Odmowa"/>
                  <ComboBoxItem Content="Inne"/>
                </ComboBox>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="210">
                <TextBlock Text="SWITCH (KLIENT RADIUS)" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbLSwitch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="230">
                <TextBlock Text="ZASADA SIECIOWA" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbLPolicy"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="140">
                <TextBlock Text="UWIERZYTELNIANIE" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbLAuth"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="190">
                <TextBlock Text="UŻYTKOWNIK / MAC" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLUser" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="150">
                <TextBlock Text="SERWER NPS" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLComputer"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="200">
                <TextBlock Text="SZUKAJ WSZĘDZIE" Style="{StaticResource Caption}"/>
                <TextBox Name="txtLSearch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Margin="4,0,16,10">
                <TextBlock Text="POKAŻ TAKŻE" Style="{StaticResource Caption}"/>
                <CheckBox Name="chkLogReq" Content="Access-Request / Challenge" Margin="0,0,0,4"/>
                <CheckBox Name="chkLogAcct" Content="Accounting"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button Name="btnLResetFilters" Content="Wyczyść filtry" Style="{StaticResource Btn}" Margin="0,0,8,0"/>
                <Button Name="btnLExport" Content="Eksport CSV" Style="{StaticResource Btn}"/>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- TABELA + SZCZEGÓŁY -->
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="14"/>
              <ColumnDefinition Width="430"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Background="{StaticResource Card}" CornerRadius="10" Padding="16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Margin="0,0,0,10">
                  <TextBlock Name="lblLCounts" DockPanel.Dock="Right" Text="" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
                  <TextBlock Text="Wpisy z logów" Style="{StaticResource CardTitle}"/>
                </DockPanel>
                <DataGrid Name="dgLog" Grid.Row="1">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Czas" Binding="{Binding TimeStr}" Width="145" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Wynik" Binding="{Binding Result}" Width="95" ElementStyle="{StaticResource ResultCell}"/>
                    <DataGridTextColumn Header="Użytkownik" Binding="{Binding User}" Width="150"/>
                    <DataGridTextColumn Header="MAC klienta" Binding="{Binding CallingStation}" Width="140" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Switch" Binding="{Binding Client}" Width="160"/>
                    <DataGridTextColumn Header="Port" Binding="{Binding NasPort}" Width="50"/>
                    <DataGridTextColumn Header="Zasada sieciowa" Binding="{Binding Policy}" Width="180"/>
                    <DataGridTextColumn Header="Uwierz." Binding="{Binding AuthType}" Width="80"/>
                    <DataGridTextColumn Header="Kod" Binding="{Binding ReasonCode}" Width="45"/>
                    <DataGridTextColumn Header="Przyczyna" Binding="{Binding Reason}" Width="*"/>
                    <DataGridTextColumn Header="Serwer NPS" Binding="{Binding Machine}" Width="110"/>
                  </DataGrid.Columns>
                </DataGrid>
              </Grid>
            </Border>

            <Border Grid.Column="2" Background="{StaticResource Card}" CornerRadius="10" Padding="16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Margin="0,0,0,10">
                  <Button Name="btnLCopy" DockPanel.Dock="Right" Content="Kopiuj" Style="{StaticResource Btn}" Padding="12,4"/>
                  <TextBlock Text="Szczegóły" Style="{StaticResource CardTitle}" VerticalAlignment="Center"/>
                </DockPanel>
                <TextBox Name="txtLDetails" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                         VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"
                         Text="Wskaż lokalizację logów i kliknij Wczytaj logi."/>
                <TextBlock Grid.Row="2" Text="Statystyki widoku" Style="{StaticResource CardTitle}" Margin="0,14,0,8"/>
                <Border Grid.Row="3" Background="{StaticResource Field}" CornerRadius="6" Padding="10" MaxHeight="260">
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
          <Border Grid.Row="0" Background="{StaticResource Card}" CornerRadius="10" Padding="16,14,16,4" Margin="0,0,0,12">
            <WrapPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="200">
                <TextBlock Text="SERWER NPS (puste = lokalny)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSServer"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="260">
                <TextBlock Text="ŹRÓDŁA ZDARZEŃ (rozdziel ;)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSProviders" Text="NPS; IAS; Microsoft-Windows-NPS" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Margin="4,0,16,10">
                <TextBlock Text="DZIENNIKI" Style="{StaticResource Caption}"/>
                <CheckBox Name="chkSSystem" Content="System" IsChecked="True" Margin="0,0,0,4"/>
                <CheckBox Name="chkSApp" Content="Application"/>
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
              <StackPanel Style="{StaticResource FieldBox}" Width="150">
                <TextBlock Text="OD (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSFrom" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="150">
                <TextBlock Text="DO (rrrr-mm-dd gg:mm)" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSTo" FontFamily="Consolas" IsEnabled="False"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="110">
                <TextBlock Text="MAKS. ZDARZEŃ" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSMax" Text="5000" FontFamily="Consolas"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button Name="btnSLoad" Content="Wczytaj zdarzenia" Style="{StaticResource BtnPrimary}" Padding="20,8"/>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- FILTRY -->
          <Border Grid.Row="1" Background="{StaticResource Card}" CornerRadius="10" Padding="16,14,16,4" Margin="0,0,0,12">
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
              <StackPanel Style="{StaticResource FieldBox}" Width="230">
                <TextBlock Text="KLIENT RADIUS (IP)" Style="{StaticResource Caption}"/>
                <ComboBox Name="cbSClient"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" Width="260">
                <TextBlock Text="SZUKAJ W TREŚCI" Style="{StaticResource Caption}"/>
                <TextBox Name="txtSSearch"/>
              </StackPanel>
              <StackPanel Style="{StaticResource FieldBox}" VerticalAlignment="Bottom" Orientation="Horizontal">
                <Button Name="btnSResetFilters" Content="Wyczyść filtry" Style="{StaticResource Btn}" Margin="0,0,8,0"/>
                <Button Name="btnSExport" Content="Eksport CSV" Style="{StaticResource Btn}"/>
              </StackPanel>
            </WrapPanel>
          </Border>

          <!-- TABELA + SZCZEGÓŁY -->
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="14"/>
              <ColumnDefinition Width="430"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Background="{StaticResource Card}" CornerRadius="10" Padding="16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Margin="0,0,0,10">
                  <TextBlock Name="lblSCounts" DockPanel.Dock="Right" Text="" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
                  <TextBlock Text="Zdarzenia systemowe" Style="{StaticResource CardTitle}"/>
                </DockPanel>
                <DataGrid Name="dgSys" Grid.Row="1">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="Czas" Binding="{Binding TimeStr}" Width="145" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Poziom" Binding="{Binding Result}" Width="95" ElementStyle="{StaticResource ResultCell}"/>
                    <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="55"/>
                    <DataGridTextColumn Header="Źródło" Binding="{Binding Source}" Width="80"/>
                    <DataGridTextColumn Header="Dziennik" Binding="{Binding LogName}" Width="80"/>
                    <DataGridTextColumn Header="Klient RADIUS" Binding="{Binding ClientIp}" Width="120" FontFamily="Consolas"/>
                    <DataGridTextColumn Header="Switch" Binding="{Binding ClientName}" Width="130"/>
                    <DataGridTextColumn Header="Wiadomość" Binding="{Binding Short}" Width="*"/>
                  </DataGrid.Columns>
                </DataGrid>
              </Grid>
            </Border>

            <Border Grid.Column="2" Background="{StaticResource Card}" CornerRadius="10" Padding="16">
              <Grid>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <DockPanel Grid.Row="0" Margin="0,0,0,10">
                  <Button Name="btnSCopy" DockPanel.Dock="Right" Content="Kopiuj" Style="{StaticResource Btn}" Padding="12,4"/>
                  <TextBlock Text="Szczegóły" Style="{StaticResource CardTitle}" VerticalAlignment="Center"/>
                </DockPanel>
                <TextBox Name="txtSDetails" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True"
                         VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"
                         Text="Kliknij Wczytaj zdarzenia."/>
                <TextBlock Grid.Row="2" Text="Statystyki widoku" Style="{StaticResource CardTitle}" Margin="0,14,0,8"/>
                <Border Grid.Row="3" Background="{StaticResource Field}" CornerRadius="6" Padding="10" MaxHeight="260">
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
    <Grid Grid.Row="2" Margin="0,12,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="260"/>
      </Grid.ColumnDefinitions>
      <TextBlock Name="lblStatus" Text="Gotowy" Foreground="{StaticResource Muted}" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
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
    param($Computer, $Start, $End, $Max)

    $ids = 6272, 6273, 6274, 6276, 6277, 6278, 6279, 6280
    $p = @{
        FilterHashtable = @{ LogName = 'Security'; Id = $ids; StartTime = $Start; EndTime = $End }
        ErrorAction     = 'Stop'
    }
    if ($Max -gt 0) { $p.MaxEvents = $Max }
    if ($Computer)  { $p.ComputerName = $Computer }

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
            MacHex         = ($calling -replace '[^0-9A-Fa-f]', '').ToUpper()
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
    param([string[]]$Paths, [string]$Mask, [datetime]$Start, [datetime]$End, [int]$Max)

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
    $portTypes = @{ '0' = 'Async'; '1' = 'Sync'; '2' = 'ISDN Sync'; '5' = 'Virtual'; '15' = 'Ethernet'; '19' = 'Wireless-802.11' }
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
    }
    $mergeKeys = 'User-Name', 'Calling-Station-Id', 'Called-Station-Id', 'NAS-Port', 'NAS-Port-Type', 'NAS-Identifier', 'NAS-IP-Address'

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
    $stat    = @{ Lines = 0; Parsed = 0; Skipped = 0; Unsupported = 0 }
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
                        if ($pt -eq '1') { if ($key) { $lastReq[$key] = $d } }
                        elseif ($pt -eq '2' -or $pt -eq '3' -or $pt -eq '11') {
                            if ($key -and $lastReq.ContainsKey($key)) {
                                $rq = $lastReq[$key]
                                foreach ($k in $mergeKeys) { if (-not $d[$k] -and $rq[$k]) { $d[$k] = $rq[$k]; $d['#Merged'] = $true } }
                            }
                        }
                        if ($time -lt $Start) { continue }
                        $queue.Enqueue($d)
                        if ($Max -gt 0 -and $queue.Count -gt $Max) { [void]$queue.Dequeue() }
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
                if ($pt -eq '1') { if ($key) { $lastReq[$key] = $d } }
                elseif ($pt -eq '2' -or $pt -eq '3' -or $pt -eq '11') {
                    if ($key -and $lastReq.ContainsKey($key)) {
                        $rq = $lastReq[$key]
                        foreach ($k in $mergeKeys) { if (-not $d[$k] -and $rq[$k]) { $d[$k] = $rq[$k]; $d['#Merged'] = $true } }
                    }
                }
                if ($time -lt $Start) { continue }
                $queue.Enqueue($d)
                if ($Max -gt 0 -and $queue.Count -gt $Max) { [void]$queue.Dequeue() }
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
            MacHex         = ($calling -replace '[^0-9A-Fa-f]', '').ToUpper()
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
        Errors      = @($errors)
    }
}

# --- Wczytywanie w tle: zdarzenia systemowe NPS (System / Application) --------------------
$script:SysLoader = {
    param($Computer, [datetime]$Start, [datetime]$End, [int]$Max, [string[]]$Logs, [string[]]$Providers)

    $prov = ($Providers | ForEach-Object { "@Name='$($_ -replace "'", '')'" }) -join ' or '
    $cond = "Provider[$prov]"
    if ($Start -gt [datetime]::MinValue) {
        $fmt = 'yyyy-MM-ddTHH:mm:ss.fffZ'
        $cond += " and TimeCreated[@SystemTime>='$($Start.ToUniversalTime().ToString($fmt))' and @SystemTime<='$($End.ToUniversalTime().ToString($fmt))']"
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

    if ($C.Name -eq 'Sys') {
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
        $C.Details.Text = $sb.ToString().TrimEnd()
        return
    }

    if ($C.Name -eq 'Ev') {
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

    if ($C.Name -eq 'Log') {
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
    $C.Details.Text = $sb.ToString().TrimEnd()
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

$win.Add_ContentRendered({ Start-EvLoad })
[void]$win.ShowDialog()
