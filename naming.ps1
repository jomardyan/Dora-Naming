# Wymagana biblioteka do GUI
Add-Type -AssemblyName PresentationFramework

# Funkcja do tworzenia GUI
function Open-DORAFileNameConfigDialog {
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="DORA: Generowanie i Zmiana Nazwy Pliku" Height="500" Width="500">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>

        <Label Grid.Row="0" Grid.Column="0" Content="KOD UKNF PODMIOTU:" Margin="5"/>
        <TextBox Grid.Row="0" Grid.Column="1" Name="CodeTextBox" Margin="5"/>

        <Label Grid.Row="1" Grid.Column="0" Content="SYMBOL MODUŁU:" Margin="5"/>
        <ComboBox Grid.Row="1" Grid.Column="1" Name="ModuleComboBox" Margin="5"/>

        <Label Grid.Row="2" Grid.Column="0" Content="SYMBOL OKRESU (yyyyMM):" Margin="5"/>
        <TextBox Grid.Row="2" Grid.Column="1" Name="PeriodTextBox" Margin="5"/>

        <Label Grid.Row="3" Grid.Column="0" Content="Wybierz Plik:" Margin="5"/>
        <TextBox Grid.Row="3" Grid.Column="1" Name="FilePathTextBox" IsReadOnly="True" Margin="5"/>
        <Button Grid.Row="3" Grid.Column="2" Content="Przeglądaj" Name="BrowseButton" Margin="5"/>

        <Label Grid.Row="4" Grid.Column="0" Content="ZNACZNIK CZASU:" Margin="5"/>
        <TextBox Grid.Row="4" Grid.Column="1" Name="TimestampTextBox" IsReadOnly="True" Margin="5"/>

        <Label Grid.Row="5" Grid.Column="0" Content="Rozszerzenie:" Margin="5"/>
        <ComboBox Grid.Row="5" Grid.Column="1" Name="ExtensionComboBox" Margin="5">
            <ComboBoxItem Content=".xlsx" IsSelected="True"/>
            <ComboBoxItem Content=".zip"/>
        </ComboBox>

        <Button Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="3" Content="Generuj i Zmień Nazwę Pliku" Name="RenameButton" Margin="5" HorizontalAlignment="Center"/>

        <Label Grid.Row="7" Grid.Column="0" Content="Wygenerowana Nazwa Pliku:" Margin="5"/>
        <TextBox Grid.Row="7" Grid.Column="1" Name="GeneratedFileNameTextBox" IsReadOnly="True" Margin="5"/>
        <Button Grid.Row="7" Grid.Column="2" Content="Kopiuj" Name="CopyButton" Margin="5"/>
    </Grid>
</Window>
"@

    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    $form = [Windows.Markup.XamlReader]::Load($reader)

    # Elementy GUI
    $CodeTextBox = $form.FindName("CodeTextBox")
    $ModuleComboBox = $form.FindName("ModuleComboBox")
    $PeriodTextBox = $form.FindName("PeriodTextBox")
    $FilePathTextBox = $form.FindName("FilePathTextBox")
    $TimestampTextBox = $form.FindName("TimestampTextBox")
    $ExtensionComboBox = $form.FindName("ExtensionComboBox")
    $BrowseButton = $form.FindName("BrowseButton")
    $RenameButton = $form.FindName("RenameButton")
    $GeneratedFileNameTextBox = $form.FindName("GeneratedFileNameTextBox")
    $CopyButton = $form.FindName("CopyButton")

    # Wypełnienie listy rozwijanej dla SYMBOL MODUŁU
    for ($i = 1; $i -le 27; $i++) {
        $ModuleComboBox.Items.Add("SPRPF{0:D2}" -f $i) | Out-Null
    }
    $ModuleComboBox.SelectedIndex = 0

    # Obsługa przycisku Przeglądaj
    $BrowseButton.Add_Click({
        $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $FileDialog.Filter = "Wszystkie pliki (*.*)|*.*"
        if ($FileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $FilePath = $FileDialog.FileName
            $FilePathTextBox.Text = $FilePath

            # Generowanie znacznika czasu na podstawie daty modyfikacji pliku
            $FileInfo = Get-Item $FilePath
            $LastWriteTime = $FileInfo.LastWriteTime
            $Timestamp = $LastWriteTime.ToString("yyyyMMddHHmmssfff")
            $TimestampTextBox.Text = $Timestamp
        }
    })

    # Obsługa kliknięcia przycisku Generuj i Zmień Nazwę
    $RenameButton.Add_Click({
        $Code = $CodeTextBox.Text
        $Module = $ModuleComboBox.SelectedItem
        $Period = $PeriodTextBox.Text
        $Timestamp = $TimestampTextBox.Text
        $Extension = $ExtensionComboBox.Text
        $OriginalFilePath = $FilePathTextBox.Text

        # Walidacja pól
        if (-not ($Code -and $Module -and $Period -and $Timestamp -and $OriginalFilePath)) {
            [System.Windows.MessageBox]::Show("Wszystkie pola muszą być wypełnione, a plik musi być wybrany!", "Błąd", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            return
        }

        if ($Period -notmatch '^\d{6}$') {
            [System.Windows.MessageBox]::Show("Symbol okresu musi mieć format yyyyMM.", "Błąd", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
            return
        }

        # Generowanie nazwy pliku
        $NewFileName = "${Code}_${Module}_${Period}_${Timestamp}${Extension}"
        $GeneratedFileNameTextBox.Text = $NewFileName

        # Zmiana nazwy pliku
        $NewFilePath = [System.IO.Path]::Combine((Get-Item $OriginalFilePath).DirectoryName, $NewFileName)
        Rename-Item -Path $OriginalFilePath -NewName $NewFilePath -Force

        [System.Windows.MessageBox]::Show("Plik został pomyślnie przemianowany:\n$NewFilePath", "Sukces", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    })

    # Obsługa przycisku Kopiuj
    $CopyButton.Add_Click({
        $GeneratedFileName = $GeneratedFileNameTextBox.Text
        if (-not [string]::IsNullOrWhiteSpace($GeneratedFileName)) {
            Set-Clipboard -Value $GeneratedFileName
            [System.Windows.MessageBox]::Show("Wygenerowana nazwa pliku została skopiowana do schowka.", "Sukces", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        } else {
            [System.Windows.MessageBox]::Show("Brak nazwy pliku do skopiowania.", "Błąd", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        }
    })

    # Wyświetlenie okna
    $form.ShowDialog() | Out-Null
}

# Wymagana biblioteka Windows Forms dla okna wyboru pliku
Add-Type -AssemblyName System.Windows.Forms

# Wywołanie funkcji
Open-DORAFileNameConfigDialog
