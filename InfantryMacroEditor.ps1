#<# Elevate to Administrator if not already, comment out if compiling the .exe with Invoke-PS2EXE
If (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    # Create a new process with elevated privileges
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"" + $myInvocation.MyCommand.Definition + "`""
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Exit
}

# Enforce .EXE to run as admin with mt.exe
# & "C:\Program Files (x86)\Windows Kits\10\bin\10.0.19041.0\x64\mt.exe" -manifest "InfantryMacroEditor.manifest" -outputresource:"InfantryMacroEditor.exe;1"

#>

$settingsFilePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("ApplicationData"), "InfantryMacroEditor", "Settings.json")
$settingsDirectory = [System.IO.Path]::GetDirectoryName($settingsFilePath)

if (-not (Test-Path $settingsDirectory)) {
    New-Item -Path $settingsDirectory -ItemType Directory

    $settings = @{
        MainUIBackgroundColor1 = 0
        MainUIBackgroundColor2 = 0
        DataGridBackgroundColor = 125
        SyntaxCommandColor = "#FF800080"
        SyntaxBongColor = "#FFFFD700" 
        SyntaxPlusQTYColor = "#FF228B22"
        SyntaxMinusQTYColor = "#FFFF0000"
        SyntaxNormalColor = "#FF000000"
        installPath = "C:\\Program Files (x86)\\Infantry Online"
    }

    $settings | ConvertTo-Json | Set-Content -Path $settingsFilePath
}

$global:settings = Get-Content -Path $settingsFilePath | ConvertFrom-Json
$global:SyntaxCommandColor = $settings.SyntaxCommandColor
$global:SyntaxBongColor = $settings.SyntaxBongColor
$global:SyntaxPlusQTYColor = $settings.SyntaxPlusQTYColor
$global:SyntaxMinusQTYColor = $settings.SyntaxMinusQTYColor
$global:SyntaxNormalColor = $settings.SyntaxNormalColor
$global:installPath = $settings.installPath

# Load the necessary assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# Define the necessary types outside of functions to avoid scoping issues
Add-Type -TypeDefinition @"
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Collections.ObjectModel;
using System.Linq;

public class FormattedPart : INotifyPropertyChanged {
    private string text;
    private string color;
    private string fontWeight;

    public string Text {
        get { return text; }
        set { text = value; OnPropertyChanged(); }
    }

    public string Color {
        get { return color; }
        set { color = value; OnPropertyChanged(); }
    }

    public string FontWeight {
        get { return fontWeight; }
        set { fontWeight = value; OnPropertyChanged(); }
    }

    public static string SyntaxCommand { get; set; }
    public static string SyntaxBong { get; set; }
    public static string SyntaxPlusQTY { get; set; }
    public static string SyntaxMinusQTY { get; set; }
    public static string SyntaxNormal { get; set; }

    static FormattedPart() {
        SyntaxCommand = "$global:SyntaxCommandColor";
        SyntaxBong = "$global:SyntaxBongColor";
        SyntaxPlusQTY = "$global:SyntaxPlusQTYColor";
        SyntaxMinusQTY = "$global:SyntaxMinusQTYColor";
        SyntaxNormal = "$global:SyntaxNormalColor";
    }

    public event PropertyChangedEventHandler PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string name = null) {
        if (PropertyChanged != null) {
            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}

public class MacroItem : INotifyPropertyChanged {
    private string keyBinding;
    private string line;
    private ObservableCollection<FormattedPart> formattedParts;

    public string KeyBinding {
        get { return keyBinding; }
        set { keyBinding = value; OnPropertyChanged(); }
    }

    public string Line {
        get { return line; }
        set {
            line = value;
            OnPropertyChanged();
            FormattedParts = FormatLine(line);
        }
    }

    public ObservableCollection<FormattedPart> FormattedParts {
        get { return formattedParts; }
        set { formattedParts = value; OnPropertyChanged(); }
    }

    public event PropertyChangedEventHandler PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string name = null) {
        if (PropertyChanged != null) {
            PropertyChanged.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    private ObservableCollection<FormattedPart> FormatLine(string line) {
        var formattedParts = new ObservableCollection<FormattedPart>();
        var word = string.Empty;
        var specialWord = string.Empty;
        bool isSpecial = false;

        for (int i = 0; i < line.Length; i++) {
            var character = line[i];

            if (character == ' ' || character == ',' || character == ':') {
                if (!string.IsNullOrEmpty(word)) {
                    AddFormattedPart(word, formattedParts);
                    word = string.Empty;
                }
                formattedParts.Add(new FormattedPart {
                    Text = character.ToString(),
                    Color = FormattedPart.SyntaxNormal,
                    FontWeight = "Normal"
                });
            } else if (character == '-') {
                if (!string.IsNullOrEmpty(word)) {
                    specialWord = word + character;
                    word = string.Empty;
                    isSpecial = true;
                } else {
                    word += character;
                }
            } else if (character == '%' && i + 1 < line.Length && char.IsDigit(line[i + 1])) {
                if (!string.IsNullOrEmpty(word)) {
                    AddFormattedPart(word, formattedParts);
                    word = string.Empty;
                }
                word += character;
                i++;
                while (i < line.Length && char.IsDigit(line[i])) {
                    word += line[i];
                    i++;
                }
                AddFormattedPart(word, formattedParts);
                word = string.Empty;
                i--;
            } else {
                word += character;
            }
        }
        if (!string.IsNullOrEmpty(word)) {
            if (isSpecial) {
                specialWord += word;
                AddFormattedPart(specialWord, formattedParts);
            } else {
                AddFormattedPart(word, formattedParts);
            }
        }
        return formattedParts;
    }

    private void AddFormattedPart(string word, ObservableCollection<FormattedPart> formattedParts) {
        var formattedPart = new FormattedPart {
            Text = word,
            Color = FormattedPart.SyntaxNormal,
            FontWeight = "Normal"
        };

        if (word.StartsWith("?") || (word.StartsWith("%") && word.Length > 1 && char.IsLetter(word[1]))) {
            formattedPart.Color = FormattedPart.SyntaxCommand; // Assuming purple color is the same as SyntaxCommand
            formattedPart.FontWeight = "Bold";
        } else if (word.StartsWith("%") && word.Skip(1).All(char.IsDigit)) {
            formattedPart.Color = FormattedPart.SyntaxBong;
            formattedPart.FontWeight = "Bold";
        } else if (word.All(char.IsDigit)) {
            formattedPart.Color = FormattedPart.SyntaxPlusQTY;
            formattedPart.FontWeight = "Bold";
        } else if (word.StartsWith("#") && word.Skip(1).All(char.IsDigit)) {
            formattedPart.Color = FormattedPart.SyntaxPlusQTY;
            formattedPart.FontWeight = "Bold";
        } else if (word.StartsWith("-") && word.Skip(1).All(char.IsDigit)) {
            formattedPart.Color = FormattedPart.SyntaxMinusQTY;
            formattedPart.FontWeight = "Bold";
        } else if (word.Contains("%")) {
            int idx = word.IndexOf('%');
            if (idx > 0) {
                AddFormattedPart(word.Substring(0, idx), formattedParts);
            }
            if (word.Length > idx + 1 && char.IsLetter(word[idx + 1])) {
                // Handle % followed by letters (A-Z)
                string remainingWord = word.Substring(idx + 1);
                int endIdx = remainingWord.TakeWhile(c => char.IsLetter(c)).Count();
                formattedParts.Add(new FormattedPart {
                    Text = "%" + remainingWord.Substring(0, endIdx),
                    Color = FormattedPart.SyntaxCommand, // Assuming purple color is the same as SyntaxCommand
                    FontWeight = "Bold"
                });
                AddFormattedPart(remainingWord.Substring(endIdx), formattedParts);
            } else {
                formattedParts.Add(new FormattedPart {
                    Text = "%",
                    Color = FormattedPart.SyntaxBong,
                    FontWeight = "Bold"
                });
                AddFormattedPart(word.Substring(idx + 1), formattedParts);
            }
            return;
        } else if (word.Contains(":") && word.Split(':')[1].StartsWith("-") && word.Split(':')[1].Skip(1).All(char.IsDigit)) {
            formattedPart.Color = FormattedPart.SyntaxMinusQTY;
        } else if (word.All(char.IsLetterOrDigit)) {
            formattedPart.Color = FormattedPart.SyntaxNormal;
            formattedPart.FontWeight = "Normal";
        }


        formattedParts.Add(formattedPart);
    }
}

public class BinaryReader {
    public byte[] buffer;
    public int position;

    public BinaryReader(byte[] buffer) {
        this.buffer = buffer;
        this.position = 0;
    }

    public int GetInt32() {
        int value = BitConverter.ToInt32(this.buffer, this.position);
        this.position += 4;
        return value;
    }

    public uint GetUint32() {
        uint value = BitConverter.ToUInt32(this.buffer, this.position);
        this.position += 4;
        return value;
    }

    public string GetString(int length) {
        string value = System.Text.Encoding.ASCII.GetString(this.buffer, this.position, length).TrimEnd('\0');
        this.position += length;
        return value;
    }
}

public class BlobEntry {
    public string name;
    public int offset;
    public int length;

    public BlobEntry(string name, int offset, int length) {
        this.name = name;
        this.offset = offset;
        this.length = length;
    }
}

public class BlobFile {
    public System.Collections.Generic.List<BlobEntry> entries;

    public BlobFile() {
        this.entries = new System.Collections.Generic.List<BlobEntry>();
    }

    public void Deserialize(byte[] buffer) {
        BinaryReader reader = new BinaryReader(buffer);
        int version = reader.GetInt32();
        uint filecount = reader.GetUint32();

        int nameLength = (version == 2) ? 32 : 14;

        for (int i = 0; i < filecount; i++) {
            string name = reader.GetString(nameLength);
            uint offset = reader.GetUint32();
            uint length = reader.GetUint32();
            this.entries.Add(new BlobEntry(name, (int)offset, (int)length));
        }
    }
}

public class Audio {
    [System.Runtime.InteropServices.DllImport("winmm.dll")]
    public static extern bool PlaySound(string fname, int Mod, int flag);
}
"@

# Define the main window XAML
$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Infantry Macro Editor" Height="600" Width="1300">
    <Grid>
        <ComboBox Name="fileComboBox" Width="200" Height="25" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <DataGrid Name="dataGrid" AutoGenerateColumns="False" Margin="10,40,10,40" CanUserAddRows="False" CanUserDeleteRows="False" AlternatingRowBackground="#F0F0F0" RowBackground="White">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Key Binding" Binding="{Binding KeyBinding}" Width="100"/>
                <DataGridTemplateColumn Header="Macro" Width="*">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <ItemsControl ItemsSource="{Binding FormattedParts}">
                               <ItemsControl.ItemsPanel>
                                    <ItemsPanelTemplate>
                                        <StackPanel Orientation="Horizontal"/>
                                    </ItemsPanelTemplate>
                                </ItemsControl.ItemsPanel>
                                <ItemsControl.ItemTemplate>
                                    <DataTemplate>
                                        <TextBlock Text="{Binding Text}" Foreground="{Binding Color}" FontWeight="{Binding FontWeight}"/>
                                    </DataTemplate>
                                </ItemsControl.ItemTemplate>
                            </ItemsControl>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                    <DataGridTemplateColumn.CellEditingTemplate>
                        <DataTemplate>
                            <TextBox Text="{Binding Line, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellEditingTemplate>
                </DataGridTemplateColumn>
            </DataGrid.Columns>
        </DataGrid>
        <Button Name="testMacroButton" Content="Test Macro" Width="100" Height="25" HorizontalAlignment="Left" Margin="10,0,0,10" VerticalAlignment="Bottom"/>
        <Button Name="saveButton" Content="Save" Width="75" Height="25" HorizontalAlignment="Right" Margin="0,0,10,10" VerticalAlignment="Bottom"/>
        <Button Name="saveAsGlobalButton" Content="Save as Global Macro" Width="150" Height="25" HorizontalAlignment="Center" Margin="0,0,10,10" VerticalAlignment="Bottom"/>
        <Button Name="settingsButton" Content="Settings" Width="100" Height="25" HorizontalAlignment="Right" Margin="0,0,90,10" VerticalAlignment="Bottom"/>
    </Grid>
</Window>
"@

# Define the settings window XAML
$settingsXaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Settings" Height="600" Width="400">
    <Grid Margin="10">
        <TextBlock Text="Main UI Background Color 1" VerticalAlignment="Top" Margin="10,10,0,0"/>
        <Slider Name="bgColor1Slider" Minimum="0" Maximum="255" Width="300" Height="25" HorizontalAlignment="Left" Margin="10,40,0,0" VerticalAlignment="Top"/>
        <TextBlock Text="Main UI Background Color 2" VerticalAlignment="Top" Margin="10,80,0,0"/>
        <Slider Name="bgColor2Slider" Minimum="0" Maximum="255" Width="300" Height="25" HorizontalAlignment="Left" Margin="10,110,0,0" VerticalAlignment="Top"/>
        <TextBlock Text="DataGrid Row Background" VerticalAlignment="Top" Margin="10,150,0,0"/>
        <Slider Name="dataGridBgSlider" Minimum="0" Maximum="255" Width="300" Height="25" HorizontalAlignment="Left" Margin="10,180,0,0" VerticalAlignment="Top"/>
        <TextBlock Text="?command Color" VerticalAlignment="Top" Margin="10,220,0,0"/>
        <Button Name="syntaxCommandColorButton" Content="Choose Color" Width="100" Height="25" HorizontalAlignment="Left" Margin="10,250,0,0" VerticalAlignment="Top"/>
        <TextBlock Text="Install Path" VerticalAlignment="Top" Margin="160,220,0,0"/>
        <TextBox Name="installPathTextBox" Width="200" Height="25" HorizontalAlignment="Left" Margin="160,250,0,0" VerticalAlignment="Top" IsReadOnly="False"/>
        <Button Name="browseInstallPathButton" Content="Browse..." Width="75" Height="25" HorizontalAlignment="Left" Margin="160,275,0,0" VerticalAlignment="Top"/>
        <TextBlock Text="%1-30 Color" VerticalAlignment="Top" Margin="10,290,0,0"/>
        <Button Name="syntaxBongColorButton" Content="Choose Color" Width="100" Height="25" HorizontalAlignment="Left" Margin="10,320,0,0" VerticalAlignment="Top"/>
        <TextBlock Text="+Qty Color" VerticalAlignment="Top" Margin="10,360,0,0"/>
        <Button Name="syntaxPlusQTYColorButton" Content="Choose Color" Width="100" Height="25" HorizontalAlignment="Left" Margin="10,390,0,0" VerticalAlignment="Top"/>
        <TextBlock Text="-Qty Color" VerticalAlignment="Top" Margin="10,430,0,0"/>
        <Button Name="syntaxMinusQTYColorButton" Content="Choose Color" Width="100" Height="25" HorizontalAlignment="Left" Margin="10,460,0,0" VerticalAlignment="Top"/>
        <Button Name="saveSettingsButton" Content="Save" Width="100" Height="25" HorizontalAlignment="Right" Margin="0,0,10,10" VerticalAlignment="Bottom"/>
    </Grid>
</Window>
"@

# Load the XAML
$reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)

# Load the settings XAML
$settingsReader = (New-Object System.Xml.XmlNodeReader ([xml]$settingsXaml))
$settingsWindow = [Windows.Markup.XamlReader]::Load($settingsReader)

# Access the ComboBox, DataGrid, and Buttons
$fileComboBox = $window.FindName("fileComboBox")
$dataGrid = $window.FindName("dataGrid")
$saveButton = $window.FindName("saveButton")
$testMacroButton = $window.FindName("testMacroButton")
$saveAsGlobalButton = $window.FindName("saveAsGlobalButton")
$settingsButton = $window.FindName("settingsButton")

# Access the settings controls
$bgColor1Slider = $settingsWindow.FindName("bgColor1Slider")
$bgColor2Slider = $settingsWindow.FindName("bgColor2Slider")
$dataGridBgSlider = $settingsWindow.FindName("dataGridBgSlider")
$syntaxCommandColorButton = $settingsWindow.FindName("syntaxCommandColorButton")
$syntaxBongColorButton = $settingsWindow.FindName("syntaxBongColorButton")
$syntaxPlusQTYColorButton = $settingsWindow.FindName("syntaxPlusQTYColorButton")
$syntaxMinusQTYColorButton = $settingsWindow.FindName("syntaxMinusQTYColorButton")
$saveSettingsButton = $settingsWindow.FindName("saveSettingsButton")
$installPathTextBox = $settingsWindow.FindName("installPathTextBox")


# Ensure controls are not null
if ($null -eq $fileComboBox -or $null -eq $dataGrid) {
    Write-Error "Failed to find the necessary controls in the XAML."
    pause
    exit
}

# Ensure settings controls are not null
if ($null -eq $bgColor1Slider -or $null -eq $bgColor2Slider -or $null -eq $dataGridBgSlider -or $null -eq $syntaxCommandColorButton -or $null -eq $syntaxBongColorButton -or $null -eq $syntaxPlusQTYColorButton -or $null -eq $saveSettingsButton) {
    Write-Error "Failed to find the necessary settings controls in the XAML."
    pause
    exit
}

# Define the registry path
$regPath = "HKCU:\SOFTWARE\HarmlessGames\Infantry\Profile0\Keyboard"

# Define initial values and increment steps
$initialAltValue = 4743680
$initialShiftValue = 4219392
$initialCtrlValue = 4481536
$initialLAltValue = 43024896  # LAlt+A
$initialLAltNumValue = 43016192  # LAlt+0
$incrementStep = 512

# Additional starting key bindings and their initial values
$initialRAltValue = 43278336  # RAlt+0
$initialRAltAValue = 43287040  # RAlt+A
$initialLShiftNumValue = 41967616  # LShift+0
$initialLShiftValue = 41976320  # LShift+A
$initialLControlNumValue = 42491904  # LControl+0
$initialLControlValue = 42500608  # LControl+A
$initialRShiftNumValue = 42229760  # RShift+0
$initialRShiftValue = 42238464  # RShift+A
$initialRControlNumValue = 42754048  # RControl+0
$initialRControlValue = 42762752  # RControl+A

# Define possible keys (A-Z and 0-9)
$keys = @([char[]](65..90) + [char[]](48..57))  # A-Z and 0-9

# Function to generate key bindings dynamically
function Generate-KeyBindings {
    param (
        [string]$modifier,
        [int]$initialValue,
        [int]$incrementStep
    )
    
    $bindings = @{}
    $currentValue = $initialValue
    
    foreach ($key in $keys) {
        $bindings[$currentValue] = "$modifier+$key"
        $currentValue += $incrementStep
    }
    
    return $bindings
}

# Function to generate key bindings dynamically for numeric keys
function Generate-NumKeyBindings {
    param (
        [string]$modifier,
        [int]$initialValue,
        [int]$incrementStep
    )
    
    $bindings = @{}
    $currentValue = $initialValue
    
    for ($i = 0; $i -le 9; $i++) {
        $bindings[$currentValue] = "$modifier+$i"
        $currentValue += $incrementStep
    }
    
    return $bindings
}

# Generate key bindings for Alt, Shift, Ctrl, and LAlt with alphanumeric characters
$altBindings = Generate-KeyBindings -modifier "Alt" -initialValue $initialAltValue -incrementStep $incrementStep
$shiftBindings = Generate-KeyBindings -modifier "Shift" -initialValue $initialShiftValue -incrementStep $incrementStep
$ctrlBindings = Generate-KeyBindings -modifier "Ctrl" -initialValue $initialCtrlValue -incrementStep $incrementStep
$laltBindings = Generate-KeyBindings -modifier "LAlt" -initialValue $initialLAltValue -incrementStep $incrementStep
$laltNumBindings = Generate-NumKeyBindings -modifier "LAlt" -initialValue $initialLAltNumValue -incrementStep $incrementStep

# Generate key bindings for additional key combinations
$raltBindings = Generate-NumKeyBindings -modifier "RAlt" -initialValue $initialRAltValue -incrementStep $incrementStep
$raltABindings = Generate-KeyBindings -modifier "RAlt" -initialValue $initialRAltAValue -incrementStep $incrementStep
$lshiftNumBindings = Generate-NumKeyBindings -modifier "LShift" -initialValue $initialLShiftNumValue -incrementStep $incrementStep
$lshiftBindings = Generate-KeyBindings -modifier "LShift" -initialValue $initialLShiftValue -incrementStep $incrementStep
$lctrlNumBindings = Generate-NumKeyBindings -modifier "LControl" -initialValue $initialLControlNumValue -incrementStep $incrementStep
$lctrlBindings = Generate-KeyBindings -modifier "LControl" -initialValue $initialLControlValue -incrementStep $incrementStep
$rshiftNumBindings = Generate-NumKeyBindings -modifier "RShift" -initialValue $initialRShiftNumValue -incrementStep $incrementStep
$rshiftBindings = Generate-KeyBindings -modifier "RShift" -initialValue $initialRShiftValue -incrementStep $incrementStep
$rctrlNumBindings = Generate-NumKeyBindings -modifier "RControl" -initialValue $initialRControlNumValue -incrementStep $incrementStep
$rctrlBindings = Generate-KeyBindings -modifier "RControl" -initialValue $initialRControlValue -incrementStep $incrementStep

# Combine all key bindings into one hashtable
$keyBindings = @{}
$keyBindings += $altBindings
$keyBindings += $shiftBindings
$keyBindings += $ctrlBindings
$keyBindings += $laltBindings
$keyBindings += $laltNumBindings
$keyBindings += $raltBindings
$keyBindings += $raltABindings
$keyBindings += $lshiftNumBindings
$keyBindings += $lshiftBindings
$keyBindings += $lctrlNumBindings
$keyBindings += $lctrlBindings
$keyBindings += $rshiftNumBindings
$keyBindings += $rshiftBindings
$keyBindings += $rctrlNumBindings
$keyBindings += $rctrlBindings

# Function to read and convert registry values
function Get-ConvertedRegistryValues {
    param (
        [string]$Path,
        [hashtable]$ValueMap,
        [array]$KeysToFilter
    )
    
    # Initialize an empty array to store the converted values
    $convertedValues = @()
    
    # Loop through each key to filter
    foreach ($key in $KeysToFilter) {
        $keyName = $key.ToString()
        try {
            # Get the value from the registry
            $keyValue = (Get-ItemProperty -Path $Path -Name $keyName).$keyName
            
            # Convert the value to the corresponding key binding
            if ($ValueMap.ContainsKey($keyValue)) {
                $keyBinding = $ValueMap[$keyValue]
                $convertedValues += [PSCustomObject]@{
                    Key     = $keyName
                    Value   = $keyValue
                    Binding = $keyBinding
                }
            } else {
                $convertedValues += [PSCustomObject]@{
                    Key     = $keyName
                    Value   = $keyValue
                    Binding = "Unknown"
                }
            }
        } catch {
            Write-Output "Failed to read key: $keyName"
        }
    }
    
    return $convertedValues
}

# Define the keys to filter
$keysToFilter = (35..46) + (53..64)

# Get and display the converted registry values
$convertedRegistryValues = Get-ConvertedRegistryValues -Path $regPath -ValueMap $keyBindings -KeysToFilter $keysToFilter

# Function to format a line with syntax highlighting
function Format-Line {
    param ($line)
    $formattedParts = New-Object System.Collections.ArrayList

    # Split the line into parts
    $words = $line -split '(\s+|(?<=\?)|(?=\?)|(?<=#)|(?=#)|(?<=:)|(?=,)|(?<=,))'
    foreach ($word in $words) {
        # Determine color and font weight
        $color, $fontWeight = if ($word -match '^\?[a-zA-Z]+') {
            $global:SyntaxCommandColor, [System.Windows.FontWeights]::Normal
        } elseif ($word -match '^-?\d+$' -and $word -lt 0) {
            $global:SyntaxMinusQTYColor, [System.Windows.FontWeights]::Bold
        } elseif ($word -match '^\d+$' -or $word -match '^#\d+$') {
            $global:SyntaxPlusQTYColor, [System.Windows.FontWeights]::Bold
        } else {
            $global:SyntaxNormalColor, [System.Windows.FontWeights]::Normal
        }

        # Create formatted part
        $formattedPart = [PSCustomObject]@{
            Text = $word
            Color = $color
            FontWeight = $fontWeight
        }
        # Add to the list
        [void]$formattedParts.Add($formattedPart)
    }

    return $formattedParts
}

# Function to load and format the content from the selected file
function Load-FileContent {
    param ($filePath)

    $fileContent = Get-Content -Path $filePath

    # Prepare data for DataGrid
    $data = [System.Collections.ObjectModel.ObservableCollection[MacroItem]]::new()
    for ($i = 0; $i -lt $keysToFilter.Count; $i++) {
        $keyName = $keysToFilter[$i].ToString()
        $registryValue = ($convertedRegistryValues | Where-Object { $_.Key -eq $keyName })
        $keyBinding = if ($registryValue) { $registryValue.Binding } else { "Unknown" }

        $line = if ($i -lt $fileContent.Count) { $fileContent[$i] } else { "" }

        $macroItem = [MacroItem]::new()
        $macroItem.KeyBinding = $keyBinding
        $macroItem.Line = $line

        $data.Add($macroItem)
    }

    $dataGrid.ItemsSource = $data
}

$filePathMap = @{}

# Load the list of .mc0 files into the ComboBox and store full paths in hashtable
$mc0Files = Get-ChildItem -Path $global:installPath -Filter '*.mc0'

foreach ($file in $mc0Files) {
    $fileComboBox.Items.Add($file.Name)
    $filePathMap[$file.Name] = $file.FullName
}

# Event handler for file selection change
$fileComboBox.add_SelectionChanged({
    $selectedFile = $fileComboBox.SelectedItem
    if ($selectedFile) {
        $filePath = $filePathMap[$selectedFile]
        if ($filePath) {
            Load-FileContent -filePath $filePath
        }
    }
})

# Function to check if running as administrator
function Test-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Add event handler for the Save button
$saveButton.Add_Click({
    # Check if running as administrator; if not, show error message and return
    if (-not (Test-Admin)) {
        [System.Windows.MessageBox]::Show("Unable to save, please relaunch the program as administrator.", "Error")
        return
    }

    $selectedFile = $fileComboBox.SelectedItem
    if ($selectedFile) {
        # Retrieve the updated content from the DataGrid
        $updatedContent = $dataGrid.ItemsSource | ForEach-Object { $_.Line }
        # Attempt to write the updated content back to the file
        try {
            Set-Content -Path "$global:installPath\$selectedFile" -Value $updatedContent -ErrorAction Stop
            [System.Windows.MessageBox]::Show("File saved successfully to $global:installPath\$selectedFile!", "Success")
        }
        catch {
            [System.Windows.MessageBox]::Show("Error saving file: $_", "Error")
        }
    } else {
        [System.Windows.MessageBox]::Show("No file selected.", "Error")
    }
})

# Function to play sound based on the macro text
function Play-SoundFromMacro {
    param ($macroText)

    # Define the .cfg file path
    $cfgPath = "C:\Program Files (x86)\Infantry Online\ctfpl.cfg"
    $cfgContent = Get-Content -Path $cfgPath

    # Extract Bong entries
    $bongEntries = @{}
    foreach ($line in $cfgContent) {
        if ($line -match "^Bong(\d+)=(.+),(.+)$") {
            $bongEntries[$matches[1]] = @{ Blob = $matches[2]; Sound = $matches[3] }
        }
    }

    # Find the Bong reference in the macro text
    if ($macroText -match "%(\d+)" -and $bongEntries.ContainsKey($matches[1])) {
        $bongNumber = $matches[1]
        $bongEntry = $bongEntries[$bongNumber]

        if ($bongEntry) {
            $blobFileName = $bongEntry.Blob
            $soundFileName = $bongEntry.Sound

            # Logic to play the sound file from the .blo file
            $bloFilePath = "C:\Program Files (x86)\Infantry Online\$blobFileName.blo"
            Write-Output "Loading .blo file: $bloFilePath"
            $bloFileContent = [System.IO.File]::ReadAllBytes($bloFilePath)

            $blobFile = [BlobFile]::new()
            $blobFile.Deserialize($bloFileContent)

            Write-Output "Blob file entries:"
            $blobFile.entries | ForEach-Object { Write-Output "$($_.name): Offset = $($_.offset), Length = $($_.length)" }

            $wavEntry = $blobFile.entries | Where-Object { $_.name -eq "$soundFileName.wav" }
            if ($wavEntry -ne $null) {
                Write-Output "Found sound file: $soundFileName.wav"
                $wavData = New-Object byte[] $wavEntry.length
                [System.Array]::Copy($bloFileContent, $wavEntry.offset, $wavData, 0, $wavEntry.length)

                $wavTempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "$soundFileName.wav")
                [System.IO.File]::WriteAllBytes($wavTempPath, $wavData)

                [Audio]::PlaySound($wavTempPath, 0, 0x0001)
            } else {
                Write-Output "Sound file $soundFileName.wav not found in .blo file."
                [System.Windows.Forms.MessageBox]::Show("Sound file not found in .blo file.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        } else {
            Write-Output "Bong entry not found."
            [System.Windows.Forms.MessageBox]::Show("Bong entry not found.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    } else {
        Write-Output "No Bong reference found in the macro text."
        [System.Windows.Forms.MessageBox]::Show("No Bong reference found in the macro text.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Add event handler for the Test Macro button
$testMacroButton.Add_Click({
    $selectedItem = $dataGrid.SelectedItem
    if ($selectedItem) {
        $macroText = $selectedItem.Line
        Play-SoundFromMacro -macroText $macroText
    } else {
        [System.Windows.MessageBox]::Show("No macro selected.", "Error")
    }
})

# Function to create the global macro update window
function Create-GlobalUpdateWindow {
    param ($updates, $keyBinding)

    $xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Global Macro Update" Height="500" Width="800">
    <Grid>
        <TextBlock Text="Globally unify values for Key Binding $keyBinding" FontSize="20" HorizontalAlignment="Center" Margin="10,10,10,10" VerticalAlignment="Top"/>
        <DataGrid Name="globalDataGrid" AutoGenerateColumns="False" Margin="10,50,10,50" CanUserAddRows="False" CanUserDeleteRows="False" AlternatingRowBackground="#F0F0F0" RowBackground="White">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Header="Select" Binding="{Binding IsChecked}" Width="50"/>
                <DataGridTextColumn Header="File Name" Binding="{Binding FileName}" Width="*"/>
                <DataGridTextColumn Header="Old Value" Binding="{Binding OldValue}" Width="*"/>
                <DataGridTextColumn Header="New Value" Binding="{Binding NewValue}" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        <Button Name="confirmButton" Content="Confirm" Width="100" Height="25" HorizontalAlignment="Right" Margin="0,0,10,10" VerticalAlignment="Bottom"/>
    </Grid>
</Window>
"@

    # Load the XAML
    $reader = (New-Object System.Xml.XmlNodeReader ([xml]$xaml))
    $globalWindow = [Windows.Markup.XamlReader]::Load($reader)

    # Access the DataGrid and Button
    $globalDataGrid = $globalWindow.FindName("globalDataGrid")
    $confirmButton = $globalWindow.FindName("confirmButton")

    # Load the updates into the DataGrid
    $globalDataGrid.ItemsSource = $updates

    # Add event handler for the Confirm button
    $confirmButton.Add_Click({
        $selectedUpdates = $updates | Where-Object { $_.IsChecked }
        $updatedFiles = @()
        $excludedFiles = @()

        foreach ($update in $selectedUpdates) {
            $fileContent = Get-Content -Path $update.FilePath
            $fileContent[$update.Index] = $update.NewValue
            try {
                Set-Content -Path $update.FilePath -Value $fileContent -ErrorAction Stop
                $updatedFiles += $update.FileName
            }
            catch {
                $excludedFiles += $update.FileName
            }
        }

        [System.Windows.MessageBox]::Show("Key binding updated in $($updatedFiles.Count) of $($updates.Count) macro sets.`nUpdated files:`n$($updatedFiles -join "`n")`nExcluded files:`n$($excludedFiles -join "`n")", "Update Summary")
        $globalWindow.Close()
    })

    $globalWindow.ShowDialog()
}

# Add event handler for the Save as Global Macro button
$saveAsGlobalButton.Add_Click({
    $selectedItem = $dataGrid.SelectedItem
    if ($selectedItem) {
        $macroText = $selectedItem.Line
        $keyBinding = $selectedItem.KeyBinding

        # Create a list to hold update details
        $updates = @()

        # Loop through each .mc0 file and prepare the update list
        foreach ($file in $mc0Files) {
            $filePath = $file.FullName
            $fileContent = Get-Content -Path $filePath

            for ($i = 0; $i -lt $keysToFilter.Count; $i++) {
                $keyName = $keysToFilter[$i].ToString()
                $registryValue = ($convertedRegistryValues | Where-Object { $_.Key -eq $keyName })
                $currentKeyBinding = if ($registryValue) { $registryValue.Binding } else { "Unknown" }

                $currentLine = if ($i -lt $fileContent.Count) { $fileContent[$i] } else { "" }

                if ($currentKeyBinding -eq $keyBinding) {
                    $update = [PSCustomObject]@{
                        IsChecked = $false
                        FileName = $file.Name
                        FilePath = $filePath
                        OldValue = $currentLine
                        NewValue = $macroText
                        Index = $i
                    }
                    $updates += $update
                }
            }
        }

        # Create and show the global update window
        Create-GlobalUpdateWindow -updates $updates -keyBinding $keyBinding
    } else {
        [System.Windows.MessageBox]::Show("No macro selected.", "Error")
    }
})

# Function to load settings from AppData
function Load-Settings {
    $settingsFilePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("ApplicationData"), "InfantryMacroEditor", "Settings.json")

    if (Test-Path $settingsFilePath) {
        $global:settings = Get-Content -Path $settingsFilePath | ConvertFrom-Json

        $bgColor1Slider.Value = $settings.MainUIBackgroundColor1
        $bgColor2Slider.Value = $settings.MainUIBackgroundColor2
        $dataGridBgSlider.Value = $settings.DataGridBackgroundColor

        $installPathTextBox.Text = $global:settings.installPath
    }
}

# Function to save settings to AppData
function Save-Settings {
    $settingsFilePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("ApplicationData"), "InfantryMacroEditor", "Settings.json")
    $settingsDirectory = [System.IO.Path]::GetDirectoryName($settingsFilePath)

    # Ensure the settings directory exists
    if (-not (Test-Path $settingsDirectory)) {
        New-Item -Path $settingsDirectory -ItemType Directory
    }

    # Update settings with current values from the UI
    $settings = @{
        MainUIBackgroundColor1 = $bgColor1Slider.Value
        MainUIBackgroundColor2 = $bgColor2Slider.Value
        DataGridBackgroundColor = $dataGridBgSlider.Value
        SyntaxCommandColor = $global:SyntaxCommandColor
        SyntaxBongColor = $global:SyntaxBongColor
        SyntaxPlusQTYColor = $global:SyntaxPlusQTYColor
        SyntaxMinusQTYColor = $global:SyntaxMinusQTYColor
        installPath = $installPathTextBox.Text
    }

    # Convert settings to JSON and save to file
    $settings | ConvertTo-Json | Set-Content -Path $settingsFilePath
}

function UpdateBackground {
    $global:colorObj1 = New-Object System.Windows.Media.Color
    $global:colorObj1.R = [byte]$bgColor1Slider.Value
    $global:colorObj1.G = [byte]0
    $global:colorObj1.B = [byte]0
    $global:colorObj1.A = [byte]255  # Alpha (transparency)

    $global:colorObj2 = New-Object System.Windows.Media.Color
    $global:colorObj2.R = [byte]0
    $global:colorObj2.G = [byte]$bgColor2Slider.Value
    $global:colorObj2.B = [byte]0
    $global:colorObj2.A = [byte]255  # Alpha (transparency)

    $brush = New-Object System.Windows.Media.LinearGradientBrush
    $brush.StartPoint = New-Object System.Windows.Point 0,0
    $brush.EndPoint = New-Object System.Windows.Point 1,1

    $gradientStop1 = New-Object System.Windows.Media.GradientStop
    $gradientStop1.Color = $global:colorObj1
    $gradientStop1.Offset = 0.0

    $gradientStop2 = New-Object System.Windows.Media.GradientStop
    $gradientStop2.Color = $global:colorObj2
    $gradientStop2.Offset = 1.0

    $brush.GradientStops.Add($gradientStop1)
    $brush.GradientStops.Add($gradientStop2)

    $window.Background = $brush
}

# Add event handler for the Settings button
$settingsButton.Add_Click({
    # Recreate the settings window every time the button is clicked
    $settingsReader = (New-Object System.Xml.XmlNodeReader ([xml]$settingsXaml))
    $settingsWindow = [Windows.Markup.XamlReader]::Load($settingsReader)

    # Access the settings controls
    $bgColor1Slider = $settingsWindow.FindName("bgColor1Slider")
    $bgColor2Slider = $settingsWindow.FindName("bgColor2Slider")
    $dataGridBgSlider = $settingsWindow.FindName("dataGridBgSlider")
    $syntaxCommandColorButton = $settingsWindow.FindName("syntaxCommandColorButton")
    $syntaxBongColorButton = $settingsWindow.FindName("syntaxBongColorButton")
    $syntaxPlusQTYColorButton = $settingsWindow.FindName("syntaxPlusQTYColorButton")
    $syntaxMinusQTYColorButton = $settingsWindow.FindName("syntaxMinusQTYColorButton")
    $saveSettingsButton = $settingsWindow.FindName("saveSettingsButton")
    $installPathTextBox = $settingsWindow.FindName("installPathTextBox")
    $browseInstallPathButton = $settingsWindow.FindName("browseInstallPathButton")

    # Load the settings into the newly created window
    Load-Settings

    # Attach event handlers to the sliders and buttons for real-time updates
    $bgColor1Slider.Add_ValueChanged({
        $global:settings.MainUIBackgroundColor1 = [int]$bgColor1Slider.Value
        UpdateBackground
    })

    $bgColor2Slider.Add_ValueChanged({
        $global:settings.MainUIBackgroundColor2 = [int]$bgColor2Slider.Value
        UpdateBackground
    })

    $dataGridBgSlider.Add_ValueChanged({
        $global:settings.DataGridBackgroundColor = [int]$dataGridBgSlider.Value
        $dataGrid.RowBackground = New-Object Windows.Media.SolidColorBrush ([Windows.Media.Color]::FromArgb(255, [int]$global:settings.DataGridBackgroundColor, [int]$global:settings.DataGridBackgroundColor, [int]$global:settings.DataGridBackgroundColor))
        $dataGrid.AlternatingRowBackground = New-Object Windows.Media.SolidColorBrush ([Windows.Media.Color]::FromArgb(255, [int]$global:settings.DataGridBackgroundColor, [int]$global:settings.DataGridBackgroundColor, [int]$global:settings.DataGridBackgroundColor))
    })

    # Attach event handlers to color buttons
    $syntaxCommandColorButton.Add_Click({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        $colorDialog.Color = [System.Drawing.ColorTranslator]::FromHtml($global:settings.SyntaxCommandColor)
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $global:settings.SyntaxCommandColor = [System.Drawing.ColorTranslator]::ToHtml($colorDialog.Color)
        }
    })

    $syntaxBongColorButton.Add_Click({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        $colorDialog.Color = [System.Drawing.ColorTranslator]::FromHtml($global:settings.SyntaxBongColor)
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $global:settings.SyntaxBongColor = [System.Drawing.ColorTranslator]::ToHtml($colorDialog.Color)
        }
    })

    $syntaxPlusQTYColorButton.Add_Click({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        $colorDialog.Color = [System.Drawing.ColorTranslator]::FromHtml($global:settings.SyntaxPlusQTYColor)
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $global:settings.SyntaxPlusQTYColor = [System.Drawing.ColorTranslator]::ToHtml($colorDialog.Color)
        }
    })

    $syntaxMinusQTYColorButton.Add_Click({
        $colorDialog = New-Object System.Windows.Forms.ColorDialog
        $colorDialog.Color = [System.Drawing.ColorTranslator]::FromHtml($global:settings.SyntaxMinusQTYColor)
        if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $global:settings.SyntaxMinusQTYColor = [System.Drawing.ColorTranslator]::ToHtml($colorDialog.Color)
        }
    })

    # Add event handler for the Browse button
    $browseInstallPathButton.Add_Click({
        # Create a new FolderBrowserDialog instance
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select the Installation Folder"

        # Set the initial selected path to the current install path
        $folderDialog.SelectedPath = $installPathTextBox.Text

        # Show the dialog and check if the user selects a folder
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            # Update the text box with the selected folder
            $installPathTextBox.Text = $folderDialog.SelectedPath
        }
    })

    # Add event handler for the Save Settings button
    $saveSettingsButton.Add_Click({
        Save-Settings
        $settingsWindow.Close()
    })

    # Show the new instance of the settings window
    $settingsWindow.ShowDialog()
})

# Load settings and apply them on startup
Load-Settings
UpdateBackground

# Show the main window
$window.ShowDialog()
