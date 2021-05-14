if ((Test-Path -LiteralPath "HKCU:\Console\%SystemRoot%_system32_cmd.exe") -ne $true) {
    New-Item "HKCU:\Console\%SystemRoot%_system32_cmd.exe" -Force -ea SilentlyContinue
};

if ((Test-Path -LiteralPath "HKCU:\Console\%SystemRoot%_sysnative_WindowsPowerShell_v1.0_powershell.exe") -ne $true) {
    New-Item "HKCU:\Console\%SystemRoot%_sysnative_WindowsPowerShell_v1.0_powershell.exe" -Force -ea SilentlyContinue
};

New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'ScreenBufferSize' -Value 655294574 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'WindowSize' -Value 1966190 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'HistoryBufferSize' -Value 25 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'InsertMode' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'QuickEdit' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'FontSize' -Value 1179648 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'FaceName' -Value 'Consolas' -PropertyType String -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'CursorSize' -Value 25 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'NumberOfHistoryBuffers' -Value 4 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'FilterOnPaste' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_system32_cmd.exe' -Name 'LineSelection' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_sysnative_WindowsPowerShell_v1.0_powershell.exe' -Name 'ScreenBufferSize' -Value 655294574 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_sysnative_WindowsPowerShell_v1.0_powershell.exe' -Name 'WindowSize' -Value 1966190 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_sysnative_WindowsPowerShell_v1.0_powershell.exe' -Name 'CursorSize' -Value 25 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_sysnative_WindowsPowerShell_v1.0_powershell.exe' -Name 'HistoryBufferSize' -Value 25 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_sysnative_WindowsPowerShell_v1.0_powershell.exe' -Name 'FontSize' -Value 1179648 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_sysnative_WindowsPowerShell_v1.0_powershell.exe' -Name 'FaceName' -Value 'Consolas' -PropertyType String -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKCU:\Console\%SystemRoot%_sysnative_WindowsPowerShell_v1.0_powershell.exe' -Name 'NumberOfHistoryBuffers' -Value 4 -PropertyType DWord -Force -ea SilentlyContinue;