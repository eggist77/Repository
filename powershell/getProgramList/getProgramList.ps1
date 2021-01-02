
# @description Get a list of installed applications
# @auther T.N.
# @version 1.0
# @since 2021-01-02
# @update 2021-01-02


$CurrentDir = Split-Path $MyInvocation.MyCommand.Path

# 1
Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
    Export-csv -path $CurrentDir"\ProgramList1.csv" -Encoding Default -NoTypeInformation

# 2
Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
    Export-csv -path $CurrentDir"\ProgramList2.csv" -Encoding Default -NoTypeInformation

# 3
Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
    Export-csv -path $CurrentDir"\ProgramList3.csv" -Encoding Default -NoTypeInformation

