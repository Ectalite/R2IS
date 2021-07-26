function memoryinfo{
    Write-Host -NoNewline "Memory Info:`r`n"
    Get-WmiObject Win32_PhysicalMemory | Select-Object BankLabel, @{Name = "Capacity (MB)"; Expression = {$_.Capacity/1MB}},ConfiguredClockSpeed, Manufacturer | out-host
        
}

function diskinfo{
    Write-Host -NoNewline "Disk Info & Space:`r`n"
    Get-WmiObject Win32_logicaldisk | Where-Object {$_.DriveType -eq 3} | Format-Table DeviceID, VolumeName, @{Name = "FreeSpace (GB)"; Expression =  {"{0:N2}" -f ($_.FreeSpace/1GB)}}, @{Name = "Size (GB)"; Expression = {"{0:N2}" -f ($_.Size/1GB)}}, @{Label=”PercentFree”;Expression= {“{0:P}” -f ($_.FreeSpace / $_.Size)}}
    Write-Host -NoNewline "Mapped logical disk:`r`n"
    Get-WmiObject Win32_logicaldisk | Where-Object {$_.DriveType -eq 4} | Format-Table DeviceID, VolumeName
}

function CSVtoXLSX($CommandPath){
    $CSVTemp = "C:\Temp\"

    New-Item -Path $CSVTemp\tmp -ItemType Directory | Out-Null

    Get-WmiObject Win32_Computersystem | Select-Object -Property Name, Manufacturer, Model, Domain, SystemType, SystemFamily| Export-Csv "$CSVTemp\tmp\System Informations.csv"
    Write-Host -NoNewline "5%.."
    Get-WmiObject Win32_Operatingsystem | Select-Object -Property Caption, Version, BuildNumber, SystemDirectory, SerialNumber | Export-Csv "$CSVTemp\tmp\OS Informations.csv"
    Write-Host -NoNewline "10%.."
    Get-WmiObject Win32_Computersystem | Select-Object -Property UserName | Export-Csv "$CSVTemp\tmp\UserName.csv"
    Write-Host -NoNewline "15%.."
    [DateTime]::Now – (Get-WmiObject -Class Win32_OperatingSystem).ConvertToDateTime((Get-WmiObject -Class Win32_OperatingSystem).LastBootUpTime)` | Select-Object -Property Days, Hours, Minutes, Seconde | Export-Csv -UseCulture -Path "$CSVTemp\tmp\Uptime.csv" -NoTypeInformation -Encoding UTF8
    Write-Host -NoNewline "20%.."
    Get-WmiObject Win32_Printer | Select-Object -Property sharename,name | Export-Csv "$CSVTemp\tmp\Printer Informations.csv"
    Write-Host -NoNewline "25%.."
    Get-WmiObject Win32_USBControllerDevice | Foreach-Object { [Wmi]$_.Dependent } | Select-Object -Property Name, Caption, Description, DeviceID | Export-Csv "$CSVTemp\tmp\USB Devices.csv"
    Write-Host -NoNewline "30%.."
    Get-WmiObject Win32_Processor |Select-Object -Property Name, NumberOfCores, NumberOfEnabledCore, NumberOfLogicalProcessors, CurrentClockSpeed, MaxClockSpeed | Export-Csv "$CSVTemp\tmp\Processor Informations.csv" 
    Write-Host -NoNewline "35%.."
    Get-WmiObject Win32_PhysicalMemory | Select-Object BankLabel, @{Name = "Capacity (MB)"; Expression = {$_.Capacity/1MB}},ConfiguredClockSpeed, Manufacturer | Export-Csv "$CSVTemp\tmp\Memory Informations.csv"
    Write-Host -NoNewline "40%.."
    Get-WmiObject Win32_logicaldisk | Where-Object {$_.DriveType -eq 3} | Select-Object -Property DeviceID, VolumeName, @{Name = "FreeSpace (GB)"; Expression =  {"{0:N2}" -f ($_.FreeSpace/1GB)}}, @{Name = "Size (GB)"; Expression = {"{0:N2}" -f ($_.Size/1GB)}}, @{Label=”PercentFree”;Expression= {“{0:P}” -f ($_.FreeSpace / $_.Size)}} | Export-Csv "$CSVTemp\tmp\Disk Informations and Space.csv"
    Write-Host -NoNewline "45%.."
    Get-WmiObject Win32_logicaldisk | Where-Object {$_.DriveType -eq 4} | Select-Object -Property DeviceID, VolumeName | Export-Csv "$CSVTemp\tmp\Mapped logical disk.csv"
    Write-Host -NoNewline "50%.."
    Get-NetAdapter –Physical | Select-Object -Property Name, InterfaceDescription, Status, MacAddress, LinkSpeed | Export-Csv "$CSVTemp\tmp\Active Network Card.csv"
    Write-Host -NoNewline "55%.."
    Get-Process | Select-Object -Property Id, Name, mainWindowtitle | Export-Csv "$CSVTemp\tmp\Process List.csv"
    Write-Host "60%"
    Write-Host "Processing Services, this will take a while"
    Get-Service | Where-Object {$_.Status -eq "Running"} | Export-Csv "$CSVTemp\tmp\Service List.csv"

    $path="$CSVTemp\tmp" #target folder
    cd $path| Out-Null

    $csvs = Get-ChildItem .\* -Include *.csv

    $outputfilename = "Computer_Informations_" + $(get-date -f yyyyMMdd)+ ".xlsx" #creates file name with date/username

    Write-Host "Creating: $outputfilename in folder $CommandPath"

    $excelapp = new-object -comobject Excel.Application
    $excelapp.sheetsInNewWorkbook = $csvs.Count
    $xlsx = $excelapp.Workbooks.Add()
    $sheet=1

    foreach ($csv in $csvs){
        $pourcentage = 60+3*$sheet
        Write-Host -NoNewline "$pourcentage%.."
        $row=1
        $column=1
        $worksheet = $xlsx.Worksheets.Item($sheet)
        $worksheet.Name = $csv.Name
        $file = (Get-Content $csv)
        foreach($line in $file) {
            $linecontents=$line -split ',(?!\s*\w+")'
            foreach($cell in $linecontents) {
                $worksheet.Cells.Item($row,$column) = $cell
                $column++
            }
            $column=1
            $row++
        }
        $objRange = $worksheet.UsedRange
        [void] $objRange.EntireColumn.Autofit()
        $sheet++
    }

    cd .. | Out-Null

    $xlsx.SaveAs($outputfilename)
    $excelapp.quit()

    Remove-Item '.\tmp' -Recurse

    Write-Host "Done"
}

Write-Host "Bienvenue dans R2IS (outil de récupérations d'informations système)."

$ComputerSelect = 0

while(-not (($ComputerSelect -eq 1) -or ($ComputerSelect -eq 2))){
    Write-Host "Types d'utilisations:"
    Write-Host "1 - Informations Ordinateur Local"
    Write-Host "2 - Informations Ordinateur Distant"
    $ComputerSelect = Read-Host -Prompt "Veulliez sélectionner le type d'utilisation"
    <#Write-Host "Vous avez choisi $ComputerSelect."#>
    
    if($ComputerSelect -eq 1){
        
        <#Screen output#>

        Write-Host -NoNewline "`r`nSystem Info:"
        Get-WmiObject Win32_Computersystem | Select-Object -Property Name, Manufacturer, Model, Domain, SystemType, SystemFamily
        Write-Host -NoNewline "OS Info:`r`n"
        Get-WmiObject Win32_Operatingsystem | Select-Object -Property Caption, Version, BuildNumber, SystemDirectory, SerialNumber
        Write-Host -NoNewline "User Login:`r`n"
        Get-WmiObject Win32_Computersystem | Select-Object -Property UserName
        Write-Host "`r`nUptime:"
        [DateTime]::Now – (Get-WmiObject -Class Win32_OperatingSystem).ConvertToDateTime((Get-WmiObject -Class Win32_OperatingSystem).LastBootUpTime)` | Format-Table Days, Hours, Minutes, Seconds -AutoSize
        Write-Host "Printer Info:"
        Get-WmiObject Win32_Printer | Format-Table sharename,name
        Write-Host -NoNewline "USB Devices:`r`n"
        Get-WmiObject Win32_USBControllerDevice | Foreach-Object { [Wmi]$_.Dependent } | Select-Object -Property Name, Caption, Description, DeviceID
        Write-Host -NoNewline "`r`nProcessor Info:"
        Get-WmiObject Win32_Processor | Format-List -Property Name, NumberOfCores, NumberOfEnabledCore, NumberOfLogicalProcessors, CurrentClockSpeed, MaxClockSpeed
        memoryinfo
        diskinfo
        Write-Host -NoNewline "Active Network Card:`r`n"
        Get-NetAdapter –Physical | Format-Table Name, InterfaceDescription, Status, MacAddress, LinkSpeed
        Write-Host -NoNewline "Process List:`r`n"
        Get-Process | Where-Object {$_.mainWindowTitle} | Format-Table Id, Name, mainWindowtitle -AutoSize
        Write-Host -NoNewline "Running Service List:`r`n"
        Get-Service | Where-Object {$_.Status -eq "Running"} | Format-Table Status, Name, DisplayName -AutoSize
        $ExcelFile = Read-Host -Prompt "Voulez-vous sauvegarder les résultats dans un fichier excel ? (Y/N)"
        if($ExcelFile -like 'Y'){
            Write-Host "Début de la sauvegarde..."
            CSVtoXLSX($mypath = Split-Path ($MyInvocation.MyCommand.Path) -Parent)
        }
    }elseif($ComputerSelect -eq 2){
        $ComputerConnectName = Read-Host -Prompt "Veuillez donner le nom de l'ordinateur distant"
        Write-Host "Veuillez vous connecter"
        <#$ComputerConnectUsername = Read-Host -Prompt "Nom d'utilisateur"
        <#$ComputerConnectPassword = Read-Host -Prompt "Mot de passe"#>
        $adminaccount = "\"
        $PASSWORD = ConvertTo-SecureString "" -AsPlainText -Force
        $UNPASSWORD = New-Object System.Management.Automation.PsCredential $adminaccount, $PASSWORD
        #$ComputerCredential = $host.ui.PromptForCredential("Connection à $ComputerConnectName", "Veuillez entrer votre nom d'utilisateur et votre mot de passe.", "", "UserName")
        $ComputerAdress = (Get-WmiObject Win32_Computersystem -Credential $UNPASSWORD -ComputerName $ComputerConnectName)
    }else{
        Write-Host "Mauvaise sélection"
    }
}


