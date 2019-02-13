$date = Get-Date -Format “yyyyMMdd”
$fileName = "c:\users\dbiermann\documents\Vipre_$date.csv"
$adList = Get-ADComputer -Filter '*' | Select -Exp Name
$getOS = Get-ADComputer -Properties OperatingSystem  -Filter {OperatingSystem -like "*Server*"} | Select-Object -ExpandProperty Name
$computers = Compare-Object $getOS $adList | Select-Object -ExpandProperty InputObject


 ForEach ($computer in $computers) {

    #ping before trying anything
    if (Test-Connection -Computername $computer -BufferSize 16 -Count 1 -Quiet){
    #search for vipre installation
        $hasVipre = Invoke-Command -ComputerName $computer -ScriptBlock {Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -like "*Vipre*"}|Select-Object DisplayName}
    #check for error befor moving on   
            if ($error -ne $null ){
            Write-Host "can't connect"
    #write to csv if query returns vipre
        } elseif ($hasVipre -eq $null){
                     $computer | Out-File -Append -Force $fileName
                }else{
                        Write-Host "Vipre Installed"
                        }
           $error.clear()       
        }
    else {
            Write-Host "$computer offline"
            #"$computer - offline/unavailable"|Out-File -Append $fileName
         }
}

#Send e-mail notification

$From = "it@ci.montrose.co.us"
$To = "dbiermann@cityofmontrose.org"
$Attachment = $fileName
$Subject = "Computers without AV"
$Body = "See attached txt file for more information."
$SMTPServer = "192.168.60.23"
$SMTPPort = "25"
Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -Attachments $Attachment –DeliveryNotificationOption OnSuccess