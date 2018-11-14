$fileName = "c:\users\dbiermann\documents\Vipre.csv"
$computers = Get-ADComputer -Filter '*' | Select -Exp Name

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