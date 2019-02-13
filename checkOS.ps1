$fileName = Read-Host -Prompt "Enter CSV Path"
$computers = Get-ADComputer -Filter '*' | Select -Exp Name
$winOS = Read-Host -Prompt "Enter OS Name wildcard included"

$scriptblock = {
    Get-WmiObject -class Win32_OperatingSystem | Select-Object -expand caption
}

 ForEach ($computer in $computers) {

    #ping before trying anything
    if (Test-Connection -Computername $computer -BufferSize 16 -Count 1 -Quiet){
    #search for OS
        $hasWinOS = Invoke-Command -ComputerName $computer -ScriptBlock $scriptblock
    #check for error befor moving on   
            if ($error -ne $null ){
            Write-Host "can't connect"
    #write to csv if query returns Windows 7
        } elseif ($hasWinOS.trim() -like $winOS){
                     Write-Host $computer "Has specified OS!"
                     $computer | Out-File -Append -Force $fileName
                     
                }else{
                        Write-Host $computer "Not OS"
                        }
           $error.clear()
                
        }
    else {
            Write-Host "$computer offline"
            #"$computer - offline/unavailable"|Out-File -Append $fileName
         }
}


#check via AD
#Get-ADComputer -Properties OperatingSystem  -Filter {OperatingSystem -like "*Windows 7*"} | Select-Object Name, Operatingsystem | Out-File "c:\users\dbiermann\desktop\win7.txt"