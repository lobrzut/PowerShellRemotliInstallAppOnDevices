cls
$ErrorActionPreference = 'silentlycontinue'
$computername = Get-Content $PSScriptRoot\losts.txt
$sourcefile = "c:\temp\APP.msi"
$adminUserName = "youknow"
$adminPassword = "youknow"
$collection = $()


foreach ($computer in $computername) 
{

$status = @{ "Host" = $computer  }


    if (Test-Connection $computer -Count 1 -ea 0 -Quiet){

        if (Test-Path -path "\\$computer\C$\Program Files (x86)\AppDir\AppNameFolder\")
        {
        $status["Stan"] = "Already Installed"
        
        }

        else{
        
        copy-item -Path $sourcefile -destination "\\$computer\C$\Windows\Temp" -PassThru
        Write-Host "Files Copied Successfully"
        C:\PSTools\psexec.exe \\$computer -s -u $adminUserName -p $adminPassword msiexec.exe /i c:\Windows\temp\APP.msi /qb /l* out.txt
        Write-Host "Installed On: $computer" -ForegroundColor Green
        $status["State"] = "Up"
        
        }

    }

    else
        {
         $status["State"] = "Down"
   
        }
        New-Object -TypeName PSObject -Property $status -OutVariable serverStatus
        $collection += $serverStatus  
       
}

   $collection | Export-Csv $PSScriptRoot\Complete.csv -NoTypeInformation
 


