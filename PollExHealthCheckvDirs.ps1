

#Default Configuration for all scripts
#Import-Module SSCExchange -Force
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.ToString()
#Read-SSCExchangeConfig $ScriptPath $ScriptName

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$counter = 0

$logfile = "D:\ExchangeAdmin\Scripts\Monitoring\logs\HeathCheck.log"


function HealthCheck ($ExchangeServer)
{
    $StartTime = Get-Date

    if ($ExchangeServer -eq "email-courriel.rcmp-grc.gc.ca")
    {
        $Health = Invoke-WebRequest "https://$($ExchangeServer)/mapi/healthcheck.htm"
    } 
    else 
    {
        $Health = Invoke-WebRequest "http://$($ExchangeServer).natl.rcmp-grc.gc.ca/mapi/healthcheck.htm"
    }

        
    if ($($Health.StatusCode) -eq 200)
    {
        $TimeTaken = (get-date).Subtract($StartTime)

        $TotalTimeTaken = $TotalTimeTaken + $TimeTaken

        If ($TimeTaken.TotalSeconds -gt 1)
        {
                
            "$StartTime - $($ExchangeServer) - Delay: $TimeTaken - MAPI" | Out-File $logfile -Append
            $counter = $counter + 1
        }
        else
        {
            #Write-Host "$StartTime - $($ExchangeServer) - Delay: $TimeTaken"
        }

    }
    else
    {
        "$StartTime - $($ExchangeServer) - Error: $($Error.code)" |  Out-File $logfile -Append
        $counter = $counter + 1
    }

    $StartTime = Get-Date

    if ($ExchangeServer -eq "email-courriel.rcmp-grc.gc.ca")
    {
        $Health = Invoke-WebRequest "https://$($ExchangeServer)/owa/healthcheck.htm"
    } 
    else 
    {
        $Health = Invoke-WebRequest "http://$($ExchangeServer).natl.rcmp-grc.gc.ca/owa/healthcheck.htm"
    }

        
    if ($($Health.StatusCode) -eq 200)
    {
        $TimeTaken = (get-date).Subtract($StartTime)

        $TotalTimeTaken = $TotalTimeTaken + $TimeTaken

        If ($TimeTaken.TotalSeconds -gt 1)
        {
                
            "$StartTime - $($ExchangeServer) - Delay: $TimeTaken - OWA" | Out-File $logfile -Append
            $counter = $counter + 1
        }
        else
        {
            #Write-Host "$StartTime - $($ExchangeServer) - Delay: $TimeTaken"
        }

    }
    else
    {
        "$StartTime - $($ExchangeServer) - Error: $($Error.code)" |  Out-File $logfile -Append
        $counter = $counter + 1
    }
}




while ($Counter -lt 100)
{
    foreach ($ExchangeServer in $ExchangeCPICMailboxServers)
    {
        HealthCheck $ExchangeServer  
    }

    HealthCheck "email-courriel.rcmp-grc.gc.ca"


    #write-host "$(get-date) - Total Time taken: $TotalTimeTaken"
    #"$(get-date) - Total Time taken: $TotalTimeTaken" |  Out-File $logfile -Append
    Start-Sleep -Seconds 10
    $TotalTimeTaken = 0
}
