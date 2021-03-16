#Default Configuration for all scripts
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.ToString()

# Force TLS1.2 protocol
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Initializing variables
$counter = 0
$OutputPath = "c:\temp\"
$LogFileName = "HealthCheck.txt"
$logfile = $OutputPath + $LogFileName

$ExchangeLoadBalancerURL = "mail.CanadaSam.ca"
$ADInternalDomainFQDN = ""

If (!(Test-Path -Path $OutputPath)){New-Item -Folder "C:\temp" -Force}


function HealthCheck ($ExchangeServer)
{
    $StartTime = Get-Date

    if ($ExchangeServer -eq "$ExchangeLoadBalancerURL")
    {
        $Health = Invoke-WebRequest "https://$($ExchangeServer)/mapi/healthcheck.htm"
    } 
    else 
    {
        $Health = Invoke-WebRequest "http://$($ExchangeServer).$ADInternalDomainFQDN/mapi/healthcheck.htm"
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

    if ($ExchangeServer -eq "$ExchangeLoadBalancerURL")
    {
        $Health = Invoke-WebRequest "https://$($ExchangeServer)/owa/healthcheck.htm"
    } 
    else 
    {
        $Health = Invoke-WebRequest "http://$($ExchangeServer).$ADInternalDomainFQDN/owa/healthcheck.htm"
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

    HealthCheck "$ExchangeLoadBalancerURL"


    #write-host "$(get-date) - Total Time taken: $TotalTimeTaken"
    #"$(get-date) - Total Time taken: $TotalTimeTaken" |  Out-File $logfile -Append
    Start-Sleep -Seconds 10
    $TotalTimeTaken = 0
}
