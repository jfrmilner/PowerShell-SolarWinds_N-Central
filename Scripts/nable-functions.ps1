



function Test-PsRemoting { 
    #http://www.leeholmes.com/blog/2009/11/20/testing-for-powershell-remoting-test-psremoting/
    param( 
        [Parameter(Mandatory = $true)] 
        $computerName 
    ) 
    
    try 
    { 
        $errorActionPreference = "Stop" 
        $result = Invoke-Command -ComputerName $computername { 1 } 
    } 
    catch 
    { 
        Write-Verbose $_ 
        return $false 
    } 
    
    ## I've never seen this happen, but if you want to be 
    ## thorough.... 
    if($result -ne 1) 
    { 
        Write-Verbose "Remoting to $computerName returned an unexpected result." 
        return $false 
    } 
    
    $true    
}

function Enable-PsRemotingPsExec { 
    param( 
        [Parameter(Mandatory = $true)] 
        $computerName 
    ) 
    
    try
    { 
        if (!(Test-PsRemoting -computerName $computerName)) {
            Write-Verbose "Enabling PSRemoting with PsExec"
            Start-Process -FilePath 'C:\Support\jfrmilner\PSTools\PsExec64.exe' -ArgumentList " \\$($computerName) -s PowerShell Enable-PSRemoting -Force " -NoNewWindow -Wait
        } 
        $errorActionPreference = "Stop" 
        $result = Invoke-Command -ComputerName $computername { 1 } 
    } 
    catch 
    { 
        Write-Verbose $_ 
        return $false 
    } 
    
    ## I've never seen this happen, but if you want to be 
    ## thorough.... 
    if($result -ne 1) 
    { 
        Write-Verbose "Remoting to $computerName returned an unexpected result." 
        return $false
    } 
    
    $true
}


function Uninstall-NAbleAgent {
<#
  .SYNOPSIS
  Uninstall N-Able Agent Installation from Windows
  .EXAMPLE
  Uninstall-NAbleAgent -computerName W10L-12345
  .NOTES
  Requires PowerShell Remoting

#>
    param (
        [Parameter(Mandatory=$true)]
        $computerName
    )

    begin {
        # Check target computer is online
        if (Test-Connection -ComputerName $computerName -Count 1) {
            Write-Host "$computerName Online" -ForegroundColor Green
        } else {
            Write-Host "$computerName Offline (not responding to ICMP)" -ForegroundColor Red
            break
        }
        # Check target computer is enabled for PS Remoting
        Write-Host "Testing PowerShell Remoting..." -ForegroundColor DarkYellow
        if (Enable-PsRemotingPsExec -computerName $computerName) {
        } else {
            Write-Host "PowerShell Remoting Needs to be enabled for $computerName" -ForegroundColor Red
            break
        }
        # Copy NcentralAssetTool.exe to target system
        try {
            Write-Host "Transferring NcentralAssetTool.exe" -ForegroundColor DarkYellow
            Copy-Item -LiteralPath "\\sccm2016\Source_Content\Application Deployment\N-Able\GroupPolicyInstaller\NcentralAssetTool.exe" \\$($computerName)\C$\Windows\Temp -Force
            Write-Host "Transfer Complete" -ForegroundColor Green
        }
        catch [system.exception] {
            Write-Host '$_ is' $_
            Write-Host '$Error[0].GetType().FullName is' $Error[0].GetType().FullName
            Write-Host '$Error[0].Exception is' $Error[0].Exception
            Write-Host '$Error[0].Exception.GetType().FullName is' $Error[0].Exception.GetType().FullName
            Write-Host '$Error[0].Exception.Message is' $Error[0].Exception.Message
        }


	}#begin
    process {
            try { 
                Invoke-Command -ComputerName $computerName -ScriptBlock {
                    # Check for Windows Agent (N-Able) installation
                    Write-Host "Checking for existing Windows Agent installation" -ForegroundColor DarkYellow
                    $app = Get-CimInstance Win32_Product -Filter "Name = 'Windows Agent'"
                    # Uninstall
                    if ($app.Caption -eq 'Windows Agent') {
                        Write-Host "Windows Agent found" -ForegroundColor Green
                        Write-Host "Uninstalling Windows Agent" -ForegroundColor DarkYellow
                        $app | Invoke-CimMethod -MethodName Uninstall
                    } else {
                        "Windows Agent not found"
                    }
                    # Clear Asset Tag Values
                    Write-Host "Clear Windows Agent Asset Tag Values" -ForegroundColor DarkYellow
                    & C:\Windows\Temp\NcentralAssetTool.exe -d
                    }
            }
            catch [system.exception] {
                Write-Host '$_ is' $_
                Write-Host '$Error[0].GetType().FullName is' $Error[0].GetType().FullName
                Write-Host '$Error[0].Exception is' $Error[0].Exception
                Write-Host '$Error[0].Exception.GetType().FullName is' $Error[0].Exception.GetType().FullName
                Write-Host '$Error[0].Exception.Message is' $Error[0].Exception.Message
            }
            finally {
            }
	}#process
    end {
        Write-Host "Uninstall-NAbleAgent Complete" -ForegroundColor Green
	}#end
}


function Install-NAbleAgent {
<#
  .SYNOPSIS
  Install N-Able Agent Installation from Windows
  .EXAMPLE
  Install-NAbleAgent -computerName W10L-12345 -customerID 397
  .NOTES
  Requires PowerShell Remoting

#>
    param (
        [Parameter(Mandatory=$true)]
        $computerName,
        [Parameter(Mandatory=$true)]
        $customerID
    )

    begin {
        # Check target computer is online
        if (Test-Connection -ComputerName $computerName -Count 1) {
            Write-Host "$computerName Online" -ForegroundColor Green
        } else {
            Write-Host "$computerName Offline (not responding to ICMP)" -ForegroundColor Red
            break
        }
        # Check target computer is enabled for PS Remoting
        Write-Host "Testing PowerShell Remoting..." -ForegroundColor DarkYellow
        if (Enable-PsRemotingPsExec -computerName $computerName) {
        } else {
            Write-Host "PowerShell Remoting Needs to be enabled for $computerName" -ForegroundColor Red
            break
        }
        # Copy Installer WindowsAgentSetup.exe to target system
        try {
            Write-Host "Transferring WindowsAgentSetup.exe" -ForegroundColor DarkYellow
            Copy-Item -LiteralPath "\\sccm2016\Source_Content\Application Deployment\N-Able\GroupPolicyInstaller\WindowsAgentSetup.exe" \\$($computerName)\C$\Windows\Temp -Force
            Write-Host "Transfer Complete" -ForegroundColor Green
        }
        catch [system.exception] {
            Write-Host '$_ is' $_
            Write-Host '$Error[0].GetType().FullName is' $Error[0].GetType().FullName
            Write-Host '$Error[0].Exception is' $Error[0].Exception
            Write-Host '$Error[0].Exception.GetType().FullName is' $Error[0].Exception.GetType().FullName
            Write-Host '$Error[0].Exception.Message is' $Error[0].Exception.Message
        }

	}#begin
    process {
            try { 
                Invoke-Command -ComputerName $computerName -ScriptBlock {
                        # Check for Windows Agent (N-Able) installation
                        Write-Host "Checking for existing Windows Agent installation" -ForegroundColor DarkYellow
                        if (Test-Path "C:\Program Files (x86)\N-able Technologies\Windows Agent\bin\agent.exe") {
                            Write-Host "Aborting: Existing Windows Agent installation Detected" -ForegroundColor Red
                            } else {
                                if (Test-Path "C:\Windows\Temp\WindowsAgentSetup.exe") {
                                #Install Windows Agent
                                Write-Host "Installing Windows Agent" -ForegroundColor DarkYellow
                                $redirectStandardError = "C:\Windows\Temp\RedirectStandardError.log"
                                $redirectStandardOutput = "C:\Windows\Temp\RedirectStandardOutput.log"
                                Remove-Item -LiteralPath $redirectStandardError -ErrorAction SilentlyContinue
                                Start-Process C:\Windows\Temp\WindowsAgentSetup.exe -ArgumentList "/s /v`" /qn CUSTOMERID=$($using:customerID) CUSTOMERSPECIFIC=1 SERVERPROTOCOL=HTTPS SERVERADDRESS=support.wirebird.co.uk SERVERPORT=443 `"" -Wait -RedirectStandardError $redirectStandardError -RedirectStandardOutput $redirectStandardOutput
                                Get-Content $redirectStandardError
                                Get-Content $redirectStandardOutput
                                
                                #Test installation agent
                                if (Test-Path "C:\Program Files (x86)\N-able Technologies\Windows Agent\bin\agent.exe") {
                                    Write-Host "Complete: Installation Windows Agent" -ForegroundColor Green
                                    } else {
                                        Write-Host "Error/Failure: Installation Windows Agent" -ForegroundColor Red
                                    }
                                #Display 10 Newest Application Event Logs
                                Write-Host "Event Log Summary" -ForegroundColor DarkYellow
                                Get-EventLog -LogName Application -Source MsiInstaller -Newest 10 | ? { $_.Message -match "Windows Agent"} | select TimeGenerated, Message
                                } else {
                                    Write-Host "Aborting: Installer file Not Found" -ForegroundColor Red
                                }
                            }
                    } -HideComputerName
            }
            catch [system.exception] {
                Write-Host '$_ is' $_
                Write-Host '$Error[0].GetType().FullName is' $Error[0].GetType().FullName
                Write-Host '$Error[0].Exception is' $Error[0].Exception
                Write-Host '$Error[0].Exception.GetType().FullName is' $Error[0].Exception.GetType().FullName
                Write-Host '$Error[0].Exception.Message is' $Error[0].Exception.Message
            }
            finally {
            }
	}#process
    end {
	}#end
}





