Function Set-1EConstants{
    [cmdletBinding()]
    Param(
    )
    #! Define Module Constants
    $Global:dbInfo += (@{'NightWatchman' = @{serverName = '1EServer';databaseName = 'AgilityFrameworkReporting'}})
}
Function Get-machine1EHistory{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2018-07-09 10:03
        .COMPONENT
            1E, Web Wake Up, Machine Status, History
        .SYNOPSIS
            Returns 1E History for specific machine
        .DESCRIPTION
            Get-machine1EHistory queries the 1E database for the historic status
            changes of a particular machine
        .PARAMETER computerName
            This mandatory string-type parameter allows the technician to 
            specify a target machine
        .PARAMETER dateBegin
            This DateTime-type parameter allows the technician to narrow the
            result set by selecting a starting date
            [Defaults to 7 days ago]
        .PARAMETER dateEnd
            This DateTime-type parameter allows the technician to narrow the
            result set by selecting a ending date
            [Defaults to current DateTime]
        .EXAMPLE
            Get-machine1EHistory -computerName 'Computer1'
            This is the simplest use case, and will return the last 7 days of
            status changes
        .EXAMPLE
            Get-machine1EHistory -computerName 'Computer1' -dateBegin (Get-Date).AddDays(-25)
            This will return the last 25 days of status changes for the
            specified machine
        .EXAMPLE
            Get-machine1EHistory -computerName 'Computer1' -dateBegin (Get-Date '2018-06-01') -dateEnd (Get-Date '2018-06-05')
            This will return the status changes for the specified machine
            between the specified dates.
        .LINK
            https://github.com/mpearon
        .LINK
            https://twitter.com/@mpearon
    #>
    Param(
        [Parameter(Mandatory = $true)]$computerName,
        $dateBegin = (Get-Date).AddDays(-7),
        $dateEnd = (Get-Date)
    )
    Begin{
        Set-1EConstants
    }
    Process{
        $computerName | ForEach-Object{
            
            $queryString = @"
                SELECT dc.NetbiosName, rc.StartTimeStamp, rc.EndTimeStamp, lsn.StateName, rc.MinutesInState
                FROM tbAFR_Dimension_ConfigurationItem AS dc
                    INNER JOIN tbNWM_Report_Consumption AS rc
                        ON dc.Id = rc.ComputerId
                    INNER JOIN tbNWM_Lookup_StateNames lsn
                        ON rc.State = lsn.State
                WHERE dc.NetbiosName = '$computerName'
                    AND rc.StartTimeStamp >= '$dateBegin'
                    AND rc.EndTimeStamp <= '$dateEnd'
                ORDER BY rc.StartTimeStamp
"@
            $connectionString = ( -join ('Data Source=',($dbInfo.NightWatchman.serverName),',1433;Initial Catalog=',($dbInfo.NightWatchman.databaseName),';Integrated Security=True;'))
            $sqlConnection = New-Object System.Data.SQLClient.SQLConnection
            $sqlConnection.ConnectionString = $ConnectionString
            $sqlCommand = $sqlConnection.CreateCommand()
            $sqlCommand.CommandText = $QueryString
            $dataAdapter = New-Object System.Data.SQLClient.SQLDataAdapter $sqlCommand
            $dataSet = New-Object System.Data.Dataset
            try{
                $records = $dataAdapter.Fill($DataSet)
                $data = $dataSet.Tables[0]
            }
            catch{
                Write-Warning (-join( 'Failure: DataSet - ',($_)[0] ))
            }
            $sqlConnection.Close()
            $compiledResults += $data
        }
    }
    End{
        return $compiledResults
    }
}
Function Start-Computer{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2018-02-22 12:06
        .COMPONENT
            1E, Web Wake Up, Magic Packet, Port 9, Night Watchman, Power, Start
        .SYNOPSIS
            Sends Magic Packet via 1E
        .DESCRIPTION
            Start-Computer leverages 1E to send a Magic Packet to the target
            machines, waking them if possible
        .PARAMETER computerName
            This string-type pipeline variable allows the user to target a
            specific or set of computers to wake
        .EXAMPLE
            Start-Computer -computerName 'Computer1'
            This will send a Magic Packet to Computer1 using named parameters
        .EXAMPLE
            Start-Computer -computerName 'Computer1','Computer2'
            This will send a Magic Packet to Computer1 and Computer2 using
            named parameters
        .EXAMPLE
            'Computer1' | Start-Computer
            This will send a Magic Packet to Computer1 using the pipeline
        .EXAMPLE
            'Computer1','Computer2' | Start-Computer
            This will send a Magic Packet to Computer1 and Computer2 using
            the pipeline
        .LINK
            https://github.com/mpearon
        .LINK
            https://twitter.com/@mpearon
    #>
    [cmdletBinding(DefaultParameterSetName = 'noPrompt')]
    Param(
        [Parameter(ValueFromPipeline = $true, ParameterSetName = 'noPrompt')]$computerName,
        [Parameter(ParameterSetName = 'prompt')][switch]$promptForList
    )
    Begin{
        Set-1EConstants
        Write-Output 'Beginning Wake Process'
        if($PSBoundParameters.promptForList){
            New-Item -ItemType File -Path $env:TEMP -Name '1E.WakeList.txt' -Value '# One machine per line - Remove this line, then save and close #' -Force | Out-Null
            Start-Process Notepad.exe -ArgumentList (-join($env:TEMP,'\1E.WakeList.txt')) -Wait
            $computerName = Get-Content (-join($env:TEMP,'\1E.WakeList.txt'))
        }
    }
    Process{
        $computerName | ForEach-Object{
            Write-Verbose (-join('Sending Magic Packet to ',$_))
            Try{
                $WakeProv = Get-wmiobject -Namespace 'root\N1E\Wakeup' -ComputerName 'SDC1E-VS1' -List -Credential $Creds -ErrorAction Stop | Where-Object { $_.name -eq 'WakeUp' }
                $WakeProv.WakeName($_) | Out-Null
                Write-Verbose (-join('Magic Packet sent to ',$_))
            }
            Catch{
                Write-Host -ForegroundColor Red (-join('Unable to wake ',$_))
                Write-Verbose (-join('Unable to send Magic Packet to ',$_))
            }
        }
    }
    End{
        Write-Output 'Wake Process Complete'
    }
}