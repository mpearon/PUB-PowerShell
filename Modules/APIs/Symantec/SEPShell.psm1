<#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2017-10-21 17:53
            Keywords:   Symantec Endpoint Protection, SEP, Antivirus, AV, API
        .COMPONENT
            Symantec Endpoint Protection, SEP, Antivirus, AV, RESTful API
        .SYNOPSIS
            Allows interface to Symantec Endpoint Protection
        .DESCRIPTION
            SEPShell allows the user to query or manipulate resources via the Symantec Endpoint Protection Manager (SEPM) RESTful API
        .EXAMPLE
            Import-Module SEPShell
        .LINK
            https://support.symantec.com/en_US/article.HOWTO125873.html
        .LINK
            https://github.com/mpearon
    #>

Function Set-SEPShellConstants{
    [cmdletBinding()]
    Param(
    )
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12;
    $Global:sepBaseURI = 'https://SEPServer:PORT/sepm/api/v1'
}
Function Get-sepAuthorization{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2017-10-21 18:28
            Keywords:   Symantec Endpoint Protection Manager, SEPM, SEP, API, Authorization, Token, Username, Password, Session, Headers
        .COMPONENT
            Symantec Endpoint Protection Manager, SEPM, SEP, API, Authorization, Token, Username, Password, Session, Headers
        .SYNOPSIS
            Acquires access token from SEPM API
        .DESCRIPTION
            Get-sepAuthorization returns a custom object containing sepSessionHeaders (Authorization=Bearer) and a sepWebSession.
            Other functions call this function at runtime, so it is not necessary to call at launch.
        .PARAMETER sepAlternateCreds
            This optional PSCredential parameter allows the user to provide credentials to the function.  If omitted, the user will be prompted.
        .EXAMPLE
            Get-sepAuthorization
            This will prompt the user for credentials, returning an object containing sepSessionHeaders (Authorization=Bearer) and a sepWebSession.
        .EXAMPLE
            Get-sepAuthorization -sepAlternateCreds $otherCredentials
            This will accept a PSCredential object, returning an object containing sepSessionHeaders (Authorization=Bearer) and a sepWebSession.
        .LINK
            https://github.com/mpearon
    #>
    [cmdletbinding()]
    Param(
        [PSCredential]$sepAlternateCreds
    )
    Set-SEPShellConstants
    if($sepAlternateCreds){
        $Global:sepCreds = $sepAlternateCreds
    }
    if($sepCreds -eq $null){
        $Global:sepCreds = (Get-Credential -Credential (($env:USERNAME).ToLower()))
    }
    $sepAuthorizationURI = '/identity/authenticate'
    $sepAuthorizationBody = @{
        username = ($sepCreds.username)
        password = ($sepCreds.GetNetworkCredential().Password)
        domain   = ('')
    }
    Write-Verbose ('Requesting token')
    Try{
        $sepReturn = Invoke-RestMethod -Uri ( -join ($sepBaseURI,$sepAuthorizationURI)) -Method POST -Body ($sepAuthorizationBody | ConvertTo-JSON) -ContentType 'application/json' -SessionVariable 'sepWebSession'
        Write-Verbose ('Token acquired')
    }
    Catch{
        Write-Verbose ('Unable to acquire token')
        Write-Host -ForegroundColor Red ('Unable to acquire token')
        Write-Host $error[0]
        break
    }
    $sepAuthorizationObject = [PSCustomObject]@{
        sepSessionHeaders = @{
            'Authorization' = ( -join ('Bearer ',$($sepReturn).Token))
        }
        sepRefreshToken = @{
            'Authorization' = ( -join ('Bearer ',$($sepReturn).refreshToken))
        }
        sepWebSession     = $sepWebSession
    }
    return ($sepAuthorizationObject)
}

Function Get-sepAdministrators{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2017-10-21 18:39
            Keywords:   Symantec Endpoint Protection Manager, SEPM, SEP, API, Administrators
        .COMPONENT
            Symantec Endpoint Protection Manager, SEPM, SEP, API, Administrators
        .SYNOPSIS
            Acquires list of administrative user accounts
        .DESCRIPTION
            Get-sepAdministrators returns a list of Administrative user accounts via the SEPM API
        .EXAMPLE
            get-sepAdministrators
            This will return an array of objects, listing the account details of SEP administrators
        .LINK
            https://github.com/mpearon
    #>
    [cmdletbinding()]
    Param()
    Set-SEPShellConstants
    $sepAuthorizationObject = Get-sepAuthorization
    $sepAdministratorsURI = '/admin-users'
    Try{
        $sepAdministrators = Invoke-RestMethod -Uri ( -join ($sepBaseURI,$sepAdministratorsURI)) -Headers ($sepAuthorizationObject.sepSessionHeaders)
        Write-Verbose ('Administrators Acquired')
    }
    Catch{
        Write-Verbose ( -join ('Unable to GET ',$sepAdministratorsURI))
        Write-Host -ForegroundColor Red ('Unable to acquire Administrators.')
        Write-Host -ForegroundColor Red $error[0]
        break
    }
    return ($sepAdministrators)
}
Function Get-sepComputers{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2017-10-21 18:59
            Keywords:   Symantec Endpoint Protection Manager, SEPM, SEP, API, Computers
        .COMPONENT
            Symantec Endpoint Protection Manager, SEPM, SEP, API, Computers
        .SYNOPSIS
            Acquires list of enrolled computers
        .DESCRIPTION
            get-sepComputers returns a list of SEP-enrolled computers via the SEPM API.
        .PARAMETER computerName
            This optional pipeline array parameter allows the user to target specific computers
        .EXAMPLE
            Get-sepComputers
            This will return an array of objects containing all of the SEP-enrolled computers.
        .EXAMPLE
            Get-sepComputers -computerName 'Computer1'
            This will return the specifics of Computer1 (using named parameter).
        .EXAMPLE
            Get-sepComputers -computerName 'Computer1','Computer2'
            This will return the specifics of Computer1 and Computer2 (using named parameter).
        .EXAMPLE
            'Computer2' | Get-sepComputers
            This will return the specifics of Computer1 (using pipeline parameter).
        .EXAMPLE
            'Computer2','Computer2' | Get-sepComputers
            This will return the specifics of Computer1 and Computer2 (using pipeline parameter).
        .LINK
            https://github.com/mpearon
    #>
    [cmdletbinding()]
    Param(
        [Parameter(ValueFromPipeline = $true)]$computerName
    )

    Begin{
        Set-SEPShellConstants
        $sepAuthorizationObject = Get-sepAuthorization
        $replaceHash = @{
            'osname'         = 'lowerOsName'
            'osflavorNumber' = 'lowerOsFlavorNumber'
            'osfunction'     = 'lowerOsFunction'
            'osmajor'        = 'lowerOsMajor'
            'osservicePack'  = 'lowerOsServicePack'
            'oslanguage'     = 'lowerOsLanguage'
            'osbitness'      = 'lowerOsBitness'
            'osminor'        = 'lowerOsMinor'
            'osversion'      = 'lowerOsVersion'
        }
    }
    Process{
        $computerName | ForEach-Object{
            Write-Verbose (-join('count: ',$computerName.count))
            if($computerName.count -lt 1 ){
                Try{
                    $sepComputers = Do{
                        $currentPage++
                        $sepComputersURI = ( -join ('/computers?pageSize=1000&pageIndex=',$currentPage))
                        Write-Verbose ( -join ('Page ',$currentPage,': ',$sepComputersURI))

                        $sepComputersResults = (Invoke-RestMethod -Uri ( -join ($sepBaseURI,$sepComputersURI)) -Headers ($sepAuthorizationObject.sepSessionHeaders) -ContentType application/json -UseBasicParsing)
                        $replaceHash.GetEnumerator() | ForEach-Object{
                            $sepComputersResults = $sepComputersResults.replace($($_.Key),$($_.Value))
                        }
                        $pageResults = ($sepComputersResults | ConvertFrom-JSON)
                        $totalPages = ($pageResults.totalPages)
                        $pageResults.content
                    }
                    Until($currentPage -ge $totalPages)
                }
                Catch{
                    Write-Verbose ( -join ('Unable to GET ',$sepComputersURI))
                    Write-Host -ForegroundColor Red ('Unable to acquire Computers.')
                    Write-Host -ForegroundColor Red $error[0]
                }
            }
            else{
                $sepComputers = $computerName | ForEach-Object{
                    Try{
                        $sepComputerURI = ( -join ('/computers?computerName=',$_))
                        Write-Verbose $sepComputerURI
                        $sepComputer = Invoke-RestMethod -Uri ( -join ($sepBaseURI,$sepComputerURI)) -Headers ($sepAuthorizationObject.sepSessionHeaders) -ContentType application/json -UseBasicParsing
                        $replaceHash.GetEnumerator() | ForEach-Object{
                            $sepComputer = $sepComputer.replace($($_.Key),$($_.Value))
                        }
                        ($sepComputer | ConvertFrom-JSON).Content
                    }
                    Catch{
                        Write-Verbose ( -join ('Unable to GET ',$sepComputerURI))
                        Write-Host -ForegroundColor Red ( -join ('Unable to acquire ',$computerName))
                        Write-Host -ForegroundColor Red $error[0]
                    }
                }
            }
        }
    }
    End{
        return ($sepComputers | Select-Object computerName,ipAddresses,macAddresses,logonUserName,@{Name='infected';Expression={ switch($_.infected){ '0'{$false}; '1'{$true} } }},@{Name='edrEnrolled';Expression={ switch($_.edrStatus){ '2'{$True}; default{$false} } }},@{Name='apOn';Expression={ switch($_.apOnOff){ '0'{$false}; '1'{$true} } }},@{Name='avEngineOn';Expression={ switch($_.avEngineOnOff){ '0'{$false}; '1'{$true} } }},@{Name='firewallOn';Expression={ switch($_.firewallOnOff){ '0'{$false}; '1'{$true} } }},profileVersion,profileSerialNo,@{Name='online';Expression={ switch($_.onlineStatus){ '0'{$false}; '1'{$true} } }},hardwareKey,@{Name='lastDeploymentDate';Expression={(Get-Date '1970-01-01').AddMilliseconds($_.lastDeploymentTime)}},@{Name='xgroup';Expression={$_.group.name}})
    }
}
Function Get-sepGroups{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2017-10-27 07:52
            Keywords:   Symantec Endpoint Protection Manager, SEPM, SEP, API, Groups
        .COMPONENT
            Symantec Endpoint Protection Manager, SEPM, SEP, API, Groups
        .SYNOPSIS
            Acquires list of Computer Groups
        .DESCRIPTION
            Get-sepGroups returns an object array of SEP Computer Groups via the SEPM API
        .EXAMPLE
            Get-sepGroups
            This will return an array of objects containing the SEP Groups
        .LINK
            https://github.com/mpearon
    #>
    [cmdletbinding()]
    Param(
        [Switch]$showAll
    )
    Set-SEPShellConstants
    $sepAuthorizationObject = Get-sepAuthorization
    $sepGroupsURI = '/groups'
    Try{
        $sepGroups = Invoke-RestMethod -Uri ( -join ($sepBaseURI,$sepGroupsURI)) -Headers ($sepAuthorizationObject.sepSessionHeaders)
        Write-Verbose ('Groups Acquired')
    }
    Catch{
        Write-Verbose ( -join ('Unable to GET ',$sepGroupsURI))
        Write-Host -ForegroundColor Red ('Unable to acquire Groups.')
        Write-Host -ForegroundColor Red $error[0]
        break
    }
    switch($showAll){
        $true   { $minimumComputerCount = '0' }
        $false  { $minimumComputerCount = '1' }
    }
    return (($sepGroups).Content | Where-Object { $_.numberOfPhysicalComputers -ge $minimumComputerCount } | Select-Object name,id,numberOfPhysicalComputers)
}
Function Get-sepPolicy{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2017-10-27 07:52
            Keywords:   Symantec Endpoint Protection Manager, SEPM, SEP, API, Computers, Groups
        .COMPONENT
            Symantec Endpoint Protection Manager, SEPM, SEP, API, Computers, Groups
        .SYNOPSIS
            Moves SEP-enrolled computer from one SEP group to another
        .DESCRIPTION
            Move-sepComputers relocates a SEP computer asset  Groups via the SEPM API
        .EXAMPLE
            Move-sepComputers
            Using the function without any parameters will cause the user to be present with lists from which to select target computer(s) and a target group.
        .EXAMPLE
            Move-sepComputers -computerName 'Test-VM'
            This use case will cause the user to be presented with a list from which to select the target group.
        .EXAMPLE
            Move-sepComputers -groupID '9876543210'
            This use case will cause the user to be presented with a list from which to select the computer(s).
        .EXAMPLE
            Move-sepComputers -computerName 'Test-VM' -groupID '9876543210'
            This use case will not cause any prompts.
        .LINK
            https://github.com/mpearon
    #>
    [cmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline = $true)]$policy
    )
    Begin{
        Set-SEPShellConstants
    }
    Process{
        $policy | ForEach-Object{
            $sepAuthorizationObject = Get-sepAuthorization
            if($policy){
                Try{
                    $sepPolicyURI = (-join('/policies/exceptions/',$policy))
                    $sepPolicy = Invoke-RestMethod -Uri ( -join ($sepBaseURI,$sepPolicyURI)) -Headers ($sepAuthorizationObject.sepSessionHeaders)
                    Write-Verbose ('Policies Acquired')
                }
                Catch{
                    Write-Verbose ( -join ('Unable to GET ',$sepPolicyURI))
                    Write-Host -ForegroundColor Red ('Unable to acquire Policies.')
                    Write-Host -ForegroundColor Red $error[0]
                    break
                }
            }
            else{
                Try{
                    $sepPolicyURI = '/policies/summary'
                    $sepPolicy = Invoke-RestMethod -Uri ( -join ($sepBaseURI,$sepPolicyURI)) -Headers ($sepAuthorizationObject.sepSessionHeaders)
                    Write-Verbose ('Policies Acquired')
                }
                Catch{
                    Write-Verbose ( -join ('Unable to GET ',$sepPolicyURI))
                    Write-Host -ForegroundColor Red ('Unable to acquire Policies.')
                    Write-Host -ForegroundColor Red $error[0]
                    break
                }
            }
            Return ($sepPolicy).Content
        }
    }
    End{
    }
}
Function Move-sepComputer{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2017-10-27 07:52
            Keywords:   Symantec Endpoint Protection Manager, SEPM, SEP, API, Computers, Groups
        .COMPONENT
            Symantec Endpoint Protection Manager, SEPM, SEP, API, Computers, Groups
        .SYNOPSIS
            Moves SEP-enrolled computer from one SEP group to another
        .DESCRIPTION
            Move-sepComputers relocates a SEP computer asset  Groups via the SEPM API
        .EXAMPLE
            Move-sepComputers
            Using the function without any parameters will cause the user to be present with lists from which to select target computer(s) and a target group.
        .EXAMPLE
            Move-sepComputers -computerName 'Test-VM'
            This use case will cause the user to be presented with a list from which to select the target group.
        .EXAMPLE
            Move-sepComputers -groupID '9876543210'
            This use case will cause the user to be presented with a list from which to select the computer(s).
        .EXAMPLE
            Move-sepComputers -computerName 'Test-VM' -groupID '9876543210'
            This use case will not cause any prompts.
        .LINK
            https://github.com/mpearon
    #>
    [CmdletBinding(DefaultParameterSetName = 'sepComputerObject')]
    Param(
        [PSCustomObject][Parameter(Mandatory,ValueFromPipeline,ParameterSetName = "sepComputerObject", Position = 0)]$sepComputerObject,
        [Parameter(Mandatory,ValueFromPipeline,ParameterSetName = "sepComputerString", Position = 0)]$computerName,
        [String][Parameter(Position = 1)]$groupID,
        [Switch][Parameter(Position = 2)]$test,
        [Switch][Parameter(Position = 2)]$listUpdate
    )
    Begin{
        Set-SEPShellConstants
        $jsonMove = @()
        if(($PSBoundParameters.listUpdate) -eq $false){
            Write-Host -ForegroundColor Yellow 'Not Updating List'
        }
        else{
            Write-Host -ForegroundColor Yellow 'Updating List'
            $global:sepComputersHash = @{}
            Get-sepComputers | ForEach-Object{
                $sepComputersHash[$_.computerName] = $_.hardwareKey
            }
            Write-Host 'List Updated'
        }
        if(-not ($PSBoundParameters.groupID)){
            $groupID = (Get-sepGroups -showAll | Out-GridView -Title 'Select Target Group' -PassThru).id
        }
        $sepAuthorizationObject = Get-sepAuthorization
        $sepComputerMoveURI = '/computers'
    }
    Process{
        if($PSCmdlet.ParameterSetName -eq 'sepComputerObject'){
            Write-Verbose 'sepComputerObject'
            $jsonMove += $sepComputerObject | ForEach-Object{
                if($_.computerName){
                    $thisComputer = $_.computerName
                }
                else{
                    $thisComputer = $_
                }
                @{
                    group       = @{
                        id = $groupID
                    }
                    hardwareKey = $sepComputersHash.$thisComputer
                }
            }
        }
        else{
            Write-Verbose 'sepComputerName'
            $jsonMove += $computerName | ForEach-Object{
                @{
                    group       = @{
                        id = $groupID
                    }
                    hardwareKey = $sepComputersHash.$_
                }
            }
        } 
    }
    End{
        $finalJson = $jsonMove | ConvertTo-Json
        if($finalJson -notmatch '^\['){
            $finalJson = ( -join ('[',$finalJson,']'))
        }
        if($test){
            return $finalJson
        }
        else{
            $sepComputerMove = Invoke-RestMethod -Uri ( -join ($sepBaseURI,$sepComputerMoveURI)) -Headers ($sepAuthorizationObject.sepSessionHeaders) -ContentType 'application/json' -Method Patch -Body $finalJson
            return ($sepComputerMove.responseMessage)
        }
    }
}