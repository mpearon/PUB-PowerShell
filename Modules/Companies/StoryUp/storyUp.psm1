Function New-ProjectDirectory{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2018-04-03 11:58
        .COMPONENT
            Project, Directory Tree
        .SYNOPSIS
            Creates project directory tree
        .DESCRIPTION
            New-ProjectDirectory creates a project-specific directory, and then
            creates several nested directories
        .PARAMETER projectTitle
            This mandatory string-type parameter allows the user to specify a
            name for the parent directory
        .PARAMETER basePath
            This mandatory string-type parameter allows the user to specify a
            base path for the directory tree
        .PARAMETER force
            This switch-type parameter allows the user to overwrite any existing
            directories. If directories have contents they will not be 
            overwritten
        .EXAMPLE
            New-ProjectDirectory -projectName 'AstonMartin' -basePath 'C:\Temp'
            This will create an AstonMartin directory in C:\Temp, and create
            nested directories
        .EXAMPLE
            New-ProjectDirectory -projectName 'AstonMartin','Jaguar' -basePath 'C:\Temp'
            This will create an AstonMartin and a Jaguar directory in C:\Temp, 
            then create nested directories in each
        .EXAMPLE
            'AstonMartin' | New-ProjectDirectory -basePath 'C:\Temp'
            This will create an AstonMartin directory in C:\Temp, and create
            then create nested directories in each
        .EXAMPLE
            'AstonMartin','Jaguar' | New-ProjectDirectory -basePath 'C:\Temp'
            This will create an AstonMartin and a Jaguar directory in C:\Temp, 
            and create nested directories
        .LINK
            https://github.com/mpearon
        .LINK
            https://twitter.com/@mpearon
    #>
    [cmdletBinding()]
    Param(
        [Parameter(ValueFromPipeline = $true, Position = 0, Mandatory = $true)][string]$projectTitle,
        [Parameter(Position = 1)]$basePath = 'E:\Users\Ricky\OneDrive - StoryUp XR\Projects',
        $directoryArray = ('3ds','Quixel','FBXs','Reference','Substances','Texture'),
        [switch]$force
    )
    Begin{
    }
    Process{
        $projectTitle | ForEach-Object{
            if($force){
                Try{
                    $projectDirectory = New-Item -Path $basePath -Name $_ -ItemType Directory -ErrorAction Stop -Force -Confirm:$false
                }
                Catch{
                    Write-Host -ForegroundColor Red (-join('[!] FAILURE: Unable to create directory: ',$_))
                    return (-join('[!] FAILURE: Unable to create directory: ',$_))
                }
                $directoryArray | ForEach-Object{
                    Try{
                        New-Item -Path $projectDirectory -ItemType Directory -Name $_ -ErrorAction Stop -Force -Confirm:$false
                    }
                    Catch{
                        Write-Host -ForegroundColor Red (-join('[!] FAILURE: Unable to create directory: ',$_))
                        return (-join('[!] FAILURE: Unable to create directory: ',$_))
                    }
                }
                #No failures detected
                return 0
            }
            else{
                Try{
                    $projectDirectory = New-Item -Path $basePath -Name $_ -ItemType Directory -ErrorAction Stop
                }
                Catch{
                    Write-Host -ForegroundColor Red (-join('[!] FAILURE: Unable to create directory: ',$projectDirectory))
                }
                $directoryArray | ForEach-Object{
                    Try{
                        New-Item -Path $projectDirectory -ItemType Directory -Name $_ -ErrorAction Stop
                    }
                    Catch{
                        Write-Host -ForegroundColor Red (-join('[!] Unable to create directory: ',$_))
                        return (-join('[!] FAILURE: Unable to create directory: ',$_))
                    }
                }
                #No failures detected
                return 0
            }
        }
    }
    End{
    }
}

Function Export-AsVideo{
    <#
        .NOTES
            Author:     Matthew A. Pearon
            Date:       2018-04-10 19:19
        .COMPONENT
            StoryUP, Handbrake, After Effects, Render, Mux
        .SYNOPSIS
            Creates MP4 from PNG Sequence
        .DESCRIPTION
            Export-AsVideo compiles an MP4 from a PNG sequence
        .PARAMETER inputPath
            This string-type parameter allows the user to provide a path
            containing a set of PNGs
        .PARAMETER workingPath
            This string-type parameter allows the user to provide a path that
            the script will use as a temporary path
        .PARAMETER outputPath
            This string-type parameter allows the user to provide a path that
            will collect the final output
        .EXAMPLE
            Export-AsVideo
            The simplest use case uses the default values of all parameters
        .EXAMPLE
            Export-AsVideo -inputPath C:\Temp\input
            This overloads the default value of the $inputPath parameter
        .EXAMPLE
            Export-AsVideo -workingPath D:\test
            This overloads the default value of the $workingPath parameter
        .EXAMPLE
            Export-AsVideo -outputPath C:\renders\results
            This overloads the default value of the $outputPath parameter
        .EXAMPLE
            Export-AsVideo -inputPath C:\temp\input -workingPath C:\temp\working -outputPath C:\temp\output
            This overloads the default value of all parameters
        .LINK
            http://docs.aenhancers.com/
        .LINK
            https://github.com/mpearon
        .LINK
            https://twitter.com/@mpearon
    #>
    [cmdletBinding()]
    Param(
        $inputPath = 'C:\temp\forRicky\pngSeq\input',
        $workingPath = 'C:\temp\forRicky\pngSeq\working',
        $outputPath = 'C:\temp\forRicky\pngSeq\output'
    )
    #Region > Constants >
    $afterEffectsPath = 'C:\Program Files\Adobe\Adobe After Effects CS5\Support Files\AfterFX.exe'
    $handbrakeCLIPath = 'C:\Program Files\HandBrake\HandBrakeCLI.exe'
    #EndRegion < Constants <
    #Region > Pre-flight >
        Try{
            Write-Host -ForegroundColor Yellow '[#] PROCESS: Pre-flight Checks - ' -NoNewline
            if(Test-Path $afterEffectsPath -ErrorAction Stop){
                # afterEffectsPath exists - nothing to do
            }
            else{
                Write-Host -ForegroundColor Red '[!] FAILURE: After Effects not present'
                return '[!] FAILURE: After Effects not present'
            }
            if(Test-Path $handbrakeCLIPath -ErrorAction Stop){
                Try{
                    Unblock-File -Path $handbrakeCLIPath
                }
                Catch{
                    Write-Host -ForegroundColor Red '[!] Error: Unable to Unblock HandbrakeCLI'
                    return '[!] Error: Unable to Unblock HandbrakeCLI'
                }
            }
            else{
                Write-Host -ForegroundColor Red '[!] FAILURE: HandbrakeCLI not present'
                return '[!] FAILURE: HandbrakeCLI not present'
            }
            if(Test-Path $inputPath -ErrorAction Stop){
                if(Get-ChildItem $inputPath -ErrorAction Stop){
                    # inputPath contains files - nothing to do
                }
                else{
                    Write-Host -ForegroundColor Red '[!] FAILURE: Input Path is empty'
                    return '[!] FAILURE: Input Path is empty'
                }
            }
            else{
                Write-Host -ForegroundColor Red '[!] FAILURE: Input Path does not exist'
                return '[!] FAILURE: Input Path does not exist'
            }
            if(Test-Path $workingPath -ErrorAction Stop){
                if(Get-ChildItem $workingPath -ErrorAction Stop){
                    Write-Host -ForegroundColor Red '[!] FAILURE: Working Path must be empty'
                    return '[!] FAILURE: Working Path must be empty'
                }
                else{
                    # workingPath empty - nothing to do
                }
            }
            else{
                Write-Host -ForegroundColor Red '[!] FAILURE: Working Path does not exist'
                return '[!] FAILURE: Working Path does not exist'
            }
            if(Test-Path $outputPath -ErrorAction Stop){
                if(Get-ChildItem $outputPath -ErrorAction Stop){
                    Write-Host -ForegroundColor Red '[!] FAILURE: Output Path must be empty'
                    return '[!] FAILURE: Output Path must be empty'
                }
                else{
                    Write-Host -ForegroundColor Green 'Complete'
                }
            }
            else{
                Write-Host -ForegroundColor Red '[!] FAILURE: Output Path does not exist'
                return '[!] FAILURE: Output Path does not exist'
            }
        }
        Catch{
            Write-Host -ForegroundColor Red '[!] FAILURE: Unable to complete pre-flight checks'
            return '[!] FAILURE: Unable to complete pre-flight checks'
        }
    #EndRegion < Pre-flight <
    #Region > Build PNG Sequence >
    $jsxBody = @"
    var inputDirectory = Folder("$( $inputPath.replace('\','//') )");
    if (inputDirectory != null){
        var fileList = inputDirectory.getFiles();
    }
    var importOptions = new ImportOptions(fileList[1]);
    importOptions.sequence = true;
    app.project.importFile(importOptions);
    var thisItem = app.project.item(1)
    var currentComposition = app.project.items.addComp( thisItem.name+'-comp', thisItem.width, thisItem.height, thisItem.pixelAspect,thisItem.duration,thisItem.frameRate);
    currentComposition.layers.add(thisItem);
    var currentRender = app.project.renderQueue.items.add(currentComposition);
    currentRender.outputModule(1).file = new File("$( (-join($workingPath.replace('\','//'),'//out.avi')) )");
    app.project.renderQueue.render();
    app.project.close(CloseOptions.DO_NOT_SAVE_CHANGES);
    app.quit(); 
"@
    $jsxPath = (-join($workingPath,'\renderScript.jsx'))
    Try{
        Write-Host -ForegroundColor Yellow '[#] PROCESS: Dynamically creating JSX - ' -NoNewline
        New-Item -itemType File -Path $jsxPath -Value $jsxBody -Force -ErrorAction Stop | Out-Null
        Write-Host -ForegroundColor Green 'Complete'
    }
    Catch{
        Write-Host -ForegroundColor Red '[!] FAILURE: Unable to create JSX'
        return '[!] FAILURE: Unable to create JSX'
    }
    Try{
        Write-Host -ForegroundColor Yellow '[#] PROCESS: Starting After Effects - ' -NoNewline
        Start-Process $afterEffectsPath -ArgumentList (-join('-r ',$jsxPath)) -Wait -ErrorAction Stop
        Write-Host -ForegroundColor Green 'Complete'
    }
    Catch{
        Write-Host -ForegroundColor Red '[!] FAILURE: Unable to launch After Effects'
        return '[!] FAILURE: Unable to launch After Effects'
    }
    #EndRegion < Build PNG Sequence <
    #Region > HandBrake >
    $handbrakeArguments = (-join('-i ',$workingPath,'\out.avi -o ',$outputPath,'\out.mp4 --no-markers --width "1920" --height "1080" --preset "Fast 1080p30" --encoder "x264"'))
    Try{
        Write-Host -ForegroundColor Yellow '[#] PROCESS: Starting Handbrake - ' -NoNewline
        #Start-Process $handbrakeCLIPath -ArgumentList $handbrakeArguments -NoNewWindow -ErrorAction Stop -Wait
        Start-Process $handbrakeCLIPath -ArgumentList $handbrakeArguments -ErrorAction Stop -Wait
        Write-Host -ForegroundColor Green 'Complete'
    }
    Catch{
        Write-Host -ForegroundColor Red '[!] FAILURE: Unable to launch HandBrakeCLI'
        return '[!] FAILURE: Unable to launch HandBrakeCLI'
    }
    #EndRegion < HandBrake <
    #Region > HandBrake >
    Try{
        Write-Host -ForegroundColor Yellow '[#] PROCESS: Starting Cleanup - ' -NoNewline
        Get-ChildItem $workingPath -ErrorAction Stop | Remove-Item -Force -ErrorAction Stop
        Write-Host -ForegroundColor Green 'Complete'
    }
    Catch{
        Write-Host -ForegroundColor Red '[!] FAILURE: Unable to perform Cleanup'
        return '[!] WARNING: Unable to perform Cleanup'
    }
    #EndRegion < Cleanup <
    
    #No failures detected
    return 0
}