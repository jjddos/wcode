#import necessary modules 
Import-Module ServerManager

#Define vars for log path, server details, directory paths and file filters for import
$LogPath = "C:\CMS\Get-WindowsServices.log"
[string] $CMSServer = "SGP1PCNCMSSRV01"
[string] $Centraldb = "CentralDB"
$ImportDirectoryName = "C:\CMS\Import\"
$fileFilter = "WindowsServices_*.csv"
$OutputDirectoryName = "C:\CMS\Output\ServicesInfo"
$cmsInventoryFile = "C:\CMS\Output\Svr.Inventory_$($env:computername).csv" 
$currentTime = Get-Date -Format ddMMMyyyy-HHmm

#function to write logs
function Write-Log {   
            #region Parameters
                    [cmdletbinding()]
                    Param(
                            [Parameter(ValueFromPipeline=$true,Mandatory=$true)] [ValidateNotNullOrEmpty()]
                            [string] $Message,
                            [Parameter()] [ValidateSet(“Error”, “Warn”, “Info”)]
                            [string] $Level = “Info”,
                            [Parameter()]
                            [Switch] $NoConsoleOut,
                            [Parameter()]
                            [String] $ConsoleForeground = 'White',
                            [Parameter()] [ValidateRange(1,30)]
                            [Int16] $Indent = 0,     
                            [Parameter()]
                            [IO.FileInfo] $Path = $LogPath,                           
                            [Parameter()]
                            [Switch] $Clobber,                          
                            [Parameter()]
                            [String] $EventLogName,                          
                            [Parameter()]
                            [String] $EventSource,                         
                            [Parameter()]
                            [Int32] $EventID = 1,
                            [Parameter()]
                            [String] $LogEncoding = "ASCII"                         
                    )                   
            #endregion
            Begin {}
            Process {
                    try {                  
                            $msg = '{0}{1} : {2} : {3}' -f (" " * $Indent), (Get-Date -Format “yyyy-MM-dd HH:mm:ss”), $Level.ToUpper(), $Message                           
                            if ($NoConsoleOut -eq $false) {
                                    switch ($Level) {
                                            'Error' { Write-Error $Message }
                                            'Warn' { Write-Warning $Message }
                                            'Info' { Write-Host ('{0}{1}' -f (" " * $Indent), $Message) -ForegroundColor $ConsoleForeground}
                                    }
                            }
                            if ($Clobber) {
                                    $msg | Out-File -FilePath $Path -Encoding $LogEncoding -Force
                            } else {
                                    $msg | Out-File -FilePath $Path -Encoding $LogEncoding -Append
                            }
                            if ($EventLogName) {
                           
                                    if (-not $EventSource) {
                                            $EventSource = ([IO.FileInfo] $MyInvocation.ScriptName).Name
                                    }
                           
                                    if(-not [Diagnostics.EventLog]::SourceExists($EventSource)) {
                                            [Diagnostics.EventLog]::CreateEventSource($EventSource, $EventLogName)
                            }
                                $log = New-Object System.Diagnostics.EventLog  
                                $log.set_log($EventLogName)  
                                $log.set_source($EventSource)                       
                                    switch ($Level) {
                                            “Error” { $log.WriteEntry($Message, 'Error', $EventID) }
                                            “Warn”  { $log.WriteEntry($Message, 'Warning', $EventID) }
                                            “Info”  { $log.WriteEntry($Message, 'Information', $EventID) }
                                    }
                            }
                    } catch {
                            throw “Failed to create log entry in: ‘$Path’. The error was: ‘$_’.”
                    }
            }    
            End {}    
            
    }

#function to execute script blocks on multiple comps in parallel 
function Run-Jobs($ComputerList,$ScriptBlock) {
    Write-Host ([string]::Format("Running Jobs on {0} computers",
         [int]$ComputerList.Count))
    
    $Job = Invoke-Command -ComputerName $ComputerList -ScriptBlock $Scriptblock -AsJob
    $terminateAt = (Get-Date).AddMinutes(15)

    do {
        $waitingFor = $null
        $waitingFor = [array]($Job.ChildJobs | where {$_.State -ne "Completed" -and $_.State -ne "Failed"})

        if ([int]$waitingFor.Count -gt 0) {
            Write-Host ([string]::Format("Waiting for {0} {1} to finish", [int]$waitingFor.Count, $(if ([int]$waitingFor.Count -eq 1) {"job"} else {"jobs"})))
            Start-Sleep 15
        }
    }
    until ((Get-Date) -ge $terminateAt -or [int]$waitingFor.Count -eq 0)

    if ((Get-Date) -ge $terminateAt) {
        $Job | Stop-Job
    }

    $Data = Receive-Job $job -ErrorAction SilentlyContinue

    # Get the list of failed jobs, remove the error (if any) for those jobs, then rerun the jobs
    # using the -IncludePortInSPN session option

    $FailedJobs = $Job.Childjobs | where {$_.State -ne "Completed"} | select -ExpandProperty Location

    $Job.Dispose()

    if ($FailedJobs -ne $null) {

        foreach ($f in $FailedJobs) {
            $ErrorStringToFind = [string]::Format("Connecting to remote server {0} failed*",
                $f)
            $ErrorToRemove = $Error | where {$_.Exception.Message -like $ErrorStringToFind}
            if ($ErrorToRemove -ne $null) {
                $Error.Remove($ErrorToRemove)
            }
        }

        $NewJob = Invoke-Command -ComputerName $FailedJobs -ScriptBlock $Scriptblock -SessionOption (New-PSSessionOption -IncludePortInSPN) -AsJob
        $terminateAt = (Get-Date).AddMinutes(5)

        Write-Host "Re-running failed jobs with the -IncludePortInSPN session option"

        do {
            $waitingFor = $null
            $waitingFor = [array]($NewJob.ChildJobs | where {$_.State -ne "Completed" -and $_.State -ne "Failed"})

            if ([int]$waitingFor.Count -gt 0) {
                Write-Host ([string]::Format("Waiting for {0} {1} to finish", [int]$waitingFor.Count, $(if ([int]$waitingFor.Count -eq 1) {"job"} else {"jobs"})))
                Start-Sleep 15
            }
        }
        until ((Get-Date) -ge $terminateAt -or [int]$waitingFor.Count -eq 0)

        if ((Get-Date) -ge $terminateAt) {
            $NewJob | Stop-Job
        }

        $Data += Receive-Job $NewJob -ErrorAction SilentlyContinue

        $FailedJobs = $null
        $FailedJobs = $NewJob.Childjobs | where {$_.State -ne "Completed"} | select -ExpandProperty Location
        $NewJob.Dispose()
    }

    Write-Output $Data
}

#function to clean SQL inputs (')
function Sanitize-SQLInput([string]$inputValue){

    $inputValue = $inputValue -replace "'","''"

    $nonsecureKeywords = @('SELECT', 'INSERT', 'DROP', 'UPDATE', 'DELETE', 'EXEC', '--', ';')
    foreach($keyword in $nonsecureKeywords){
        $inputValue = $inputValue -replace $keyword,""
    }
    return $inputValue
}

Write-Log -Message "Script Started at $(Get-Date)  " -NoConsoleOut -Clobber -Path $LogPath

#gets services on each machine and checks permissions of exec and folder
$ScriptBlock = { 
$resultTask = @()
$resultService = @()
$LogPath = "C:\CMS\Get-WindowsServices.log"

#checks the format of service image paths
function Test-ImagePath($path)
{
    
    if($path.StartsWith('"') -or $path.StartsWith("\??")) {return $true}
    if(-not $path.Contains(" ")){return $true}
    $execPath = $path -split ' -| /', 2 | Select-Object -First 1
    $execPath = $execPath.Trim(" ")
    $segments = $execPath -split '\\(?=[^\\]+$)'
    if($segments[-1].Contains(' ') -and -not $segments[1].StartsWith('"')){
        return $false
    }
    
    return $true
}

#gets permissions of a file or folder
function get-Permissions($path, $groupsToCheck, $isFile) {
   
    $acl = Get-Acl $path 
    if($acl){
        :get foreach($permission in $acl.Access)
        {
            $identity = $permission.IdentityReference
             
            if($groupsToCheck -contains $identity)
            {
                
                switch ($permission.FileSystemRights){
                'ReadAndExecute, Synchronize' { continue get }
                '-1610612736' { continue get }
                'FullControl'{return $false}
                '268435456' {return $false}
                Default { 
                    if(-not $isFile){ continue get}
                    else{ return $false }}
                }
            }
        }
        return $true
    }
    else{return $false}
    
}

#corrects the file paths by replacing systemroot, adding C:windows, removing \??\
function Check-BinaryPath($path)
{
    $defaultSystemRoot = $env:SystemRoot+'\\'
    $path = $path -replace '^\\\?\?\\','' -replace '%windir%','\systemroot' -replace '%systemroot%','\systemroot'
    if($path -match '^\\systemroot\\'){
        $path = $path -replace '^\\systemroot\\', $defaultSystemRoot
    }

    if($path -match '^system32\\' -and -not ($path -match '^[a-zA-Z]:')){
        $path = $defaultSystemRoot + $path
    }
    
    return $path
}

#groups for permission checks 
$groupsToCheck = @("Everyone", "BUILTIN\Users", "$env:USERDOMAIN\Domain users", "NT AUTHORITY\Authenticated Users")
#exe possible names to catch
$exeToCheck = @("powershell.exe", "pwsh.exe", "pwsh", "powershell")
#pattern to identify file extensions .exe .sys
$extpattern = '(\.\w+)(?=\s|$)'

#opening registry key to get service info
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$env:COMPUTERNAME) 
$regkey = $reg.OpenSubKey('SYSTEM\CurrentControlSet\Services')
$services = $regkey.GetSubKeyNames()

$schldTasks = Get-ScheduledTask

#iteration for each scheduled task
    foreach($task in $schldTasks)
    {
        $sctaskObject = "" | select Domain,Computer,TaskName,State,Executable,ExecPass,FolderPass,AllPass,Comments
        $sctaskObject.TaskName = $task.TaskName
        $sctaskObject.State = $task.State
       # $sctaskObject.Executable = $task.Actions.Execute
        $sctaskObject.AllPass = 1
        $sctaskObject.Computer = $env:COMPUTERNAME
        $sctaskObject.Domain = $env:USERDOMAIN
        $sctaskObject.Comments = ''
        $failedGroups = @()
        $failedFolders = @()
        
        $psfailedGroups = @()
        $psfailedFolders = @()
        $psComments = @()
        $psFailPath = @()
        :get foreach($action in $task.Actions)
        {
            $taskexe = $action.Execute

            if([string]::IsNullOrEmpty($taskexe)){
                $sctaskObject.Comments = "Scheduled task has null executable path" 
                $sctaskObject.AllPass = "NA"
                $sctaskObject.ExecPass = "NA"
                $sctaskObject.FolderPass = "NA"
                $resultTask += $sctaskObject
                continue get
            }

            $taskexecPath= Check-BinaryPath -path $taskexe
            $taskexeFilename = [System.IO.Path]::GetFileName(($taskexecPath))

            if($taskexecPath -eq $taskexeFilename){
                $taskexecPath = where.exe $taskexecPath
            }

            if(Test-Path $taskexecPath){
                $taskexecFolder = [System.IO.Path]::GetDirectoryName($taskexecPath)
                if(($exeToCheck -contains $taskexeFilename) -and $action.Arguments){
                        $psPaths = $action.Arguments -split ' ' | Where-Object {$_.EndsWith('.ps1')}
                        $psComments += "Scheduled task has action to start powershell. $($action.Arguments)"
                }
                foreach($groupname in $groupsToCheck){
                    $result = get-Permissions -path $taskexecPath -groupsToCheck $groupname -isFile $true
                    $folderResult = get-Permissions -path $taskexecFolder -groupsToCheck $groupname -isFile $false

                    foreach($pspath in $psPaths)
                    {
                        if(Test-Path $pspath){
                            $psFolder = [System.IO.Path]::GetDirectoryName($pspath)
                            $psresult = get-Permissions -path $pspath -groupsToCheck $groupname -isFile $true
                            $psfolderResult = get-Permissions -path $psFolder -groupsToCheck $groupname -isFile $false

                            if(-not $psfolderResult){
                                $sctaskObject.AllPass = 0
                                $psfailedFolders += $groupname
                            }
                            if(-not $psresult){
                                $psfailedGroups += $groupname
                                $sctaskObject.AllPass = 0 
                            }
                                
                        }
                        else{
                            $sctaskObject.AllPass = 0
                            $psFailPath += $pspath
                        }
                    }

                    if(-not $folderResult){
                        $sctaskObject.AllPass = 0
                        $failedFolders += $groupname 
                    }
                    if(-not $result){
                        $failedGroups += $groupname
                        $sctaskObject.AllPass = 0 
                    }
                }
            }
            else{
                $sctaskObject.AllPass = 0
                $sctaskObject.Comments = "The specfied path ''$taskexecPath'' does not exist. Please check the path."
            }
        }

        $sctaskObject.ExecPass = [int](-not ($psfailedGroups.Count + $failedGroups.Count))
        $sctaskObject.FolderPass = [int](-not ($psfailedFolders.Count + $failedFolders.Count))
        
        $psComments += ($psFailPath.Count -gt 0),"" -ne "The specfied path $($psFailPath | Select-Object -Unique | ForEach-Object { $_ -join ', '}) does not exist. Please check the path."
        if(-not $sctaskObject.ExecPass -or -not $sctaskObject.FolderPass){
            if([int](-not ($failedGroups.Count + $failedFolders.Count)) -gt 0) {$sctaskObject.Comments = "Groups not passed: $($failedGroups -join ', '), Folders not passed: $($failedFolders -join ', ')"}
            if([int](-not ($psfailedGroups.Count + $psfailedFolders.Count)) -gt 0){
            $psComments += "Groups not passed for arguments: $($psfailedGroups -join ', '), Folders not passed for arguments: $($psfailedFolders -join ', ')"
            }
        }
        $sctaskObject.Comments += $psComments
        $resultTask += $sctaskObject
    }

#iteration for each service located in registry to check path permissions
    foreach($service in $services)
    {
        $serviceObject = "" | select Domain,Computer,ServiceName,DisplayName,ImagePath,ExecPass,FolderPass,AllPass,Comments
        $serviceObject.ServiceName = $service
        $serviceObject.DisplayName = $($regkey.OpenSubKey($service).GetValue('DisplayName'))
        $serviceObject.AllPass = 1
        $serviceObject.Computer = $env:COMPUTERNAME
        $serviceObject.Domain = $env:USERDOMAIN
        $serviceObject.Comments = ''

        $failedGroups = @()
        $failedFolders = @()
        $serviceBinaryPath = $($regkey.OpenSubKey($service).GetValue('ImagePath'))
        $serviceObject.ImagePath = $serviceBinaryPath
        
        #for some services the imagepath might be empty, no need to check permissions
        if([string]::IsNullOrEmpty($serviceBinaryPath)){
            $serviceObject.Comments = "Service has null Imagepath" 
            $serviceObject.AllPass = "NA"
            $serviceObject.ExecPass = "NA"
            $serviceObject.FolderPass = "NA"
            $resultService += $serviceObject
            continue
        }

        #if imagepath is ok, then checks the permissions
        if(Test-ImagePath -path $serviceBinaryPath){
            $serviceBinaryPath = $serviceBinaryPath.Replace('"','')
            $binaryExt = [regex]::Match($serviceBinaryPath,$extpattern).Value
            if($binaryExt){
                #removes all data after the file extension  eg. -k blablabla
                $extPos = $serviceBinaryPath.LastIndexOf($binaryExt)
                $binaryPath = $serviceBinaryPath.Substring(0,$extPos + $binaryExt.Length)
                $binaryPath= Check-BinaryPath -path $binaryPath
                #gets the folder of file
                $binaryFolder = [System.IO.Path]::GetDirectoryName($binaryPath)
                if(Test-Path $binaryPath){
                    foreach($groupname in $groupsToCheck){
                        $result = get-Permissions -path $binaryPath -groupsToCheck $groupname -isFile $true
                        $folderResult = get-Permissions -path $binaryFolder -groupsToCheck $groupname -isFile $false

                        if(-not $folderResult){
                            $serviceObject.AllPass = 0
                            $failedFolders += $groupname 
                        }
                        if(-not $result){
                            $failedGroups += $groupname
                            $serviceObject.AllPass = 0 
                        }
                    }
                }
                else{
                    $serviceObject.AllPass = 0
                    $serviceObject.Comments = "The specfied path ''$binaryPath'' does not exist. Please check the path."
                }
                
            }
        }
        else {
            $serviceObject.AllPass = 0
            $serviceObject.Comments = "Image path is incorrect and it has failed the Test-ImagePath"
        }

        $serviceObject.ExecPass = [int](-not $failedGroups.Count)
        $serviceObject.FolderPass = [int](-not $failedFolders.Count)
        if(-not $serviceObject.ExecPass -or -not $serviceObject.FolderPass){
            $serviceObject.Comments = "Groups not passed: $($failedGroups -join ', '), Folders not passed: $($failedFolders -join ', ')"
        }
        $resultService += $serviceObject
    }

    #returns the accumulated results from the scriptblock
    $returnValue = New-Object PSObject -Property @{
        resultTask = $resultTask 
        resultService = $resultService
    }

    return $returnValue
}


#try to get active computer inventory from Svr.Inventory_(env:computername).csv or sql server 
try {
    $colComputers = $null
     #if the script runs in SGP1PCNCMSSRV01 then get all inventory from DB, for other cases need to get inventory using the csv file.
    if($CMSServer -ne $env:COMPUTERNAME){
        Write-Log -Message "Getting computer list from Inventory csv..." -NoConsoleOut -Path $LogPath
        if(Test-Path $cmsInventoryFile){
        $inventory = Import-Csv "C:\CMS\Output\Svr.Inventory_$($env:computername).csv" | Where {$_.Active -eq 1 }
        $colComputers = $inventory.Name
        }
        else {$colComputers = $env:COMPUTERNAME} 
        
    }
    else {
        Write-Log -Message "Running the query to get Inventory..." -NoConsoleOut -Path $LogPath
        $querys = "SELECT Name FROM [CentralDB].[Svr].[Inventory] WHERE [Active]=1 AND [Domain]='SGPCN.LOCAL' and name='SGP1PCNCMSSRV01'"
        $inventory = Invoke-Sqlcmd -ServerInstance $CMSServer -Database $env:CentralDB -Query $querys
        $colComputers = $inventory.Name
    }

    #if no computers are found
    if ($colComputers -eq $null) {
        Write-Log -Message  "There are either no servers in the database, or there was a problem reading from the database.`nWriting Log"  -NoConsoleOut -Path $LogPath
        exit
    } 
    else {
         Write-Log -Message  ([string]::Format("Script set to run on {0} servers",$colComputers.count)) -NoConsoleOut -Path $LogPath
         #logging the total number of PCs 
         }
    }
catch {       
    $colComputers = "ERROR"  
    $ex = $_.Exception
    write-log -Level Error -Message "$ex.Message on while pulling inventory data"  -NoConsoleOut -Path $LogPath 
    Write-Log -Level Error -Message "Failed to retrieve Inventory List" -NoConsoleOut -Path $LogPath
}

#runs the scriptblock 
$servicesScan  = Run-Jobs $colComputers $Scriptblock

try {
if($CMSServer -ne $env:COMPUTERNAME){
     #Export $returnValue into C:\CMS\Import with name starting as WindowsServices_
    $servicesScan | Export-Csv -Path "C:\CMS\Import\WindowsServices_$($env:USERDOMAIN).csv" -Force -notypeinformation -Append
    Write-Log -Message "Service info file has been exported to the output folder" -NoConsoleOut -Path $LogPath
   
    #get data from all WindowsServices_*.csv files exported from OTS/SGPCN/PMCS/DMZ 
    $directory = Get-ChildItem -File -Filter $fileFilter -Path $ImportDirectoryName
    $fileDirectory = $directory | Select-Object -ExpandProperty Fullname
    $importAll = $fileDirectory | Import-Csv

    #import files which have problems with check
    $importFiles = $importAll  | Where-Object  {$_.AllPass -ne '1'}
    Write-Log -Message "Imported $($importFiles.Count) services from the WindowsServices csv file" -NoConsoleOut -Path $LogPath
    
    Move-Item -path "$ImportDirectoryName$fileFilter" -Destination $OutputDirectoryName 
    Get-ChildItem -File -Filter $fileFilter -Path $OutputDirectoryName | Rename-Item -NewName {$_.BaseName+ '_' + $currentTime + $_.Extension -replace 'WindowsServices', 'WS' }
    Write-Log -Message " $($importAll.Count) Services info file has been moved" -NoConsoleOut -Path $LogPath
    $importAll | Export-Csv -Path "C:\CMS\Reports\WindowsServices_$currentTime.xlsx"
    #move csv files to C:\CMS\Output\ServicesInfo with changed name - WS_<DomainName>_<currenttime>.csv - to keep for additional info
    Write-Log -Message " $($importAll.Count) Services info file has been exported to the report folder" -NoConsoleOut -Path $LogPath

  try{
    #loop the imported data from the result of scriptblock
    foreach ($d in $importFiles){

        #if the record doesn't exist in DB table then need to insert, else just need to update. Checks by servicename and servername 
        $sqlCommand = $null
        $sqlCommand = [string]::Format("MERGE INTO [ServicePermission] As Target
                USING (SELECT '{2}' AS ServiceName, '{1}' AS ServerName) AS Source 
                    ON Target.ServiceName = Source.ServiceName 
                    AND Target.ServerName = Source.ServerName
                    WHEN MATCHED THEN 
                        UPDATE SET  [ImagePath] = '{4}', [FileExecPass] = '{5}',[FolderPass] = '{6}',[AllPass] = '{7}',[Comments] = '{8}', ScanDate = GETDATE()
                    WHEN NOT MATCHED BY TARGET THEN 
                        INSERT ([ServerID],[Domain],[ServerName],[ServiceName],[DisplayName],[ImagePath],[FileExecPass],[FolderPass],[AllPass],[Comments])
	    		VALUES 
                ((SELECT ID FROM [CentralDB].[Svr].[Inventory] WHERE [Active]=1 AND NAME='{1}' ),'{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}');",
                $d.Domain, $d.Computer, $d.ServiceName, (Sanitize-SQLInput -Input $d.DisplayName), $d.ImagePath,$d.ExecPass, $d.FolderPass, $d.AllPass, $d.Comments
            )

        Invoke-Sqlcmd -ServerInstance $CMSServer -Database $env:CentralDB -Query $sqlCommand

        Write-Log -Message "Service $($d.ServiceName) in $($d.Computer) has been inserted" -NoConsoleOut -Path $LogPath
        
    }}
        catch
        {$colComputers = "ERROR"  
    $ex = $_.Exception
    write-log -Level Error -Message "$ex  "  -NoConsoleOut -Path $LogPath 
    Write-Log -Level Error -Message "Failed to push data" -NoConsoleOut -Path $LogPath
    }}
else{
        #this is for ECS/OTS/PMCS/DMZ domains
        #Export $returnValue into C:\CMS\Output with name starting as WindowsShares_ 
        $filepath ="C:\CMS\Output\WindowsTasks_$($env:COMPUTERNAME).csv"
        if(Test-Path -path $filepath){ Remove-Item -path $filepath -Force }
        $servicesScan.resultTask | Export-Csv -Path $filepath -Force -notypeinformation -Append
        Write-Log -Message "Scheduled Tasks info have been exported to the output folder" -NoConsoleOut -Path $LogPath

        $filepath ="C:\CMS\Output\WindowsServices_$($env:COMPUTERNAME).csv"
        if(Test-Path -path $filepath){ Remove-Item -path $filepath -Force }
        $servicesScan.resultService | Export-Csv -Path $filepath -Force -notypeinformation -Append
        Write-Log -Message "Services info have been exported to the output folder" -NoConsoleOut -Path $LogPath
    }
}
catch{
    $colComputers = "ERROR"   
    $ex = $_.Exception
    write-log -Level Error -Message "$ex. While pushing data"  -NoConsoleOut -Path $LogPath 
    Write-Log -Level Error -Message "Failed to export and insert data into DB" -NoConsoleOut -Path $LogPath}


Write-Log -Message "Script Ended at $(Get-Date) "  -NoConsoleOut -Path $LogPath
