param(
    [string]$psex, 
    [string]$endSrv, 
    [string]$endCmp, 
    [string]$comp,
    [string]$us,
    [string]$pass,
    [string]$cmd,
    [string]$dirInst,
    [string]$dirSftw,
    [string]$execNome,
    [string]$dirExec,
    [string]$dirShtc,
    [string]$userShtc
)

$resultOutput = "Success"
$resultCode = "0"

$configInstalation = @{
    psexec = $psex
    endServer = $endSrv
    endComputer = $endCmp
    computer = $comp
    user = $us
    password = $pass
    command = $cmd
    dirInstalation = $dirInst
    dirSoftware = $dirSftw
    executavelNome = $execNome
    dirExecutable = $dirExec
    dirShortcut = $dirShtc
    userShortcut = $userShtc
}


function setShortcut(){
    param (
        [string]$SourceExe,
        [string]$userShortcut
    )
$return = "0"

    try{
        $dirE = $configInstalation['dirExecutable']
        $command1 = '($WScriptObj = New-Object -ComObject ("""WScript.Shell"""))'
        $command2 = '($dir = """C:\Users\' + $userShortcut + '\Desktop\NewShortcut.lnk""")'
        $command3 = '($shortcut = $WscriptObj.CreateShortcut($dir))'
        $command4 = '($shortcut.TargetPath = """' + $dirE + '""")'
        $command5 = '($shortcut.Save())'
        $subCommandShortcut = "Invoke-Command -ScriptBlock {$command1,$command2,$command3, $command4, $command5}"

        $Shortcut = $configInstalation['computer'] + $configInstalation['user'] + $configInstalation['password'] + $configInstalation['command'] + $subCommandShortcut 
        $processShortcut = startProcess -argument $Shortcut  
        
        if($processShortcut[1] -ne "0"){
            $return = "MIS010"
        }        
    }catch{
        $return = "MIS010"
    }
    return $return
}


#Remove substring de string.
function Remove-Substring {
    param(
        [string]$String,
        [string]$Word,
        [int]$StartIndex
    )

    return $String.Remove($String.IndexOf($Word,$StartIndex),$Word.Length)
}

function getAllUsers{
    $return = "0"
    try{
        $subCommandUser = "Get-ChildItem '\Users' -Directory | Format-List -Property Name"

        $User = $configInstalation['computer'] + $configInstalation['user'] + $configInstalation['password'] + $configInstalation['command'] + $subCommandUser 
        $processUser = startProcess -argument $User 
        
        if($processUser -ne "1"){
            $gUser = $processUser[0].Split(" : ")
            $newUser = "0"
            foreach($gU in $gUser){
                if(![string]::IsNullOrWhiteSpace($gU)){
                    if($gU -like "*Name*"){
                        $newUser = Remove-Substring -String $gU -Word "Name"
                    }
                    $resutltSetShortcut = setShortcut -SourceExe $configInstalation['dirExecutable'] -userShortcut $newUser
                    
                    if($resutltSetShortcut -ne "0"){                       
                        $return = $resutltSetShortcut
                        break
                    }                 
                }
            }
        }   
    }catch{
        $return = "MIS010"
        Write-Host $_.ScriptStackTrace
    }
    return $return
}