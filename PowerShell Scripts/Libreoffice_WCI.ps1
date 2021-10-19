#Variables of Argument 1, Argument 2, $exe, $path, and a little ascii of Zero Two ^^
$param1=$args[0]
$param2=$args[1]
$exe = "C:\Program Files\LibreOffice\program\"
$path = "$HOME\Desktop"
$ascii="⣿⣿⣿⣿⣯⣿⣿⠄⢠⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡟⠈⣿⣿⣿⣿⣿⣿⣆⠄
⢻⣿⣿⣿⣾⣿⢿⣢⣞⣿⣿⣿⣿⣷⣶⣿⣯⣟⣿⢿⡇⢃⢻⣿⣿⣿⣿⣿⢿⡄
⠄⢿⣿⣯⣏⣿⣿⣿⡟⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⣧⣾⢿⣮⣿⣿⣿⣿⣾⣷
⠄⣈⣽⢾⣿⣿⣿⣟⣄⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⣝⣯⢿⣿⣿⣿⣿
⣿⠟⣫⢸⣿⢿⣿⣾⣿⢿⣿⣿⢻⣿⣿⣿⢿⣿⣿⣿⢸⣿⣼⣿⣿⣿⣿⣿⣿⣿
⡟⢸⣟⢸⣿⠸⣷⣝⢻⠘⣿⣿⢸⢿⣿⣿⠄⣿⣿⣿⡆⢿⣿⣼⣿⣿⣿⣿⢹⣿
⡇⣿⡿⣿⣿⢟⠛⠛⠿⡢⢻⣿⣾⣞⣿⡏⠖⢸⣿⢣⣷⡸⣇⣿⣿⣿⢼⡿⣿⣿
⣡⢿⡷⣿⣿⣾⣿⣷⣶⣮⣄⣿⣏⣸⣻⣃⠭⠄⠛⠙⠛⠳⠋⣿⣿⣇⠙⣿⢸⣿
⠫⣿⣧⣿⣿⣿⣿⣿⣿⣿⣿⣿⠻⣿⣾⣿⣿⣿⣿⣿⣿⣿⣷⣿⣿⣹⢷⣿⡼⠋
⠄⠸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣷⣦⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡟⣿⣿⣿⠄⠄
⠄⠄⢻⢹⣿⠸⣿⣿⣿⣿⣿⣷⣿⣿⣿⣿⣿⣿⣿⣿⣿⡿⣼⣿⣿⣿⣿⡟⠄⠄
⠄⠄⠈⢸⣿⠄⠙⢿⣿⣿⣹⣿⣿⣿⣿⣟⡃⣽⣿⣿⡟⠁⣿⣿⢻⣿⣿⢿⠄⠄
⠄⠄⠄⠘⣿⡄⠄⠄⠙⢿⣿⣿⣾⣿⣷⣿⣿⣿⠟⠁⠄⠄⣿⣿⣾⣿⡟⣿⠄⠄
⠄⠄⠄⠄⢻⡇⠸⣆⠄⠄⠈⠻⣿⡿⠿⠛⠉⠄⠄⠄⠄⢸⣿⣇⣿⣿⢿⣿⠄"


#Test argument if they are Writer, Calc, Impress, or All
if (($param1 -inotlike "writer") -and ($param1 -inotlike "calc") -and ($param1 -inotlike "impress") -and ($param1 -inotlike "all")){
Write-Host "";Write-host -BackgroundColor Red "[ERROR]"
    Write-host -BackgroundColor Red  "Please, as second argument only enter one of this options:";Write-Host ""
    Write-Host -BackgroundColor Red "1. Writer";Write-Host -BackgroundColor Red "2. Calc";Write-Host -BackgroundColor Red "3. Impress";Write-Host -BackgroundColor Red "4. All";"";$ascii;exit}

#Test if argument 2 is create or delete
if (($param2 -inotlike "create") -and ($param2 -inotlike "delete")){
    "";Write-host -BackgroundColor Red "[ERROR]"
    Write-host -BackgroundColor Red  "Please, enter as second argument one option";"";
    Write-host -BackgroundColor Red  "1. Create"
    Write-host -BackgroundColor Red  "2. Delete"
    ;"";$ascii;exit
}


#Hard part but it ended up being a copy paste

#Writer 
if (($param1 -ilike "writer") -and ($param2 -ilike "create")){""

	$file = Get-ChildItem $path | Where-Object Name -Like "Writer.lnk" ; if (!!$file){ Write-Host "Writer Shortcut exists"}
    else{Write-Host -BackgroundColor Red "Writer Shortcut doesen't exists, so it has been created"}
	$exepath = "C:\Program Files\LibreOffice\program\swriter.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Writer Executable exists"}else{Write-Host -BackgroundColor Red "Writer Executable  doesen't exists"}
    $SourceFilePath = "$exe\swriter.exe";$ShortcutPath = "$HOME\Desktop\Writer.lnk";$WScriptObj = New-Object -ComObject ("WScript.Shell")
    $shortcut = $WscriptObj.CreateShortcut($ShortcutPath);$shortcut.TargetPath = $SourceFilePath;$shortcut.Save()

}elseif(($param1 -ilike "Writer") -and ($param2 -ilike "delete")){""

    $file = Get-ChildItem $path | Where-Object Name -Like "Writer.lnk" ; if (!!$file){ Write-Host "Writer Shortcut exists, so it's gonna be deleted"; Remove-Item $HOME\Desktop\writer.lnk}
    else{Write-Host -BackgroundColor Red "Writer Shortcut doesen't exists"}
	$exepath = "C:\Program Files\LibreOffice\program\swriter.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Writer Executable exists"}else{Write-Host -BackgroundColor Red "Writer Executable  doesen't exists"}}


#Calc
if (($param1 -ilike "calc") -and ($param2 -ilike "create")){""

	$file = Get-ChildItem $path | Where-Object Name -Like "Calc.lnk" ; if (!!$file){ Write-Host "Calc Shortcut exists"}
    else{Write-Host -BackgroundColor Red "Calc Shortcut doesen't exists, so it has been created"}
	$exepath = "C:\Program Files\LibreOffice\program\scalc.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Calc Executable exists"}else{Write-Host -BackgroundColor Red "Calc Executable  doesen't exists"}
    $SourceFilePath = "$exe\scalc.exe";$ShortcutPath = "$HOME\Desktop\Calc.lnk";$WScriptObj = New-Object -ComObject ("WScript.Shell")
    $shortcut = $WscriptObj.CreateShortcut($ShortcutPath);$shortcut.TargetPath = $SourceFilePath;$shortcut.Save()

}elseif(($param1 -ilike "calc") -and ($param2 -ilike "delete")){""

    $file = Get-ChildItem $path | Where-Object Name -Like "Calc.lnk" ; if (!!$file){ Write-Host "Calc Shortcut exists, so it's gonna be deleted"; Remove-Item $HOME\Desktop\Calc.lnk}
    else{Write-Host -BackgroundColor Red "Calc Shortcut doesen't exists"}
	$exepath = "C:\Program Files\LibreOffice\program\scalc.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Calc Executable exists"}else{Write-Host -BackgroundColor Red "Calc Executable  doesen't exists"}}


#Impress
if (($param1 -ilike "impress") -and ($param2 -ilike "create")){""

	$file = Get-ChildItem $path | Where-Object Name -Like "Impress.lnk" ; if (!!$file){ Write-Host "Impress Shortcut exists"}
    else{Write-Host -BackgroundColor Red "Impress Shortcut doesen't exists, so it has been created"}
	$exepath = "C:\Program Files\LibreOffice\program\simpress.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Impress Executable exists"}else{Write-Host -BackgroundColor Red "Impress Executable  doesen't exists"}
    $SourceFilePath = "$exe\simpress.exe";$ShortcutPath = "$HOME\Desktop\Impress.lnk";$WScriptObj = New-Object -ComObject ("WScript.Shell")
    $shortcut = $WscriptObj.CreateShortcut($ShortcutPath);$shortcut.TargetPath = $SourceFilePath;$shortcut.Save()

}elseif(($param1 -ilike "impress") -and ($param2 -ilike "delete")){""

    $file = Get-ChildItem $path | Where-Object Name -Like "Impress.lnk" ; if (!!$file){ Write-Host "Impress Shortcut exists, so it's gonna be deleted"; Remove-Item $HOME\Desktop\Impress.lnk}
    else{Write-Host -BackgroundColor Red "Impress Shortcut doesen't exists"}
	$exepath = "C:\Program Files\LibreOffice\program\simpress.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Impress Executable exists"}else{Write-Host -BackgroundColor Red "Impress Executable  doesen't exists"}}


if (($param1 -ilike "all") -and ($param2 -ilike "create")){""

    #Writer Create
	$file = Get-ChildItem $path | Where-Object Name -Like "Writer.lnk" ; if (!!$file){ Write-Host "Writer Shortcut exists"}
    else{Write-Host -BackgroundColor Red "Writer Shortcut doesen't exists, so it has been created"}
	$exepath = "C:\Program Files\LibreOffice\program\swriter.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Writer Executable exists"}else{Write-Host -BackgroundColor Red "Writer Executable  doesen't exists"}
    $SourceFilePath = "$exe\swriter.exe";$ShortcutPath = "$HOME\Desktop\Writer.lnk";$WScriptObj = New-Object -ComObject ("WScript.Shell")
    $shortcut = $WscriptObj.CreateShortcut($ShortcutPath);$shortcut.TargetPath = $SourceFilePath;$shortcut.Save()

    #Calc Create
    $file = Get-ChildItem $path | Where-Object Name -Like "Calc.lnk" ; if (!!$file){ Write-Host "Calc Shortcut exists"}
    else{Write-Host -BackgroundColor Red "Calc Shortcut doesen't exists, so it has been created"}
	$exepath = "C:\Program Files\LibreOffice\program\scalc.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Calc Executable exists"}else{Write-Host -BackgroundColor Red "Calc Executable  doesen't exists"}
    $SourceFilePath = "$exe\scalc.exe";$ShortcutPath = "$HOME\Desktop\Calc.lnk";$WScriptObj = New-Object -ComObject ("WScript.Shell")
    $shortcut = $WscriptObj.CreateShortcut($ShortcutPath);$shortcut.TargetPath = $SourceFilePath;$shortcut.Save()

    #Impress Create
    $file = Get-ChildItem $path | Where-Object Name -Like "Impress.lnk" ; if (!!$file){ Write-Host "Impress Shortcut exists"}
    else{Write-Host -BackgroundColor Red "Impress Shortcut doesen't exists, so it has been created"}
	$exepath = "C:\Program Files\LibreOffice\program\simpress.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Impress Executable exists"}else{Write-Host -BackgroundColor Red "Impress Executable  doesen't exists"}
    $SourceFilePath = "$exe\simpress.exe";$ShortcutPath = "$HOME\Desktop\Impress.lnk";$WScriptObj = New-Object -ComObject ("WScript.Shell")
    $shortcut = $WscriptObj.CreateShortcut($ShortcutPath);$shortcut.TargetPath = $SourceFilePath;$shortcut.Save()


}elseif(($param1 -ilike "all") -and ($param2 -ilike "delete")){""

    #Writer Delete
    $file = Get-ChildItem $path | Where-Object Name -Like "Writer.lnk" ; if (!!$file){ Write-Host "Writer Shortcut exists, so it's gonna be deleted"; Remove-Item $HOME\Desktop\writer.lnk}
    else{Write-Host -BackgroundColor Red "Writer Shortcut doesen't exists"}
	$exepath = "C:\Program Files\LibreOffice\program\swriter.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Writer Executable exists"}else{Write-Host -BackgroundColor Red "Writer Executable  doesen't exists"}

    #Calc Delete
    $file = Get-ChildItem $path | Where-Object Name -Like "Calc.lnk" ; if (!!$file){ Write-Host "Calc Shortcut exists, so it's gonna be deleted"; Remove-Item $HOME\Desktop\Calc.lnk}
    else{Write-Host -BackgroundColor Red "Calc Shortcut doesen't exists"}
	$exepath = "C:\Program Files\LibreOffice\program\scalc.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Calc Executable exists"}else{Write-Host -BackgroundColor Red "Calc Executable  doesen't exists"}

    #Impress Create
    $file = Get-ChildItem $path | Where-Object Name -Like "Impress.lnk" ; if (!!$file){ Write-Host "Impress Shortcut exists, so it's gonna be deleted"; Remove-Item $HOME\Desktop\Impress.lnk}
    else{Write-Host -BackgroundColor Red "Impress Shortcut doesen't exists"}
	$exepath = "C:\Program Files\LibreOffice\program\simpress.exe"
	$test = Test-Path $exepath -PathType Leaf ; if ($test){Write-Host "Impress Executable exists"}else{Write-Host -BackgroundColor Red "Impress Executable  doesen't exists"}}

"";$ascii
