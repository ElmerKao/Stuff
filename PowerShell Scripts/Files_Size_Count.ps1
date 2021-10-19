gci -force 'C:\'-ErrorAction SilentlyContinue | ? { $_ -is [io.directoryinfo] } | % {
$len = 0
$cont=(Get-ChildItem -Force -Recurse -ErrorAction SilentlyContinue  "C:\$_").count
gci -recurse -force $_.fullname -ErrorAction SilentlyContinue | % { $len += $_.length }
$_.fullname, '{0:N2} MB' -f ($len / 1mb)
Write-host " * Files and Directories in $_ $cont";""
}
