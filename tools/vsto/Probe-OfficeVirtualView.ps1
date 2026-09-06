param([Parameter(Mandatory=$true)][string]$OutputPath)
$key='HKCU:\Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto'
[pscustomobject]@{
 ProcessId=$PID
 LoadBehavior=(Get-ItemProperty $key -ErrorAction SilentlyContinue).LoadBehavior
 TrustedKeyExists=(Test-Path "$env:LOCALAPPDATA\ChuanHoa\Development\trusted-key.xml")
 AppVModules=@((Get-Process -Id $PID).Modules | Where-Object ModuleName -match 'AppV' | ForEach-Object ModuleName)
} | Export-Clixml -LiteralPath $OutputPath
