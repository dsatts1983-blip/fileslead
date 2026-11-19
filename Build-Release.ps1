$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $csc)) { throw 'The Windows .NET Framework C# compiler was not found.' }
$exe = Join-Path $root 'WebsitesDoIt-Lead-Finder.exe'
& $csc /nologo /target:winexe /platform:anycpu /optimize+ /reference:System.Windows.Forms.dll /out:$exe /resource:"$(Join-Path $root 'LeadFinder.ps1'),LeadFinder.ps1" (Join-Path $root 'Launcher.cs')
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $exe)) { throw 'The Windows executable could not be built.' }
Write-Host "Built $exe"
