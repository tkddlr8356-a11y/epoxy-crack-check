param(
    [Parameter(Mandatory = $true)][string]$CasesCsv,   # same input list as solve_case1.ps1 (bilinear rows only)
    [Parameter(Mandatory = $true)][string]$OutCsv,     # same result schema as solve_case1.ps1 (method = ANSYS)
    [Parameter(Mandatory = $true)][string]$WorkDir,    # one sub-folder per case is created here (fresh folder recommended)
    [switch]$KeepBig,                                   # keep .db/.rst etc. (default: deleted after extraction)
    [switch]$DryRun                                     # only write ds.dat / run.inp per case (no solve) - for checking
)
# ANSYS MAPDL batch run of the user's Case 1 model (Case 3.wbpz system G, written out as ds.dat):
# concrete plane-stress strip 6,000 x 200 x 1 mm (4 bands of 50 mm), right end = 4 nonlinear COMBIN39 springs
# (one per band, 50 mm2), uniform temperature -20 C, both ends fixed.
# Per case the 4 spring tables are replaced by the bilinear epoxy curve: point 2 (yield) and point 3 (end).
# Only E_c (MP,EX,1) and the crack width w (spring displacement = strain x w) can be changed; L_c 6000 mm,
# dT -20 C, alpha 1e-5 are fixed in the model. One licence seat -> cases run one after another (~6 s each).
$ErrorActionPreference = "Stop"
$inv = [Globalization.CultureInfo]::InvariantCulture
$enc = [Text.Encoding]::GetEncoding(1252)
$exe = "C:\Program Files\ANSYS Inc\v251\ANSYS\bin\winx64\ANSYS251.exe"
if (-not (Test-Path -LiteralPath $exe)) { throw "ANSYS 2025 R1 not found at $exe" }
$asset = Join-Path $PSScriptRoot "..\assets"
$dsText = [IO.File]::ReadAllText((Join-Path $asset "ds_case1.dat"), $enc)
$post = [IO.File]::ReadAllText((Join-Path $asset "post_case1.inp"))
function Num($v, $def) { if ($v -eq $null -or "$v".Trim() -eq '') { return [double]$def }; return [double]::Parse("$v".Trim(), $inv) }
function Has($r, $n) { return ($r.PSObject.Properties[$n] -ne $null) -and ("$($r.$n)".Trim() -ne '') }
New-Item -ItemType Directory -Force $WorkDir | Out-Null

$rows = Import-Csv -LiteralPath $CasesCsv -Encoding UTF8
$out = New-Object System.Collections.ArrayList
foreach ($r in $rows) {
    if (Has $r 'curve') { throw "$($r.name): the ANSYS template holds a 3-point (bilinear) spring table only - use solve_case1.ps1 for multi-point curves" }
    if ((Has $r 'Lc') -and (Num $r.Lc 6000) -ne 6000) { throw "$($r.name): L_c is fixed at 6000 mm in the ANSYS model" }
    if ((Has $r 'dT') -and [math]::Abs((Num $r.dT 20)) -ne 20) { throw "$($r.name): dT is fixed at -20 C in the ANSYS model" }
    $sy = Num $r.sy 0; $fu = Num $r.fu 0; $E1 = Num $r.E1 100; $eu = (Num $r.eu 100) / 100.0
    $fck = $(if (Has $r 'fck') { Num $r.fck 0 } else { 0 })
    $Ec = $(if (Has $r 'Ec') { Num $r.Ec 0 } elseif ($fck -gt 0) { 8500 * [math]::Pow($fck + 4, 1.0 / 3.0) } else { 30000 })
    $limits = @()
    if (Has $r 'ft') { foreach ($p in ("$($r.ft)" -split '[;/ ]+')) { if ($p.Trim()) { $limits += [double]::Parse($p.Trim(), $inv) } } }
    elseif ($fck -gt 0) { $limits += [math]::Round(0.30 * [math]::Pow($fck, 2.0 / 3.0), 2) }
    $w = Num $r.w 0.2
    $ey = $sy / $E1
    $l3 = "rmod,_sid,3," + ($ey * $w).ToString("0.000000e+00", $inv) + "," + ($sy * 50).ToString("0.000000e+00", $inv)
    $l5 = "rmod,_sid,5," + ($eu * $w).ToString("0.000000e+00", $inv) + "," + ($fu * 50).ToString("0.000000e+00", $inv)
    $txt = [regex]::Replace($dsText, '(?m)^rmod,_sid,3,[^\r\n]*', $l3)
    $txt = [regex]::Replace($txt, '(?m)^rmod,_sid,5,[^\r\n]*', $l5)
    $txt = [regex]::Replace($txt, '(?m)^MP,EX,1,[^,\r\n]*,', ('MP,EX,1,' + $Ec.ToString('R', $inv) + ','))
    if (([regex]::Matches($txt, [regex]::Escape($l3))).Count -ne 4) { throw "$($r.name): spring tables not replaced (template changed?)" }
    $safe = ($r.name -replace '[^\w\-\.]', '_')
    $d = Join-Path $WorkDir $safe
    New-Item -ItemType Directory -Force $d | Out-Null
    [IO.File]::WriteAllText("$d\ds.dat", $txt, $enc)
    [IO.File]::WriteAllText("$d\run.inp", "/INPUT,ds,dat`r`n" + $post, [Text.Encoding]::ASCII)
    if ($DryRun) { "{0}: prepared {1} (yield {2} / end {3})" -f $r.name, $d, $l3, $l5; continue }
    if (Test-Path "$d\results.txt") { Remove-Item -LiteralPath "$d\results.txt" }
    $t0 = Get-Date
    $p = Start-Process -FilePath $exe -ArgumentList "-b", "-np", "2", "-dir", "`"$d`"", "-j", "G", "-i", "`"$d\run.inp`"", "-o", "`"$d\G.out`"" -PassThru -NoNewWindow -Wait
    $secs = [int]((Get-Date) - $t0).TotalSeconds
    $sigs = @{}; $F = @{}; $ux = @{}
    if (-not (Test-Path "$d\results.txt")) { throw "$($r.name): no results.txt - check $d\G.out (licence busy? the solver seat allows one run at a time)" }
    foreach ($ln in (Get-Content "$d\results.txt")) {
        $t = $ln.Trim() -split '\s+'
        if ($t[0] -eq 'SIG') { $sigs[[int][double]::Parse($t[1], $inv)] = [double]::Parse($t[2], $inv) }
        if ($t[0] -eq 'SPR') { $kk = [int][double]::Parse($t[1], $inv); $F[$kk] = [double]::Parse($t[2], $inv); $ux[$kk] = [double]::Parse($t[3], $inv) }
    }
    $g = Get-Content "$d\G.out" -ErrorAction SilentlyContinue
    $nErr = @($g | Select-String -Pattern '^\s*\*\*\* ERROR').Count
    $nNc = @($g | Select-String -Pattern 'SOLUTION NOT CONVERGED|NOT CONVERGED AFTER|DID NOT CONVERGE').Count
    if (-not $KeepBig) { Get-ChildItem -LiteralPath $d -File | Where-Object { $_.Extension -in '.db', '.rdb', '.rst', '.ldhi', '.ce', '.DSP', '.cnd', '.gst', '.mntr' } | Remove-Item -Force }
    $sigMax = ($sigs.Values | Measure-Object -Maximum).Maximum
    $strain = 100 * (-($ux.Values | Measure-Object -Average).Average) / $w
    $rupture = $strain / 100 -gt $eu + 1e-9
    $verd = @()
    foreach ($ft in $limits) {
        $near = [math]::Abs($sigMax - $ft) -le 0.005
        $crack = ($sigMax -gt $ft) -and -not $near
        $c1 = $(if ($near) { '한계' } elseif ($crack) { '균열' } else { '' })
        $verd += ($ft.ToString('0.0##', $inv) + ':' + $(if ($c1 -and $rupture) { $c1 + '·파단' } elseif ($c1) { $c1 } elseif ($rupture) { '파단' } else { 'O' }))
    }
    $o = [ordered]@{
        name = $r.name; method = 'ANSYS (MAPDL)'
        sy = $sy; fu = $fu; E1 = $E1; eu_pct = 100 * $eu; E2 = ($fu - $sy) / ($eu - $ey)
        curve = ((100 * $ey).ToString('0.###', $inv) + ':' + $sy + ';' + (100 * $eu).ToString('0.###', $inv) + ':' + $fu)
        Ec = $Ec; fck = $(if ($fck -gt 0) { $fck } else { '' }); w = $w; Lc = 6000; dT = 20; alpha = 1e-5; cal = ''
        sigma_c0 = ''; k = ''
        sig1 = $sigs[1]; sig2 = $sigs[2]; sig3 = $sigs[3]; sig4 = $sigs[4]; sig5 = $sigs[5]
        F1 = $F[1]; F2 = $F[2]; F3 = $F[3]; F4 = $F[4]; Ftotal = ($F[1] + $F[2] + $F[3] + $F[4])
        strain_pct = $strain; branch = $(if ($strain / 100 -lt $ey) { '탄성 (항복 전)' } elseif ($rupture) { '끝점 너머 (외삽)' } else { '항복 후' })
        rupture = $(if ($rupture) { '파단' } else { '-' })
        limits = (($limits | ForEach-Object { $_.ToString('0.0##', $inv) }) -join ';')
        verdicts = ($verd -join ';'); emin = ''
        ansys_secs = $secs; ansys_errors = $nErr; ansys_notconv = $nNc
    }
    [void]$out.Add([pscustomobject]$o)
    "{0}: sigma_c = {1:N4} MPa, F1 = {2:N2} N, strain = {3:N2} %, {4} | {5} ({6} s, errors {7}, not-converged {8})" -f $r.name, $sigMax, $F[1], $strain, $o.rupture, $o.verdicts, $secs, $nErr, $nNc
}
$out | Export-Csv -LiteralPath $OutCsv -NoTypeInformation -Encoding UTF8
"saved $OutCsv ($($out.Count) cases)"
