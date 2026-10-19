param(
    [Parameter(Mandatory = $true)][string]$CasesCsv,   # input list (see SKILL.md "입력 CSV")
    [Parameter(Mandatory = $true)][string]$OutCsv      # result list (read by build_excel.ps1)
)
# Case 1 (관통균열 · 양단 구속 · 균일 온도 하강) closed-form equilibrium.
# Concrete line  : sigma_c = c*Ec*alpha*dT - c*(Ec*w/Lc)*eps_e     (eps_e = epoxy strain, fraction)
# Epoxy curve    : piecewise linear through (0,0) and the given points; beyond the last point the last
#                  segment is extended (same as ANSYS COMBIN39), so a rupture case still gets a number.
# Equilibrium    : sigma_c = sigma_e (n = A_e/A_c = 1). Spring force per spring = sigma x 50 mm2 (4 springs, 200 x 1 mm section).
# c (cal)        : 1.0004 reproduces the ANSYS 2D model (G, Case 3.wbpz) to <= 0.001 MPa; 1.0 = pure 1D theory.
$ErrorActionPreference = "Stop"
$inv = [Globalization.CultureInfo]::InvariantCulture
function Num($v, $def) { if ($v -eq $null -or "$v".Trim() -eq '') { return [double]$def }; return [double]::Parse("$v".Trim(), $inv) }
function Has($r, $n) { return ($r.PSObject.Properties[$n] -ne $null) -and ("$($r.$n)".Trim() -ne '') }

$rows = Import-Csv -LiteralPath $CasesCsv -Encoding UTF8
$out = New-Object System.Collections.ArrayList
foreach ($r in $rows) {
    # ---- concrete ----
    $fck = $(if (Has $r 'fck') { Num $r.fck 0 } else { 0 })
    $Ec = $(if (Has $r 'Ec') { Num $r.Ec 0 } elseif ($fck -gt 0) { 8500 * [math]::Pow($fck + 4, 1.0 / 3.0) } else { 30000 })
    $limits = @()
    if (Has $r 'ft') { foreach ($p in ("$($r.ft)" -split '[;/ ]+')) { if ($p.Trim()) { $limits += [double]::Parse($p.Trim(), $inv) } } }
    elseif ($fck -gt 0) { $limits += [math]::Round(0.30 * [math]::Pow($fck, 2.0 / 3.0), 2) }
    $w = Num $r.w 0.2; $Lc = Num $r.Lc 6000; $dT = [math]::Abs((Num $r.dT 20)); $alpha = Num $r.alpha 1e-5; $cal = Num $r.cal 1.0004
    $s0 = $cal * $Ec * $alpha * $dT          # full-restraint stress
    $k = $cal * $Ec * $w / $Lc               # MPa per unit epoxy strain
    # ---- epoxy curve (strain fraction, stress MPa) ----
    $pts = New-Object System.Collections.ArrayList; [void]$pts.Add(@(0.0, 0.0))
    $bilinear = $false; $sy = $null; $fu = $null; $E1 = $null; $eu = $null
    if (Has $r 'curve') {
        foreach ($pair in ("$($r.curve)" -split ';')) { if ($pair.Trim()) { $xy = $pair.Split(':'); [void]$pts.Add(@(([double]::Parse($xy[0].Trim(), $inv) / 100.0), [double]::Parse($xy[1].Trim(), $inv))) } }
        if ($pts[1][0] -gt 0) { $E1 = $pts[1][1] / $pts[1][0] }
        $sy = $pts[1][1]; $fu = $pts[$pts.Count - 1][1]; $eu = $pts[$pts.Count - 1][0]
        $bilinear = ($pts.Count -eq 3)
    } else {
        $sy = Num $r.sy 0; $fu = Num $r.fu 0; $E1 = Num $r.E1 100; $eu = (Num $r.eu 100) / 100.0
        $ey0 = $sy / $E1
        if ($ey0 -ge $eu) { throw "$($r.name): yield strain $([math]::Round(100*$ey0,3)) % >= end strain" }
        [void]$pts.Add(@($ey0, $sy)); [void]$pts.Add(@($eu, $fu)); $bilinear = $true
    }
    for ($i = 1; $i -lt $pts.Count; $i++) { if ($pts[$i][0] -le $pts[$i - 1][0]) { throw "$($r.name): curve strains must increase" } }
    # ---- intersection ----
    $eps = $null; $seg = -1
    for ($i = 0; $i -lt $pts.Count - 1; $i++) {
        $x0 = $pts[$i][0]; $y0 = $pts[$i][1]; $x1 = $pts[$i + 1][0]; $y1 = $pts[$i + 1][1]
        $g0 = $y0 - ($s0 - $k * $x0); $g1 = $y1 - ($s0 - $k * $x1)
        if ($g0 -le 0 -and $g1 -ge 0) { $sl = ($y1 - $y0) / ($x1 - $x0); $eps = ($s0 - $y0 + $sl * $x0) / ($sl + $k); $seg = $i; break }
    }
    if ($eps -eq $null) {
        $n = $pts.Count; $x0 = $pts[$n - 2][0]; $y0 = $pts[$n - 2][1]; $x1 = $pts[$n - 1][0]; $y1 = $pts[$n - 1][1]
        $sl = ($y1 - $y0) / ($x1 - $x0)
        if ($sl + $k -le 0) { throw "$($r.name): no equilibrium (last segment falls faster than the concrete line)" }
        $eps = ($s0 - $y0 + $sl * $x0) / ($sl + $k); $seg = $n - 2
    }
    $sig = $s0 - $k * $eps
    $rupture = $eps -gt ($eu + 1e-12)
    $E2 = $(if ($bilinear) { ($fu - $sy) / ($eu - $sy / $E1) } else { $null })
    # ---- verdicts and minimum end strain per concrete limit ----
    $verd = @(); $emin = @()
    foreach ($ft in $limits) {
        # within +/-0.005 MPa of the limit = '한계' (borderline; the exact stress is in the table)
        $near = [math]::Abs($sig - $ft) -le 0.005
        $crack = ($sig -gt $ft) -and -not $near
        $c1 = $(if ($near) { '한계' } elseif ($crack) { '균열' } else { '' })
        $v = $(if ($c1 -and $rupture) { $c1 + '·파단' } elseif ($c1) { $c1 } elseif ($rupture) { '파단' } else { 'O' })
        $verd += ($ft.ToString('0.0##', $inv) + ':' + $v)
        if ($bilinear) {
            $ey = $sy / $E1
            if ($sy -ge $ft) { $emin += ($ft.ToString('0.0##', $inv) + ':불가') }
            else {
                $ereq = ($s0 - $ft) / $k
                $t2 = ($s0 - $fu) / $k
                $t1 = 0.0; if ($fu -gt $ft) { $e2max = ($ft - $sy) / ($ereq - $ey); $t1 = $ey + ($fu - $sy) / $e2max }
                $em = [math]::Max([math]::Max($t1, $t2), $ey)
                $emin += ($ft.ToString('0.0##', $inv) + ':' + (100 * $em).ToString('0.0', $inv))
            }
        }
    }
    $o = [ordered]@{
        name = $r.name; method = '식 (1D 평형)'
        sy = $sy; fu = $fu; E1 = $E1; eu_pct = 100 * $eu; E2 = $E2
        curve = (($pts | Select-Object -Skip 1 | ForEach-Object { (100 * $_[0]).ToString('0.###', $inv) + ':' + $_[1].ToString('0.###', $inv) }) -join ';')
        Ec = $Ec; fck = $(if ($fck -gt 0) { $fck } else { '' }); w = $w; Lc = $Lc; dT = $dT; alpha = $alpha; cal = $cal
        sigma_c0 = $s0; k = $k
        sig1 = $sig; sig2 = $sig; sig3 = $sig; sig4 = $sig; sig5 = $sig
        F1 = $sig * 50; F2 = $sig * 50; F3 = $sig * 50; F4 = $sig * 50; Ftotal = $sig * 200
        strain_pct = 100 * $eps; branch = $(if ($seg -eq 0) { '탄성 (항복 전)' } elseif ($eps -gt $eu) { '끝점 너머 (외삽)' } else { '항복 후' })
        rupture = $(if ($rupture) { '파단' } else { '-' })
        limits = (($limits | ForEach-Object { $_.ToString('0.0##', $inv) }) -join ';')
        verdicts = ($verd -join ';'); emin = ($emin -join ';')
    }
    [void]$out.Add([pscustomobject]$o)
    "{0}: sigma_c = {1:N4} MPa, F = {2:N2} N/spring, strain = {3:N2} %, {4} | {5}" -f $r.name, $sig, ($sig * 50), (100 * $eps), $o.rupture, $o.verdicts
}
$out | Export-Csv -LiteralPath $OutCsv -NoTypeInformation -Encoding UTF8
"saved $OutCsv ($($out.Count) cases)"
