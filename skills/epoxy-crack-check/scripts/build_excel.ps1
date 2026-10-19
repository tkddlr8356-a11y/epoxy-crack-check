param(
    [Parameter(Mandatory = $true)][string[]]$ResultsCsv,   # one or more result CSVs (solve_case1 / run_ansys_case1)
    [Parameter(Mandatory = $true)][string]$OutXlsx,        # target workbook; an existing file is never overwritten (_2, _3 ...)
    [string]$Title = "Case 1 (관통균열 · 양단 구속 · 온도 하강) – 에폭시 물성별 콘크리트 응력 · 에폭시 힘 · 판정"
)
# Writes the result workbook in the layout used for the earlier Case 1 deliverables:
# cases as columns; 입력 물성 / 콘크리트 응력1-5 / 에폭시 힘1-4 / 최대 변형률 / 에폭시 파단 / 인장강도별 판정 / 필요 최소 변형률.
$ErrorActionPreference = "Stop"
$inv = [Globalization.CultureInfo]::InvariantCulture
$rows = @(); foreach ($f in $ResultsCsv) { $rows += @(Import-Csv -LiteralPath $f -Encoding UTF8) }
if ($rows.Count -eq 0) { throw "no result rows" }
$target = $OutXlsx; $n = 2
while (Test-Path -LiteralPath $target) { $target = [IO.Path]::Combine([IO.Path]::GetDirectoryName($OutXlsx), [IO.Path]::GetFileNameWithoutExtension($OutXlsx) + "_$n" + [IO.Path]::GetExtension($OutXlsx)); $n++ }
New-Item -ItemType Directory -Force ([IO.Path]::GetDirectoryName($target)) | Out-Null

function D($v) { if ($v -eq $null -or "$v".Trim() -eq '') { return $null }; return [double]::Parse("$v", $inv) }
function MapOf([string]$s) { $h = [ordered]@{}; if ($s) { foreach ($p in ($s -split ';')) { if ($p) { $kv = $p.Split(':'); $h[$kv[0]] = $kv[1] } } }; return $h }
$GRAY = 15921906; $SHADE = 14277081; $RED = 13421823; $GREEN = 13434828; $YEL = 10092543   # F2F2F2 / D9D9D9 / FFCCCC / CCFFCC / FFFF99 (BGR)

# all concrete limits that appear
$allLim = New-Object System.Collections.ArrayList
foreach ($r in $rows) { foreach ($k in (MapOf $r.verdicts).Keys) { if (-not $allLim.Contains($k)) { [void]$allLim.Add($k) } } }
$allLim = @($allLim | Sort-Object { [double]::Parse($_, $inv) })

$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false; $xl.Visible = $false
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $ws = $wb.Worksheets.Item(1); $ws.Name = "결과"
    $ws.Cells.Font.Name = "Times New Roman"; $ws.Cells.Font.Size = 11
    $ws.Cells.Item(2, 2).Value2 = $Title; $ws.Cells.Item(2, 2).Font.Bold = $true; $ws.Cells.Item(2, 2).Font.Size = 13
    $c0 = 3   # first case column
    $r = 4
    $ws.Cells.Item($r, 2).Value2 = "구분"
    for ($i = 0; $i -lt $rows.Count; $i++) { $ws.Cells.Item($r, $c0 + $i).Value2 = $rows[$i].name }
    $hdr = $ws.Range($ws.Cells.Item($r, 2), $ws.Cells.Item($r, $c0 + $rows.Count - 1)); $hdr.Interior.Color = $GRAY; $hdr.Font.Bold = $true
    $r++
    function Row([string]$label, [scriptblock]$val, [string]$fmt) {
        $script:ws.Cells.Item($script:r, 2).Value2 = $label
        for ($i = 0; $i -lt $script:rows.Count; $i++) {
            $v = & $val $script:rows[$i]
            $cell = $script:ws.Cells.Item($script:r, $script:c0 + $i)
            # numbers go in through .Formula (invariant text): PowerShell's COM binder breaks when Value2 gets a string first and a double later
            if ($v -is [double]) { $cell.Formula = $v.ToString('R', $script:inv); if ($fmt) { $cell.NumberFormat = $fmt } } elseif ($v -ne $null) { $cell.NumberFormat = "@"; $cell.Formula = [string]$v }
        }
        $script:r++
        return ($script:r - 1)
    }
    function Section([string]$t) { $script:ws.Cells.Item($script:r, 2).Value2 = $t; $script:ws.Cells.Item($script:r, 2).Font.Bold = $true; $script:ws.Range($script:ws.Cells.Item($script:r, 2), $script:ws.Cells.Item($script:r, $script:c0 + $script:rows.Count - 1)).Interior.Color = $script:GRAY; $script:r++ }

    Section "입력 – 에폭시"
    [void](Row "항복강도 σ_ey (MPa)" { param($x) D $x.sy } $null)
    [void](Row "인장강도 f_eu (MPa)" { param($x) D $x.fu } $null)
    [void](Row "초기 탄성계수 E_e1 (MPa)" { param($x) D $x.E1 } $null)
    [void](Row "파단 변형률 ε_eu (%)" { param($x) D $x.eu_pct } $null)
    [void](Row "항복 후 기울기 E_e2 (MPa)" { param($x) D $x.E2 } "0.000")
    [void](Row "곡선 점 (ε % : σ MPa)" { param($x) $x.curve } $null)
    Section "입력 – 콘크리트 · 모델"
    [void](Row "콘크리트 탄성계수 E_c (MPa)" { param($x) D $x.Ec } "#,##0")
    [void](Row "설계기준강도 f_ck (MPa)" { param($x) D $x.fck } "0")
    [void](Row "균열폭 w (mm)" { param($x) D $x.w } "0.00#")
    [void](Row "구속 길이 L_c (mm) / 온도 하강 ΔT (℃)" { param($x) ("{0:#,0} / −{1}" -f (D $x.Lc), (D $x.dT)) } $null)
    [void](Row "해석 방법" { param($x) $x.method } $null)
    Section "결과"
    $firstLim = { param($x) $m = MapOf $x.verdicts; if ($m.Count -gt 0) { return [double]::Parse(@($m.Keys)[0], $inv) } else { return $null } }
    $sigRows = @()
    for ($k = 1; $k -le 5; $k++) { $kk = $k; $sigRows += (Row ("콘크리트 응력 σ_c" + $k + " (MPa)") ([scriptblock]::Create("param(`$x) D `$x.sig$kk")) "0.0000") }
    for ($k = 1; $k -le 4; $k++) { $kk = $k; [void](Row ("에폭시 힘 " + $k + " (N, 스프링당)") ([scriptblock]::Create("param(`$x) D `$x.F$kk")) "0.00") }
    [void](Row "에폭시 응력 σ_e (MPa) = 힘 / 50 mm²" { param($x) $f = D $x.F1; if ($f -ne $null) { $f / 50 } else { $null } } "0.0000")
    [void](Row "최대 변형률 (%)" { param($x) D $x.strain_pct } "0.0")
    [void](Row "평형 위치" { param($x) $x.branch } $null)
    $rupRow = Row "에폭시 파단" { param($x) $x.rupture } $null
    # shading: concrete stress above the first limit = gray (= 균열), rupture = red
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $lim = & $firstLim $rows[$i]
        if ($lim -ne $null) { foreach ($sr in $sigRows) { $cv = $ws.Cells.Item($sr, $c0 + $i).Value2; if ($cv -ne $null -and [double]$cv -gt $lim) { $ws.Cells.Item($sr, $c0 + $i).Interior.Color = $SHADE } } }
        if ($rows[$i].rupture -eq '파단') { $ws.Cells.Item($rupRow, $c0 + $i).Interior.Color = $RED }
    }
    if ($allLim.Count -gt 0) {
        Section "판정 – 콘크리트 인장강도별 (O = 콘크리트 응력 ≤ 인장강도 이고 에폭시 미파단, 한계 = 인장강도 ± 0.005 MPa)"
        foreach ($L in $allLim) {
            $LL = $L
            $rr = Row ("f_ct = " + $L + " MPa") ([scriptblock]::Create("param(`$x) (MapOf `$x.verdicts)['$LL']")) $null
            for ($i = 0; $i -lt $rows.Count; $i++) { $cv = $ws.Cells.Item($rr, $c0 + $i).Value2; if ($cv) { $ws.Cells.Item($rr, $c0 + $i).Interior.Color = $(if ($cv -eq 'O') { $GREEN } elseif ($cv -eq '한계') { $YEL } else { $RED }) } }
        }
        $anyMin = @($rows | Where-Object { $_.emin }).Count -gt 0
        if ($anyMin) {
            Section "참고 – 필요 최소 변형률 ε_eu,min (%) : 이 항복 · 인장강도 · E_e1 에서 균열 · 파단을 모두 막는 가장 작은 끝점 변형률"
            foreach ($L in $allLim) {
                $LL = $L
                $rr = Row ("f_ct = " + $L + " MPa") ([scriptblock]::Create("param(`$x) `$v = (MapOf `$x.emin)['$LL']; if (`$v -and `$v -ne '불가') { [double]::Parse(`$v, [Globalization.CultureInfo]::InvariantCulture) } else { `$v }")) "0.0"
                for ($i = 0; $i -lt $rows.Count; $i++) { if ($ws.Cells.Item($rr, $c0 + $i).Value2 -ne $null) { $ws.Cells.Item($rr, $c0 + $i).Interior.Color = $YEL } }
            }
        }
    }
    $used = $ws.Range($ws.Cells.Item(4, 2), $ws.Cells.Item($r - 1, $c0 + $rows.Count - 1))
    $used.Borders.LineStyle = 1; $used.Borders.Weight = 2; $used.HorizontalAlignment = -4108; $used.VerticalAlignment = -4108
    $ws.Range($ws.Cells.Item(4, 2), $ws.Cells.Item($r - 1, 2)).HorizontalAlignment = -4131
    $ws.Range($ws.Cells.Item(4, 2), $ws.Cells.Item($r - 1, 2)).Font.Name = "굴림"
    $ws.Columns.Item(2).ColumnWidth = 44
    for ($i = 0; $i -lt $rows.Count; $i++) { $ws.Columns.Item($c0 + $i).ColumnWidth = 16 }
    $ws.Cells.Item(4, 2).Font.Name = "굴림"

    # ---- 해석조건 sheet ----
    $wc = $wb.Worksheets.Add([Type]::Missing, $ws); $wc.Name = "해석조건"
    $lines = @(
        "해석 모델 : Case 1 관통균열 – 콘크리트 6,000 × 200 × 1 mm (50 mm 띠 4개) + 균열부 에폭시 비선형 스프링 4개 (스프링 1개 = 50 × 1 mm 단면), 양단 고정, 전 구간 온도 하강",
        "식 (1D 평형) : 콘크리트 구속선 σ_c = σ_c0 − k·ε_e,  σ_c0 = c·E_c·α·ΔT,  k = c·E_c·w / L_c  (c = 1.0004 : ANSYS 2D 모델 보정, 1 이면 1D 이론)",
        "            에폭시 곡선 = (0,0) 과 입력 점을 잇는 꺾은선, 끝점 너머는 마지막 기울기로 연장 (ANSYS COMBIN39 와 같음) → 교점이 평형 해",
        "에폭시 힘 = 에폭시 응력 × 50 mm² (스프링당), 최대 변형률 = 균열 벌어짐 / 균열폭 w",
        "판정 : 콘크리트 균열 = 콘크리트 응력 > 인장강도 f_ct (± 0.005 MPa 이내는 한계),  에폭시 파단 = 최대 변형률 > 파단 변형률 ε_eu (끝점)",
        "필요 최소 변형률 (bilinear 만) : ε_eu,min = max{ ε_ey + (f_eu − σ_ey)(ε_req − ε_ey)/(f_ct − σ_ey) [f_eu > f_ct 일 때],  (σ_c0 − f_eu)/k },  ε_req = (σ_c0 − f_ct)/k,  σ_ey ≥ f_ct 이면 불가",
        "ANSYS (요청 시) : 사용자 모델 G (Case 3.wbpz) 의 ds.dat 를 MAPDL 로 풀고 σ_c1~5 = SX (x = 6,000, y = 200~0), 힘 = COMBIN39 스프링 힘. L_c 6,000 mm, ΔT −20 ℃ 고정, bilinear 만.",
        "색 : 회색 응력 = 첫 번째 인장강도 초과 (균열),  빨강 = 균열/파단,  초록 = O,  노랑 = 한계 / 필요 최소 변형률",
        "검증 : 식 vs ANSYS – (2,5) E 100 끝점 100/150/200 % : 5.0015/4.6664/4.3986 vs 5.0016/4.6665/4.3987 MPa,  예영준 (12.5, 17, E 250) : 5.9785 vs 5.9786 MPa"
    )
    for ($i = 0; $i -lt $lines.Count; $i++) { $wc.Cells.Item(2 + $i, 2).Value2 = $lines[$i] }
    $wc.Cells.Font.Name = "굴림"; $wc.Columns.Item(2).ColumnWidth = 160
    $ws.Activate()
    $wb.SaveAs($target, 51)
    $wb.Close($false)
    "saved $target"
} finally {
    $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
