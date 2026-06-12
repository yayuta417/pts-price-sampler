# ============================================================
# PTS株価 1分値サンプリング・ポーラー
# ------------------------------------------------------------
# 株探(kabutan.jp)の銘柄ページに表示されるPTS最新約定値+約定時刻を
# 毎分取得し、データリポジトリのCSVへ追記する。
# GitHub Actions (ubuntu-latest / pwsh) での実行を想定。PS5.1でも動作。
#
# 出力(データリポジトリ内):
#   PTS/1分値/{code}_PTS_1分値_YYYY-MM-DD.csv  (取得時刻,PTS約定時刻,PTS価格)
#   PTS/PTS日次サマリー.csv (日付,セッション,始値,高値,安値,終値,サンプル数)
#
# セッションラベルと集計ウィンドウ(約定時刻ベース):
#   寄り前 = 08:20〜09:00 / 夜間 = 16:30〜23:59
# ============================================================
param(
    [string]$DataDir = "./data",
    [string]$Code = $env:STOCK_CODE,
    [string]$StartHHmm = "16:25",
    [string]$EndHHmm = "00:01",
    [string]$SessionLabel = "夜間",
    [int]$MaxMinutes = 345,
    [int]$TestMinutes = 0
)
$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"

if (-not $Code) { Write-Output "STOCK_CODE が未設定です"; exit 1 }

try { $script:jstZone = [TimeZoneInfo]::FindSystemTimeZoneById("Asia/Tokyo") }
catch { $script:jstZone = [TimeZoneInfo]::FindSystemTimeZoneById("Tokyo Standard Time") }
function Get-JstNow { return [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $script:jstZone) }

# 追記はBOMなしUTF-8、新規作成時のみBOM付き(Excel互換)
$script:encNew = "UTF8"
if ($PSVersionTable.PSVersion.Major -ge 6) { $script:encNew = "utf8BOM" }

$ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"

function Get-PtsSample {
    param($code)
    try {
        $r = Invoke-WebRequest -Uri ("https://kabutan.jp/stock/?code={0}" -f $code) -Headers @{ "User-Agent" = $ua } -TimeoutSec 20 -UseBasicParsing
        $m = [regex]::Match($r.Content, 'kabuka1">PTS</div>\s*<div class="kabuka2">([\d,\.]+)円</div>\s*<div class="kabuka3">(\d{1,2}:\d{2})\s+(\d{2}/\d{2})</div>')
        if (-not $m.Success) { return $null }
        $price = [double]($m.Groups[1].Value -replace ",", "")
        $now = Get-JstNow
        $vd = [datetime]::ParseExact(("{0}/{1} {2}" -f $now.Year, $m.Groups[3].Value, $m.Groups[2].Value), "yyyy/MM/dd HH:mm", [Globalization.CultureInfo]::InvariantCulture)
        if ($vd -gt $now.AddDays(2)) { $vd = $vd.AddYears(-1) }   # 年跨ぎ補正
        return [PSCustomObject]@{ Sampled = $now; VenueTime = $vd; Price = $price }
    } catch {
        Write-Output ("fetch error: {0}" -f $_.Exception.Message)
        return $null
    }
}

function Format-Price {
    param($v)
    $d = [math]::Round([double]$v, 2)
    if ($d -eq [math]::Floor($d)) { return [string][long]$d }
    return [string]$d
}

function Push-DataRepo {
    param($msg)
    git -C $DataDir add PTS | Out-Null
    $st = git -C $DataDir status --porcelain
    if (-not $st) { return }
    git -C $DataDir commit -m $msg | Out-Null
    for ($i = 0; $i -lt 3; $i++) {
        git -C $DataDir pull --rebase origin main
        git -C $DataDir push origin main
        if ($LASTEXITCODE -eq 0) { Write-Output "pushed: $msg"; return }
        Start-Sleep -Seconds (10 + $i * 15)
    }
    Write-Output "push failed after retries"
}

# ---------------- ウィンドウ決定 ----------------
$now = Get-JstNow
$isTest = ($TestMinutes -gt 0)
if ($isTest) {
    $startAt = $now
    $endAt = $now.AddMinutes($TestMinutes)
} else {
    $startAt = $now.Date + [TimeSpan]::Parse($StartHHmm)
    $endAt = $now.Date + [TimeSpan]::Parse($EndHHmm)
    if ($endAt -le $startAt) { $endAt = $endAt.AddDays(1) }
    if ($now -ge $endAt) { Write-Output "ウィンドウ終了済みのため何もしません"; exit 0 }
}
$hardStop = $now.AddMinutes($MaxMinutes)
if ($endAt -gt $hardStop) { $endAt = $hardStop }

$sessionDate = $startAt.ToString("yyyy-MM-dd")
$csvDir = Join-Path $DataDir "PTS/1分値"
New-Item -ItemType Directory -Force -Path $csvDir | Out-Null
$csvPath = Join-Path $csvDir ("{0}_PTS_1分値_{1}.csv" -f $Code, $sessionDate)
if (-not (Test-Path $csvPath)) {
    Set-Content -Path $csvPath -Value "取得時刻,PTS約定時刻,PTS価格" -Encoding $script:encNew
}

Write-Output ("開始: session={0} window={1}〜{2} (JST) csv={3}" -f $SessionLabel, $startAt.ToString("HH:mm"), $endAt.ToString("MM/dd HH:mm"), $csvPath)

# 開始時刻まで待機
while ((Get-JstNow) -lt $startAt) { Start-Sleep -Seconds 15 }

# ---------------- サンプリングループ ----------------
$count = 0
$lastPush = 0
while ((Get-JstNow) -lt $endAt) {
    $s = Get-PtsSample $Code
    if ($null -ne $s) {
        Add-Content -Path $csvPath -Value ("{0},{1},{2}" -f $s.Sampled.ToString("yyyy-MM-dd HH:mm:ss"), $s.VenueTime.ToString("yyyy-MM-dd HH:mm"), (Format-Price $s.Price)) -Encoding UTF8
        $count++
        if ($count % 10 -eq 0) { Write-Output ("{0} 件目 price={1} venue={2}" -f $count, $s.Price, $s.VenueTime.ToString("HH:mm")) }
    }
    if (($count - $lastPush) -ge 20) {
        Push-DataRepo ("PTS収集(途中) {0} {1}" -f $sessionDate, $SessionLabel)
        $lastPush = $count
    }
    $sec = 60 - ((Get-JstNow).Second)
    if ($sec -lt 5) { $sec += 60 }
    if ((Get-JstNow).AddSeconds($sec) -ge $endAt -and (Get-JstNow) -ge $endAt) { break }
    Start-Sleep -Seconds $sec
}
Write-Output ("ループ終了: {0}件取得" -f $count)

# ---------------- 日次サマリー(セッション内の約定時刻でフィルタ) ----------------
if (-not $isTest) {
    $winMap = @{ "寄り前" = @("08:20", "09:00"); "夜間" = @("16:30", "23:59") }
    if ($winMap.ContainsKey($SessionLabel)) {
        $w = $winMap[$SessionLabel]
        $wStart = [datetime]::ParseExact(("{0} {1}" -f $sessionDate, $w[0]), "yyyy-MM-dd HH:mm", [Globalization.CultureInfo]::InvariantCulture)
        $wEnd = [datetime]::ParseExact(("{0} {1}" -f $sessionDate, $w[1]), "yyyy-MM-dd HH:mm", [Globalization.CultureInfo]::InvariantCulture)
        $rows = @(Import-Csv $csvPath | Where-Object {
            $vt = [datetime]::ParseExact($_."PTS約定時刻", "yyyy-MM-dd HH:mm", [Globalization.CultureInfo]::InvariantCulture)
            $vt -ge $wStart -and $vt -le $wEnd
        })
        if ($rows.Count -gt 0) {
            $prices = @($rows | ForEach-Object { [double]$_."PTS価格" })
            $summary = [PSCustomObject]@{
                日付 = $sessionDate; セッション = $SessionLabel
                始値 = (Format-Price $prices[0]); 高値 = (Format-Price ($prices | Measure-Object -Maximum).Maximum)
                安値 = (Format-Price ($prices | Measure-Object -Minimum).Minimum); 終値 = (Format-Price $prices[$prices.Count - 1])
                サンプル数 = $rows.Count
            }
            $sumPath = Join-Path $DataDir "PTS/PTS日次サマリー.csv"
            $existing = @{}
            if (Test-Path $sumPath) {
                foreach ($r0 in @(Import-Csv $sumPath)) { $existing[($r0.日付 + "|" + $r0.セッション)] = $r0 }
            }
            $existing[($sessionDate + "|" + $SessionLabel)] = $summary
            $lines = New-Object System.Collections.ArrayList
            [void]$lines.Add("日付,セッション,始値,高値,安値,終値,サンプル数")
            foreach ($k in ($existing.Keys | Sort-Object)) {
                $r1 = $existing[$k]
                [void]$lines.Add(("{0},{1},{2},{3},{4},{5},{6}" -f $r1.日付, $r1.セッション, $r1.始値, $r1.高値, $r1.安値, $r1.終値, $r1.サンプル数))
            }
            Set-Content -Path $sumPath -Value $lines -Encoding $script:encNew
            Write-Output ("サマリー更新: {0} {1} 始値{2} 終値{3}" -f $sessionDate, $SessionLabel, $summary.始値, $summary.終値)
        } else {
            Write-Output "セッションウィンドウ内の約定なし(サマリー更新スキップ)"
        }
    }
}

Push-DataRepo ("PTS収集 {0} {1} ({2}件)" -f $sessionDate, $SessionLabel, $count)
Write-Output "完了"
exit 0
