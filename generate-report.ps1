# ============================================================
# Abrio — Weekly Shopify Revenue Report Generator
# Laeuft lokal und via GitHub Actions (keine Secrets im Code)
# ============================================================
param(
    [string]$ClientId     = $env:SHOPIFY_CLIENT_ID,
    [string]$ClientSecret = $env:SHOPIFY_CLIENT_SECRET
)

if (-not $ClientId -or -not $ClientSecret) {
    Write-Error "SHOPIFY_CLIENT_ID und SHOPIFY_CLIENT_SECRET muessen gesetzt sein."
    exit 1
}

$ROOT     = $PSScriptRoot
$TEMPLATE = Join-Path $ROOT "Templates\abrio-report-template.html"
$REVIEWS  = Join-Path $ROOT "Reviews"
$INDEX    = Join-Path $ROOT "index.html"
$deDE     = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")

function de($n) { [string]::Format($deDE, "{0:N2}", [double]$n) }
function deM($n) { "$(de $n)" }

# ── Schritt 1: Datum berechnen ──────────────────────────────
$today   = Get-Date
$dow     = [int]$today.DayOfWeek; if ($dow -eq 0) { $dow = 7 }
$monday  = $today.AddDays(-($dow - 1) - 7)
$sunday  = $monday.AddDays(6)

$START_DATE = $monday.ToString("yyyy-MM-dd")
$END_DATE   = $sunday.ToString("yyyy-MM-dd")
$LABEL_VON  = $monday.ToString("d. MMMM yyyy", $deDE)
$LABEL_BIS  = $sunday.ToString("d. MMMM yyyy", $deDE)
$HEUTE_DE   = $today.ToString("d. MMMM yyyy", $deDE)
$HEUTE_KURZ = $today.ToString("d. MMM. yyyy", $deDE)

# ISO-Wochennummer der Vorwoche
$jan4      = [datetime]::new($monday.Year, 1, 4)
$startOfW1 = $jan4.AddDays(-([int]$jan4.DayOfWeek - 1))
if ($monday -lt $startOfW1) {
    $jan4      = [datetime]::new($monday.Year - 1, 1, 4)
    $startOfW1 = $jan4.AddDays(-([int]$jan4.DayOfWeek - 1))
}
$KW       = [int](($monday - $startOfW1).Days / 7) + 1
$KW_STR   = $KW.ToString("00")
$YEAR     = $monday.Year
$DATEINAME = "revenue-report-KW$KW_STR-$YEAR.html"

Write-Host "KW$KW_STR | $START_DATE – $END_DATE | $DATEINAME"

# ── Schritt 2 & 3: Tokens holen + Bestelldaten abfragen ────
$body = "grant_type=client_credentials&client_id=$ClientId&client_secret=$ClientSecret"

$storeDefs = @(
    @{ domain="development-plus-store.myshopify.com"; name="Dev Plus" },
    @{ domain="no-cosmetics-dev.myshopify.com";       name="No Cosmetics" },
    @{ domain="barfgold-dev.myshopify.com";           name="Barfgold" },
    @{ domain="hofmanns-dev.myshopify.com";           name="Hofmanns" }
)

$results = @()
foreach ($s in $storeDefs) {
    # Token via Client Credentials
    $tr    = Invoke-RestMethod -Uri "https://$($s.domain)/admin/oauth/access_token" `
                               -Method POST -Body $body `
                               -ContentType "application/x-www-form-urlencoded"
    $token   = $tr.access_token
    $headers = @{ "X-Shopify-Access-Token" = $token; "Content-Type" = "application/json" }

    # Bestellungen abfragen (mit Pagination)
    $allOrders = @()
    $cursor    = $null
    $hasNext   = $true
    while ($hasNext) {
        $after = if ($cursor) { ", after: \`"$cursor\`"" } else { "" }
        $gql   = '{"query":"{ orders(first:250' + $after + ', query:\"created_at:>=' + $START_DATE + ' created_at:<=' + $END_DATE + '\") { edges { node { totalPriceSet { shopMoney { amount } } } } pageInfo { hasNextPage endCursor } } }"}'
        $r     = Invoke-RestMethod -Uri "https://$($s.domain)/admin/api/2024-04/graphql.json" `
                                   -Method POST -Headers $headers -Body $gql
        $edges   = $r.data.orders.edges
        $hasNext = $r.data.orders.pageInfo.hasNextPage
        $cursor  = $r.data.orders.pageInfo.endCursor
        if ($edges) { $allOrders += $edges }
    }

    $revenue = [math]::Round(($allOrders | ForEach-Object { [double]$_.node.totalPriceSet.shopMoney.amount } | Measure-Object -Sum).Sum, 2)
    if (-not $revenue) { $revenue = 0.0 }
    $orders  = $allOrders.Count
    $aov     = if ($orders -gt 0) { [math]::Round($revenue / $orders, 2) } else { 0.0 }

    $results += [PSCustomObject]@{ Domain=$s.domain; Name=$s.name; Revenue=$revenue; Orders=$orders; AOV=$aov }
    Write-Host "  $($s.name): €$revenue | $orders Bestellungen"
}

# ── Schritt 4: Metriken berechnen ──────────────────────────
$totalRevenue = [math]::Round(($results | Measure-Object -Property Revenue -Sum).Sum, 2)
$totalOrders  = [int]($results | Measure-Object -Property Orders -Sum).Sum
$totalAOV     = if ($totalOrders -gt 0) { [math]::Round($totalRevenue / $totalOrders, 2) } else { 0.0 }
$activeStores = ($results | Where-Object { $_.Orders -gt 0 }).Count

foreach ($r in $results) {
    $r | Add-Member -NotePropertyName Share -NotePropertyValue (
        if ($totalRevenue -gt 0) { [math]::Round($r.Revenue / $totalRevenue * 100, 2) } else { 0.0 }
    )
}

Write-Host "GESAMT: €$totalRevenue | $totalOrders Bestellungen | AOV €$totalAOV | $activeStores/4 aktiv"

# ── Schritt 5: HTML generieren ──────────────────────────────
$html = Get-Content $TEMPLATE -Raw -Encoding UTF8

# Standard-Platzhalter
$map = @{
    '{{REPORT_TITLE}}'       = "Multi-Store Revenue Report KW$KW_STR $YEAR"
    '{{DATE_RANGE}}'         = "$LABEL_VON – $LABEL_BIS"
    '{{HEADER_META_LINE2}}'  = "4 Stores abgefragt · $activeStores Aktiv"
    '{{REPORT_EYEBROW}}'     = "E-Commerce Analytics · Shopify Admin API · KW$KW_STR $YEAR"
    '{{REPORT_H1_PLAIN}}'    = 'Multi-Store'
    '{{REPORT_H1_ACCENT}}'   = "Revenue Report KW$KW_STR"
    '{{REPORT_SUBTITLE}}'    = "Umsatz, Bestellungen &amp; Kennzahlen $LABEL_VON – $LABEL_BIS"
    '{{KPI1_LABEL}}'         = 'Gesamtumsatz'
    '{{KPI1_VALUE}}'         = "€ $(de $totalRevenue)"
    '{{KPI1_SUB}}'           = "alle Stores · KW$KW_STR $YEAR"
    '{{KPI2_LABEL}}'         = 'Bestellungen gesamt'
    '{{KPI2_VALUE}}'         = "$totalOrders"
    '{{KPI2_SUB}}'           = 'über alle aktiven Stores'
    '{{KPI3_LABEL}}'         = 'Ø Bestellwert (AOV)'
    '{{KPI3_VALUE}}'         = "€ $(de $totalAOV)"
    '{{KPI3_SUB}}'           = 'Gesamtumsatz / Gesamtbestellungen'
    '{{KPI4_LABEL}}'         = 'Aktive Stores'
    '{{KPI4_VALUE}}'         = "$activeStores / 4"
    '{{KPI4_SUB}}'           = 'mit Bestellungen im Zeitraum'
    '{{CHART1_TITLE}}'       = 'Umsatz je Store'
    '{{CHART2_TITLE}}'       = 'Umsatzanteile (%)'
    '{{CHART3_TITLE}}'       = 'Bestellungen je Store'
    '{{CHART4_TITLE}}'       = 'Ø Bestellwert (AOV) je Store'
    '{{TABLE_TITLE}}'        = "Store-Kennzahlen — $LABEL_VON bis $LABEL_BIS"
    '{{TABLE_TAG}}'          = 'Shopify Admin API'
    '{{COL3}}'               = 'Umsatz (€)'
    '{{COL4}}'               = 'Bestellungen'
    '{{COL5}}'               = 'AOV (€)'
    '{{COL6}}'               = 'Umsatzanteil'
    '{{GENERATED_DATE}}'     = $HEUTE_DE
    "'{{LABEL_1}}'"          = "'$($results[0].Name)'"
    "'{{LABEL_2}}'"          = "'$($results[1].Name)'"
    "'{{LABEL_3}}'"          = "'$($results[2].Name)'"
    "'{{LABEL_4}}'"          = "'$($results[3].Name)'"
}
foreach ($k in $map.Keys) { $html = $html.Replace($k, $map[$k]) }

# Chart-Daten
$d1 = ($results | ForEach-Object { $_.Revenue }) -join ', '
$d2 = ($results | ForEach-Object { $_.Share })   -join ', '
$d3 = ($results | ForEach-Object { $_.Orders })  -join ', '
$d4 = ($results | ForEach-Object { $_.AOV })     -join ', '
$html = $html -replace 'const DATA_1\s*=\s*\[.*?\];.*', "const DATA_1 = [$d1]; // Umsatz"
$html = $html -replace 'const DATA_2\s*=\s*\[.*?\];.*', "const DATA_2 = [$d2]; // Anteile %"
$html = $html -replace 'const DATA_3\s*=\s*\[.*?\];.*', "const DATA_3 = [$d3]; // Bestellungen"
$html = $html -replace 'const DATA_4\s*=\s*\[.*?\];.*', "const DATA_4 = [$d4]; // AOV"

# Chart_Scripts Kommentar entfernen
$html = $html -replace '(?s)/\*\s*═+\s*\{\{CHART_SCRIPTS\}\}.*?Example — replace with real data \*/', "/* KW$KW_STR $YEAR — generiert am $HEUTE_DE */"

# Tabellenzeilen generieren
$rows = ""
foreach ($r in $results) {
    $active = $r.Orders -gt 0
    $badge  = if ($active) { '<span class="badge badge-active">Aktiv</span>' } else { '<span class="badge badge-inactive">Inaktiv</span>' }
    $muted  = if (-not $active) { ' style="color:var(--text-muted)"' } else { '' }
    $aovStr = if ($active) { de $r.AOV } else { '—' }
    $valCls = if ($active) { ' val-highlight' } else { '' }
    $rows  += @"

            <tr>
              <td>
                <div class="td-store-name">$($r.Name)</div>
                <div class="td-store-domain">$($r.Domain)</div>
              </td>
              <td>$badge</td>
              <td class="right num$valCls"$muted>$(de $r.Revenue)</td>
              <td class="right num"$muted>$($r.Orders)</td>
              <td class="right num"$muted>$aovStr</td>
              <td>
                <span class="share-pct"$(if (-not $active){' style="color:var(--text-muted)"'})>$(de $r.Share) %</span>
                <div class="share-bar"><div class="share-fill" style="width:$($r.Share)%"></div></div>
              </td>
            </tr>
"@
}
$rows += @"

            <tr class="row-total">
              <td><div class="td-store-name">Gesamt</div></td>
              <td><span class="badge badge-active">$activeStores / 4 Aktiv</span></td>
              <td class="right num val-highlight">$(de $totalRevenue)</td>
              <td class="right num">$totalOrders</td>
              <td class="right num">$(de $totalAOV)</td>
              <td><span class="share-pct">100,00 %</span></td>
            </tr>
"@
$html = $html.Replace('            {{TABLE_ROWS}}', $rows)

$outPath = Join-Path $REVIEWS $DATEINAME
$html | Set-Content $outPath -Encoding UTF8
Write-Host "Report gespeichert: $outPath"

# ── Schritt 6: index.html aktualisieren ────────────────────
$idx = Get-Content $INDEX -Raw -Encoding UTF8

# Neuestes-Badge von alter Karte entfernen
$idx = $idx -replace 'class="report-card report-card--latest"', 'class="report-card"'
$idx = $idx -replace '\s*<span class="card-badge-new">Neuester</span>', ''

# Neue Karte nach <!-- REPORTS_START --> einfügen
$newCard = @"

    <a class="report-card report-card--latest" href="Reviews/$DATEINAME">
      <div class="card-top">
        <span class="card-kw">KW$KW_STR</span>
        <span class="card-badge-new">Neuester</span>
      </div>
      <div>
        <div class="card-title">Multi-Store Revenue Report</div>
        <div class="card-range">$LABEL_VON – $LABEL_BIS</div>
      </div>
      <div class="card-meta">
        <div class="card-metric">
          <span class="card-metric-val">€ $(de $totalRevenue)</span>
          <span class="card-metric-lbl">Umsatz</span>
        </div>
        <div class="card-metric">
          <span class="card-metric-val">$totalOrders</span>
          <span class="card-metric-lbl">Bestellungen</span>
        </div>
        <div class="card-metric">
          <span class="card-metric-val">€ $(de $totalAOV)</span>
          <span class="card-metric-lbl">AOV</span>
        </div>
        <div class="card-metric">
          <span class="card-metric-val">$activeStores / 4</span>
          <span class="card-metric-lbl">Aktive Stores</span>
        </div>
      </div>
      <div class="card-arrow">
        Report öffnen
        <svg viewBox="0 0 24 24"><path d="M5 12h14M12 5l7 7-7 7"/></svg>
      </div>
    </a>
"@
$idx = $idx -replace '(<!-- REPORTS_START -->)', "`$1$newCard"

# Anzahl Report-Cards zählen
$reportCount = ([regex]::Matches($idx, 'class="report-card')).Count

# Stats-Bar ersetzen
$statsNew = @"
  <div class="stats-bar">
    <div class="stat-item">
      <span class="stat-value">$reportCount</span>
      <span class="stat-label">Reports gesamt</span>
    </div>
    <div class="stat-item">
      <span class="stat-value">KW$KW_STR</span>
      <span class="stat-label">Aktuellste Woche</span>
    </div>
    <div class="stat-item">
      <span class="stat-value">4</span>
      <span class="stat-label">Verbundene Stores</span>
    </div>
    <div class="stat-item">
      <span class="stat-value">$HEUTE_KURZ</span>
      <span class="stat-label">Zuletzt aktualisiert</span>
    </div>
  </div>
"@
$idx = $idx -replace '(?s)<!-- STATS_START -->.*?<!-- STATS_END -->', "<!-- STATS_START -->`n$statsNew  <!-- STATS_END -->"

# Footer-Datum
$idx = $idx -replace '<!-- FOOTER_DATE_START -->.*?<!-- FOOTER_DATE_END -->', "<!-- FOOTER_DATE_START -->Zuletzt aktualisiert: <strong>$HEUTE_DE</strong><!-- FOOTER_DATE_END -->"

$idx | Set-Content $INDEX -Encoding UTF8
Write-Host "index.html aktualisiert ($reportCount Reports)"

Write-Host ""
Write-Host "✅ Fertig: KW$KW_STR $YEAR | €$(de $totalRevenue) | $totalOrders Bestellungen | $activeStores/4 Stores aktiv"
