
# dashboard_generator.ps1
# Script para extrair dados do Excel e gerar o JSON multi-período para o Dashboard Criffer

$excelPath = Join-Path $PSScriptRoot "Valores Recuperados.xlsx"
$outputPath = Join-Path $PSScriptRoot "data.js"

Write-Host "Iniciando extração de dados do Excel..." -ForegroundColor Cyan

function Get-MonthName($m) {
    switch ($m) {
        1 { return "Janeiro" }
        2 { return "Fevereiro" }
        3 { return "Março" }
        4 { return "Abril" }
        5 { return "Maio" }
        6 { return "Junho" }
        7 { return "Julho" }
        8 { return "Agosto" }
        9 { return "Setembro" }
        10 { return "Outubro" }
        11 { return "Novembro" }
        12 { return "Dezembro" }
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $wb = $excel.Workbooks.Open($excelPath)
    $ws = $wb.Sheets.Item(1)
    
    $rowCount = $ws.UsedRange.Rows.Count
    $allData = @()

    $cleanMoney = {
        param($val)
        if (!$val) { return 0 }
        $cleaned = $val -replace '[^0-9,.]', '' 
        if ($cleaned -match "\." -and $cleaned -match ",") {
            $cleaned = $cleaned -replace "\.", ""
            $cleaned = $cleaned -replace ",", "."
        } elseif ($cleaned -match ",") {
            $cleaned = $cleaned -replace ",", "."
        }
        return [double]$cleaned
    }

    for ($i = 2; $i -le $rowCount; $i++) {
        $cobrador = $ws.Cells.Item($i, 1).Text.Trim()
        if ([string]::IsNullOrWhiteSpace($cobrador)) { continue }

        $vencStr = $ws.Cells.Item($i, 9).Text.Trim()
        $pagStr = $ws.Cells.Item($i, 10).Text.Trim()

        if ([string]::IsNullOrWhiteSpace($vencStr) -or [string]::IsNullOrWhiteSpace($pagStr)) { continue }

        try {
            $dVenc = [datetime]::Parse($vencStr, [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR"))
            $dPag = [datetime]::Parse($pagStr, [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR"))
            
            $diasAtraso = ($dPag - $dVenc).Days
            if ($diasAtraso -lt 0) { $diasAtraso = 0 }

            $periodo = "$(Get-MonthName $dPag.Month)-$($dPag.Year)"
            $periodoSort = $dPag.ToString("yyyyMM")

            $allData += [PSCustomObject]@{
                Periodo     = $periodo
                PeriodoSort = $periodoSort
                Cobrador    = $cobrador
                Cliente     = $ws.Cells.Item($i, 3).Text.Trim()
                VOriginal   = &$cleanMoney $ws.Cells.Item($i, 5).Text
                Juros       = &$cleanMoney $ws.Cells.Item($i, 6).Text
                VRecuperado = &$cleanMoney $ws.Cells.Item($i, 7).Text
                DiasAtraso  = $diasAtraso
                Filial      = $ws.Cells.Item($i, 11).Text.Trim()
                Motivo      = $ws.Cells.Item($i, 13).Text.Trim()
            }
        } catch { continue }
    }

    $wb.Close($false)
    $excel.Quit()
}
catch {
    Write-Host "Erro: $_" -ForegroundColor Red
    if ($excel) { $excel.Quit() }
    exit
}

Write-Host "Processando métricas por período..." -ForegroundColor Yellow

$metricasPorPeriodo = @{}
$periodosUnicos = $allData | Select-Object Periodo, PeriodoSort -Unique | Sort-Object PeriodoSort

$ytdRecuperado = @{} # Dicionário para acumulado por ano

foreach ($p in $periodosUnicos) {
    $periodoNome = $p.Periodo
    $data = $allData | Where-Object { $_.Periodo -eq $periodoNome }
    $anoAtual = $p.PeriodoSort.ToString().Substring(0, 4)

    $totalRecuperado = ($data | Measure-Object VRecuperado -Sum).Sum
    $totalJuros = ($data | Measure-Object Juros -Sum).Sum
    $totalOriginal = ($data | Measure-Object VOriginal -Sum).Sum
    $totalRegistros = $data.Count
    $atrasoMedio = ($data | Measure-Object DiasAtraso -Average).Average
    
    # Acumular YTD
    if (!$ytdRecuperado.ContainsKey($anoAtual)) { $ytdRecuperado[$anoAtual] = 0 }
    $ytdRecuperado[$anoAtual] += $totalRecuperado

    $aging = @{
        "ate_30"   = ($data | Where-Object { $_.DiasAtraso -le 30 }).Count
        "31_90"    = ($data | Where-Object { $_.DiasAtraso -gt 30 -and $_.DiasAtraso -le 90 }).Count
        "91_120"   = ($data | Where-Object { $_.DiasAtraso -gt 90 -and $_.DiasAtraso -le 120 }).Count
        "121_180"  = ($data | Where-Object { $_.DiasAtraso -gt 120 -and $_.DiasAtraso -le 180 }).Count
        "mais_180" = ($data | Where-Object { $_.DiasAtraso -gt 180 }).Count
    }

    $agingValues = @{
        "ate_30"   = ($data | Where-Object { $_.DiasAtraso -le 30 } | Measure-Object VRecuperado -Sum).Sum
        "31_90"    = ($data | Where-Object { $_.DiasAtraso -gt 30 -and $_.DiasAtraso -le 90 } | Measure-Object VRecuperado -Sum).Sum
        "91_120"   = ($data | Where-Object { $_.DiasAtraso -gt 90 -and $_.DiasAtraso -le 120 } | Measure-Object VRecuperado -Sum).Sum
        "121_180"  = ($data | Where-Object { $_.DiasAtraso -gt 120 -and $_.DiasAtraso -le 180 } | Measure-Object VRecuperado -Sum).Sum
        "mais_180" = ($data | Where-Object { $_.DiasAtraso -gt 180 } | Measure-Object VRecuperado -Sum).Sum
    }

    $porCobrador = $data | Group-Object Cobrador | ForEach-Object {
        [PSCustomObject]@{
            Nome = $_.Name
            Qtd = $_.Count
            Valor = ($_.Group | Measure-Object VRecuperado -Sum).Sum
            Juros = ($_.Group | Measure-Object Juros -Sum).Sum
            AtrasoMedio = ($_.Group | Measure-Object DiasAtraso -Average).Average
        }
    } | Sort-Object Valor -Descending

    $porFilial = $data | Group-Object Filial | ForEach-Object {
        [PSCustomObject]@{
            Nome = $_.Name
            Qtd = $_.Count
            VOriginal = ($_.Group | Measure-Object VOriginal -Sum).Sum
            Juros = ($_.Group | Measure-Object Juros -Sum).Sum
            VRecuperado = ($_.Group | Measure-Object VRecuperado -Sum).Sum
        }
    } | Sort-Object VRecuperado -Descending

    $porMotivo = $data | Group-Object Motivo | ForEach-Object {
        $nomeMotivo = if ($_.Name) { $_.Name } else { "NÃO INFORMADO" }
        [PSCustomObject]@{
            Nome = $nomeMotivo
            Qtd = $_.Count
            Valor = ($_.Group | Measure-Object VRecuperado -Sum).Sum
        }
    } | Sort-Object Valor -Descending

    $abcData = $data | Group-Object Cliente | ForEach-Object {
        [PSCustomObject]@{ Cliente = $_.Name; Valor = ($_.Group | Measure-Object VRecuperado -Sum).Sum }
    } | Sort-Object Valor -Descending

    $acumulado = 0
    $curvaABC = $abcData | ForEach-Object {
        $acumulado += $_.Valor
        $pct = 0
        if ($totalRecuperado -gt 0) { $pct = ($acumulado / $totalRecuperado) * 100 }
        
        $cat = "C"
        if ($pct -le 80) { $cat = "A" }
        elseif ($pct -le 95) { $cat = "B" }
        
        [PSCustomObject]@{ Cliente = $_.Cliente; Valor = $_.Valor; PctAcum = $pct; Categoria = $cat }
    }

    $avgAtrasoRaw = if ($atrasoMedio) { $atrasoMedio } else { 0 }
    $metricasPorPeriodo[$periodoNome] = @{
        kpis = @{
            totalRecuperado = $totalRecuperado
            totalJuros = $totalJuros
            totalOriginal = $totalOriginal
            totalRegistros = $totalRegistros
            atrasoMedio = [math]::Round($avgAtrasoRaw, 0)
            ytdRecuperado = $ytdRecuperado[$anoAtual]
        }
        aging = $aging
        agingValues = $agingValues
        porCobrador = $porCobrador
        porFilial = $porFilial
        porMotivo = $porMotivo
        curvaABC = $curvaABC | Select-Object -First 20
    }
}

$totalAnoRecuperado = ($allData | Measure-Object VRecuperado -Sum).Sum

$finalObj = @{
    atualizadoEm = (Get-Date).ToString("dd/MM/yyyy HH:mm")
    totalAnoRecuperado = $totalAnoRecuperado
    periodos = @($periodosUnicos | Sort-Object PeriodoSort -Descending | Select-Object -ExpandProperty Periodo)
    dados = $metricasPorPeriodo
}

$jsonBody = $finalObj | ConvertTo-Json -Depth 6
# Escapar caracteres especiais para garantir compatibilidade universal
$jsonBody = $jsonBody -replace 'ç', '\u00e7' -replace 'ã', '\u00e3' -replace 'Ç', '\u00c7' -replace 'Ã', '\u00c3'

$utf8NoBOM = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($outputPath, "const dashData = $jsonBody;", $utf8NoBOM)

Write-Host "Arquivo data.js multi-período gerado com sucesso!" -ForegroundColor Green
