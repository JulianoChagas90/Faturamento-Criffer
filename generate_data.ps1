$currentDir = "C:\Users\juliano.chagas\OneDrive - CFF DESENVOLVIMENTO DE PRODUTOS ELETRONICOS LTDA\Documentos\Dashboard teste"
$excelFile = Get-Item "$currentDir\Relat*rio de Controle Gerencial V2908.xlsx"
$excelPath = $excelFile.FullName
$outputPath = "$currentDir\data.js"

Write-Host "Abrindo Excel (Modo Otimizado): $excelPath"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath, [Type]::Missing, $true) # Open ReadOnly
$sheet = $workbook.Sheets.Item(1)

Write-Host "Lendo dados para a memoria..."
$usedRange = $sheet.UsedRange
$data = $usedRange.Value2
$rowCount = $usedRange.Rows.Count
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Processando $rowCount linhas..."

$colMap = @{ Transacao = 2; Cliente = 5; Vendedor = 7; Cancelado = 10; Status = 11; NF = 15; Modelo = 17; Data = 19; Valor = 32; Pais = 38; Estado = 39; CFOP = 43 }
$vendasCFOPs = @("5101","5102","6101","6102","5107","6107","5108","6108","5118","6118","5119","6119","5922","6922")
$exportCFOPs = @("7101","7102")

$allRows = New-Object System.Collections.Generic.List[PSCustomObject]
for ($i = 2; $i -le $rowCount; $i++) {
    $canc = "$($data[$i, $colMap.Cancelado])".Trim()
    if ($canc -ne "N") { continue }
    
    # Em Value2, datas podem vir como numericos (OADate) ou strings dependendo do formato.
    # No Excel, as datas costumam ser OADate em Value2.
    $valData = $data[$i, $colMap.Data]
    $dataRaw = ""
    if ($valData -is [double]) {
        $dataRaw = [DateTime]::FromOADate($valData).ToString("dd/MM/yyyy")
    } else {
        $dataRaw = "$valData".Trim()
    }

    if ($dataRaw -notmatch "^\d{2}/\d{2}/\d{4}") { continue }
    
    $dParts = $dataRaw -split "/"
    $anoMes = "$($dParts[2])-$($dParts[1])"
    
    $trans = "$($data[$i, $colMap.Transacao])".Trim()
    $status = "$($data[$i, $colMap.Status])".Trim()
    $modelo = "$($data[$i, $colMap.Modelo])".Trim()
    $cfop   = "$($data[$i, $colMap.CFOP])".Trim()
    $valor  = $data[$i, $colMap.Valor]
    if ($valor -eq $null) { $valor = 0 }
    
    $seg = "outro"
    if ($trans -like "*Devolu*" -and ($trans -like "*Sa?da*" -or $trans -like "*Sada*")) { 
        if ($status -eq "Autorizado") { $seg = "devolucoes" } 
    }
    elseif ($trans -like "*Nota Fiscal de Sa?da*" -or $trans -like "*Nota Fiscal de Sada*") {
        if ($modelo -eq "NFS-e") { if ($status -eq "Autorizado") { $seg = "servicos" } }
        elseif ($modelo -eq "FAT") { $seg = "locacao" }
        elseif ($modelo -eq "NFe (55)") {
            if ($status -eq "Autorizado") {
                if ($cfop -in $exportCFOPs) { $seg = "exportacao" }
                elseif ($cfop -in $vendasCFOPs) { $seg = "vendas" }
            }
            elseif (($status -eq "" -or $status -eq $null) -and ($cfop -eq "" -or $cfop -eq $null)) { $seg = "vendas" }
        }
    }
    
    if ($seg -eq "outro") { continue }
    
    $allRows.Add([PSCustomObject]@{ 
        AnoMes=$anoMes; 
        Segmento=$seg; 
        Cliente="$($data[$i, $colMap.Cliente])".Trim(); 
        Vendedor="$($data[$i, $colMap.Vendedor])".Trim(); 
        NF="$($data[$i, $colMap.NF])".Trim(); 
        Valor=[double]$valor; 
        Estado="$($data[$i, $colMap.Estado])".Trim(); 
        Data=$dataRaw 
    })
}

Write-Host "Calculando estatisticas..."
$nfInfo = @{}
foreach ($r in $allRows) { 
    if ($r.Vendedor -ne "" -and $r.Vendedor -ne "-Nenhum vendedor-") { 
        $nfInfo[$r.NF] = @{ Vendedor = $r.Vendedor; Estado = $r.Estado } 
    } 
}

function Get-SummaryData($targetRows) {
    if ($targetRows.Count -eq 0) { return @{ resumo=@{}; vendedores=@(); estados=@(); decadas=@{}; faixas=@(); periodoStr="" } }
    
    $resumo = @{ 
        vendas = @{t=0;n=@{};c=@{}}; 
        servicos = @{t=0;n=@{};c=@{}}; 
        locacao = @{t=0;n=@{};c=@{}}; 
        exportacao = @{t=0;n=@{};c=@{}}; 
        devolucoes = @{t=0;n=@{};c=@{}} 
    }
    $vencDict = @{}; $estDict = @{}; $decadas = @{ d1=0; d2=0; d3=0 }; $cliDict = @{}; $dates = @()
    
    foreach ($r in $targetRows) {
        if (($r.Vendedor -eq "" -or $r.Vendedor -eq "-Nenhum vendedor-") -and $nfInfo[$r.NF]) { 
            $r.Vendedor = $nfInfo[$r.NF].Vendedor; $r.Estado = $nfInfo[$r.NF].Estado 
        }
        $seg = $r.Segmento; 
        $resumo.$seg.t += $r.Valor; 
        $resumo.$seg.n[$r.NF] = 1; 
        $resumo.$seg.c[$r.Cliente] = 1
        
        if ($seg -ne "devolucoes") {
            $vName = if ($r.Vendedor -eq "" -or $r.Vendedor -eq "-Nenhum vendedor-") { "NAO ATRIBUIDO" } else { $r.Vendedor }
            if (-not $vencDict[$vName]) { $vencDict[$vName] = @{ total=0; nfs=@{} } }
            $vencDict[$vName].total += $r.Valor; $vencDict[$vName].nfs[$r.NF] = 1
            
            $uf = if ($seg -eq "exportacao") { "EX" } else { $r.Estado }
            if (-not $estDict[$uf]) { $estDict[$uf] = @{ total=0; segmento=@{vendas=0;servicos=0;locacao=0;exportacao=0} } }
            $estDict[$uf].total += $r.Valor; $estDict[$uf].segmento.$seg += $r.Valor
            
            if ($r.Data -match "^(\d{2})/") { 
                $dia = [int]$Matches[1]; 
                $dates += $r.Data; 
                if ($dia -le 10) { $decadas.d1 += [double]$r.Valor } 
                elseif ($dia -le 20) { $decadas.d2 += [double]$r.Valor } 
                else { $decadas.d3 += [double]$r.Valor } 
            }
            if (-not $cliDict[$r.Cliente]) { $cliDict[$r.Cliente] = 0 }; $cliDict[$r.Cliente] += $r.Valor
        }
    }
    
    $resFinal = @{}; foreach ($k in $resumo.Keys) { $resFinal[$k] = @{ total=[double]$resumo.$k.t; countNFs=$resumo.$k.n.Count; countClientes=$resumo.$k.c.Count } }
    $vendFinal = @(); foreach ($k in $vencDict.Keys) { $vendFinal += @{ nome=$k; total=[double]$vencDict[$k].total; nfs=$vencDict[$k].nfs.Count } }
    
    $regMap = @{ 
        'SP'='Sudeste'; 'RJ'='Sudeste'; 'MG'='Sudeste'; 'ES'='Sudeste'; 
        'PR'='Sul'; 'SC'='Sul'; 'RS'='Sul'; 
        'MT'='Centro-Oeste'; 'MS'='Centro-Oeste'; 'GO'='Centro-Oeste'; 'DF'='Centro-Oeste'; 
        'BA'='Nordeste'; 'PE'='Nordeste'; 'CE'='Nordeste'; 'RN'='Nordeste'; 'PB'='Nordeste'; 'AL'='Nordeste'; 'SE'='Nordeste'; 'PI'='Nordeste'; 'MA'='Nordeste'; 
        'AM'='Norte'; 'PA'='Norte'; 'RO'='Norte'; 'TO'='Norte'; 'AC'='Norte'; 'RR'='Norte'; 'AP'='Norte'; 
        'EX'='Exportacao' 
    }
    
    $estFinal = @(); foreach ($k in $estDict.Keys) { $estFinal += @{ uf=$k; regiao=$regMap[$k]; total=[double]$estDict[$k].total; segmento=$estDict[$k].segmento } }
    
    $faixas = @( 
        @{ n='F1 >=R$50k'; l=50000; h=1e20; c=0; t=0 }; 
        @{ n='F2 R$20k-49k'; l=20000; h=49999; c=0; t=0 }; 
        @{ n='F3 R$10k-19k'; l=10000; h=19999; c=0; t=0 }; 
        @{ n='F4 R$5k-9k'; l=5000; h=9999; c=0; t=0 }; 
        @{ n='F5 ate R$4.9k'; l=0; h=4999; c=0; t=0 } 
    )
    foreach ($cVal in $cliDict.Values) { 
        foreach ($f in $faixas) { 
            if ($cVal -ge $f.l -and $cVal -le $f.h) { $f.c++; $f.t += [double]$cVal; break } 
        } 
    }
    
    $sortedDates = $dates | Sort-Object; 
    $pStr = if ($sortedDates) { "$($sortedDates[0]) a $($sortedDates[-1])" } else { "" }
    
    return @{ resumo=$resFinal; vendedores=$vendFinal | Sort-Object total -Descending; estados=$estFinal | Sort-Object total -Descending; decadas=$decadas; faixas=$faixas; periodoStr=$pStr }
}

$periods = @($allRows | Select-Object -ExpandProperty AnoMes -Unique | Sort-Object)
$monthsData = @{}
foreach ($pm in $periods) {
    Write-Host "Compilando mes $pm..."
    $monthlyRows = $allRows | Where-Object { $_.AnoMes -eq $pm }
    $yyyy = $pm.Split("-")[0]
    $ytdRows = $allRows | Where-Object { $_.AnoMes.StartsWith($yyyy) -and $_.AnoMes -le $pm }
    $monthsData[$pm] = @{ mensal = Get-SummaryData $monthlyRows; ytd = Get-SummaryData $ytdRows }
}

$menu = @()
foreach ($pm in ($periods | Sort-Object -Descending)) {
    $yy = $pm.Split("-")[0]; $mm = $pm.Split("-")[1]
    $mName = @{ "01"="Janeiro"; "02"="Fevereiro"; "03"="Marco"; "04"="Abril"; "05"="Maio"; "06"="Junho"; "07"="Julho"; "08"="Agosto"; "09"="Setembro"; "10"="Outubro"; "11"="Novembro"; "12"="Dezembro" }[$mm]
    $menu += @{ id=$pm; label="$mName/$yy" }
}

$res = @{ periods = $monthsData; menu = $menu }
"const DASHBOARD_DATA = $($res | ConvertTo-Json -Depth 10);" | Out-File -FilePath $outputPath -Encoding utf8
Write-Host "Concluido! Data.js pronto com $rowCount linhas processadas."
