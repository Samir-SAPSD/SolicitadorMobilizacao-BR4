param(
    [Parameter(Mandatory=$false)]
    [string]$ExcelPath = ".\DadosParaImportar.xlsx",
    
    [Parameter(Mandatory=$false)]
    [string]$SheetName = "PESSOAS"
)

[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

# Importar módulo de regras de Analista Responsável
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$analistRulesPath = Join-Path $scriptDir "Apply-AnalistRules.ps1"
if (Test-Path $analistRulesPath) {
    . $analistRulesPath
}

function Convert-ExcelColumnToIndex {
    param([string]$ColumnLetters)
    $sum = 0
    foreach ($ch in $ColumnLetters.ToUpper().ToCharArray()) {
        $sum = ($sum * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $sum
}

function Get-OpenXmlCellText {
    param(
        [System.Xml.XmlElement]$Cell,
        [array]$SharedStrings,
        [System.Xml.XmlNamespaceManager]$Ns
    )

    $cellType = $Cell.GetAttribute("t")
    $valueNode = $Cell.SelectSingleNode("x:v", $Ns)
    $inlineNode = $Cell.SelectSingleNode("x:is/x:t", $Ns)

    if ($cellType -eq "inlineStr" -and $inlineNode) {
        return $inlineNode.InnerText
    }

    if (-not $valueNode) {
        return $null
    }

    $raw = $valueNode.InnerText
    if ($cellType -eq "s") {
        $idx = 0
        if ([int]::TryParse($raw, [ref]$idx) -and $idx -ge 0 -and $idx -lt $SharedStrings.Count) {
            return $SharedStrings[$idx]
        }
        return $raw
    }

    return $raw
}

function Read-ExcelOpenXml {
    param(
        [string]$Path,
        [string]$WorksheetName,
        [bool]$IncludeEmptyColumns
    )

    Add-Type -AssemblyName System.IO.Compression

    $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $zip = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
    try {
        $workbookEntry = $zip.GetEntry("xl/workbook.xml")
        if (-not $workbookEntry) { throw "Arquivo workbook.xml não encontrado no xlsx." }

        $relsEntry = $zip.GetEntry("xl/_rels/workbook.xml.rels")
        if (-not $relsEntry) { throw "Arquivo workbook.xml.rels não encontrado no xlsx." }

        [xml]$workbookXml = New-Object System.Xml.XmlDocument
        $wbStream = $workbookEntry.Open()
        try { $workbookXml.Load($wbStream) } finally { $wbStream.Dispose() }

        [xml]$relsXml = New-Object System.Xml.XmlDocument
        $relsStream = $relsEntry.Open()
        try { $relsXml.Load($relsStream) } finally { $relsStream.Dispose() }

        $wbNs = New-Object System.Xml.XmlNamespaceManager($workbookXml.NameTable)
        $wbNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $wbNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        $relNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
        $relNs.AddNamespace("pr", "http://schemas.openxmlformats.org/package/2006/relationships")

        $sheetNode = $workbookXml.SelectSingleNode("//x:sheets/x:sheet[@name='$WorksheetName']", $wbNs)
        if (-not $sheetNode) {
            $sheetNames = @()
            $workbookXml.SelectNodes("//x:sheets/x:sheet", $wbNs) | ForEach-Object { $sheetNames += $_.GetAttribute("name") }
            throw "Aba '$WorksheetName' não encontrada. Abas disponíveis: $($sheetNames -join ', ')"
        }

        $relId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        if ([string]::IsNullOrWhiteSpace($relId)) { throw "Relacionamento da aba '$WorksheetName' não encontrado." }

        $targetNode = $relsXml.SelectSingleNode("//pr:Relationship[@Id='$relId']", $relNs)
        if (-not $targetNode) { throw "Target da aba '$WorksheetName' não encontrado em workbook.xml.rels." }

        $target = $targetNode.GetAttribute("Target")
        if ([string]::IsNullOrWhiteSpace($target)) { throw "Target da aba '$WorksheetName' vazio." }

        if ($target.StartsWith("/")) {
            $sheetPath = $target.TrimStart('/')
        } elseif ($target.StartsWith("xl/")) {
            $sheetPath = $target
        } else {
            $sheetPath = "xl/$target"
        }

        $sheetEntry = $zip.GetEntry($sheetPath)
        if (-not $sheetEntry) { throw "Worksheet XML não encontrado: $sheetPath" }

        $sharedStrings = @()
        $ssEntry = $zip.GetEntry("xl/sharedStrings.xml")
        if ($ssEntry) {
            [xml]$ssXml = New-Object System.Xml.XmlDocument
            $ssStream = $ssEntry.Open()
            try { $ssXml.Load($ssStream) } finally { $ssStream.Dispose() }

            $ssNs = New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
            $ssNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

            $ssXml.SelectNodes("//x:si", $ssNs) | ForEach-Object {
                $parts = @()
                $_.SelectNodes(".//x:t", $ssNs) | ForEach-Object { $parts += $_.InnerText }
                $sharedStrings += ($parts -join "")
            }
        }

        [xml]$sheetXml = New-Object System.Xml.XmlDocument
        $sheetStream = $sheetEntry.Open()
        try { $sheetXml.Load($sheetStream) } finally { $sheetStream.Dispose() }

        $sheetNs = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
        $sheetNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        $rows = $sheetXml.SelectNodes("//x:sheetData/x:row", $sheetNs)
        if (-not $rows -or $rows.Count -lt 2) {
            throw "Planilha vazia ou apenas cabeçalho."
        }

        $headerMap = @{}
        $headerRow = $rows[0]
        foreach ($cell in $headerRow.SelectNodes("x:c", $sheetNs)) {
            $ref = $cell.GetAttribute("r")
            if ($ref -match '^([A-Za-z]+)') {
                $colIdx = Convert-ExcelColumnToIndex -ColumnLetters $matches[1]
                $headerText = Get-OpenXmlCellText -Cell $cell -SharedStrings $sharedStrings -Ns $sheetNs
                if (-not [string]::IsNullOrWhiteSpace("$headerText")) {
                    $headerMap[$colIdx] = "$headerText"
                }
            }
        }

        $items = @()
        for ($r = 1; $r -lt $rows.Count; $r++) {
            $rowNode = $rows[$r]
            $obj = New-Object PSCustomObject
            $hasData = $false

            $valueByCol = @{}
            foreach ($cell in $rowNode.SelectNodes("x:c", $sheetNs)) {
                $ref = $cell.GetAttribute("r")
                if ($ref -match '^([A-Za-z]+)') {
                    $colIdx = Convert-ExcelColumnToIndex -ColumnLetters $matches[1]
                    $valueByCol[$colIdx] = Get-OpenXmlCellText -Cell $cell -SharedStrings $sharedStrings -Ns $sheetNs
                }
            }

            foreach ($colIdx in ($headerMap.Keys | Sort-Object)) {
                $header = $headerMap[$colIdx]
                $cellValue = $null
                if ($valueByCol.ContainsKey($colIdx)) { $cellValue = $valueByCol[$colIdx] }

                if ($IncludeEmptyColumns) {
                    $obj | Add-Member -MemberType NoteProperty -Name $header -Value $cellValue -Force
                } elseif (-not [string]::IsNullOrWhiteSpace("$cellValue")) {
                    $obj | Add-Member -MemberType NoteProperty -Name $header -Value "$cellValue" -Force
                }

                if (-not [string]::IsNullOrWhiteSpace("$cellValue")) {
                    $hasData = $true
                }
            }

            if ($hasData) {
                $items += $obj
            }
        }

        return $items
    }
    finally {
        if ($zip) { $zip.Dispose() }
        if ($fileStream) { $fileStream.Dispose() }
    }
}

function Resolve-SharePointDefaultValue {
    param(
        $Field,
        $DefaultValue
    )

    if ([string]::IsNullOrWhiteSpace("$DefaultValue")) {
        return $null
    }

    if ($Field.TypeAsString -notmatch "DateTime") {
        return $DefaultValue
    }

    $dv = ("$DefaultValue").Trim()
    if ($dv -match '^\[today\]$' -or $dv -match '^today\(\)$' -or $dv -match '^=today\(\)$') {
        return (Get-Date).Date
    }

    try {
        return [DateTime]::Parse($dv)
    }
    catch {
        # Default de data nao parseavel localmente: deixa o campo ausente para o SharePoint aplicar o default nativo.
        return $null
    }
}

# Configurações
$TestMode = $false # Altere para $false para usar a lista de produção
$SiteUrl = "https://vestas.sharepoint.com/sites/CC-ControleService-BR"

# Centralização da lista alvo (conforme solicitado, a lista é a mesma para Pessoas e Equipamentos)
if ($TestMode) {
    Write-Host "--- MODO DE TESTE ATIVADO ---" -ForegroundColor Yellow
    $ListId = "ea1e6a2e-8df6-4171-825e-1b7ecfbea7a0"
} else {
    $ListId = "2d72b0f5-d3a3-4add-a8b0-3f94de786223"
}

if (-not $ListId) {
    Write-Error "ID da lista não configurado."
    exit
}

# Configura TLS 1.2 (necessário para PSGallery)
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# Verifica/Instala o provedor NuGet antes de tentar instalar módulos
if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
    Write-Warning "Instalando provedor NuGet..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
}

# Lógica de Versão PnP: Windows PowerShell 5.1 detectado.
$TargetPnPVersion = $null
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "Ambiente: Windows PowerShell 5.1 detectado."
    Write-Warning "Forçando uso da versão legacy 1.12.0 do PnP.PowerShell (versões 2.0+ requerem PowerShell 7)."
    $TargetPnPVersion = "1.12.0"
}

# Verifica se a versão correta está instalada
$IsInstalled = if ($TargetPnPVersion) {
    Get-Module -ListAvailable -Name PnP.PowerShell | Where-Object { $_.Version -eq $TargetPnPVersion }
} else {
    Get-Module -ListAvailable -Name PnP.PowerShell
}

if (-not $IsInstalled) {
    $vMsg = if ($TargetPnPVersion) { " v$TargetPnPVersion" } else { "" }
    Write-Warning "Instalando PnP.PowerShell$vMsg..."
    try {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        
        $InstallArgs = @{
            Name = "PnP.PowerShell"
            Scope = "CurrentUser"
            Force = $true
            AllowClobber = $true
            ErrorAction = "Stop"
        }
        if ($TargetPnPVersion) { $InstallArgs["RequiredVersion"] = $TargetPnPVersion }

        Install-Module @InstallArgs
    }
    catch {
        Write-Error "ERRO CRÍTICO: Não foi possível instalar o módulo. Detalhes: $_"
        Write-Host "Execute manualmente:" -ForegroundColor Yellow
        $cmd = "Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force"
        if ($TargetPnPVersion) { $cmd += " -RequiredVersion $TargetPnPVersion" }
        Write-Host $cmd -ForegroundColor Yellow
        exit 1
    }
}

# Importação Explícita da Versão Correta
try {
    if ($TargetPnPVersion) {
        # Tenta carregar a versão específica
        Import-Module PnP.PowerShell -RequiredVersion $TargetPnPVersion -ErrorAction Stop
    } else {
        Import-Module PnP.PowerShell -ErrorAction Stop
    }
}
catch {
    Write-Error "ERRO CRÍTICO: Falha ao importar PnP.PowerShell. Detalhes: $_"
    Write-Host "Se você está no PowerShell 5.1 e a atualização falhou, tente instalar o PowerShell 7 manualmente: https://aka.ms/PS7" -ForegroundColor Yellow
    exit 1
}

# Conectar ao SharePoint
# O parâmetro -Interactive abrirá uma janela para login via navegador (MFA suportado)
# Se preferir usar ClientId/ClientSecret ou credenciais, altere este comando.
try {
    Write-Host "Conectando ao site: $SiteUrl" -ForegroundColor Cyan
    Connect-PnPOnline -Url $SiteUrl -UseWebLogin -ErrorAction Stop -WarningAction SilentlyContinue
    Write-Host "Conectado com sucesso!" -ForegroundColor Green
    
    if ($TestMode) {
        # try {
        #     # Tenta obter usuário, se falhar usa um padrão
        #     try { $currentUserEmail = (Get-PnPCurrentUser).Email } catch { $currentUserEmail = "sapsd@vestas.com" }
            
        #     Write-Host "Modo de Teste: Regenerando dados (Smart Generator)..." -ForegroundColor Yellow
        #     .\Smart-Generate-TestData.ps1 -UserEmail $currentUserEmail
        # }
        # catch {
        #     Write-Warning "Não foi possível gerar dados de teste: $_"
        # }
        Write-Host "Modo de Teste: Usando arquivo Excel existente." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Falha ao conectar ao SharePoint: $_"
    exit
}

# Mapeamento de Campos
Write-Host "Mapeando campos da lista..." -ForegroundColor Cyan

# LISTA MESTRE PARA LOOKUPS DE "PARQUE" - Definida pelo usuário
$ParqueLookupListId = "678f10f9-8d46-404b-a451-70dfe938a1ee"

try {
    # Obter TODOS os campos para mapear Excel Title -> SharePoint InternalName
    $AllFields = Get-PnPField -List $ListId | Select-Object InternalName, Title, TypeAsString, LookupList, LookupField, Required, DefaultValue
    $LookupFields = $AllFields | Where-Object { $_.TypeAsString -eq "Lookup" }
    $LookupCache = @{} # Cache: "FieldName:Value" -> ID
    $LookupDatasetCache = @{} # Cache: "ListId|Field" -> Itens da lista de lookup
    $LookupExactIndexCache = @{} # Cache: "ListId|Field" -> Hashtable(normalizedValue -> Id)
    $FieldMap = @{} # Cache: lower(title|internalname) -> FieldInfo

    foreach ($f in $AllFields) {
        if (-not [string]::IsNullOrWhiteSpace("$($f.InternalName)")) {
            $FieldMap[$f.InternalName.ToLowerInvariant()] = $f
        }
        if (-not [string]::IsNullOrWhiteSpace("$($f.Title)")) {
            $FieldMap[$f.Title.ToLowerInvariant()] = $f
        }
    }
    
    if ($AllFields) {
        Write-Host "Campos da lista carregados: $($AllFields.Count)" -ForegroundColor Gray
    }
}
catch {
    Write-Warning "Não foi possível obter campos da lista. A importação pode falhar se os nomes das colunas não forem exatos."
}

# Configurações do Arquivo Excel
$ExcelFilePath = $ExcelPath # Caminho recebido via parâmetro

# Se o parâmetro SheetName vier vazio, define padrão
if ([string]::IsNullOrWhiteSpace($SheetName)) { $SheetName = "PESSOAS" }

# Ler dados do Excel
if (Test-Path $ExcelFilePath) {
    Write-Host "Lendo arquivo Excel: $ExcelFilePath (Aba: $SheetName)" -ForegroundColor Cyan
    $ItensParaAdicionar = @()
    $readSuccess = $false

    # 1. TENTATIVA PRIORITÁRIA: LEITURA OPENXML (rápido e sem dependências externas)
    try {
        $ItensParaAdicionar = @(Read-ExcelOpenXml -Path (Resolve-Path $ExcelFilePath).Path -WorksheetName $SheetName -IncludeEmptyColumns $false)
        if ($ItensParaAdicionar -and $ItensParaAdicionar.Count -gt 0) {
            Write-Host "Leitura via OpenXML bem sucedida! Itens: $($ItensParaAdicionar.Count)" -ForegroundColor Green
            $readSuccess = $true
        }
    } catch {
        Write-Warning "Falha na leitura OpenXML ($($_.Exception.Message))."
    }

    # 2. TENTATIVA SECUNDÁRIA: VIA COM (EXCEL INSTALADO)
    if (-not $readSuccess) {
        try {
            Write-Host "Tentando leitura via Excel COM..." -ForegroundColor Gray
            $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            
            $workbook = $excel.Workbooks.Open((Resolve-Path $ExcelFilePath).Path)
            
            try {
                $worksheet = $workbook.Worksheets.Item($SheetName)
            } catch {
                $allSheets = foreach($s in $workbook.Worksheets) { $s.Name }
                throw "Aba '$SheetName' não encontrada. Abas disponíveis: $($allSheets -join ', ')"
            }

            $usedRange = $worksheet.UsedRange
            $rowCount = $usedRange.Rows.Count
            $colCount = $usedRange.Columns.Count
            
            if ($rowCount -lt 2) { throw "Planilha vazia ou apenas cabeçalho." }

            $valueArray = $usedRange.Value2
            $headers = @()
            for ($c = 1; $c -le $colCount; $c++) {
                $headers += $valueArray[1, $c]
            }

            $ItensParaAdicionar = @()
            for ($r = 2; $r -le $rowCount; $r++) {
                $obj = New-Object PSCustomObject
                $hasData = $false
                for ($c = 1; $c -le $colCount; $c++) {
                    $val = $valueArray[$r, $c]
                    if (-not [string]::IsNullOrWhiteSpace("$val")) {
                        $val = "$val"
                        $header = $headers[$c-1]
                        if (-not [string]::IsNullOrWhiteSpace("$header")) {
                            $obj | Add-Member -MemberType NoteProperty -Name $header -Value $val -Force
                            $hasData = $true
                        }
                    }
                }
                if ($hasData) { $ItensParaAdicionar += $obj }
            }
            
            Write-Host "Leitura via COM bem sucedida! Itens: $($ItensParaAdicionar.Count)" -ForegroundColor Green
            $readSuccess = $true
        }
        catch {
            Write-Warning "Falha na leitura via COM ($($_.Exception.Message))."
        }
        finally {
            try { if ($usedRange) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) } } catch {}
            try { if ($worksheet) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) } } catch {}
            try { if ($workbook) { $workbook.Close($false); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) } } catch {}
            try { if ($excel) { $excel.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } } catch {}
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }

    # 3. TENTATIVA TERCIÁRIA: INSTALAR/USAR IMPORT-EXCEL
    if (-not $readSuccess) {
        Write-Warning "Tentando fallback para módulo Import-Excel..."

        # Verifica/Instala Import-Excel apenas se necessário
        if (-not (Get-Module -ListAvailable -Name Import-Excel)) {
            Write-Warning "O módulo Import-Excel não foi encontrado. Tentando instalar..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            
            try {
                Install-Module -Name Import-Excel -Repository PSGallery -Scope CurrentUser -Force -ErrorAction Stop
                Import-Module Import-Excel -ErrorAction Stop
            }
            catch {
                try {
                    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
                        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force -ErrorAction SilentlyContinue
                    }
                    Install-Package -Name Import-Excel -Source "https://www.powershellgallery.com/api/v2" -Scope CurrentUser -Force -ErrorAction Stop
                } catch {
                     Write-Error "FALHA CRÍTICA: Não foi possível instalar Import-Excel e a leitura via COM falhou."
                     exit 1
                }
            }
        }
        
        if (-not (Get-Module -Name Import-Excel)) { Import-Module Import-Excel -ErrorAction SilentlyContinue }

        try {
            $ItensParaAdicionar = @(Import-Excel -Path $ExcelFilePath -WorksheetName $SheetName -ErrorAction Stop)
            if (-not $ItensParaAdicionar) { Write-Warning "Nenhum dado encontrado na aba '$SheetName'." }
            else { Write-Host "Leitura via Import-Excel bem sucedida!" -ForegroundColor Green }
        } catch {
            Write-Error "ERRO AO LER EXCEL: $_"
            exit 1
        }
    }
}
else {
    Write-Error "Arquivo Excel não encontrado: $ExcelFilePath"
    Write-Host "Por favor, crie um arquivo Excel com as colunas correspondentes ao SharePoint (ex: Title)."
    exit 1
}

# === APLICAR REGRAS DE PREENCHIMENTO AUTOMÁTICO DO ANALISTA RESPONSÁVEL ===
Write-Host ""
Write-Host "Aplicando Regras de Preenchimento - Analista Responsável" -ForegroundColor Cyan

# Validar se a função Apply-AnalistRules foi carregada
if (Get-Command -Name Apply-AnalistRules -ErrorAction SilentlyContinue) {
    try {
        $ItensParaAdicionar = @(Apply-AnalistRules -Items $ItensParaAdicionar -ParqueLookupListId $ParqueLookupListId -SiteUrl $SiteUrl -DetailedLog)
        Write-Host "Regras aplicadas com sucesso!" -ForegroundColor Green
    } catch {
        Write-Warning "Erro ao aplicar regras de Analista: $_"
    }
} else {
    Write-Warning "Módulo Apply-AnalistRules não disponível. Pulando preenchimento automático do Analista."
}

Write-Host ""

# Loop para preparar e validar os itens antes do envio (evita envio parcial)
$ExecutionReport = @()
$PreparedItems = @()
$BlockingErrors = @()
$RequiredFields = $AllFields | Where-Object { $_.Required -eq $true }

for ($rowIndex = 0; $rowIndex -lt $ItensParaAdicionar.Count; $rowIndex++) {
    $Row = $ItensParaAdicionar[$rowIndex]
    $lineNum = $rowIndex + 2 # Linha 1 = cabeçalho

    try {
        # Converte a linha do Excel (PSCustomObject) para Hashtable
        $ItemValues = @{}
        $Row.PSObject.Properties | Where-Object { $_.MemberType -eq 'NoteProperty' } | ForEach-Object {
            $val = $_.Value
            $colName = $_.Name.Trim() # Remove espaços extras do nome da coluna
            
            # CORREÇÃO DE NOME DE COLUNA (Remover escape extra do Excel)
            # O Excel/Import-Excel às vezes escapa o underscore como _x005F_
            if ($colName -match '_x005F_') {
                $colName = $colName -replace '_x005F_', '_'
            }

            # Verifica se não é nulo e se, convertido para string, não é vazio ou apenas espaços
            if ($null -ne $val -and "$val".Trim().Length -gt 0) {
                
                # 1. Identificar o campo SharePoint correto (por Title ou InternalName)
                $fieldInfo = $null
                $colNameKey = $colName.ToLowerInvariant()
                if ($FieldMap.ContainsKey($colNameKey)) {
                    $fieldInfo = $FieldMap[$colNameKey]
                }
                if ($fieldInfo) {
                    $realColName = $fieldInfo.InternalName

                    # 2. Se for Lookup, resolve o ID
                    if ($fieldInfo.TypeAsString -match "Lookup") {
                        if ($val -match '^\d+$') {
                            $ItemValues[$realColName] = $val
                        }
                        else {
                            $cacheKey = "${realColName}:${val}"
                            if ($LookupCache.ContainsKey($cacheKey)) {
                                $ItemValues[$realColName] = $LookupCache[$cacheKey]
                            }
                            else {
                                # Write-Host "Resolvendo Lookup '$colName' ($realColName) para valor '$val'..." -NoNewline -ForegroundColor Gray
                                try {
                                    $targetListId = $fieldInfo.LookupList
                                    $targetInternalField = $fieldInfo.LookupField
                                    
                                    # Fallback "Parque"
                                    if ($realColName -match "Parque" -or $colName -match "Parque") {
                                        $targetListId = $ParqueLookupListId
                                        $targetInternalField = "Title"
                                    }

                                    $searchVal = "$val".Trim()
                                    $datasetKey = "$targetListId|$targetInternalField"

                                    if (-not $LookupDatasetCache.ContainsKey($datasetKey)) {
                                        $allItems = Get-PnPListItem -List $targetListId -PageSize 2000 -ErrorAction SilentlyContinue
                                        if (-not $allItems) { $allItems = @() }
                                        $LookupDatasetCache[$datasetKey] = $allItems

                                        $exactIndex = @{}
                                        foreach ($li in $allItems) {
                                            $fv = $li.FieldValues[$targetInternalField]
                                            if ($null -ne $fv) {
                                                $normalized = ("$fv").Trim().ToLowerInvariant()
                                                if (-not [string]::IsNullOrWhiteSpace($normalized) -and -not $exactIndex.ContainsKey($normalized)) {
                                                    $exactIndex[$normalized] = $li.Id
                                                }
                                            }
                                        }
                                        $LookupExactIndexCache[$datasetKey] = $exactIndex
                                    }

                                    $foundId = $null
                                    $searchNorm = $searchVal.ToLowerInvariant()

                                    if ($LookupExactIndexCache.ContainsKey($datasetKey) -and $LookupExactIndexCache[$datasetKey].ContainsKey($searchNorm)) {
                                        $foundId = $LookupExactIndexCache[$datasetKey][$searchNorm]
                                    }

                                    if (-not $foundId) {
                                        $allItems = $LookupDatasetCache[$datasetKey]
                                        $foundItem = $allItems | Where-Object { $_.FieldValues[$targetInternalField] -ilike "*$searchVal*" } | Select-Object -First 1
                                        if ($foundItem) { $foundId = $foundItem.Id }
                                    }

                                    if ($foundId) {
                                        $LookupCache[$cacheKey] = $foundId
                                        $ItemValues[$realColName] = $foundId
                                        # Write-Host " [OK ID: $foundId]" -ForegroundColor Green
                                    } else {
                                        Write-Host " [Parque não encontrado]" -ForegroundColor Red
                                    }
                                } catch {
                                    Write-Host " [Erro]" -ForegroundColor Red
                                }
                            }
                        }
                    } 
                    elseif ($fieldInfo.TypeAsString -match "DateTime") {
                        # 3. Tratamento especial para Datas (Excel via COM retorna números OLE Automation)
                        try {
                            if ($val -match '^\d+(\.\d+)?$') {
                                # É um número (ex: 45302.5), converte de OADate
                                $ItemValues[$realColName] = [DateTime]::FromOADate([double]$val)
                            } else {
                                # Tenta converter string de data
                                $ItemValues[$realColName] = [DateTime]::Parse("$val")
                            }
                        }
                        catch {
                            $BlockingErrors += "Linha ${lineNum}: valor de data invalido para o campo '$realColName': '$val'."
                        }
                    }
                    else {
                        # 4. Campo comum: Mapeia para o InternalName correto
                        $ItemValues[$realColName] = $val
                    }
                } else {
                    $normalizedColName = Normalize-TextForCompare -Text $colName
                    if ($normalizedColName.Contains("ANALISTA") -and $normalizedColName.Contains("RESP")) {
                        $ItemValues["AnalistaRespons_x00e1_vel"] = $val
                    }
                    if ($colName -ieq "Title") { $ItemValues["Title"] = $val }
                }
            }
        }

        # GARANTIA DE 'TITLE': Se a coluna Title estiver vazia mas houver colunas comuns de Equipamento, mapeia para Title
        if (-not $ItemValues.ContainsKey("Title")) {
            $possibleTitleCols = @("Nome", "Equipamento", "Modelo", "Descricao", "Tag", "Serial")
            foreach ($col in $possibleTitleCols) {
                # Procura nas chaves originais da linha do Excel
                $foundCol = $Row.PSObject.Properties | Where-Object { $_.Name -ieq $col } | Select-Object -ExpandProperty Name -First 1
                if ($foundCol -and $Row.$foundCol) {
                    $ItemValues["Title"] = $Row.$foundCol
                    # Write-Host " [Auto-mapeamento $foundCol -> Title]" -ForegroundColor Gray
                    break
                }
            }
        }

        # Preencher campos obrigatórios vazios com DefaultValue do SharePoint
        foreach ($field in $RequiredFields) {
            $internalName = $field.InternalName
            $defaultValue = $field.DefaultValue
            $isEmpty = (-not $ItemValues.ContainsKey($internalName)) -or ([string]::IsNullOrWhiteSpace("" + $ItemValues[$internalName]))
            if ($isEmpty -and -not [string]::IsNullOrWhiteSpace($defaultValue)) {
                $resolvedDefault = Resolve-SharePointDefaultValue -Field $field -DefaultValue $defaultValue
                if ($null -ne $resolvedDefault) {
                    $ItemValues[$internalName] = $resolvedDefault
                }
            }
        }

        # Bloqueia se existir campo obrigatório sem valor e sem default
        $missingRequired = @()
        foreach ($field in $RequiredFields) {
            $internalName = $field.InternalName
            $fieldTitle = $field.Title
            $defaultValue = $field.DefaultValue
            $isEmpty = (-not $ItemValues.ContainsKey($internalName)) -or ([string]::IsNullOrWhiteSpace("" + $ItemValues[$internalName]))
            $hasDefault = -not [string]::IsNullOrWhiteSpace("$defaultValue")

            if ($isEmpty -and -not $hasDefault) {
                $missingRequired += "$fieldTitle ($internalName)"
            }
        }

        if ($missingRequired.Count -gt 0) {
            $BlockingErrors += "Linha ${lineNum}: campos obrigatórios sem valor e sem default no SharePoint: $($missingRequired -join ', ')"
            continue
        }

        # Se não houver nenhuma coluna com dados, pula a linha
        if ($ItemValues.Count -eq 0) {
            Write-Warning "Linha sem dados encontrada no Excel. Ignorando..."
            continue
        }

        $PreparedItems += [PSCustomObject]@{
            LinhaExcel = $lineNum
            Values = $ItemValues
        }
    }
    catch {
        $errorDetail = $_.Exception.Message
        $BlockingErrors += "Linha ${lineNum}: erro ao preparar dados para envio: $errorDetail"
    }
}

if ($BlockingErrors.Count -gt 0) {
    Write-Host "" 
    Write-Host "UPLOAD CANCELADO: Foram encontrados campos obrigatórios sem valor e sem default." -ForegroundColor Red
    foreach ($be in $BlockingErrors) {
        Write-Host " - $be" -ForegroundColor Red
    }
    exit 1
}

if ($PreparedItems.Count -eq 0) {
    Write-Error "Nenhum item válido para envio foi preparado."
    exit 1
}

foreach ($prepared in $PreparedItems) {
    try {
        $ItemValues = $prepared.Values

        # DEBUG: Mostrar chaves sendo enviadas (opcional)
        # Write-Host "Chaves: $($ItemValues.Keys -join ', ')" -ForegroundColor Gray

        Write-Host "Adicionando item: $($ItemValues.Title)..." -NoNewline
        
        # Adiciona o item à lista usando o ID da lista
        $novoItem = Add-PnPListItem -List $ListId -Values $ItemValues -ErrorAction Stop
        
        Write-Host " [OK] (ID: $($novoItem.Id))" -ForegroundColor Green
        Write-Host "--- RESULT: SUCCESS ---" -ForegroundColor Gray

        # Adiciona ao relatório
        $reportTitle = "Item"
        if ($ItemValues.Title) { $reportTitle = $ItemValues.Title }
        
        $ExecutionReport += [PSCustomObject]@{
            "Linha" = $prepared.LinhaExcel
            "Item"  = $reportTitle
            "ID"    = $novoItem.Id
            "Status" = "Sucesso"
        }
    }
    catch {
        Write-Host ' [ERRO]' -ForegroundColor Red
        $errorDetail = $_.Exception.Message
        Write-Error "Erro ao adicionar item: $errorDetail"
        Write-Host '--- RESULT: ERROR ---' -ForegroundColor Gray

        # Adiciona ao relatório de erro
        $reportTitle = "Erro na linha"
        $ExecutionReport += [PSCustomObject]@{
            "Linha" = $prepared.LinhaExcel
            "Item"  = $reportTitle
            "ID"    = "N/A"
            "Status" = "Erro: $errorDetail"
        }
    }
}

Write-Host ""
Write-Host '=== RELATORIO DE IMPORTACAO ===' -ForegroundColor Cyan
if ($ExecutionReport) {
    $ExecutionReport | Format-Table -Property Linha, Item, ID, Status -AutoSize | Out-String | Write-Host
} else {
    Write-Host 'Nenhum item foi processado.' -ForegroundColor Yellow
}

Write-Host '=== FINAL SUMMARY ===' -ForegroundColor Cyan
Write-Host 'Processo finalizado.' -ForegroundColor Cyan

$hasUploadErrors = $ExecutionReport | Where-Object { "$($_.Status)" -like "Erro:*" } | Select-Object -First 1
if ($hasUploadErrors) {
    exit 1
}

exit 0
