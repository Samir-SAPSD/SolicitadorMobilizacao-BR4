param(
    [Parameter(Mandatory=$false)]
    [string]$ExcelPath = ".\DadosParaImportar.xlsx",
    
    [Parameter(Mandatory=$false)]
    [string]$SheetName = "PESSOAS"
)

# Configurações
$TestMode = $true # Altere para $false para usar a lista de produção
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

# Verifica se o módulo PnP.PowerShell está instalado
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Warning "O módulo PnP.PowerShell não foi encontrado. Tentando instalar..."
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
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
    $AllFields = Get-PnPField -List $ListId | Select-Object InternalName, Title, TypeAsString, LookupList, LookupField
    $LookupFields = $AllFields | Where-Object { $_.TypeAsString -eq "Lookup" }
    $LookupCache = @{} # Cache: "FieldName:Value" -> ID
    
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

# # Verifica se o módulo Import-Excel está instalado
# if (-not (Get-Module -ListAvailable -Name Import-Excel)) {
#     Write-Warning "O módulo Import-Excel não foi encontrado. Tentando instalar..."
#     Install-Module -Name Import-Excel -Scope CurrentUser -Force
# }

# Ler dados do Excel
if (Test-Path $ExcelFilePath) {
    Write-Host "Lendo arquivo Excel: $ExcelFilePath (Aba: $SheetName)" -ForegroundColor Cyan
    try {
        $ItensParaAdicionar = Import-Excel -Path $ExcelFilePath -WorksheetName $SheetName -ErrorAction Stop
    }
    catch {
        Write-Error "Erro ao ler a aba '$SheetName'. Verifique se ela existe no arquivo."
        exit
    }
}
else {
    Write-Error "Arquivo Excel não encontrado: $ExcelFilePath"
    Write-Host "Por favor, crie um arquivo Excel com as colunas correspondentes ao SharePoint (ex: Title)."
    exit
}

# Loop para adicionar os itens
$ExecutionReport = @()

foreach ($Row in $ItensParaAdicionar) {
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
                $fieldInfo = $AllFields | Where-Object { $_.InternalName -eq $colName -or $_.Title -eq $colName }
                if ($fieldInfo) {
                    if ($fieldInfo -is [array]) { $fieldInfo = $fieldInfo[0] }
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
                                    $allItems = Get-PnPListItem -List $targetListId -PageSize 500 -ErrorAction SilentlyContinue
                                    $foundItem = $null
                                    
                                    if ($allItems) {
                                        $foundItem = $allItems | Where-Object { $_.FieldValues[$targetInternalField] -ieq $searchVal } | Select-Object -First 1
                                        if (!$foundItem) {
                                            $foundItem = $allItems | Where-Object { $_.FieldValues[$targetInternalField] -ilike "*$searchVal*" } | Select-Object -First 1
                                        }
                                    }

                                    if ($foundItem) {
                                        $foundId = $foundItem.Id
                                        $LookupCache[$cacheKey] = $foundId
                                        $ItemValues[$realColName] = $foundId
                                        Write-Host " [OK ID: $foundId]" -ForegroundColor Green
                                    } else {
                                        Write-Host " [Não encontrado]" -ForegroundColor Red
                                    }
                                } catch {
                                    Write-Host " [Erro]" -ForegroundColor Red
                                }
                            }
                        }
                    } else {
                        # 3. Campo comum: Mapeia para o InternalName correto
                        $ItemValues[$realColName] = $val
                    }
                } else {
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

        # Se não houver nenhuma coluna com dados, pula a linha
        if ($ItemValues.Count -eq 0) {
            Write-Warning "Linha sem dados encontrada no Excel. Ignorando..."
            continue
        }

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
            "Linha" = $ExecutionReport.Count + 1
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
            "Linha" = $ExecutionReport.Count + 1
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
