#Requires -Modules PnP.PowerShell
# If this is your first time using ImportExcel, install it:
# Install-Module ImportExcel -Scope CurrentUser -Force

# ==== PARAMETERS ====
$siteUrl   = "https://vestas.sharepoint.com/sites/CC-Subcontractors-BR"
$listGuid  = "205a1e3b-9c65-4733-b67f-0effd21b7953"

# Salva na raiz do projeto (pasta pai de onde este script está)
$projectRoot = Split-Path -Parent $PSScriptRoot
$outputXlsx = Join-Path $projectRoot "DefaultView-Data.xlsx"

# ==== CONNECTION ====
# If you use MFA, you may prefer: -Interactive
Connect-PnPOnline -Url $siteUrl -UseWebLogin

# ==== METADATA: DEFAULT VIEW COLUMNS ====
Write-Host "Obtendo colunas da View Padrão..." -ForegroundColor Cyan

# 1. Get the Default View
$view = Get-PnPView -List $listGuid | Where-Object { $_.DefaultView -eq $true }
if (-not $view) { throw "Nenhuma view padrão encontrada na lista." }

# 2. Get All Fields (to look up Display Names later)
$allFields = Get-PnPField -List $listGuid

# 3. Map "View Fields" (which might be pseudo-fields like LinkTitle) to "Real Fields"
$pseudoMap = @{
    "LinkTitle"          = "Title"
    "LinkTitleNoMenu"    = "Title"
    "LinkFilename"       = "FileLeafRef"
    "LinkFilenameNoMenu" = "FileLeafRef"
    "Edit"               = "" 
    "DocIcon"            = "" 
}

# 4. Build the Columns List based strictly on the View Order
$usedNames = @{ }
$columns = @()

# === ADIÇÃO: Força a inclusão da coluna ID no início ===
$columns += [pscustomobject]@{ Internal = "ID"; Title = "ID" }
$usedNames["ID"] = 1

foreach ($vf in $view.ViewFields) {
    $realInternalName = if ($pseudoMap.ContainsKey($vf)) { $pseudoMap[$vf] } else { $vf }
    if ([string]::IsNullOrEmpty($realInternalName)) { continue }
    if ($realInternalName -eq "ID") { continue }

    $fDef = $allFields | Where-Object { $_.InternalName -eq $realInternalName }
    $title = if ($fDef -and -not [string]::IsNullOrWhiteSpace($fDef.Title)) { $fDef.Title } else { $realInternalName }

    if ($usedNames.ContainsKey($title)) {
        $usedNames[$title]++
        $title = "$title ($($usedNames[$title]))"
    } else {
        $usedNames[$title] = 1
    }

    $columns += [pscustomobject]@{ Internal = $realInternalName; Title = $title }
}

Write-Host "Colunas encontradas: $($columns.Count)" -ForegroundColor Gray

# ==== COLLECTING ITEMS (OPTIMIZED) ====
Write-Host "Baixando itens ..." -ForegroundColor Cyan
$fieldsToLoad = $columns.Internal | Select-Object -Unique
$items = Get-PnPListItem -List $listGuid -Fields $fieldsToLoad

# ==== TRANSFORMING TO OBJECTS (OPTIMIZED) ====
Write-Host "Processando $($items.Count) registros..." -ForegroundColor Cyan

# Otimização: Lista Genérica para performance
$dataList = New-Object System.Collections.Generic.List[object]

foreach ($item in $items) {
    $o = [ordered]@{}
    
    # Otimização: Acesso direto ao dicionário de valores (evita chamadas COM)
    $fieldValues = $item.FieldValues
    
    foreach ($col in $columns) {
        $key = $col.Internal
        $val = $null
        
        # Tenta pegar do dicionário (Rápido)
        if ($fieldValues.ContainsKey($key)) {
            $val = $fieldValues[$key]
        } else {
            # Fallback (Lento, para colunas computadas)
            try { $val = $item[$key] } catch {}
        }

        # Lógica In-Line (Remove overhead de função)
        if ($null -eq $val) {
            $o[$col.Title] = ""
        }
        else {
            # Verificações de tipo otimizadas
            if ($val -is [Microsoft.SharePoint.Client.FieldLookupValue]) {
                $o[$col.Title] = $val.LookupValue
            }
            elseif ($val -is [Microsoft.SharePoint.Client.FieldUserValue]) {
                $o[$col.Title] = $val.LookupValue
            }
            elseif ($val -is [string]) {
                $o[$col.Title] = $val
            }
            elseif ($val -is [DateTime]) {
                $o[$col.Title] = $val.ToString("yyyy-MM-dd HH:mm")
            }
            elseif ($val -is [bool]) {
                $o[$col.Title] = if ($val) { "Sim" } else { "Não" }
            }
            elseif ($val -is [Microsoft.SharePoint.Client.FieldUrlValue]) {
                $o[$col.Title] = if ($val.Description) { $val.Description } else { $val.Url }
            }
            # Taxonomy (Verificação por nome para evitar erro de assembly)
            elseif ($val.GetType().Name -like "*TaxonomyFieldValue*") {
                try { $o[$col.Title] = $val.Label } catch { $o[$col.Title] = $val.ToString() }
            }
            # Arrays (Multi-choice/User)
            elseif ($val -is [System.Collections.IEnumerable]) {
                $parts = @()
                foreach ($sub in $val) {
                    if ($sub -is [Microsoft.SharePoint.Client.FieldLookupValue]) { $parts += $sub.LookupValue }
                    elseif ($sub -is [Microsoft.SharePoint.Client.FieldUserValue]) { $parts += $sub.LookupValue }
                    elseif ($sub.GetType().Name -like "*TaxonomyFieldValue*") { try { $parts += $sub.Label } catch { $parts += $sub.ToString() } }
                    else { $parts += $sub.ToString() }
                }
                $o[$col.Title] = $parts -join "; "
            }
            else {
                $o[$col.Title] = $val.ToString()
            }
        }
    }
    $dataList.Add([pscustomobject]$o)
}

# ==== EXPORT TO EXCEL ====
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

Write-Host "Gerando Excel..." -ForegroundColor Cyan
$dataList |
    Export-Excel `
        -Path $outputXlsx `
        -WorksheetName "DefaultView" `
        -TableName "ViewData" `
        -TableStyle Medium6 `
        -AutoSize `
        -AutoFilter `
        -FreezeTopRow `
        -BoldTopRow `
        -ClearSheet 

Write-Host "File generated at: $outputXlsx" -ForegroundColor Green