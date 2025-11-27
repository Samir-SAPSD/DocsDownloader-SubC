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
    "Edit"               = "" # Skip edit button
    "DocIcon"            = "" # Skip icon
}

# 4. Build the Columns List based strictly on the View Order
$usedNames = @{ }
$columns = @()

# === ADIÇÃO: Força a inclusão da coluna ID no início ===
$columns += [pscustomobject]@{
    Internal = "ID"
    Title    = "ID"
}
$usedNames["ID"] = 1
# ======================================================

foreach ($vf in $view.ViewFields) {
    # Resolve real internal name
    $realInternalName = if ($pseudoMap.ContainsKey($vf)) { $pseudoMap[$vf] } else { $vf }
    
    # Skip if mapped to empty (like Edit button)
    if ([string]::IsNullOrEmpty($realInternalName)) { continue }

    # === ADIÇÃO: Pula se for ID (pois já adicionamos manualmente acima) ===
    if ($realInternalName -eq "ID") { continue }
    # =====================================================================

    # Find the field definition to get the Display Title
    $fDef = $allFields | Where-Object { $_.InternalName -eq $realInternalName }
    
    # Determine Title (Display Name or Internal if not found)
    $title = if ($fDef -and -not [string]::IsNullOrWhiteSpace($fDef.Title)) { $fDef.Title } else { $realInternalName }

    # Handle duplicate titles in the header
    if ($usedNames.ContainsKey($title)) {
        $usedNames[$title]++
        $title = "$title ($($usedNames[$title]))"
    } else {
        $usedNames[$title] = 1
    }

    $columns += [pscustomobject]@{
        Internal = $realInternalName
        Title    = $title
    }
}

Write-Host "Colunas encontradas: $($columns.Count)" -ForegroundColor Gray

# ==== COLLECTING ITEMS ====
# We explicitly request the fields we identified to ensure we get the data
$fieldsToLoad = $columns.Internal | Select-Object -Unique
$items = Get-PnPListItem -List $listGuid -PageSize 2000 -Fields $fieldsToLoad

# ==== FUNCTION TO NORMALIZE VALUES (ROBUST) ====
# Adapted from exportSharepoint.ps1 to handle more types
function Resolve-SharePointField {
    param([Parameter(ValueFromPipeline=$true)]$Value)
    if ($null -eq $Value) { return "" }

    $typeName = $Value.GetType().Name
    switch ($typeName) {
        "FieldLookupValue" { return $Value.LookupValue }
        "FieldUserValue"   { return $Value.LookupValue }
        "FieldUrlValue"    { return if ($Value.Description) { $Value.Description } else { $Value.Url } }
        "DateTime"         { return $Value.ToString("yyyy-MM-dd HH:mm") }
        "Boolean"          { return if ($Value) { "Sim" } else { "Não" } }

        # Managed Metadata (if applicable)
        "TaxonomyFieldValue"          { return $Value.Label }
        "TaxonomyFieldValueCollection"{ return ($Value | ForEach-Object { $_.Label }) -join "; " }

        default {
            if ($Value -is [System.Array]) {
                return ($Value | ForEach-Object { Resolve-SharePointField $_ }) -join "; "
            }
            return $Value.ToString()
        }
    }
}

# ==== TRANSFORMING TO OBJECTS ====
$data = foreach ($item in $items) {
    $o = [ordered]@{}
    foreach ($col in $columns) {
        try {
            # Access property safely using Internal Name
            $val = $item[$col.Internal]
            # Map to Display Name in the output object
            $o[$col.Title] = Resolve-SharePointField $val
        } catch {
            # If field is missing in the item object (e.g. computed column), leave blank
            $o[$col.Title] = ""
        }
    }
    [pscustomobject]$o
}

# ==== EXPORT TO EXCEL ====
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

$data |
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