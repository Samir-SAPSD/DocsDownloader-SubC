#Requires -Modules PnP.PowerShell

param(
    [Parameter(Mandatory=$true)]
    [string]$Ids,  # IDs separados por vírgula (ex: "1,5,10")

    [string]$TargetDir = "$env:USERPROFILE\Downloads"
)

# ==== CONFIGURAÇÃO ====
$siteUrl   = "https://vestas.sharepoint.com/sites/CC-Subcontractors-BR"
$listaGuid = "205a1e3b-9c65-4733-b67f-0effd21b7953"

# ==== CONEXÃO ====
try {
    Write-Host "Conectando ao SharePoint ($siteUrl)..." -ForegroundColor Cyan
    
    # Força uma nova conexão interativa e guarda o objeto de conexão
    # Isso isola este script de outros contextos que possam estar abertos
    $conn = Connect-PnPOnline -Url $siteUrl -UseWebLogin
    
    Write-Host "Conexão estabelecida com sucesso!" -ForegroundColor Green

    # Mapeia colunas para usar no nome da pasta
    $fields = Get-PnPField -List $listaGuid -Connection $conn
    $fEmpresa = $fields | Where-Object { $_.Title -eq "EMPRESA" -or $_.Title -eq "COMPANY" } | Select-Object -First 1
    
    $fIdent = $fields | Where-Object { $_.InternalName -eq "NOME" -or $_.Title -eq "IDENTIFICAÇÃO" -or $_.Title -eq "IDENTIFICACAO" } | Select-Object -First 1

    $fEquip   = $fields | Where-Object { $_.Title -eq "EQUIPAMENTO" -or $_.Title -eq "EQUIPMENT" } | Select-Object -First 1
}
catch {
    Write-Error "ERRO CRÍTICO AO CONECTAR: $_"
    Write-Host "Detalhes do erro: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Pressione qualquer tecla para sair..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# ==== PROCESSAMENTO ====
$idList = $Ids -split ","

if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
}

foreach ($id in $idList) {
    # Limpa espaços em branco
    $id = $id.Trim()
    if ([string]::IsNullOrWhiteSpace($id)) { continue }

    try {
        Write-Host "Processando Item ID: $id..." -ForegroundColor Cyan
        
        # 1. Obtém o item (sem o parâmetro -Includes que estava falhando)
        $item = Get-PnPListItem -List $listaGuid -Id $id -Connection $conn -ErrorAction Stop

        # 2. Carrega explicitamente a propriedade AttachmentFiles usando Get-PnPProperty
        # Isso funciona em todas as versões do PnP PowerShell
        $attachments = Get-PnPProperty -ClientObject $item -Property AttachmentFiles -Connection $conn

        if ($attachments.Count -gt 0) {
            # Resolve metadados para nome da pasta
            $valEmpresa = if ($fEmpresa) { $item[$fEmpresa.InternalName] } else { "Empresa" }
            
            $valIdent = if ($fIdent) { $item[$fIdent.InternalName]  } else { "Identificacao" }

            $valEquip   = if ($fEquip)   { $item[$fEquip.InternalName] }   else { "Equipamento" }

            # Helper para extrair texto de Lookup/Choice
            function Get-Str($v) { if ($v -is [Microsoft.SharePoint.Client.FieldLookupValue]) { return $v.LookupValue } return [string]$v }
            
            $sEmpresa = Get-Str $valEmpresa
            $sIdent   = Get-Str $valIdent
            $sEquip   = Get-Str $valEquip
            $sDate    = Get-Date -Format "yyyyMMdd"

            # Sanitiza para nome de pasta (remove caracteres inválidos)
            $folderName = "{0}_{1}_{2}_{3}" -f $sIdent, $sEquip, $sEmpresa, $sDate
            $folderName = $folderName -replace '[\\/:*?"<>|]', '' -replace '\s+', '_'
            
            $itemDir = Join-Path $TargetDir $folderName
            if (-not (Test-Path $itemDir)) { New-Item -ItemType Directory -Path $itemDir -Force | Out-Null }

            foreach ($file in $attachments) {
                $fileName = $file.FileName
                $serverRelativeUrl = $file.ServerRelativeUrl
                
                Write-Host "  Baixando: $fileName" -ForegroundColor Green
                
                # IMPORTANTE: Passamos -Connection $conn explicitamente
                Get-PnPFile -Url $serverRelativeUrl -Path $itemDir -FileName $fileName -AsFile -Force -Connection $conn
            }
        } else {
            Write-Host "  Sem anexos." -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "Erro ao processar ID $id : $_" -ForegroundColor Red
    }
}

Write-Host "Concluído. Arquivos salvos em: $TargetDir" -ForegroundColor Yellow

# Aguarda 2 segundos para você ver a mensagem de sucesso e fecha o terminal
Start-Sleep -Seconds 2
Stop-Process -Id $