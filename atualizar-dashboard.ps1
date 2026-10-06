param(
    [string]$Arquivo = 'P:\COMERCIAL\Propostas em Elaboração\Controle de Proposta\03 - CONTROLE DE PROPOSTAS - 2026.xlsb',
    [string]$Branch = 'main',
    [switch]$SemPausa
)
$ErrorActionPreference = 'Stop'
try {
    if (-not (Test-Path -LiteralPath $Arquivo -PathType Leaf)) {
        throw "Planilha não encontrada: $Arquivo. Verifique a unidade P: ou informe -Arquivo com um caminho local."
    }
    if ([IO.Path]::GetExtension($Arquivo) -ne '.xlsb') { throw 'Selecione a planilha no formato .xlsb.' }
    $token = $env:GITHUB_TOKEN
    if (-not $token) {
        $segredo = Read-Host 'Informe o token do GitHub com acesso de escrita ao repositório' -AsSecureString
        $token = [Net.NetworkCredential]::new('', $segredo).Password
    }
    if (-not $token) { throw 'Token não informado.' }
    $headers = @{ Authorization = "Bearer $token"; Accept = 'application/vnd.github+json'; 'User-Agent' = 'ProTrafo-Dashboard'; 'X-GitHub-Api-Version' = '2022-11-28' }
    $api = 'https://api.github.com/repos/protrafocomercial/Dashboard-Comercial/contents/planilha.xlsb'
    $ref = [Uri]::EscapeDataString($Branch)
    $bytes = [IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $Arquivo).Path)
    if ($bytes.Length -lt 1000 -or $bytes[0] -ne 80 -or $bytes[1] -ne 75) { throw 'O arquivo não parece ser uma planilha XLSB válida.' }
    Write-Host 'Enviando planilha atualizada...' -ForegroundColor Cyan
    $anterior = Invoke-RestMethod -Uri "$api`?ref=$ref" -Headers $headers
    $body = @{ message = "Atualização da base comercial - $(Get-Date -Format 'dd/MM/yyyy HH:mm')"; content = [Convert]::ToBase64String($bytes); sha = $anterior.sha; branch = $Branch } | ConvertTo-Json
    $resultado = Invoke-RestMethod -Uri $api -Headers $headers -Method Put -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($body))
    Write-Host "Planilha enviada. Commit: $($resultado.commit.sha)" -ForegroundColor Green
    Write-Host 'Reabra o dashboard para carregar os dados. A publicação do site pode levar alguns instantes.'
} catch {
    Write-Host "Falha na atualização: $($_.Exception.Message)" -ForegroundColor Red
    if (-not $SemPausa) { [void](Read-Host 'Pressione ENTER para fechar') }
    exit 1
} finally {
    $token = $null
    $headers = $null
}
if (-not $SemPausa) { [void](Read-Host 'Pressione ENTER para fechar') }
