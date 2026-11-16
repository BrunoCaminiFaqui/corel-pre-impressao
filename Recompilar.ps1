$ErrorActionPreference = 'Stop'
try {
    $framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319'
    if (-not (Test-Path -LiteralPath (Join-Path $framework 'csc.exe'))) {
        $framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319'
    }
    $compilador = Join-Path $framework 'csc.exe'
    if (-not (Test-Path -LiteralPath $compilador)) { throw 'Compilador .NET Framework 4 nao encontrado.' }

    function Compilar([string]$saida, [string[]]$fontes) {
        if (Test-Path -LiteralPath $saida) { Remove-Item -LiteralPath $saida -Force }
        $argumentos = @('/nologo', '/target:library', '/platform:anycpu', '/optimize+', ('/out:' + $saida),
            '/r:System.dll', '/r:System.Core.dll', '/r:Microsoft.CSharp.dll',
            ('/r:' + (Join-Path $framework 'WPF\WindowsBase.dll')),
            ('/r:' + (Join-Path $framework 'WPF\PresentationCore.dll')),
            ('/r:' + (Join-Path $framework 'WPF\PresentationFramework.dll')),
            ('/r:' + (Join-Path $framework 'System.Xaml.dll'))) + $fontes
        # saida do compilador direto na tela (nao entra no valor de retorno da funcao)
        & $compilador @argumentos | Out-Host
        if ($LASTEXITCODE -ne 0) {
            if (Test-Path -LiteralPath $saida) { Remove-Item -LiteralPath $saida -Force }
            return $false
        }
        (Get-FileHash -LiteralPath $saida -Algorithm SHA256).Hash |
            Set-Content -LiteralPath ([IO.Path]::ChangeExtension($saida, '.sha256')) -Encoding ASCII
        return $true
    }

    # Painel de separacoes (obrigatorio)
    $ok = Compilar (Join-Path $PSScriptRoot 'painel\SeparacoesCorel.dll') @(
        (Join-Path $PSScriptRoot 'fonte\Separacoes.cs'), (Join-Path $PSScriptRoot 'fonte\MotorSeparacoes.cs'))
    if (-not $ok) { throw 'A compilacao do painel de separacoes falhou. Nenhum arquivo foi instalado no Corel.' }
    Write-Host 'Painel Separacoes CMYK + Pantone: compilado.' -ForegroundColor Green

    # Painel Trocar cores (se falhar, as separacoes sao instaladas mesmo assim)
    $okCores = Compilar (Join-Path $PSScriptRoot 'painel_cores\TrocarCoresCorel.dll') @(
        (Join-Path $PSScriptRoot 'fonte\TrocarCores.cs'))
    if ($okCores) { Write-Host 'Painel Trocar cores: compilado.' -ForegroundColor Green }
    else { Write-Host 'ATENCAO: a compilacao do painel Trocar cores falhou (veja a mensagem acima). Ele nao sera instalado.' -ForegroundColor Yellow }
    exit 0
} catch {
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
