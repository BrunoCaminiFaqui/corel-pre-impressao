param([string]$CorelExe, [switch]$Remover, [string]$Registro)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2
$idPainel = '0a903144-d766-44d9-a345-27143b2e26c6'

# Registro local de cada tentativa, inclusive quando falta permissao ou o Corel esta aberto.
$gravandoRegistro = $false
$caminhoRegistro = $null
try {
    $pastaRegistro = Join-Path $env:LOCALAPPDATA 'Imprell\SeparacoesCorel'
    New-Item -ItemType Directory -Path $pastaRegistro -Force | Out-Null
    if ($Registro) { $caminhoRegistro = $Registro }
    else {
    $caminhoRegistro = Join-Path $pastaRegistro ('instalacao-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.log')
    }
    Start-Transcript -LiteralPath $caminhoRegistro -Force | Out-Null
    $gravandoRegistro = $true
} catch {
    # A falta de registro nao impede a instalacao nem esconde o erro principal.
    Write-Host ('Nao foi possivel iniciar o registro: ' + $_.Exception.Message) -ForegroundColor Yellow
}


function Escolher-Corel {
    if ($CorelExe) { return $CorelExe }
    # mesmo Corel escolhido na instalacao das separacoes
    $salvo = Join-Path $env:LOCALAPPDATA 'Imprell\SeparacoesCorel\corel-escolhido.txt'
    if (Test-Path -LiteralPath $salvo) {
        $caminhoSalvo = ([string](Get-Content -LiteralPath $salvo -Raw)).Trim()
        if ($caminhoSalvo -and (Test-Path -LiteralPath $caminhoSalvo -PathType Leaf)) { return $caminhoSalvo }
    }
    $achados = New-Object 'System.Collections.Generic.List[string]'
    function Adicionar-Candidato([string]$arquivo) {
        if ($arquivo -and (Test-Path -LiteralPath $arquivo -PathType Leaf)) {
            $normal = (Get-Item -LiteralPath $arquivo).FullName
            if ([IO.Path]::GetFileName($normal) -ieq 'CorelDRW.exe' -and -not $achados.Contains($normal)) {
                $achados.Add($normal)
            }
        }
    }
    foreach ($chave in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\CorelDRW.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\CorelDRW.exe',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\CorelDRW.exe')) {
        if (Test-Path -LiteralPath $chave) { Adicionar-Candidato ((Get-Item -LiteralPath $chave).GetValue('')) }
    }
    $raizes = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432) | Where-Object { $_ } | Select-Object -Unique
    foreach ($raiz in $raizes) {
        $pastaCorel = Join-Path $raiz 'Corel'
        if (-not (Test-Path -LiteralPath $pastaCorel)) { continue }
        foreach ($suite in @(Get-ChildItem -LiteralPath $pastaCorel -Directory -ErrorAction SilentlyContinue)) {
            foreach ($bin in @('Programs64', 'Programs')) {
                Adicionar-Candidato (Join-Path $suite.FullName ($bin + '\CorelDRW.exe'))
            }
        }
    }
    if ($achados.Count -eq 1) { return $achados[0] }
    if ($achados.Count -gt 1) {
        Write-Host 'Foram encontradas estas instalacoes:'
        for ($i = 0; $i -lt $achados.Count; $i++) { Write-Host ('{0}. {1}' -f ($i + 1), $achados[$i]) }
        $escolha = Read-Host 'Digite o numero da instalacao que voce usa'
        $numero = 0
        if (-not [int]::TryParse($escolha, [ref]$numero) -or $numero -lt 1 -or $numero -gt $achados.Count) {
            throw 'Selecao invalida. Nenhum arquivo foi instalado.'
        }
        return $achados[$numero - 1]
    }
    Add-Type -AssemblyName System.Windows.Forms
    $dialogo = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $dialogo.Title = 'Selecione CorelDRW.exe na pasta Programs64 ou Programs do Corel'
        $dialogo.Filter = 'CorelDRAW (CorelDRW.exe)|CorelDRW.exe'
        $dialogo.CheckFileExists = $true
        if ($dialogo.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw 'Instalacao cancelada.' }
        return $dialogo.FileName
    } finally { $dialogo.Dispose() }
}

try {
    $identidade = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identidade)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $lancador = if ($Remover) { 'Desinstalar.cmd' } else { 'Instalar.cmd' }
        throw ('Feche esta janela. Clique com o botao direito em ' + $lancador + ' e escolha Executar como administrador.')
    }
    if (@(Get-Process -Name CorelDRW -ErrorAction SilentlyContinue).Count -gt 0) {
        throw 'Salve suas artes e feche o CorelDRAW antes de instalar ou remover o painel.'
    }
    $exeEscolhido = Escolher-Corel
    $exe = Get-Item -LiteralPath $exeEscolhido
    if ($exe.Name -ine 'CorelDRW.exe') { throw 'O arquivo selecionado precisa ser CorelDRW.exe.' }
    $versao = $exe.VersionInfo.FileMajorPart
    if ($versao -lt 23) { throw 'Esta versao do pacote foi preparada para CorelDRAW 2021 ou posterior, para Windows.' }
    Write-Host ('Corel selecionado: {0} (versao interna {1})' -f $exe.FullName, $versao)
    try {
        $null = $exe
    } catch { }
    $addons = Join-Path $exe.DirectoryName 'Addons'
    $destino = Join-Path $addons 'TrocarCoresCorel'
    $origem = Join-Path $PSScriptRoot 'painel_cores'

    if (Test-Path -LiteralPath $destino) {
        $pastaDestino = Get-Item -LiteralPath $destino
        if (($pastaDestino.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'O destino e um link de diretorio. Nada foi alterado.'
        }
        $identificador = Join-Path $destino 'painel.id'
        if (-not (Test-Path -LiteralPath $identificador) -or (Get-Content -LiteralPath $identificador -Raw).Trim() -ne $idPainel) {
            throw 'Ja existe uma pasta TrocarCoresCorel de outra origem. Nada foi substituido.'
        }
    }
    if (-not $Remover) {
        foreach ($nome in @('TrocarCoresCorel.dll', 'AppUI.xslt', 'UserUI.xslt', 'CorelDrw.addon', 'painel.id')) {
            if (-not (Test-Path -LiteralPath (Join-Path $origem $nome) -PathType Leaf)) {
                throw ('Pacote incompleto: falta {0}. Extraia todo o ZIP antes de instalar.' -f $nome)
            }
        }
        $hashEsperado = (Get-Content -LiteralPath (Join-Path $origem 'TrocarCoresCorel.sha256') -Raw).Trim().Split(' ')[0]
        $hashReal = (Get-FileHash -LiteralPath (Join-Path $origem 'TrocarCoresCorel.dll') -Algorithm SHA256).Hash
        if ($hashReal -ine $hashEsperado) { throw 'A DLL nao corresponde ao pacote. Baixe e extraia o ZIP novamente.' }
        $nomeAssembly = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $origem 'TrocarCoresCorel.dll'))
        if ($nomeAssembly.Name -ne 'TrocarCoresCorel') { throw 'A biblioteca nao corresponde a este painel.' }
        [xml]$xmlApp = Get-Content -LiteralPath (Join-Path $origem 'AppUI.xslt') -Raw
        [xml]$xmlUsuario = Get-Content -LiteralPath (Join-Path $origem 'UserUI.xslt') -Raw
    }

    $backup = $null
    if (Test-Path -LiteralPath $destino) {
        $baseBackup = Join-Path $env:LOCALAPPDATA 'Imprell\SeparacoesCorel\Backups-TrocarCores'
        New-Item -ItemType Directory -Path $baseBackup -Force | Out-Null
        $backup = Join-Path $baseBackup ((Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
    }
    if ($Remover) {
        if ($backup) { Move-Item -LiteralPath $destino -Destination $backup }
        Write-Host 'Painel Trocar cores removido.' -ForegroundColor Green
        if ($backup) { Write-Host ('Copia de seguranca: ' + $backup) }
        exit 0
    }

    # Preparacao local e copia somente para a pasta deste complemento.
    New-Item -ItemType Directory -Path $addons -Force | Out-Null
    $estagio = Join-Path $addons ('TrocarCoresCorel-preparacao-' + [guid]::NewGuid().ToString('N'))
    $substituido = $false
    $novoDestino = $false
    try {
        New-Item -ItemType Directory -Path $estagio | Out-Null
        # O marcador .addon so e copiado depois dos demais arquivos.
        foreach ($nome in @('TrocarCoresCorel.dll', 'AppUI.xslt', 'UserUI.xslt', 'painel.id', 'TrocarCoresCorel.sha256')) {
            Copy-Item -LiteralPath (Join-Path $origem $nome) -Destination (Join-Path $estagio $nome)
            Unblock-File -LiteralPath (Join-Path $estagio $nome)
        }
        if (@(Get-Process -Name CorelDRW -ErrorAction SilentlyContinue).Count -gt 0) {
            throw 'O Corel foi aberto durante a instalacao. Feche-o e tente novamente.'
        }
        if ($backup) { Move-Item -LiteralPath $destino -Destination $backup; $substituido = $true }
        Move-Item -LiteralPath $estagio -Destination $destino
        $novoDestino = $true
        Copy-Item -LiteralPath (Join-Path $origem 'CorelDrw.addon') -Destination (Join-Path $destino 'CorelDrw.addon')
    } catch {
        if ($novoDestino -and (Test-Path -LiteralPath $destino)) { Remove-Item -LiteralPath $destino -Recurse -Force }
        if ($substituido -and $backup -and (Test-Path -LiteralPath $backup)) {
            Move-Item -LiteralPath $backup -Destination $destino
        }
        throw
    } finally {
        if (Test-Path -LiteralPath $estagio) { Remove-Item -LiteralPath $estagio -Recurse -Force }
    }
    Write-Host ''
    Write-Host 'Painel Trocar cores instalado.' -ForegroundColor Green
    Write-Host 'Abra o Corel: Janela > Janelas de encaixe > Trocar cores.'
    if ($backup) { Write-Host ('Versao anterior preservada em: ' + $backup) }
    exit 0
} catch {
    Write-Host ''
    Write-Host ('Nao foi possivel concluir: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
} finally {
    if ($gravandoRegistro) {
        Write-Host ''
        Write-Host ('Registro desta tentativa: ' + $caminhoRegistro)
        try { Stop-Transcript | Out-Null } catch { }
    }
}
