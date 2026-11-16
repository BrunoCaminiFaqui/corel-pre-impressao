$ErrorActionPreference = 'Stop'
$painel = Join-Path $PSScriptRoot '..\painel'
function Transformar([string]$arquivo, [string]$xml) {
    $t = New-Object System.Xml.Xsl.XslCompiledTransform
    $t.Load($arquivo)
    $entrada = New-Object System.IO.StringReader($xml)
    $leitor = [System.Xml.XmlReader]::Create($entrada)
    $saida = New-Object System.IO.StringWriter
    try { $t.Transform($leitor, $null, $saida); return $saida.ToString() }
    finally { $leitor.Dispose(); $entrada.Dispose(); $saida.Dispose() }
}
$base = '<uiConfig><items/><dockers/><commandBars><commandBarData guid="3eaa9bbe-28fd-4672-9128-02974ee96332"><menu><item guidRef="existing-item"/></menu></commandBarData></commandBars><frame/></uiConfig>'
$app = Transformar (Join-Path $painel 'AppUI.xslt') $base
$user = Transformar (Join-Path $painel 'UserUI.xslt') $app
[xml]$resultado = Transformar (Join-Path $painel 'UserUI.xslt') $user
$hospedeiros = $resultado.SelectNodes('//itemData[@type="wpfhost"]')
if ($hospedeiros.Count -ne 1) { throw 'Host WPF ausente ou duplicado.' }
if ($hospedeiros[0].hostedType -ne 'Addons\SeparacoesCorel\SeparacoesCorel.dll,SeparacoesCorel.DockerUI') { throw 'DLL ou classe incorreta.' }
if ($resultado.SelectNodes('//dockerData').Count -ne 1) { throw 'Painel ausente ou duplicado.' }
if ($resultado.SelectSingleNode('//dockerData/container/item').guidRef -ne $hospedeiros[0].guid) { throw 'Vinculo de GUID incorreto.' }
if ($resultado.SelectNodes('//menu/item[@guidRef="c9846b2b-cf26-4411-a4a5-d8a8987d0e5d"]').Count -ne 1) { throw 'Menu ausente ou duplicado.' }
if ($resultado.SelectNodes('//menu/item[@guidRef="existing-item"]').Count -ne 1) { throw 'Comando existente removido.' }
Write-Output 'PASSOU: transformacoes XML, DLL/classe, GUIDs, preservacao de comandos e menu sem duplicacao.'
Write-Output 'Este teste nao substitui abrir o painel dentro do CorelDRAW.'
