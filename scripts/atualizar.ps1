# Wrapper da atualizacao do Painel Comercial Rodar Mutual.
# Roda apenas segunda-feira e quarta-feira. Chamado pela Tarefa Agendada "ao fazer logon".

$ErrorActionPreference = "Stop"
$dow = (Get-Date).DayOfWeek   # Monday / Wednesday / ...
$log = Join-Path $PSScriptRoot "atualizar.log"

if ($dow -ne "Monday" -and $dow -ne "Wednesday") {
    Add-Content $log ("[" + (Get-Date -Format "s") + "] pulado: hoje eh $dow (so roda seg/qua)")
    exit 0
}

# evita rodar duas vezes no mesmo dia (se logar varias vezes)
$hoje = (Get-Date).ToString("yyyy-MM-dd")
$marca = Join-Path $PSScriptRoot ".rodou_$hoje.txt"
if (Test-Path $marca) {
    Add-Content $log ("[" + (Get-Date -Format "s") + "] pulado: ja rodou hoje ($hoje)")
    exit 0
}

$node = "C:\Program Files\nodejs\node.exe"
$script = Join-Path $PSScriptRoot "atualizar.js"

Add-Content $log ("[" + (Get-Date -Format "s") + "] iniciando ($dow)...")
& $node $script
if ($LASTEXITCODE -eq 0) {
    Set-Content $marca "ok"
    # limpa marcas de dias anteriores
    Get-ChildItem $PSScriptRoot -Filter ".rodou_*.txt" | Where-Object { $_.Name -ne ".rodou_$hoje.txt" } | Remove-Item -Force -ErrorAction SilentlyContinue
}
