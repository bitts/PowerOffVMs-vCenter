<#
#########################################################################################################
SCRIPT PARA POWERON DE VIRTUAL MACHINES NO VCENTER
Create by Bitts
.VERSIONS
	1.0v [05/08/2021] - Desligando maquinas virtuais contidas em arquivo csv
 	1.1v [20/08/2021] - Adicionado autenticação por relação de confiança.
 	1.2v [15/01/2025] - Atualização do script

.SYNOPSIS
    Liga máquinas virtuais em um ambiente VMware vCenter
    respeitando a ordem inversa do desligamento.

.DESCRIPTION
    O script conecta-se ao vCenter, obtém a lista de VMs,
    inverte a ordem e liga as VMs aguardando VMware Tools
    antes de prosseguir para a próxima. 

 .REQUIREMENTS
    - PowerShell 5.1+
    - VMware PowerCLI
    - Acesso ao vCenter
.
#########################################################################################################
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$vCenterServer,

    [Parameter(Mandatory = $true)]
    [PSCredential]$Credential,

    [int]$ToolsTimeoutSeconds = 180
)

# Importar configurações
$configPath = Join-Path $PSScriptRoot "config.psd1"
$config = Import-PowerShellDataFile -Path $configPath

Write-Host "🔌 Conectando ao vCenter $vCenterServer..." -ForegroundColor Cyan
Connect-VIServer -Server $vCenterServer -Credential $Credential -ErrorAction Stop

# Buscar VMs desligadas, excluindo VMs críticas
$vmList = Get-VM | Where-Object {
    $_.PowerState -eq 'PoweredOff' -and
    ($config.ExcludedVMs -notcontains $_.Name)
}

# Inverter a ordem
$vmList = $vmList | Sort-Object Name -Descending

Write-Host "🖥️ VMs encontradas para startup: $($vmList.Count)" -ForegroundColor Yellow

foreach ($vm in $vmList) {

    Write-Host "➡️ Ligando VM: $($vm.Name)" -ForegroundColor Cyan
    Start-VM -VM $vm -Confirm:$false | Out-Null

    $elapsed = 0
    do {
        Start-Sleep -Seconds 5
        $elapsed += 5
        $vm = Get-VM -Name $vm.Name
        $toolsStatus = $vm.ExtensionData.Guest.ToolsStatus
    }
    while (
        $toolsStatus -ne "toolsOk" -and
        $elapsed -lt $ToolsTimeoutSeconds
    )

    if ($toolsStatus -eq "toolsOk") {
        Write-Host "   ✔ VMware Tools OK" -ForegroundColor Green
    } else {
        Write-Host "   ⚠ VMware Tools não responderam no tempo esperado" -ForegroundColor Yellow
    }

    Start-Sleep -Seconds $config.DelayBetweenVMsSeconds
}

Disconnect-VIServer -Server $vCenterServer -Confirm:$false
Write-Host "✅ Startup das VMs finalizado." -ForegroundColor Green
