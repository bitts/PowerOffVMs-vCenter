<#
#########################################################################################################
SCRIPT PARA SHUTDOWN DE VIRTUAL MACHINES NO VCENTER
Create by Bitts
.VERSIONS
	 1.0v [05/08/2021] - Desligando maquinas virtuais contidas em arquivo csv
 	1.1v [20/08/2021] - Adicionado autenticação por relação de confiança.
 	1.2v [15/01/2025] - Atualização do script

.SYNOPSIS
    Desliga máquinas virtuais em um ambiente VMware vCenter usando PowerCLI.

.DESCRIPTION
    O script conecta-se a um vCenter Server, identifica VMs ligadas,
    tenta desligamento gracioso via VMware Tools e,
    se necessário, força o desligamento.
    Permite excluir VMs críticas e adicionar delays entre operações.# 

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

    [int]$GracefulTimeoutSeconds = 120
)

# Importar configurações
$configPath = Join-Path $PSScriptRoot "config.psd1"
$config = Import-PowerShellDataFile -Path $configPath

Write-Host "🔌 Conectando ao vCenter $vCenterServer..." -ForegroundColor Cyan
Connect-VIServer -Server $vCenterServer -Credential $Credential -ErrorAction Stop

# Buscar VMs ligadas, excluindo VMs críticas
$vmList = Get-VM | Where-Object {
    $_.PowerState -eq 'PoweredOn' -and
    ($config.ExcludedVMs -notcontains $_.Name)
}

Write-Host "🖥️ VMs encontradas para desligamento: $($vmList.Count)" -ForegroundColor Yellow

foreach ($vm in $vmList) {

    Write-Host "➡️ Processando VM: $($vm.Name)" -ForegroundColor Cyan

    if ($vm.ExtensionData.Guest.ToolsStatus -eq "toolsOk") {

        Write-Host "   ✔ VMware Tools OK - desligamento gracioso" -ForegroundColor Green
        Shutdown-VMGuest -VM $vm -Confirm:$false

        $elapsed = 0
        while ($vm.PowerState -ne "PoweredOff" -and $elapsed -lt $GracefulTimeoutSeconds) {
            Start-Sleep -Seconds 5
            $elapsed += 5
            $vm = Get-VM -Name $vm.Name
        }

        if ($vm.PowerState -ne "PoweredOff") {
            Write-Host "   ⚠ Timeout atingido - forçando desligamento" -ForegroundColor Red
            Stop-VM -VM $vm -Confirm:$false
        }

    } else {
        Write-Host "   ⚠ VMware Tools ausente - desligamento forçado" -ForegroundColor Red
        Stop-VM -VM $vm -Confirm:$false
    }

    Start-Sleep -Seconds $config.DelayBetweenVMsSeconds
}

Disconnect-VIServer -Server $vCenterServer -Confirm:$false
Write-Host "✅ Processo finalizado." -ForegroundColor Green
