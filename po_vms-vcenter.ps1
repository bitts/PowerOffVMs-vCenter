##########################################################################################################
#SCRIPT PARA SHUTDOWN DE VIRTUAL MACHINES NO VCENTER
#Create by 2º Ten Marcelo Valvassori BITTENCOURT - 1º CTA - SGO
#USO:
#CRIAR UM ARQUIVO TXT NO FORMATO ABAIXO COM O NOME DE VMLIST.TXT
#
#Maquina
#SGO_MV_WIN10_192
#CCOP-WIN14
#
# defina as variáveis abaixo
# 
#
# 1.0v [05/08/2021] - Desligando maquinas virtuais contidas em arquivo csv
#.
#########################################################################################################

#ajustar para funcionamento
$vmUser = "1cta-bittencourt" 
$vmPswd = "" 
$vmCenter = "vsphere.qgcms.local" 
$vmList = "C:\Users\Administrator\Documents\vmlist.txt"

#não modificar
$data = Get-Date -Format "dd-MM-yyyy"

#não implementado -> <command>  | Out-File $logs
$logs = "C:\Users\Administrator\Documents\desligamentoVMs-$($data).txt"

Function sendMail{
	Param
    (
	 [Parameter(Mandatory=$true, Position=0)]
         [string] $Subject,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $Body
    )
	$From = "clonesvmware@1cta.eb.mil.br"
	$To = "sgo@1cta.eb.mil.br"
	#$Subject = "Script de Desligamento de VMS em execução"
	#$Body = "O script de desligamento do datacenter foi executado"
	$SMTPServer = "bombur.1cta.eb.mil.br"
	Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -SmtpServer $SMTPServer
}

Function Load-PowerCLI{

	if (Get-Command "*find-module"){
		$PCLIver =  (Find-Module "*vmware.powercli").Version.major
	}
	
 	($PCLIver| Out-String).Split(".")[0]

 	if ($PCLIver -ge 10){
  		$PCLI = "vmware.powercli"
		try {
			Import-Module -Name $PCLI
		} Catch {
			Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
			Exit 1
		}
    	} elseIf ($PCLIver -ge "6") {
  		$PCLI = "VMware.VimAutomation.Core"
		if ((Get-Module -Name $PCLI -ErrorAction SilentlyContinue) -eq $null){
		    	try {
				Import-Module $PCLI
		    	} Catch {
				Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
				Exit 1
		    	}
        	}
    	} elseIf ($PCLIver -ge "5") {
  		$PCLI = "VMware.VimAutomation.Core"
		if ((Get-PSSnapin $PCLI -ErrorAction "SilentlyContinue") -eq $null) {
		    Try {
			Add-PSSnapin $PCLI
		    } Catch {
			Write-Host "There is a problem loading the Powershell module. It is not possible to continue."
			Exit 1
		    }
		}
    	} else {
		Write-Host "This version of PowerCLI seems to be unsupported. Please upgrade to the latest version of PowerCLI and try again."
    	}
}

function PowerOffListVM{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string] $VIUser,
         [Parameter(Mandatory=$true, Position=1)]
         [string] $VIPwd,
         [Parameter(Mandatory=$true, Position=2)]
         [string] $vCenter,
         [Parameter(Mandatory=$true, Position=3)]
         [string] $vmlistfile
    )

    Try { 
        Write-Host "Conectando ao vCenter utilizando o usuario [$VIUser]"
		Connect-VIServer $vCenter -User $VIUser -Password $VIPwd -Force | Out-Null
        Write-Host ""
	} Catch {
		# capture any failure and display it in the error section, then end the script with a return
		# code of 1 so that CU sees that it was not successful.
		Write-Error "Unable to connect to the vCenter server. Please correct and re-run the script." -ErrorAction Continue
		Write-Error $Error[1] -ErrorAction Continue
		Exit 1
	}

    	try{

	#$vmlistfile = "C:\Users\Administrator\Documents\vmlist.txt"

		Import-Csv $vmlistfile -UseCulture | %{
		Get-VM -Name $_.Maquina |
		Select Name,
			VMHost,
			@{N='Datacenter';E={Get-Datacenter -VM $_ | Select -ExpandProperty Name}},
			@{N='Cluster';E={Get-Cluster -VM $_ | Select -ExpandProperty Name}},
			NumCpu, MemoryGB,ProvisionedSpaceGB,
			@{N='Path';E={
				$current = Get-View $_.ExtensionData.Parent
				$path = $_.Name
				do {
					$parent = $current
					if($parent.Name -ne "vm"){$path =  $parent.Name + "\" + $path}
					$current = Get-View $current.Parent
				} while ($current.Parent -ne $null)
				[string]::Join('\',($path.Split('\')[0..($path.Split('\').Count-2)]))           

			}},
			FolderId
		} | foreach {
		    #Desligando maquina
		    $spaceHD = [math]::Round($($_.ProvisionedSpaceGB),2)
		    Write-Host "Propriedades da maquina: `n - DataCenter: $($_.Datacenter) `n - CLUSTER: [$($_.Cluster)] `n - nCPU: [$($_.NumCpu)] `n - Memory: [$($_.MemoryGB)] `n - ProvisionedSpace: [$($spaceHD)GB] `n - Path: [$($_.Path)]"
		    PowerOFFVM -machineName $($_.Name)
		}

		Disconnect-VIServer $vCenter -Confirm:$false

		Write-Host "Ação de desligamento executada."

	}Catch {
		Write-Error "Erro no processo de desligamento da maquina." -ErrorAction Continue
		Write-Error $Error[1] -ErrorAction Continue
		Exit 1
	}
}

Function PowerOFFVM{

	Param
	(
		[Parameter(Mandatory=$true, Position=0)]
		[string] $machineName
	)


	Try {
		$machine = Get-VM -Name $machineName
	} Catch {
		# capture any failure and display it in the error section, then end the script with a return
		# code of 1 so that CU sees that it was not successful.
		Write-Error "Há uma VM em execução que está executando o VMTools, mas o vCenter não tem seu hostname. Investigue este problema e corrija." -ErrorAction Continue
		Write-Error $Error[1] -ErrorAction Continue
		Exit 1
	}

	If ($machine -ne $null) {
        		
		Write-Host "Desligando a Maquina Virtual: $machineName"

		#Generate a view for each vm to determine power state  
		$vm = Get-View -ViewType VirtualMachine -Filter @{"Name" = $machineName}  

		#If vm is powered on then VMware Tools status is checked  
		if ($vm.Runtime.PowerState -ne "PoweredOff") {
            		Write-Host "Processando VM: ++ $machineName ++ ..." 
		   	if ($vm.config.Tools.ToolsVersion -ne 0) { 
				Write-Warning "Maquina sem VMTools instalado."

				#Commando de Shutdown...                 
			   	Shutdown-VMGuest $machineName -Confirm:$false

				Write-Host "Executando shutdown da VM [$machineName] via VMware Tools. Aguarde..."
			} else {  
				#Commando de Shutdown...          
				Stop-VM $machineName -Confirm:$false 

				Write-Host "Executando shutdown da VM [$machineName] via Force Stop. Aguarde..."      
			}
           		sleep 5
		} else{
            		Write-Warning "Maquina já encontra-se desligada"
        	}
	} else {
		# capture any failure and display it in the error section, then end the script with a return
		# code of 1 so that CU sees that it was not successful.
		Write-Error "O computador que você deseja reiniciar não parece ser gerenciado por este vCenter. Por favor verifique e tente novamente." -ErrorAction Continue
		Write-Error $Error[1] -ErrorAction Continue
		Exit 1
	}
    	Write-Host "----------------------------------------------`n"
}

Function ShutDownRestantes{
	$waittime = 60
	
	# For each of the powered on VMs with running VM Tools
	Foreach ($VM in (Get-VMHost | Get-VM | Where {$_.PowerState -eq "poweredOn" -and $_.Guest.State -eq "Running"})){
		# Shutdown Guest
		write-host "Shutting down da VM [$VM] sem VMTools instalado..."
		$VM | Shutdown-VMGuest -Confirm:$false
	}

	Write-host "Pausing for $waittime seconds to allow VMs to shutdown cleanly"
	sleep $waittime
	
	#Force Poweroff of any VMs still running
	Foreach ($VM in (Get-VMHost | Get-VM | Where {$_.PowerState -eq "poweredOn"})){
		  # Power off Guest 
		  write-host "Powering off [$VM] via Force Stop"
		  $VM | Stop-VM -Confirm:$false
	}
	
	$Time = (Get-Date).TimeofDay
	do {
		# Wait for the VMs to poweroff
		sleep 1.0
		$timeleft = $waittime - ($Newtime.seconds)
		$numvms = (Get-VMHost | Get-VM | Where { $_.PowerState -eq "poweredOn" }).Count
		Write-host "Waiting for shutdown of $numvms VMs or until $timeleft seconds"
		$Newtime = (Get-Date).TimeofDay - $Time
	} until ((@(Get-VMHost | Get-VM | Where { $_.PowerState -eq "poweredOn" }).Count) -eq 0 -or ($Newtime).Seconds -ge $waittime)

}


Function ShuttingHosts{
	# Shutdown the ESX Hosts
	Write-host "Shutting down hosts"
	Get-VMHost | Foreach {Get-View $_.ID} | Foreach {$_.ShutdownHost_Task($TRUE)}
}

Clear
sendMail -Subject "Script de Desligamento de VMS em execução" -Body "[$($data)] O script de desligamento do datacenter foi executado"
Load-PowerCLI
PowerOffListVM -VIUser $vmUser -VIPwd $vmPswd -vCenter $vmCenter -vmlistfile $vmList
sendMail -Subject "Processo de desligando dos Hosts em execução" -Body "[$($data)] O processo de desligamento do DC esta em execução e todas as MV da lista de prioridade tiveram seu processo de desligamento iniciado. Iniciando processo de desligamento das MV e hosts restantes..."
ShutDownRestantes
ShuttingHosts

