##########################################################################################################
#SCRIPT PARA SHUTDOWN DE VIRTUAL MACHINES NO VCENTER
#
#USO:
#CRIAR UM ARQUIVO TXT NO FORMATO ABAIXO COM O NOME DE VMLIST.TXT
#2º
#Maquina
#SGO_MV_WIN10_192
#CCOP-WIN14
#
# defina as variáveis abaixo
# 
#
# 1.0v [05/08/2021] - Desligando maquinas virtuais contidas em arquivo csv
#
#########################################################################################################


##########################################################################################################
#SCRIPT PARA SHUTDOWN DE VIRTUAL MACHINES NO VCENTER
#
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
#VERSÃO 1.0
#FALTA MELHORAR A PARTE DE LOG....
#########################################################################################################


$vmUser = "1cta-bittencourt" 
$vmPswd = "" 
$vmCenter = "vsphere.qgcms.local" 
$vmList = "C:\Users\Administrator\Documents\vmlist.txt"

Function Load-PowerCLI (){

    if (Get-Command "*find-module"){
 		$PCLIver =  (Find-Module "*vmware.powercli").Version.major
    }
	
 	($PCLIver| Out-String).Split(".")[0]

 	if ($PCLIver -ge 10){
  		$PCLI = "vmware.powercli"
        try {
    		Import-Module -Name $PCLI
        } Catch {
		  	Write-Host "Ocorreu um problema ao carregar o módulo Powershell. Não é possível continuar."
	  		Exit 1
		}
    } elseIf ($PCLIver -ge "6") {
  		$PCLI = "VMware.VimAutomation.Core"
        if ((Get-Module -Name $PCLI -ErrorAction SilentlyContinue) -eq $null){
            try {
		  		Import-Module $PCLI
            } Catch {
				Write-Host "Ocorreu um problema ao carregar o módulo Powershell. Não é possível continuar."
				Exit 1
            }
        }
    } elseIf ($PCLIver -ge "5") {
  		$PCLI = "VMware.VimAutomation.Core"
        if ((Get-PSSnapin $PCLI -ErrorAction "SilentlyContinue") -eq $null) {
            Try {
                Add-PSSnapin $PCLI
            } Catch {
                Write-Host "Ocorreu um problema ao carregar o módulo Powershell. Não é possível continuar."
                Exit 1
            }
        }
    } else {
        Write-Host "Esta versão do PowerCLI parece não ser compatível. Atualize para a versão mais recente do PowerCLI e tente novamente."
    }
}

function PowerOffListVM
{
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

		$vmlistfile = "C:\Users\Administrator\Documents\vmlist.txt"

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
            Write-Host "Propriedades da maquina: $($_.Datacenter), CLUSTER: [$($_.Cluster)], nCPU: [$($_.NumCpu)], Memory: [$($_.MemoryGB)], ProvisionedSpaceGB: [$($_.ProvisionedSpaceGB)], Path: [$($_.Path)]"
            PowerOFFVM -machineName $($_.Name)
		}

		Disconnect-VIServer $vCenter -Confirm:$false

		Write-Host "Ação de desligamento executada."

	}Catch {
		Write-Error "=> Erro no processo de desligamento da maquina." -ErrorAction Continue
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
		   if ($vm.config.Tools.ToolsVersion -ne 0) { 

			   Write-Host "Processando VM: ++ $machineName ++ ..."
			   Write-Warning "Maquina sem VMTools instalado."
			   
			   #Commando de Shutdown...                 
			   Shutdown-VMGuest $machineName -Confirm:$false

			   Write-Host "Executando shutdown da VM ++ $machineName ++ via VMware Tools. Aguardando... (30s)"
			   sleep 30

		   }  
		   else {  

				Write-Host "Processando VM: ++ $machineName ++ ..."

				#Commando de Shutdown...          
			    Stop-VM $machineName -Confirm:$false 

				Write-Host "Executando shutdown da VM: ++ $machineName ++ via Force Stop. Aguardando... (30s)"
				sleep 30        
		   }  
		}  

	} else {
		# capture any failure and display it in the error section, then end the script with a return
		# code of 1 so that CU sees that it was not successful.
		Write-Error "O computador que você deseja reiniciar não parece ser gerenciado por este vCenter. Por favor verifique e tente novamente." -ErrorAction Continue
		Write-Error $Error[1] -ErrorAction Continue
		Exit 1
	}
    Write-Host ""
	
}

Clear
Load-PowerCLI
PowerOffListVM -VIUser $vmUser -VIPwd $vmPswd -vCenter $vmCenter -vmlistfile $vmList
