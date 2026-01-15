# 🛑 PowerOffVMs-vCenter

Script **PowerShell / VMware PowerCLI** para desligamento controlado de máquinas virtuais em ambientes **VMware vSphere / vCenter Server**.

Ideal para:
- automação de manutenção
- rotinas de desligamento de datacenter
- integração com UPS
- shutdown ordenado de ambientes VMware

---

## 🚀 Funcionalidades

✔️ Conexão segura ao vCenter  
✔️ Desligamento gracioso via VMware Tools  
✔️ Fallback para desligamento forçado  
✔️ Exclusão de VMs críticas (ex.: vCenter, AD, DNS)  
✔️ Delay configurável entre desligamentos  
✔️ Compatível com ambientes reais (prod/lab)

---

## 📦 Requisitos

- PowerShell 5.1 ou superior  
- VMware PowerCLI  
- Acesso administrativo ao vCenter  

Instalação do PowerCLI:

```powershell
Install-Module VMware.PowerCLI -Scope CurrentUser
```

## ▶️ Como executar
```powershell
$cred = Get-Credential

.\PowerOff-VMs.ps1 `
  -vCenterServer vcsa01.seudominio.local `
  -Credential $cred

```

## 🧠 Como funciona
1. Conecta ao vCenter
2. Lista VMs ligadas
3. Remove VMs críticas da lista
4. Tenta desligamento gracioso (VMware Tools)
5. Força desligamento se necessário
6. Aguarda intervalo configurado
7. Desconecta do vCenter


## ⚠️ Boas práticas em produção

- Nunca desligue o vCenter antes das demais VMs
- Garanta VMware Tools instalado nas VMs
- Teste sempre em ambiente de homologação
- Evite execução em clusters com DRS ativo sem planejamento


## 🧩 Possíveis melhorias futuras

- Seleção por tags ou pastas
- Geração de logs em arquivo
- Modo dry-run
- Integração com sistemas de monitoramento


## 🔎 Palavras-chave

PowerCLI, VMware, vSphere, vCenter, PowerShell, shutdown VM, automation, datacenter maintenance
