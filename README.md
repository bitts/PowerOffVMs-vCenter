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
