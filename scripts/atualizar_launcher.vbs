' Lancador silencioso do atualizador do Painel Comercial Rodar Mutual.
' Colocado na pasta Inicializar do Windows -> roda a cada logon (sem janela).
' O atualizar.ps1 decide se hoje eh seg/qua e se ja rodou hoje.
Set sh = CreateObject("WScript.Shell")
ps1 = "C:\Users\eduar\Documents\CLAUDE\Dashboard riskin\Dashboard-Comercial-Rodar\scripts\atualizar.ps1"
sh.Run "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & ps1 & """", 0, False
