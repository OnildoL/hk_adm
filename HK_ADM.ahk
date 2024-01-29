#SingleInstance, Force

Gui, Font, s12,  Verdana
Gui, Add, Tab2,, Informacoes|Relatorios|Scripts|Usuario

FileReadLine, varUser,       %A_ScriptDir%\env.ahk, 2
FileReadLine, varPassword,   %A_ScriptDir%\env.ahk, 3
FileReadLine, varTerminal,   %A_ScriptDir%\env.ahk, 4
FileReadLine, varSeller,     %A_ScriptDir%\env.ahk, 5
FileReadLine, varStore,      %A_ScriptDir%\env.ahk, 6
FileReadLine, varNewSession, %A_ScriptDir%\env.ahk, 7

RegExMatch(varUser, "(\d+)",          userFound)
RegExMatch(varPassword, "(\d+)",      passwordFound)
RegExMatch(varTerminal, "(NF\d+)",    terminalFounnd)
RegExMatch(varSeller, "(\d+)",        sellerFound)
RegExMatch(varStore, "(\d+)",         storeFound)
RegExMatch(varNewSession, "(\D{5}$)", newSessionFound)

Gui, Add, Text,,         Dados do usuario registrado
Gui, Add, Text, y+1,     -----------------------------------------------------

Gui, Add, Text, section,         Usuario:
Gui, Add, Text,,                 Loja:
Gui, Add, Text,,                 Terminal:
Gui, Add, Text,,                 Vendedor:
Gui, Add, Text, ys c0f6be9,      %userFound%
Gui, Add, Text, c0f6be9,         %storeFound%
Gui, Add, Text, c0f6be9,         %terminalFounnd%
Gui, Add, Text, c0f6be9,         %sellerFound%

;===============================================================================
;=============================       Relatorios      ===========================
;===============================================================================

Gui, Tab, 2

Gui, Add, Text,,   Selecione um arquivo de log a ser aberto
Gui, Add, ListBox, vLog gLog w372 r6

Loop, %A_ScriptDir%\src\scripts\logs\*.* {
  GuiControl,, Log, %A_LoopFileFullPath%
}

Gui, Add, Button, -Default x32 y+8, Abrir_log
Gui, Add, Button, -Default x+10,    Atualizar
Gui, Add, Button, -Default x+10,    Limpar_logs

;===============================================================================
;=============================        Scripts       ============================
;===============================================================================

Gui, Tab, 3

Gui, Add, Text,,   Selecione um script a ser executado
Gui, Add, ListBox, vScript gScript w372 r6

Gui, Add, Radio,  vSessionOption, Criar uma sessao do script
Gui, Add, Radio,, Utilizar uma sessao ja criada

Gui, Add, Button, -Default x32 y+8, Executar
Gui, Add, Button, -Default x+10, Aplicar_alteracoes

Loop, %A_ScriptDir%\src\scripts\*.* {
  GuiControl,, Script, %A_LoopFileFullPath%
}

;===============================================================================
;=============================        Usuario       ============================
;===============================================================================

Gui, Tab, 4

Gui, Add, Text,,         Atualizar dados do usuario
Gui, Add, Text, y+1,     -----------------------------------------------------

gui, add, Text,  section, Usuario
gui, add, Text,,          Senha
gui, add, Text,,          Terminal
gui, add, Text,,          Vendedor
gui, add, Text,,          Loja

Gui, Add, Edit, ys vUser  w278,          %userFound%
Gui, Add, Edit, vPassword w278 Password, %passwordFound%
Gui, Add, Edit, vTerminal w278,          %terminalFounnd%
Gui, Add, Edit, vSeller   w278,          %sellerFound%
Gui, Add, Edit, vStore    w278,          %storeFound%

Gui, Add, Button, -Default x32 y+10 w370, Salvar

Gui, Tab

Gui, Add, Text, xm, HK_ADM Created by Onildo.
Gui, Add, Text, x+5 cRed, v0.1.2-alpha

Gui, Show

return

;===============================================================================

Log:
if (A_GuiEvent != "DoubleClick") {
  return
}

ButtonAbrir_log:
GuiControlGet, Log

Run, %Log%,, UseErrorLevel

if (ErrorLevel = "ERROR") {
  MsgBox, Nao foi possivel iniciar o arquivo especificado.
}
return

ButtonAtualizar:
Gui, Destroy
Run, %A_ScriptDir%\HK_ADM.ahk
return

ButtonLimpar_logs:
MsgBox, 4,, Deseja realmente deletar todos os arquivos de logs?
IfMsgBox, No
return
Loop, %A_ScriptDir%\src\scripts\logs\*.* {
  FileDelete, %A_LoopFileFullPath%
}
Gui, Destroy
Run, %A_ScriptDir%\HK_ADM.ahk
return
;===============================================================================

Script:
if (A_GuiEvent != "DoubleClick") {
  return
}

ButtonExecutar:
GuiControlGet, Script

MsgBox, 4,, Deseja realmente executar o script abaixo?`n`n%Script%
IfMsgBox, No
return

Run, %Script%,, UseErrorLevel
if (ErrorLevel = "ERROR") {
  MsgBox, Nao foi possivel iniciar o arquivo especificado.
}
return

ButtonAplicar_alteracoes:
Gui, Submit

sessionOption := SessionOption == 1 ? "true" : "false"

changeUserData = 
(
LINK_WEBSITE_NERUS = http://leitura.nerus.com.br
USER = %userFound%
PASSWORD = %passwordFound%
TERMINAL = %terminalFounnd%
SELLER = %sellerFound%
STORE = %storeFound%
NEW_SESSION := %sessionOption%
)

FileDelete %A_ScriptDir%\env.ahk
FileAppend, %changeUserData%, %A_ScriptDir%\env.ahk

Run, %A_ScriptDir%\HK_ADM.ahk
return

;===============================================================================

ButtonSalvar:
Gui, Submit

StringUpper, outTerminal, Terminal

changeUserData = 
(
LINK_WEBSITE_NERUS = http://leitura.nerus.com.br
USER = %User%
PASSWORD = %Password%
TERMINAL = %outTerminal%
SELLER = %Seller%
STORE = %Store%
NEW_SESSION := %newSessionFound%
)

FileDelete %A_ScriptDir%\env.ahk
FileAppend, %changeUserData%, %A_ScriptDir%\env.ahk

MsgBox, Dados do usuario atualizados com sucesso.

Run, %A_ScriptDir%\HK_ADM.ahk
return

GuiClose:
GuiEscape:
ExitApp