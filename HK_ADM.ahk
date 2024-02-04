#SingleInstance, Force

#Include, %A_ScriptDir%\env.ahk

; Gui +AlwaysOnTop
Gui, Font, s12,  Verdana
Gui, Add, Tab2,, Informacoes|Opcoes|Relatorios|Scripts|Usuario

FileReadLine, varUser,       %A_ScriptDir%\env.ahk, 2
FileReadLine, varPassword,   %A_ScriptDir%\env.ahk, 3
FileReadLine, varTerminal,   %A_ScriptDir%\env.ahk, 4
FileReadLine, varSeller,     %A_ScriptDir%\env.ahk, 5
FileReadLine, varStore,      %A_ScriptDir%\env.ahk, 6
FileReadLine, varNewSession, %A_ScriptDir%\env.ahk, 7

FileReadLine, varGetSpeed, %A_ScriptDir%\src\scripts\configurations\settings.ahk, 2

RegExMatch(varUser, "(\d+)",          userFound)
RegExMatch(varPassword, "(\d+)",      passwordFound)
RegExMatch(varTerminal, "(NF\d+)",    terminalFounnd)
RegExMatch(varSeller, "(\d+)",        sellerFound)
RegExMatch(varStore, "(\d+)",         storeFound)
RegExMatch(varNewSession, "(\D{5}$)", newSessionFound)

RegExMatch(varGetSpeed, "(\d+)", speedFound)

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
;=============================         Opcoes        ===========================
;===============================================================================

Gui, Tab, 2

Gui, Add, Text,,         Planilha principal
Gui, Add, Text, y+1,     -----------------------------------------------------
Gui, Add, Button, -Default x32 y+8, Abrir planilha

Gui, Add, Text,,         Velocidade do script
Gui, Add, Text, y+1,     -----------------------------------------------------
Gui, Add, Edit
Gui, Add, UpDown, vSpeed Range1-100, %speedFound%
Gui, Add, Button, -Default x32 y+8, Atualizar velocidade

;===============================================================================
;=============================       Relatorios      ===========================
;===============================================================================

Gui, Tab, 3

Gui, Add, Text,,   Selecione um arquivo de log a ser aberto
Gui, Add, ListBox, vLog gLog w372 r6

Loop, %A_ScriptDir%\src\scripts\logs\*.* {
  GuiControl,, Log, %A_LoopFileFullPath%
}

Gui, Add, Button, -Default x32 y+8, Abrir log
Gui, Add, Button, -Default x+10,    Atualizar
Gui, Add, Button, -Default x+10,    Limpar logs

;===============================================================================
;=============================        Scripts       ============================
;===============================================================================

Gui, Tab, 4

Gui, Add, Text,,   Selecione um script a ser executado
Gui, Add, ListBox, vScript gScript w372 r6

Gui, Add, CheckBox, vLogin,         Realizar login nerus
Gui, Add, CheckBox, vSessionOption, Criar uma sessao

GuiControl,, Login, % LOGIN ? "1" : "0"
GuiControl,, SessionOption, % NEW_SESSION ? "1" : "0"

Gui, Add, Button, -Default x32 y+8, Executar
Gui, Add, Button, -Default x+10, Aplicar alteracoes

Loop, %A_ScriptDir%\src\scripts\*.* {
  GuiControl,, Script, %A_LoopFileFullPath%
}

;===============================================================================
;=============================        Usuario       ============================
;===============================================================================

Gui, Tab, 5

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
Gui, Add, Text, x+5 cRed, v0.1.5-alpha

Gui, Show

return

;===============================================================================
ButtonAbrirplanilha:
Run, %A_ScriptDir%\src\scripts\spreadsheets\center_panel.xlsb
return
;===============================================================================

Log:
if (A_GuiEvent != "DoubleClick") {
  return
}

ButtonAbrirlog:
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

ButtonLimparlogs:
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

ButtonAplicaralteracoes:
Gui, Submit

sessionOption := SessionOption == 1 ? "true" : "false"
newLogin      := Login == 1 ? "true" : "false"
scriptdir     := "%" . "A_ScriptDir" . "%"

changeUserData = 
(
LINK_WEBSITE_NERUS = http://leitura.nerus.com.br
USER = %userFound%
PASSWORD = %passwordFound%
TERMINAL = %terminalFounnd%
SELLER = %sellerFound%
STORE = %storeFound%
NEW_SESSION := %sessionOption%
SPREADSHEET_PATH = %scriptdir%\spreadsheets\center_panel.xlsb
LOGIN := %newLogin%
)

FileDelete %A_ScriptDir%\env.ahk
FileAppend, %changeUserData%, %A_ScriptDir%\env.ahk

Run, %A_ScriptDir%\HK_ADM.ahk
return

;===============================================================================

ButtonAtualizarvelocidade:
Gui, Submit

changeSettingsData = 
(
#SingleInstance, Force
SetKeyDelay, %Speed%
)

FileDelete %A_ScriptDir%\src\scripts\configurations\settings.ahk
FileAppend, %changeSettingsData%, %A_ScriptDir%\src\scripts\configurations\settings.ahk

Run, %A_ScriptDir%\HK_ADM.ahk
return

;===============================================================================

ButtonSalvar:
Gui, Submit

StringUpper, outTerminal, Terminal

newLogin  := Login == 1 ? "true" : "false"
scriptdir := "%" . "A_ScriptDir" . "%"

changeUserData = 
(
LINK_WEBSITE_NERUS = http://leitura.nerus.com.br
USER = %User%
PASSWORD = %Password%
TERMINAL = %outTerminal%
SELLER = %Seller%
STORE = %Store%
NEW_SESSION := %newSessionFound%
SPREADSHEET_PATH = %scriptdir%\spreadsheets\center_panel.xlsb
LOGIN := %newLogin%
)

FileDelete %A_ScriptDir%\env.ahk
FileAppend, %changeUserData%, %A_ScriptDir%\env.ahk

MsgBox, Dados do usuario atualizados com sucesso.

Run, %A_ScriptDir%\HK_ADM.ahk
return

GuiClose:
GuiEscape:
ExitApp