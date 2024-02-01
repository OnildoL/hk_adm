#Include, %A_ScriptDir%\configurations\settings.ahk
#Include, %A_ScriptDir%\xpaths\xpath_price_products.ahk
#Include, %A_ScriptDir%\xpaths\xpath_nerus.ahk
#Include, %A_ScriptDir%\..\..\env.ahk
#Include, %A_ScriptDir%\classes\Site.ahk
#Include, %A_ScriptDir%\classes\Excel.ahk
#Include, %A_ScriptDir%\classes\FieldsNerus.ahk
#Include, %A_ScriptDir%\classes\Logger.ahk

Gui, +AlwaysOnTop -Caption

SysGet, ScreenW, 0
SysGet, ScreenH, 1

Gui, Add, Progress, vProgressBar w416 cED1C29
Gui, Add, Text, vProgressText wp

GuiX := (ScreenW - 450) / 2
GuiY := ScreenH - 100

Gui, Show, x%GuiX% y%GuiY% w450

F1::

logger := new Logger()

try {
  site  := new Site()
  excel := new Excel()

  excel.toConnect(SPREADSHEET_PATH, "input")

  if (NEW_SESSION) {
    web    := site.createSession(LINK_WEBSITE_NERUS)
    page   := web.page
    chrome := web.chrome
  } else {
    web    := site.getSessionInstance()
    page   := web.page
    chrome := web.chrome
  }

  nf          := excel.capture("A2", "input", "number")

  fieldsNerus := new FieldsNerus(page, nf)

  if (LOGIN) {
    fieldsNerus.toFillIn(XPATH_LOGIN_LABEL,    XPATH_LOGIN_INPUT,    USER)
    fieldsNerus.toFillIn(XPATH_PASSWORD_LABEL, XPATH_PASSWORD_INPUT, PASSWORD, false)
  }

  logger.setLogFile("hkadm_error")
  
  total_item_count := excel.capture("A1", "input", "number")

  remaining_items := total_item_count

  GuiControl,, ProgressBar, 0
  GuiControl,, ProgressText, Restam: %remaining_items% itens

  loop % total_item_count 
  {
    if nf is integer 
    {
      repeatPrice:
      code  := excel.capture("B4", "input")
      price := excel.capture("E4", "input")

      code_windows := fieldsNerus.getValueFromField(XPATH_CODIGO_ATUAL_JANELA_PRECIFICACAO)

      if (code != code_windows) {
        MsgBox, 4,, Codigo da planilha diferente do codigo atual da janela de precificacao. Deseja continuar mesmo assim?
        IfMsgBox, No
          goto, repeatPrice
      }

      fieldsNerus.toFillIn(XPATH_PRECO_REFERENCIA_ATUAL_LABEL, XPATH_PRECO_REFERENCIA_ATUAL_INPUT, price)

      fieldsNerus.compareNerusTextMessageAndPressKey(XPATH_CONFIRMAR, "Confirma", "ENTER")

      sleep, 300
      excel.toClean("4:4", "input")

      remaining_items--

      GuiControl,, ProgressBar, % (total_item_count - remaining_items) / total_item_count * 100
      GuiControl,, ProgressText, Restam: %remaining_items% itens
    } else {
      msgbox, 16,, O numero da nota fiscal deve ser preenchido na celula "A2".
      logger.addError("O numero da nota fiscal deve ser preenchido na celula 'A2'.")
      goto, exitScript
    }
  }

  msgbox, 64,, Processo concluido!
  goto, exitScript
} catch e {
  message := % "Exception what: " e.what " file: " e.file . " line:" e.line " message: " e.message " extra:" e.extra
  logger.addError(message)
  msgbox, 16,, Falha ao tentar executar o script.
  goto, exitScript
}

F10::
pause
return

F11::
reload
return

F12::
exitScript:
excel.endConnection("input", false)
; chrome.QuitAllSessions()
; chrome.Driver.Exit()
exitapp