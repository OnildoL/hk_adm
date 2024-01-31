#Include, %A_ScriptDir%\configurations\settings.ahk
#Include, %A_ScriptDir%\xpaths\xpath_taxing_products.ahk
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

  row := 1

  loop % row
  {
    xpath_current_row = //*[@id="root"]/div[2]/div[4]/div/div[2]/div/div[2]/div[1]/table/tbody/tr[%i%]
    current_row      := page.getElementsbyXpath(xpath_current_row)[0].innerText

    if (current_row == "") {
      row := row - 1
      break
    }

    row := row + 1
  }

  loop % total_item_count 
  {
    if nf is integer 
    {
      base_ipi  := excel.capture("F4", "input")
      base_icms := excel.capture("G4", "input")
      icms      := excel.capture("H4", "input")
      ipi       := excel.capture("I4", "input")

      checkIfValuesMatch:
      product_value := fieldsNerus.getValueFromField(XPATH_VALOR_PRODUTO_INPUT)

      if (product_value != base_icms) {
        msgbox, 16,, O valor bruto desse produto nao corresponde ao valor da linha atual na planilha.
        goto, checkIfValuesMatch
      }

      fieldsNerus.detectFieldAndPressKey(XPATH_VALOR_DESCONTO_LABEL,  "ENTER")
      fieldsNerus.detectFieldAndPressKey(XPATH_VALOR_FRETE_LABEL,     "ENTER")
      fieldsNerus.detectFieldAndPressKey(XPATH_OUTRAS_DESPESAS_LABEL, "ENTER")
      fieldsNerus.detectFieldAndPressKey(XPATH_VALOR_SEGURO_LABEL,    "ENTER")

      fieldsNerus.toFillIn(XPATH_BASE_CALCULO_ICMS_LABEL, XPATH_BASE_CALCULO_ICMS_INPUT, base_icms)
      fieldsNerus.toFillIn(XPATH_ALIQUOTA_ICMS_LABEL, XPATH_ALIQUOTA_ICMS_INPUT,         icms)

      fieldsNerus.detectFieldAndPressKey(XPATH_VALOR_ICMS_LABEL, "ENTER")
      fieldsNerus.compareNerusTextMessageAndPressKey(XPATH_MESSAGEM_VALOR_INFORMADO, "informado nao corresponde", "ENTER")
      fieldsNerus.detectFieldAndPressKey(XPATH_VALOR_ICMS_LABEL, "ENTER")

      fieldsNerus.toFillIn(XPATH_BASE_CALCULO_IPI_LABEL, XPATH_BASE_CALCULO_IPI_INPUT, base_ipi)

      if (ipi != "0.00") {
        fieldsNerus.toFillIn(XPATH_ALIQUOTA_IPI_LABEL, XPATH_ALIQUOTA_IPI_INPUT, ipi)
        fieldsNerus.detectFieldAndPressKey(XPATH_VALOR_IPI_LABEL, "ENTER")
        fieldsNerus.compareNerusTextMessageAndPressKey(XPATH_MESSAGEM_VALOR_INFORMADO, "informado nao corresponde", "ENTER")
      } else  {
        fieldsNerus.detectFieldAndPressKey(XPATH_ALIQUOTA_IPI_LABEL, "ENTER")
      }

      fieldsNerus.detectFieldAndPressKey(XPATH_VALOR_IPI_LABEL,       "ENTER")
      fieldsNerus.detectFieldAndPressKey(XPATH_BASE_CALCULO_ST_LABEL, "ENTER")
      fieldsNerus.detectFieldAndPressKey(XPATH_VALOR_ST_LABEL,        "ENTER")
      fieldsNerus.detectFieldAndPressKey(XPATH_EDITA_DIFAL_LABEL,     "ENTER")
      fieldsNerus.detectFieldAndPressKey(XPATH_CST_ICMS_LABEL,        "ENTER")
      fieldsNerus.detectFieldAndPressKey(XPATH_CFOP_LABEL,            "ENTER")
      
      windowsClosed := fieldsNerus.checkIfWindowHasBeenClosed(XPATH_TAB_FORM)

      if (windowsClosed) {
        sleep, 700
        send, {DOWN}  
        sleep, 700
        send, {ENTER} 
      }

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