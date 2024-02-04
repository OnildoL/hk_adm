#Include %A_ScriptDir%\classes\Logger.ahk

class FieldsNerus {
  page   := ""
  nf     := ""
  logger := ""

  __new(page, nf := "null") {
    this.page   := page
    this.nf     := nf
    this.logger := new Logger()
    this.logger.setLogFile("fields_nerus")
  }

  getsTheValueOfAnIntegerFromaField(xpath) {
    repeatGetsTheValueOfAnIntegerFromaField:
    result := this.page.getElementsByXpath(xpath)[0].value

    if result is integer 
    {
      return result
    } else {
      goto, repeatGetsTheValueOfAnIntegerFromaField
    }
  }

  getValueFromField(xpath) {
    repeatGetValueFromField:
    result := this.page.getElementsByXpath(xpath)[0].value

    if (result != "") {
      return result
    } else {
      goto, repeatGetValueFromField
    }
  }

  getInnerTextFromField(xpath) {
    repeatGetInnerTextFromField:
    result := this.page.getElementsByXpath(xpath)[0].innerText

    if (result != "") {
      return result
    } else {
      goto, repeatGetInnerTextFromField
    }
  }

  insertMonetaryValue(xpath_label, xpath_input, value) {
    focused := this.detectFocusedField(xpath_label)

    if (focused) {
      this.logger.addInfo("NF: " . this.nf . " Field: " . xpath_input . " Value: " . value)

      valueWithoutPoint := StrReplace(value, ".", "")

      this.insertValueIntoField(xpath_input, valueWithoutPoint)
    }
  }

  compareNerusTextMessageAndPressKey(xpath, value, key) {
    repeatCompareNerusTextMessageAndPressKey:
    compare := this.page.getElementsbyXpath(xpath)[0].innerText

    if (compare == value) {
      send, { %key% }
    } else {
      if InStr(compare, value) {
        send, { %key% }
        return
      }
      goto, repeatCompareNerusTextMessageAndPressKey
    }
  }

  checkIfWindowHasBeenClosed(xpath) {
    repeatCheckIfWindowHasBeenClosed:
    compare := this.page.getElementsbyXpath(xpath)[0].innerHTML

    if (compare == "") {
      return true
    } else {
      goto, repeatCheckIfWindowHasBeenClosed
    }
  }

  detectFieldAndPressKey(xpath, key_value) {
    this.logger.addInfo("NF: " . this.nf . " Field: " . xpath . " Key: " . key_value)

    focused := this.detectFocusedField(xpath)

    if (focused) {
      sleep, 400
      send, { %key_value% }
    }
  }

  detectFocusedField(xpath) {
    repeatDetectFocusedField:
    class        := this.page.GetElementsByXpath(xpath)[0].class
    desiredClass := ["Mui-focused", "active", "Mui-selected", "Mui-focusVisible", "Mui-checked"]

    for _, currentClass in desiredClass {
      if InStr(class, currentClass) {
        return true
      }
    }

    goto, repeatDetectFocusedField
  }

  insertValueIntoField(xpath, value) {
    repeatInsertValueIntoField:
    send, {SHIFTDOWN}{END}{SHIFTUP}{DEL}
    send, % value
    
    fieldValue := this.page.getElementsbyXpath(xpath)[0].value

    if (fieldValue == value) {
      send, {ENTER}
    } else {
      send, {HOME}
      goto, repeatInsertValueIntoField
    }
  }

  toFillIn(xpath_label, xpath_input, value, show_log := true) {
    if (show_log) {
      this.logger.addInfo("NF: " . this.nf . " Field: " . xpath_label . " Value: " . value)
    }

    focused := this.detectFocusedField(xpath_label)

    if (focused) {
      this.insertValueIntoField(xpath_input, value)
    }
  }
}