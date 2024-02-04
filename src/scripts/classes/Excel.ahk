#Include %A_ScriptDir%\classes\Logger.ahk

class Excel {
  static instance := ComObjCreate("Excel.Application")
  logger          := ""

  __new() {
    this.logger := new Logger()
    this.logger.setLogFile("hkadm_error")
  }

  createNewTab(tabName) {
    create := Excel.instance
    newTab := create.Worksheets.Add()
    newTab.Name := tabName
  }

  copyTab(existingTabName, newTabName) {
    try {
      existingTab := this.getTab(existingTabName)
      
      newTab := Excel.instance.Worksheets.Add()
      newTab.Name := newTabName
      
      existingTab.UsedRange.Copy(newTab.Range("A1"))
    } catch e {
      message := "Exception: " e.what " file: " e.file . " line:" e.line " message: " e.message " extra:" e.extra
      this.logger.addError(message)
      MsgBox, 16,, Falha ao tentar copiar a aba.
    }
  }

  deleteTab(tabName) {
    tab := this.getTab(tabName)
    Excel.instance.DisplayAlerts := false
    tab.Delete()
    Excel.instance.DisplayAlerts := true
  }

  getTab(tabName) {
    tab := Excel.instance.Sheets(tabName)
    return tab
  }

  isWorksheetOpen(path) {
    RegexMatch(path, "[^\\]+\.xlsx", match)
    isWindowExists := WinExist(match . " - Excel")
    return isWindowExists
  }

  toConnect(path, tabName, visibility := true) {
    try {
      Excel.instance.visible := visibility ? true : false
      Excel.instance.Workbooks.Open(path)
      Excel.instance.Sheets(tabName).Activate()
    } catch e {
      message := % "Exception what: " e.what " file: " e.file . " line:" e.line " message: " e.message " extra:" e.extra
      this.logger.addError(message)
      MsgBox, 16,, Falha ao tentar abrir a planilha.
    }
  }

  toSave() {
    Excel.instance.ActiveWorkbook.save()
  }

  endConnection(tabName, save := true) {
    tab := this.getTab(tabName)
    tab.Parent.Close(SaveChanges := save)
    Excel.instance.Quit()
  }

  capture(cell, tabName, type := "string", decimalPlaces := 0) {
    tab := this.getTab(tabName)

    switch type{
      case "string" :
        return tab.Range(cell).Value
      case "number" :
        return Round(tab.Range(cell).Value, decimalPlaces)
      default:
        return
    }
  }

  toWrite(cell, tabName, value) {
    tab := this.getTab(tabName)
    tab.Range(cell).Value := value
  }

  toClean(cell, tabName) {
    tab := this.getTab(tabName)
    tab.Range(cell).Delete
  }
}
