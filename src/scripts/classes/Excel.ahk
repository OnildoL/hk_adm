class Excel {
  static instance := ComObjCreate("Excel.Application")

  __new() {

  }

  createNewTab(tabName) {
    create := Excel.instance
    newTab := create.Worksheets.Add()
    newTab.Name := tabName
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
    Excel.instance.visible := visibility ? true : false
    Excel.instance.Workbooks.Open(path)
    Excel.instance.Sheets(tabName).Activate()
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
