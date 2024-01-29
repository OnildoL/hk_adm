class Logger {
  log_file := "adms_log_"

  __new() {

  }

  setLogFile(file_name) {
    this.log_file := file_name . "_"
  }

  addLog(message) {
    file := this.log_file . A_DD . "-" . A_MM . "-" . A_YYYY
    FileAppend, %message%, %A_ScriptDir%\logs\%file%.txt
  }

  addError(message) {
    text := "[" . A_DD . "-" . A_MM . "-" . A_YYYY . "][" . A_Hour . ":" . A_Min . ":" . A_Sec . "][error]: " . message "`n"
    this.addLog(text)
  }

  addWarning(message) {
    text := "[" . A_DD . "-" . A_MM . "-" . A_YYYY . "][" . A_Hour . ":" . A_Min . ":" . A_Sec . "][warning]: " . message "`n"
    this.addLog(text)
  }

  addInfo(message) {
    text := "[" . A_DD . "-" . A_MM . "-" . A_YYYY . "][" . A_Hour . ":" . A_Min . ":" . A_Sec . "][info]: " . message "`n"
    this.addLog(text)
  }
}