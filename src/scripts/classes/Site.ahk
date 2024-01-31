#Include %A_ScriptDir%\libraries\web\Rufaydium.ahk
#Include %A_ScriptDir%\classes\Logger.ahk

class Site {
  logger := ""

  __new() {
    this.logger := new Logger()
    this.logger.setLogFile("hkadm_error")
  }

  createSession(link) {
    try {
      chrome := new Rufaydium()
      page   := chrome.NewSession()

      page.Navigate(link)

      return { page: page, chrome: chrome }
    } catch e {
      message := % "Exception what: " e.what " file: " e.file . " line:" e.line " message: " e.message " extra:" e.extra
      this.logger.addError(message)
      MsgBox, 16,, Falha ao tentar criar uma sessao.
    }
  }

  getSessionInstance(tab := 1) {
    try {
      chrome  := new Rufaydium()
      session := chrome.getSession(tab)

      return { page: session, chrome: chrome }
    } catch e {
      message := % "Exception what: " e.what " file: " e.file . " line:" e.line " message: " e.message " extra:" e.extra
      this.logger.addError(message)
      MsgBox, 16,, Falha ao tentar utilizar uma sessao existente, certifique-se de que realmente existe uma sessao criada.
    }
  }
}