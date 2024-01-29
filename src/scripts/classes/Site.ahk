#Include %A_ScriptDir%\libraries\web\Rufaydium.ahk

class Site {
  __new() {
  }

  createSession(link) {
    chrome := new Rufaydium()
    page   := chrome.NewSession()

    page.Navigate(link)

    return page
  }

  getSessionInstance(tab := 1) {
    chrome  := new Rufaydium()
    session := chrome.getSession(tab)

    return session
  }
}