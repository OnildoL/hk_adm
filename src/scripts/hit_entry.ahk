#SingleInstance, Force

#Include, %A_ScriptDir%\classes\Logger.ahk

logger := new Logger()
logger.addInfo("Testando logger")
