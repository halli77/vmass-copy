#Persistent
#SingleInstance, Force

global scriptname := "VMASS-Copy"
global version := "1.19"

global theHeap := []   ; Working Stack
global theBackup := []   ; Backup Stack
global isAlwaysOnTop := 1   ; Fenster immer oben?

; ################################# GUI aufbauen

Gui, +AlwaysOnTop

; Gui, Add, ListBox, x10 y13 w95 h520 ReadOnly vUiHeap , 
Gui, Add, ListBox, x10 w300 h280 Multi vUiHeap , 


Gui, Add, Radio, x20 y320 vModusVmass, vmass (10 Elemente mit Tab)
Gui, Add, Radio, x20 y340 vModusStrgY Checked, Schnellstartkürzel ({1} ist ein Eintrag der Liste)


Gui, Add, Edit, x40 y360 w200 vSsk, {1} vm
; Gui, Add, Edit, xp y120 w162 vSskChIban, {1} gv sper {Enter}


Gui, Add, CheckBox, x40 y385 Checked vSendCtrlY, Suchfeld aktivieren (Strg-Y)?
Gui, Add, CheckBox, x40 y400 Checked vSendEnter, Enter senden nach Schnellstartkürzel?
Gui, Add, CheckBox, x40 y415 vChIban gChIbanChecked, CH-IBAN in die Zwischenablage berechnen?

Gui, Add, GroupBox, x10 y300 w300 h140, Einfügemodus


; Gui, Add, GroupBox, x115 y160 w194 h165, Parameter für freie Eingabe

Gui, Add, Text, x20 y470 w500 h50 , [F2]: markierten Text (z.B. aus Excel) in Liste übernehmen`n[F3]: Wert aus Liste verarbeiten

; Gui, Add, Button, x152 y529 w190 h30 gGuiClose , Anwendung beenden






; ##################### Status Bar


Gui, Add, StatusBar, , Mit F2 neue Werte in die Liste übernehmen


; ##################### Menü aufbauen


Menu, FileMenu, Add, Fenster immer im Vordergrund, menuToggleWindow
Menu, FileMenu, ToggleCheck, Fenster immer im Vordergrund
Menu, FileMenu, Add
Menu, FileMenu, Add, Zwischenablage in Liste übernehmen, addClipboardToHeap
Menu, FileMenu, Add, Doppelte Einträge löschen, deleteDuplicatesInHeap
Menu, FileMenu, Add, Liste sortieren, sortTheHeap
Menu, FileMenu, Add, Liste leeren, clearTheHeap
; Menu, FileMenu, Add, Ausgewählte Elemente löschen, deleteElementsFromHeap
Menu, FileMenu, Add 
Menu, FileMenu, Add, Letzte Aktion rückgängig machen, restoreTheHeap
Menu, FileMenu, Add 
Menu, FileMenu, Add, Über dieses &Programm, showInfoBox
Menu, FileMenu, Add 
Menu, FileMenu, Add, B&eenden, closeApp

Menu, MyMenuBar, Add, &Programm, :FileMenu
Gui, Menu, MyMenuBar

Gui, Show, w320 h560, %scriptname% %version%


Menu, Tray, Tip, %scriptname% %version%


return


menuToggleWindow:
  if isAlwaysOnTop {
    isAlwaysOnTop := 0
    Gui, -AlwaysOnTop
  } else {
    isAlwaysOnTop := 1
    Gui, +AlwaysOnTop
  }
  Menu, FileMenu, ToggleCheck, Fenster immer im Vordergrund
  return

ChIbanChecked:
  GuiControlGet, chkChIban, , ChIban 
  if chkChIban {
    GuiControl, , Ssk, {1} gv sper  
    showTooltip("Berechnet die CH-IBAN des aktuellen Elements `nund speichert diese in der Zwischenablage. `nEinfügen mit Strg-V.", 8000)
  }
  return


GuiClose:
CloseApp()


; ########### Shortcuts definieren

F2::
addSelectionToHeap()
return

F3::
handlePasteRequest()
return

;F9::  ; debug
;Reload
;return

; ########### Programm




handlePasteRequest() {
  modus := ""

  GuiControlGet, modus, , ModusVmass
  if modus {
    sendtoVmass()
    return
  }
  
  GuiControlGet, modus, , ModusStrgY
  if modus {
    GuiControlGet, tmp, , Ssk
    GuiControlGet, chkStrgY, , SendCtrlY
    GuiControlGet, chkEnter, , SendEnter
    GuiControlGet, chkChIban, , ChIban 
    
    sendToStrgY(tmp, chkStrgY, chkEnter, chkChIban )
      
    return
  }
  
   

}



sendToStrgY(command, withStrgY := 0, withEnter := 1, withChIban := 0) {
  ; keys_to_send := "^y^v gv sper{Enter}"
  
  txtStatus := ""
  
  if theHeap.Length() = 0  {
    MsgBox, 48, %scriptname%, Die Verarbeitungsliste ist leer.
    return
  }   
  
  
  element := getElementFromHeap()
  stringtoSend := StrReplace(command, "{1}", element) 
  
  if withStrgY {
    SendInput, ^y
    sleep 200
  }

  SendInput, %stringToSend%
  
  if withEnter {
    SendInput, {Enter}
  }
  

  if withChIban {
    chIban := toChIban(element)
    Clipboard := chIban
    txtStatus := chIban . " (Strg-V)"
  }

  updateWindow(txtStatus)
  return
}




updateWindow(txtStatusBar := "") {
  liste := ""
  for i, element in theHeap
  {
    liste .= "|" . element
  }
  liste := Trim(liste, "|")

  GuiControl,, UiHeap, |
  GuiControl,, UiHeap, %liste%

  if (txtStatusBar = "") {
      SB_SetText("Anzahl: " . theHeap.length())
    } else {
      SB_SetText("Anzahl: " . theHeap.length() . " | " . txtStatusBar)
  }
}






showTooltip(text, showFor:=4000) {
  Tooltip %text%
  sleep %showFor%
  ToolTip 
}

showTraytip(text) {
  TrayTip %scriptname%, %text%, , 16
}


closeApp() {
  if theHeap.Length() > 0
  {
    MsgBox, 52, %scriptname%, Es sind noch Einträge in der Verarbeitungsliste. Soll die Anwendung trotzdem beendet werden?
    IfMsgBox, No
      return
  }
  ExitApp
}

; =================== Funktionen für die Verwaltung des Heaps =========================



deleteElementsFromHeap() {
  elementsToDelete := []
  testArray := []
  newArray := []
  stringElementsToDelete := ""
  countElementsToDelete := 0
  
  GuiControlGet, stringElementsToDelete, , UiHeap 
  elementsToDelete := StrSplit(stringElementsToDelete, "|")
  countElementsToDelete := elementsToDelete.Length()

  if countElementsToDelete = 0   
  {
    MsgBox, 48, %scriptname%, Die Verarbeitungsliste ist leer.
    return
  }

  for k, v in elementsToDelete  {
    MsgBox, %k% -> %v%
    testArray[v] := true
  }

  for k, v in theHeap  {
    if ! testArray.hasKey(v)  
      newArray.Push(v)
  }

  theHeap := newArray

  updateWindow(countElementsToDelete . " Element(e) gelöscht")

  

}






clearTheHeap() { 
  theHeap := []
  ; theBackup := []
  updateWindow()
}







backupTheHeap() {
  theBackup := theHeap.Clone()
}

restoreTheHeap() {
  if theBackup.Length() > 0 {
      theHeap := theBackup
      theBackup := []
      updateWindow("Letze(s) Element(e) wiederhergestellt")
    } else {
      updateWindow("Backup-Liste ist leer.")
  }
  
}

getElementsFromHeap(n=0) {
  elements := []
  element := ""
  i := 0

  backupTheHeap()

  while (i < n and theHeap.Length() > 0)
  {
    element := theHeap.RemoveAt(theHeap.MinIndex())
    elements.Push(element)
    i += 1
  }
  return elements
}


getElementFromHeap() {
  elements := []
  elements := getElementsFromHeap(1)
  return elements.Pop()
}


deleteDuplicatesInHeap() {
  testArray := []
  newArray := []
  
  numbersOfDuplicates := theHeap.length()

  for k, v in theHeap {
    if !testArray.hasKey(v) {
      testArray[v] := true
      newArray.Push(v)
    }
  }
  
  backupTheHeap()
  theHeap := newArray
  numbersOfDuplicates := numbersOfDuplicates - theHeap.length()
  updateWindow(numbersOfDuplicates . " doppelte Einträge entfernt")

}


sortTheHeap() {
  newArray := []
  delimiterChar := "|"
  workingString := "" 
  
  if theHeap.Length() = 0  {
    MsgBox, 48, %scriptname%, Die Verarbeitungsliste ist leer.
    return
  } 
  
  
  for k, v in theHeap {
    workingString := workingString . v . delimiterChar
  }
  
  
  Sort, workingString, N D%delimiterChar%
  
  
  newArray := StrSplit(workingString, delimiterChar)
  
  backupTheHeap()
  theHeap := newArray
  
  updateWindow("Liste wurde aufsteigend sortiert")
}





addClipboardToHeap() {
  
  backupTheHeap()
  
  clip := Trim(Clipboard, " <>`n`r`t`b`a`f`v")
  clipArray := StrSplit(clip, ["`n", "`t"])
  for i, element in clipArray
  {
    theHeap.Push(Trim(element, " <>`n`r`t`b`a`f`v"))    
  }

  updateWindow()
  Clipboard := ""
}

addSelectionToHeap() {
  ;Copy Selection to Clipboard
  SendInput ^c
  Sleep 500
    
  addClipboardToHeap()
  
}

; ==================================================================================


toChIban(accountnr, bankcode:="89202", spaces:=1) {
 
    strAcc := Format("{:010}", accountnr) 
    strBankcode := Format("{:05}", bankcode)
    dummy := strBankcode . "00" . strAcc . "121700"
   
    pz := ""
    while StrLen(dummy) > 0 {
      current := SubStr(dummy, 1, 9 - StrLen(pz))
      pz := Mod(pz . current, 97)
      dummy := StrReplace(dummy, current, "", , 1)
      ;MsgBox current=%current% / pz=%pz% / dummy=%dummy%
    }
    pz := 98 - pz
    tmp := "CH" . Format("{:02}", pz)  strBankcode . "00" . strAcc  
    
    if spaces {
        iban := SubStr(tmp, 1, 4) . " " . SubStr(tmp, 5, 4) . " " . SubStr(tmp, 9, 4) . " " . SubStr(tmp, 13, 4) . " " . SubStr(tmp, 17, 4) . " " . SubStr(tmp, 21, 1)
    } else {
        iban := tmp
    }

    return iban
}




sendToVmass() {
  if theHeap.Length() = 0  {
    MsgBox, 48, %scriptname%, Die Verarbeitungsliste ist leer.
    return
  } 
  
  tmp := []
  tmp := getElementsFromHeap(10)
  
  while tmp.Length() > 0
  {
    element := tmp.RemoveAt(tmp.MinIndex())
    SendInput, %element%{Tab}

  }

  updateWindow()
 
  return
}  



; ################################## Info-Box


showInfoBox() {

welcome = 
(
Daten in Verarbeitungsliste übernehmen:
 * In Excel Zeilen oder Spalten mit Daten markieren und F2 drücken

Daten aus Verarbeitungsliste verarbeiten:
 * Mit F3 wird die im Hauptfenster ausgewählte Aktion ausgeführt
 * vmass: 10 Einträge werden an der Cursorposition mit TAB getrennt eingefügt (für vmass)
 * Schnellstartkürzel: Das im Eingabefeld hinterlegte Schnellstartkürzel wird mit einem Element der Liste ausgeführt
    * Option >Strg-Y senden< fokussiert das Suchfeld
    * Option >Enter senden< löst das Schnellstartkürzel automatisch aus
    * Option >CH-IBAN< stellt die CH-IBAN eines Elements in die Zwischenablage ein
 
 Haftungsausschluss:
 * Der Autor überimmt keinerlei Haftung für Schäden, die aus der Nutzung dieses Tools entstehen.
 
 Erstellt mit AutoHotkey.
)

  MsgBox 64, %scriptname% (c) 2020 Michael Hallmann (Version %version%), %welcome%
}
