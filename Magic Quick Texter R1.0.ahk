;vt,vk,jsr,srj,
#Requires AutoHotkey v2
#SingleInstance Force

;Footer controls
Pause::Pause


; ===============
; Load Hotkeys & Hotstrings from Excel
; Excel columns: Type | Name | Key | Actions
; ===============

; --- Excel file path ---
excelFile := A_ScriptDir "\Shortcuts.xlsx"
splashImage := A_ScriptDir "\splash.png"

; Check if files exist
if !FileExist(excelFile) {
    MsgBox "❌ Excel file not found: " excelFile
    ExitApp
}

; Show splash screen
splashGui := ""
if FileExist(splashImage) {
    try {
        splashGui := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Loading Hotkeys")
        splashGui.Add("Picture", "w300 h200", splashImage)
        splashGui.Show("w320 h220")
    } catch {
        ; Fallback to text splash if image fails
        splashGui := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Loading Hotkeys")
        splashGui.Add("Text", "Center w280 h60", "Loading hotkeys and hotstrings from Excel...`nPlease wait...")
        splashGui.Show("w300 h100")
    }
} else {
    ; Create text splash using Gui
    splashGui := Gui("+AlwaysOnTop -MaximizeBox -MinimizeBox", "Loading Hotkeys")
    splashGui.Add("Text", "Center w280 h60", "Loading hotkeys and hotstrings from Excel...`nPlease wait...")
    splashGui.Show("w300 h100")
}

; --- Initialize Excel COM ---
try {
    xl := ComObject("Excel.Application")
    xl.Visible := false
    wb := xl.Workbooks.Open(excelFile)
    
    ; Select specific worksheet
    sheetName := "vt-vk-mm-ht-hotkeys"
    try {
        ws := wb.Worksheets(sheetName)
    } catch {
        ; If sheet doesn't exist, show available sheets and exit
        availableSheets := ""
        Loop wb.Worksheets.Count {
            availableSheets .= wb.Worksheets(A_Index).Name . "`n"
        }
        wb.Close(false)
        xl.Quit()
        MsgBox "❌ Sheet '" sheetName "' not found!`n`nAvailable sheets:`n" availableSheets
        ExitApp
    }
} catch Error as e {
    MsgBox "❌ Please Close Shortcuts.xlsx And Try Again "
	;MsgBox "❌ Failed to Read Contents From Excel: " e.Message
    ExitApp
}

; --- Find last row with data ---
lastRow := 1
try {
    ; Find last used row in column A
    lastUsedCell := ws.Cells(ws.Rows.Count, 1).End(-4162) ; xlUp = -4162
    lastRow := lastUsedCell.Row
} catch {
    lastRow := 100  ; Fallback to check first 100 rows
}

hotkeyCount := 0
hotstringCount := 0
hotstringList := ""  ; Track loaded hotstrings
hotkeyList := ""     ; Track loaded hotkeys

; --- Process each row ---
Loop lastRow - 1 {  ; Skip header row
    currentRow := A_Index + 1
    
    ; Get cell values safely
    itemType := ""
    itemName := ""
    itemKey := ""
    itemAction := ""
    
    try {
        cellValue := ws.Cells(currentRow, 1).Value
        itemType := cellValue ? Trim(String(cellValue)) : ""
        
        cellValue := ws.Cells(currentRow, 2).Value
        itemName := cellValue ? Trim(String(cellValue)) : ""
        
        cellValue := ws.Cells(currentRow, 3).Value
        itemKey := cellValue ? Trim(String(cellValue)) : ""
        
        cellValue := ws.Cells(currentRow, 4).Value
        itemAction := cellValue ? Trim(String(cellValue)) : ""
    } catch {
        continue
    }
    
    ; Skip empty rows
    if (itemType = "" || itemKey = "" || itemAction = "")
        continue
    
    ; Create hotstrings
    if (itemType = "Hotstring") {
        try {
            ; Clean the key - only remove whitespace
            cleanKey := Trim(itemKey)
            
            ; Skip if cleanKey is empty after cleaning
            if (cleanKey = "")
                continue
                
            ; Build proper hotstring format
            hotstringFormat := ":?*:" . cleanKey 
			;. "::"
            
            ; Create hotstring with proper callback
            Hotstring(hotstringFormat, HotstringCallback.Bind(itemAction))
            hotstringCount++
            
            ; Add to display list (show clean key without formatting)
            hotstringList .= cleanKey . " → " . itemAction . "`n"
            
        } catch Error as e {
            MsgBox "⚠ Failed to create hotstring '" itemName "' (Key: '" itemKey "'): " e.Message
        }
    }
    
    ; Create hotkeys
    else if (itemType = "Hotkey") {
        try {
            Hotkey(itemKey, HotkeyCallback.Bind(itemAction))
            hotkeyCount++
            hotkeyList .= itemKey . " → " . itemAction . "`n"
        } catch Error as e {
            MsgBox "⚠ Failed to create hotkey '" itemName "': " e.Message
        }
    }
}

; --- Cleanup Excel ---
try {
    wb.Close(false)
    xl.Quit()
} catch {
    ; Silent cleanup
}

; Show results with detailed list
resultMsg := "✅ Loaded:`n• " hotstringCount " Hotstrings`n• " hotkeyCount " Hotkeys`n`n"

if (hotstringCount > 0) {
    resultMsg .= "HOTSTRINGS:`n" . hotstringList . "`n"
}

if (hotkeyCount > 0) {
    resultMsg .= "HOTKEYS:`n" . hotkeyList
}
; remove the splash screen
splashGui.Destroy()
;Display the loaded Hotkeys and Hot strings
MsgBox resultMsg

; Add hotkey to show list anytime (Ctrl+Alt+L)
Hotkey("^!l", ShowHotstringList)


; ===============
; Show Hotstring/Hotkey List Function
; ===============
ShowHotstringList(*) {
    listMsg := "📋 LOADED HOTSTRINGS & HOTKEYS`n`n"
    
    if (hotstringList != "") {
        listMsg .= "HOTSTRINGS (" . hotstringCount . "):`n" . hotstringList . "`n"
    } else {
        listMsg .= "No hotstrings loaded.`n`n"
    }
    
    if (hotkeyList != "") {
        listMsg .= "HOTKEYS (" . hotkeyCount . "):`n" . hotkeyList
    } else {
        listMsg .= "No hotkeys loaded."
    }
    
    listMsg .= "`n`nPress Ctrl+Alt+L anytime to show this list."
    
    MsgBox listMsg
}

; ===============
; Callback Functions
; ===============
; --- Hotstring Callback Functions ---
HotstringCallback(action, *) {
    try {
        ; Convert escape sequences (\n → Enter, \t → Tab, etc.)
        action := StrReplace(action, "\n", "`n")
        action := StrReplace(action, "\t", "`t")
        action := StrReplace(action, "\\", "\")

        delayMs := 10   ; default delay
        text := action

        ; --- Check if action starts with DE<number> ---
        if RegExMatch(action, "i)^DE(\d+)\s+(.+)", &m) {
            delayMs := Integer(m[1])   ; extract number
            text    := m[2]            ; remaining text
			SendHotstringTextWithDelay(text, delayMs)
        }
        else if InStr(action, "Send ") = 1 {
            text := ExtractTextFromCommand(action, "Send")
            if (text != "")
                Send text
		}
		else {
        ; Send the replacement text with chosen delay
        SendHotstringTextWithDelay(text, delayMs)
		}
    } catch Error as e {
        MsgBox "❌ Error executing hotstring:`n'" action "'`n`nError: " e.Message
    }
}
SendHotstringTextWithDelay(text, delayMs := 30) {
    ; Regex with case-insensitive option (?i)
    pattern := "(?i)\{(Enter|Tab|Up|Down|Left|Right|Space)\}"
    pos := 1

    while pos <= StrLen(text) {
        if RegExMatch(text, pattern, &m, pos) {
            ; Send normal text before the special key
            if (m.Pos > pos) {
                normalPart := SubStr(text, pos, m.Pos - pos)
                for char in StrSplit(normalPart) {
                    Sendtext char
                    Sleep delayMs
                }
            }
            ; Send the matched key in proper {Case}
            key := Format("{:T}", m[1])   ; normalize text case (TitleCase)
            Send "{" key "}"
            Sleep delayMs
            pos := m.Pos + m.Len
        } else {
            ; Remaining normal text
            normalPart := SubStr(text, pos)
            for char in StrSplit(normalPart) {
                Sendtext char
                Sleep delayMs
            }
            break
        }
    }
}
/*
  ; --- Synthetic Send function with delay Helper Function for Hot String Delayed Typing ---
SendHotstringTextWithDelay(text, delayMs := 10) {  ; default 100 ms
    oldDelay := A_KeyDelay
    SetKeyDelay(delayMs)
    SendEvent(text)
    SetKeyDelay(oldDelay)
}
*/

HotkeyCallback(action, *) {
    ExecuteAction(action)
}

; ===============
; HotkeyCallback Action Execution Function
; ===============
ExecuteAction(action) {
    try {
        ; Convert common escape sequences
        ;action := StrReplace(action, "\n", "`n")
        ;action := StrReplace(action, "\t", "`t")
        ;action := StrReplace(action, "\\", "\")
        
        ; Handle DelayedSend command (simple space-based parsing)
        if InStr(action, "DE") = 1 {
            ; Remove "DelayedSend " from start
            params := SubStr(action, 3) ; Skip "DelayedSend DE"
            
            ; Find first space to separate delay from text
            firstSpace := InStr(params, " ")
            if (firstSpace > 0) {
                delayStr := SubStr(params, 1, firstSpace - 1)
                textPart := SubStr(params, firstSpace + 1)
                
                if (IsNumber(delayStr)) {
                    delay := Integer(delayStr)
                    textToSend := ExtractQuotedText(textPart)
                    HotkeySendTextWithDelay(textToSend, delay)
                    return
                }
            }
        }
    
        ; Handle other commands...
        else if InStr(action, "Send ") = 1 {
            text := ExtractTextFromCommand(action, "Send")
            if (text != "")
                Send text
        } else if InStr(action, "Run ") = 1 {
            text := ExtractTextFromCommand(action, "Run")
            if (text != "")
                Run text
        } else if InStr(action, "MsgBox ") = 1 {
            text := ExtractTextFromCommand(action, "MsgBox")
            if (text != "")
                MsgBox text
        } else if InStr(action, "Sleep ") = 1 {
            sleepStr := ExtractTextFromCommand(action, "Sleep")
            if (sleepStr != "" && IsNumber(sleepStr))
                Sleep Integer(sleepStr)
        } else if InStr(action, "Click") = 1 {
            clickText := ExtractTextFromCommand(action, "Click")
            if (clickText != "") {
                parts := StrSplit(clickText, ",")
                if (parts.Length >= 2) {
                    x := Trim(parts[1])
                    y := Trim(parts[2])
                    if (IsNumber(x) && IsNumber(y))
                        Click Integer(x), Integer(y)
                } else {
                    Click
                }
            } else {
                Click
            }
        } else {
            ;Send action
			delayMs := 30   ; default delay
			HotkeySendTextWithDelay(action, delayMs)
        }
        
    } catch Error as e {
        MsgBox "❌ Error executing action:`n'" action "'`n`nError: " e.Message
    }
}
; Helper function to extract quoted text
ExtractQuotedText(text) {
    text := Trim(text)
    if (SubStr(text, 1, 1) = '"' && SubStr(text, -1) = '"' && StrLen(text) > 1) {
        return SubStr(text, 2, StrLen(text) - 2)
    } else if (SubStr(text, 1, 1) = "'" && SubStr(text, -1) = "'" && StrLen(text) > 1) {
        return SubStr(text, 2, StrLen(text) - 2)
    }
    return text
}

; Helper function to extract text from commands
ExtractTextFromCommand(action, command) {
    startPos := StrLen(command) + 1
    
    ; Skip whitespace after command
    while (startPos <= StrLen(action) && (SubStr(action, startPos, 1) = " " || SubStr(action, startPos, 1) = "`t")) {
        startPos++
    }
    
    if (startPos > StrLen(action))
        return ""
    
    return ExtractQuotedText(SubStr(action, startPos))
}

; Helper function to check if a string is a number
IsNumber(str) {
    try {
        Integer(str)
        return true
    } catch {
        return false
    }
}

; Helper function to send text with custom delay between keystrokes
; Send text with custom key delay
HotkeySendTextWithDelay(text, delayMs := 30) {
    ; Regex with case-insensitive option (?i)
    pattern := "(?i)\{(Enter|Tab|Up|Down|Left|Right|Space)\}"
    pos := 1

    while pos <= StrLen(text) {
        if RegExMatch(text, pattern, &m, pos) {
            ; Send normal text before the special key
            if (m.Pos > pos) {
                normalPart := SubStr(text, pos, m.Pos - pos)
                for char in StrSplit(normalPart) {
                    Sendtext char
                    Sleep delayMs
                }
            }
            ; Send the matched key in proper {Case}
            key := Format("{:T}", m[1])   ; normalize text case (TitleCase)
            Send "{" key "}"
            Sleep delayMs
            pos := m.Pos + m.Len
        } else {
            ; Remaining normal text
            normalPart := SubStr(text, pos)
            for char in StrSplit(normalPart) {
                Sendtext char
                Sleep delayMs
            }
            break
        }
    }
}

; --- Demo hotkeys ---
;F1::SendTextWithDelay("Hello{Tab}World{Enter}Next{Space}Line", 120)
;F2::SendTextWithDelay("Testing slow typing...{Enter}Line 2", 200)


; built in hotstrings action

datefunctioncall(*) {
    CurrentDateTime := FormatTime(, "dd-MM-yyyy")
    SendInput(CurrentDateTime)
}
Timefunctioncall(*) {
    CurrentDateTime := FormatTime(, "hh:mm:ss")
    SendInput(CurrentDateTime)
}
; built in hotstrings
Hotstring(":*:--d", datefunctioncall)
Hotstring(":*:--t", Timefunctioncall)
; --- in Built Hotkey  ---
; Show only active hotstrings (Ctrl+Alt+H)
^!h::MsgBox "Active Hotstrings:`n" . ListHotkeys()
^!e:: Run excelFile
^!r::Reload

;Tray Menu Area
;excelFile := A_ScriptDir "\Shortcuts.xlsx"

; --- Create Tray Menu ---
A_TrayMenu.Delete() ; clear default items
A_TrayMenu.Add("Open Excel File", (*) => OpenExcelFile())
A_TrayMenu.Add("Help | Readme", AppInfoMenu)
A_TrayMenu.Add("Reload Script", reloadfunction)
A_TrayMenu.Add("Exit App", (*) => ExitApp())
A_TrayMenu.Default := "Open Excel File" ; double-click tray = open Excel

reloadfunction(*)
{
MsgBox "Save And Close Shortcuts Excel file If it is open"
Reload()
}
; --- Functions ---
OpenExcelFile() {
    global excelFile
    if !FileExist(excelFile) {
        MsgBox "Shortcuts Excel file not found:`n" excelFile
        return
    }

run excelFile 
}

AppInfoMenu(*) {
    appInfoGui := Gui("+AlwaysOnTop", "Help - Magic Quick Texter")
    appInfoGui.OnEvent("Close", AppInfoGui_Close)
    appInfoGui.SetFont("s18", "Verdana")
    appInfoGui.AddText(, "Magic Quick Texter")
    appInfoGui.SetFont("s10", "Verdana")
    appInfoGui.AddText("w500 h200", "Help Info")
    appInfoGui.Show()
}

AppInfoGui_Close(*) {
    ; Gui automatically closes when this event handler returns
}
