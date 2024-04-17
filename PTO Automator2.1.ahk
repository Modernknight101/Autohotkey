SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Gui, Show, x0 y0 h600 w600, PTO Tracker
Gui, font, Cblack s16
Gui, Add, Button,x0 y0, EmailAll

gui, add, text,x200 y0, Today's date
gui, add, text,x0, Employee Email

gui, add, text,x302 y43, Balances (Hrs)
Gui, font, Cblack s12
Gui, Add, Edit,x0 w300 h30 vEdit1  
Gui, Add, Edit, w300 h30 vEdit2 
Gui, Add, Edit,w300 h30 vEdit3
Gui, Add, Edit, w300 h30 vEdit4 
Gui, Add, Edit, w300 h30 vEdit5 
Gui, Add, Edit, x302 y92  w50 h30 vEdit6
Gui, Add, Edit, w50 h30 vEdit7
Gui, Add, Edit, w50 h30 vEdit8
Gui, Add, Edit, w50 h30 vEdit9
Gui, Add, Edit, w50 h30 vEdit10
Gui, Add, Edit,x330 y0 w100 h30 vEdit11
Gui, Add, Edit, x354 y92  w50 h30 vEdit12
Gui, Add, Edit, w50 h30 vEdit13
Gui, Add, Edit, w50 h30 vEdit14
Gui, Add, Edit, w50 h30 vEdit15
Gui, Add, Edit, w50 h30 vEdit16
Gui, Add, Edit, x406 y92  w50 h30 vEdit17
Gui, Add, Edit, w50 h30 vEdit18
Gui, Add, Edit, w50 h30 vEdit19
Gui, Add, Edit, w50 h30 vEdit20
Gui, Add, Edit, w50 h30 vEdit21
Gui, Add, Edit, x458 y92  w50 h30 vEdit22
Gui, Add, Edit, w50 h30 vEdit23
Gui, Add, Edit, w50 h30 vEdit24
Gui, Add, Edit, w50 h30 vEdit25
Gui, Add, Edit, w50 h30 vEdit26
Gui, font, Cblack s16
Gui, Add, Button,x0 w170 h35, LoadCells:1-5
Gui, Add, Button,x170 y302 w170 h35, LoadCells:6-10
Gui, Add, Button,x340 y302 w170 h35, LoadCells:11-15
Gui, Add, Button,x0 y340 w170 h35, LoadCells:16-20
Gui, Add, Button,x170 y340 w170 h35, LoadCells:21-25
Gui, Add, Button,x340 y340 w170 h35, LoadCells:26-30
Gui, Add, Button,x0 y378 w170 h35, LoadCells:31-35
Gui, Add, Button,x170 y378 w170 h35, LoadCells:36-40
Gui, Add, Button,x340 y378 w170 h35, LoadCells:41-45
Gui, Add, Button,x0 y416 w170 h35, LoadCells:46-50
Gui, Add, Text,, Spreadsheet Location: C:>AHK>Time

Gui, add, Button,x510 y90 w80 h33 ,Email1	
Gui, add, Button,y+10 w80 h33,Email2
Gui, add, Button,y+9 w80 h33,Email3
Gui, add, Button,y+9 w80 h33,Email4
Gui, add, Button,y+9 w80 h33,Email5


ButtonLoadCells:1-5:
Guicontrol,,Edit1, ; - Add a semicolon or leave blank at the end to blank the text
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;
Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(2, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(2, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(2, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(2, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(2, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(3, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(3, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(3, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(3, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(3, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(4, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(4, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(4, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(4, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(4, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(5, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(5, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(5, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(5, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(5, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(6, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(6, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(6, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(6, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(6, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return

ButtonLoadCells:6-10:
Guicontrol,,Edit1, ;  
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;

Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(7, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(7, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(7, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(7, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(7, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(8, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(8, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(8, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(8, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(8, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(9, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(9, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(9, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(9, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(9, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(10, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(10, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(10, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(10, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(10, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(11, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(11, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(11, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(11, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(11, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return

ButtonLoadCells:11-15:

Guicontrol,,Edit1, ;  
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;

Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(12, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(12, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(12, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(12, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(12, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(13, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(13, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(13, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(13, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(13, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(14, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(14, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(14, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(14, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(14, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(15, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(15, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(15, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(15, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(15, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(16, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(16, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(16, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(16, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(16, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return

ButtonLoadCells:16-20:

Guicontrol,,Edit1, ;  
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;

Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(17, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(17, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(17, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(17, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(17, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(18, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(18, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(18, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(18, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(18, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(19, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(19, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(19, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(19, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(19, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(20, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(20, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(20, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(20, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(20, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(21, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(21, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(21, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(21, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(21, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return

ButtonLoadCells:21-25:

Guicontrol,,Edit1, ;  
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;

Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(22, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(22, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(22, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(22, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(22, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(23, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(23, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(23, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(23, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(23, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(24, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(24, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(24, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(24, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(24, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(25, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(25, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(25, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(25, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(25, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(26, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(26, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(26, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(26, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(26, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return


ButtonLoadCells:26-30:

Guicontrol,,Edit1, ;  
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;

Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(27, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(27, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(27, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(27, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(27, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(28, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(28, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(28, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(28, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(28, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(29, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(29, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(29, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(29, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(29, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(30, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(30, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(30, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(30, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(30, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(31, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(31, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(31, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(31, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(31, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return

ButtonLoadCells:31-35:

Guicontrol,,Edit1, ;  
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;

Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(32, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(32, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(32, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(32, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(32, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(33, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(33, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(33, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(33, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(33, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(34, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(34, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(34, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(34, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(34, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(35, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(35, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(35, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(35, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(35, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(36, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(36, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(36, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(36, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(36, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return

ButtonLoadCells:36-40:

Guicontrol,,Edit1, ;  
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;

Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(37, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(37, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(37, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(37, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(37, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(38, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(38, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(38, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(38, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(38, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(39, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(39, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(39, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(39, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(39, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(40, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(40, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(40, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(40, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(40, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(41, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(41, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(41, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(41, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(41, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return


ButtonLoadCells:41-45:

Guicontrol,,Edit1, ;  
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;

Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(42, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(42, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(42, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(42, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(42, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(43, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(43, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(43, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(43, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(43, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(44, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(44, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(44, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(44, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(44, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(45, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(45, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(45, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(45, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(45, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(46, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(46, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(46, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(46, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(46, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return

ButtonLoadCells:46-50:

Guicontrol,,Edit1, ;  
Guicontrol,,Edit2, ;
Guicontrol,,Edit3, ;
Guicontrol,,Edit4, ;
Guicontrol,,Edit5, ;
Guicontrol,,Edit6, ;
Guicontrol,,Edit7, ;
Guicontrol,,Edit8, ;
Guicontrol,,Edit9, ;
Guicontrol,,Edit10, ;
Guicontrol,,Edit12, ;
Guicontrol,,Edit13, ;
Guicontrol,,Edit14, ;
Guicontrol,,Edit15, ;
Guicontrol,,Edit16, ;
Guicontrol,,Edit17, ;
Guicontrol,,Edit18, ;
Guicontrol,,Edit19, ;
Guicontrol,,Edit20, ;
Guicontrol,,Edit21, ;
Guicontrol,,Edit22, ;
Guicontrol,,Edit23, ;
Guicontrol,,Edit24, ;
Guicontrol,,Edit25, ;
Guicontrol,,Edit26, ;

Excel := ComObjCreate("Excel.Application")
Path := "C:\AHK\Time.xlsx"
; Open the Excel file
Workbook := Excel.Workbooks.Open(Path)

; Get a reference to the first sheet in the workbook
Sheet := Workbook.Worksheets(1)

; Loop through the cells in the first column of the sheet
	
    ; Get the value of the cell
    CellValue := Sheet.Cells(47, 1).Value
	
    ; Paste the cell value to the GUI
    ControlSend, Edit1, %CellValue%
	
	    ; Get the value of the cell
    CellValue := Sheet.Cells(47, 2).Value
	RoundedValue := Round(CellValue, 2)
    ; Paste the cell value to the GUI
    ControlSend, Edit6, %RoundedValue%
	
	
    CellValue := Sheet.Cells(47, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit12, %RoundedValue%
	
	CellValue := Sheet.Cells(47, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit17, %RoundedValue%
	
	CellValue := Sheet.Cells(47, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit22, %RoundedValue%
	
	CellValue := Sheet.Cells(48, 1).Value
    ControlSend, Edit2, %CellValue%
	
	CellValue := Sheet.Cells(48, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit7, %RoundedValue%
	
	CellValue := Sheet.Cells(48, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit13, %RoundedValue%
	
	CellValue := Sheet.Cells(48, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit18, %RoundedValue%
	
	CellValue := Sheet.Cells(48, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit23, %RoundedValue%
	
	CellValue := Sheet.Cells(49, 1).Value
    ControlSend, Edit3, %CellValue%
	
	CellValue := Sheet.Cells(49, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit8, %RoundedValue%
	
	CellValue := Sheet.Cells(49, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit14, %RoundedValue%
	
	CellValue := Sheet.Cells(49, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit19, %RoundedValue%
	
	CellValue := Sheet.Cells(49, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit24, %RoundedValue%
	
	CellValue := Sheet.Cells(50, 1).Value
    ControlSend, Edit4, %CellValue%
	
	CellValue := Sheet.Cells(50, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit9, %RoundedValue%
	
	CellValue := Sheet.Cells(50, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit15, %RoundedValue%
	
	CellValue := Sheet.Cells(50, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit20, %RoundedValue%
	
	CellValue := Sheet.Cells(50, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit25, %RoundedValue%
	
	CellValue := Sheet.Cells(51, 1).Value
    ControlSend, Edit5, %CellValue%
	
	CellValue := Sheet.Cells(51, 2).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit10, %RoundedValue%
	
	CellValue := Sheet.Cells(51, 3).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit16, %RoundedValue%
	
	CellValue := Sheet.Cells(51, 4).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit21, %RoundedValue%
	
	CellValue := Sheet.Cells(51, 5).Value
	RoundedValue := Round(CellValue, 2) 
    ControlSend, Edit26, %RoundedValue%
	
; Close the Excel file and release the Excel object
Workbook.Close()
Excel.Quit()
Excel := ""
Return

Edit1:
Send,{Raw}%A_GuiEvent%
Return

ButtonEmail1:
	gui,submit,nohide ;updates gui variable
	
	Send, {Raw}%A_GuiEvent%

; Check if the edit box contains a period
If (InStr(Edit1, ".") > 0)
{
		DetectHiddenWindows, On
		Process, Exist, outlook.exe
		If !ErrorLevel
			{
			Run outlook.exe
			Sleep 15000
			}
		outlookApp := ComObjActive("Outlook.Application")
	
		olMailItem := 0
		MailItem := outlookApp.CreateItem(olMailItem)
		MailItem.TO :=(Edit1)
		MailItem.Subject :="current PTO balance " 
	
		;*****************************************************
		MailItem.HTMLBody := " 
		<html>Hello,<br> 
		<br> 		Your leave balances as of "(Edit11)" are "(Edit6)" annual leave hours,
		<br> "(Edit12)" sick leave hours, "(Edit17)" personal leave hours, and "(Edit22)" compensatory hours. <br>
		
		<br>Please note, this may not include the most recently submitted leave requests if they were received by Payroll in last 2 business days.
		<br>
		<br> Thank you<br>
		<br>
	
		 
		<br>Joshua E Parkhurst – Accounting Technician III, Payroll
		<br>Colorado School for the Deaf and the Blind
		<br>33 North Institute Street, Colorado Springs, CO 80903
		<br>719-578-2111; jparkhurst@csdb.org
		<br>Check out our new website design at www.csdb.org
		</html>"
		;***********************************************

		MailItem.Display
		Mailitem.Send
	
}
Return

ButtonEmail2:
	gui,submit,nohide ;updates gui variable
	
	Send, {Raw}%A_GuiEvent%

; Check if the edit box contains a period
If (InStr(Edit2, ".") > 0)
{
		DetectHiddenWindows, On
		Process, Exist, outlook.exe
		If !ErrorLevel
			{
			Run outlook.exe
			Sleep 15000
			}
		outlookApp := ComObjActive("Outlook.Application")
	
		olMailItem := 0
		MailItem := outlookApp.CreateItem(olMailItem)
		MailItem.TO :=(Edit2)
		MailItem.Subject :="current PTO balance " 
	
		;*****************************************************
		MailItem.HTMLBody := " 
		<html>Hello,<br> 
		<br> 		Your leave balances as of "(Edit11)" are "(Edit7)" annual leave hours,
		<br> "(Edit13)" sick leave hours, "(Edit18)" personal leave hours, and "(Edit23)" compensatory hours. <br>
		
		<br>Please note, this may not include the most recently submitted leave requests if they were received by Payroll in last 2 business days.
		<br>
		<br> Thank you<br>
		<br>
	
		 
		<br>Joshua E Parkhurst – Accounting Technician III, Payroll
		<br>Colorado School for the Deaf and the Blind
		<br>33 North Institute Street, Colorado Springs, CO 80903
		<br>719-578-2111; jparkhurst@csdb.org
		<br>Check out our new website design at www.csdb.org
		</html>"
		;***********************************************

		MailItem.Display
		Mailitem.Send
	
}
Return

ButtonEmail3:
	gui,submit,nohide ;updates gui variable
	
	Send, {Raw}%A_GuiEvent%

; Check if the edit box contains a period
If (InStr(Edit3, ".") > 0)
{
		DetectHiddenWindows, On
		Process, Exist, outlook.exe
		If !ErrorLevel
			{
			Run outlook.exe
			Sleep 15000
			}
		outlookApp := ComObjActive("Outlook.Application")
	
		olMailItem := 0
		MailItem := outlookApp.CreateItem(olMailItem)
		MailItem.TO :=(Edit3)
		MailItem.Subject :="current PTO balance " 
	
		;*****************************************************
		MailItem.HTMLBody := " 
		<html>Hello,<br> 
		<br> 		Your leave balances as of "(Edit11)" are "(Edit8)" annual leave hours,
		<br> "(Edit14)" sick leave hours, "(Edit19)" personal leave hours, and "(Edit24)" compensatory hours. <br>
		
		<br>Please note, this may not include the most recently submitted leave requests if they were received by Payroll in last 2 business days.
		<br>
		<br> Thank you<br>
		<br>
		 
		<br>Joshua E Parkhurst – Accounting Technician III, Payroll
		<br>Colorado School for the Deaf and the Blind
		<br>33 North Institute Street, Colorado Springs, CO 80903
		<br>719-578-2111; jparkhurst@csdb.org
		<br>Check out our new website design at www.csdb.org
		</html>"
		;***********************************************

		MailItem.Display
		Mailitem.Send
	
}
Return

ButtonEmail4:
	gui,submit,nohide ;updates gui variable
	
	Send, {Raw}%A_GuiEvent%

; Check if the edit box contains a period
If (InStr(Edit4, ".") > 0)
{
		DetectHiddenWindows, On
		Process, Exist, outlook.exe
		If !ErrorLevel
			{
			Run outlook.exe
			Sleep 15000
			}
		outlookApp := ComObjActive("Outlook.Application")
	
		olMailItem := 0
		MailItem := outlookApp.CreateItem(olMailItem)
		MailItem.TO :=(Edit4)
		MailItem.Subject :="current PTO balance " 
	
		;*****************************************************
		MailItem.HTMLBody := " 
		<html>Hello,<br> 
		<br> 		Your leave balances as of "(Edit11)" are "(Edit9)" annual leave hours,
		<br> "(Edit15)" sick leave hours, "(Edit20)" personal leave hours, and "(Edit25)" compensatory hours. <br>
		
		<br>Please note, this may not include the most recently submitted leave requests if they were received by Payroll in last 2 business days.
		<br>
		<br> Thank you<br>
		<br>
		 
		<br>Joshua E Parkhurst – Accounting Technician III, Payroll
		<br>Colorado School for the Deaf and the Blind
		<br>33 North Institute Street, Colorado Springs, CO 80903
		<br>719-578-2111; jparkhurst@csdb.org
		<br>Check out our new website design at www.csdb.org
		</html>"
		;***********************************************

		MailItem.Display
		Mailitem.Send
	
}
Return

ButtonEmail5:
	gui,submit,nohide ;updates gui variable
	
	Send, {Raw}%A_GuiEvent%

; Check if the edit box contains a period
If (InStr(Edit5, ".") > 0)
{
		DetectHiddenWindows, On
		Process, Exist, outlook.exe
		If !ErrorLevel
			{
			Run outlook.exe
			Sleep 15000
			}
		outlookApp := ComObjActive("Outlook.Application")
	
		olMailItem := 0
		MailItem := outlookApp.CreateItem(olMailItem)
		MailItem.TO :=(Edit5)
		MailItem.Subject :="current PTO balance " 
	
		;*****************************************************
		MailItem.HTMLBody := " 
		<html>Hello,<br> 
		<br> 		Your leave balances as of "(Edit11)" are "(Edit10)" annual leave hours,
		<br> "(Edit16)" sick leave hours, "(Edit21)" personal leave hours, and "(Edit26)" compensatory hours. <br>
		
		<br>Please note, this may not include the most recently submitted leave requests if they were received by Payroll in last 2 business days.
		<br>
		<br> Thank you<br>
		<br>
		 
		<br>Joshua E Parkhurst – Accounting Technician III, Payroll
		<br>Colorado School for the Deaf and the Blind
		<br>33 North Institute Street, Colorado Springs, CO 80903
		<br>719-578-2111; jparkhurst@csdb.org
		<br>Check out our new website design at www.csdb.org
		</html>"
		;***********************************************

		MailItem.Display
		Mailitem.Send
	
}
Return

ButtonEmailAll:
gui,submit,nohide
Sleep, 1500
ControlClick, Email1
Sleep, 1500
ControlClick, Email2
Sleep, 1500
ControlClick, Email3
Sleep, 1500
ControlClick, Email4
Sleep, 1500
ControlClick, Email5


Return
	
GuiClose:
ExitApp



