VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frame1 
   Caption         =   "Discard Tissue"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15150
   OleObjectBlob   =   "Frame1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frame1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************
'Program:   Discard Tissue Workbook
'Form:      Frame1 Discard Tissue
'Author:    Aaron Perkins
'Date:      10/28/2025
'Version:   1.4.2
'**********************************

'ToDo: Fix entry box scanning bin. Doesn't work and put "Start Bin" when bin is scanned

'Declaration to map virtual keys to work within Citrix
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Const VK_RETURN = &HD
Private Const VK_TAB = &H9
Private Const VK_NUMLOCK = &H90
Private Const VK_DOWN = &H28
Private Const VK_SHIFT = &H10
Private Const KEYEVENTF_KEYUP = &H2

Private Sub EnterButton()
'============
'Enter Button
'============

    Dim mvkEnter As Double
    
    'Map the Enter key
    mvkEnter = MapVirtualKey(VK_RETURN, 0)
    
    'Simulate Enter key press
    keybd_event VK_RETURN, mvkEnter, 0, 0
    'keybd_event VK_RETURN, 0, 0, 0 '***Redundant Key Press that causes error***
    keybd_event VK_RETURN, mvkEnter, KEYEVENTF_KEYUP, 0
    
End Sub

Private Sub TabButton()
'==========
'Tab Button
'==========

    Dim mvkTab As Double
    
    'Map the Tab key
    mvkTab = MapVirtualKey(VK_TAB, 0)
    
    'Simulate Tab key press
    keybd_event VK_TAB, mvkTab, 0, 0
    'keybd_event VK_TAB, 0, 0, 0 '***Redundant Key Press***
    keybd_event VK_TAB, mvkTab, KEYEVENTF_KEYUP, 0
    
End Sub

Private Sub ShiftTabButton()
'==================
'Shift + Tab Button
'==================

    Dim mvkShift As Double
    Dim mvkTab As Double
    
    'Map the Shift and Tab keys
    mvkShift = MapVirtualKey(VK_SHIFT, 0)
    mvkTab = MapVirtualKey(VK_TAB, 0)
    
    'Simulate Shift key press and hold
    keybd_event VK_SHIFT, mvkShift, 0, 0
    
    'Simulate Tab key press
    keybd_event VK_TAB, mvkTab, 0, 0
    keybd_event VK_TAB, mvkTab, KEYEVENTF_KEYUP, 0
    
    'Release Shift key
    keybd_event VK_SHIFT, mvkShift, KEYEVENTF_KEYUP, 0
    
End Sub

Private Sub NumberLock()
'==================
'Number Lock Button
'==================

    Dim mvkNumLock As Double
    
    'Map the Num Lock key
    mvkNumLock = MapVirtualKey(VK_NUMLOCK, 0)
    
    'Simulate Num Lock key press
    keybd_event VK_NUMLOCK, mvkNumLock, 0, 0
    'keybd_event VK_NUMLOCK, 0, 0, 0 '***Redundant Key Press***
    keybd_event VK_NUMLOCK, mvkNumLock, KEYEVENTF_KEYUP, 0
    
End Sub

Sub SendStringToCoPath(ByVal inputString As String)
'=======================
'Send string as keyboard
'=======================

    Dim wsh As Object
    Dim i As Integer
    Dim char As String
    Dim vkCode As Long
    Dim scanCode As Long
    
    ' Create WScript.Shell object for AppActivate
    Set wsh = CreateObject("WScript.Shell")
    
    ' Process each character in the input string
    For i = 1 To Len(inputString)
        char = Mid(inputString, i, 1)
        
        ' Get virtual key code and scan code
        vkCode = GetVirtualKeyCode(char)
        If vkCode = 0 Then
            'MsgBox "Unsupported character: " & char, vbExclamation '***Taken out to prevent interruption***
            Exit Sub
        End If
        scanCode = MapVirtualKey(vkCode, 0)
        
        ' Handle shift for special characters if needed
        If NeedsShift(char) Then
            keybd_event VK_SHIFT, MapVirtualKey(VK_SHIFT, 0), 0, 0 ' Press Shift
        End If
        
        ' Simulate key press and release
        keybd_event vkCode, scanCode, 0, 0
        keybd_event vkCode, scanCode, KEYEVENTF_KEYUP, 0
        
        ' Release Shift if used
        If NeedsShift(char) Then
            keybd_event VK_SHIFT, MapVirtualKey(VK_SHIFT, 0), KEYEVENTF_KEYUP, 0
        End If
    Next i
    
    Set wsh = Nothing
    
End Sub

Private Function GetVirtualKeyCode(ByVal char As String) As Long
'====================
'Character Conversion
'====================

    ' Convert character to virtual key code
    Select Case LCase(char)
        Case "a" To "z"
            GetVirtualKeyCode = Asc(UCase(char))
        Case "0" To "9"
            GetVirtualKeyCode = Asc(char)
        Case "-"
            GetVirtualKeyCode = &HBD ' VK_OEM_MINUS
        Case ";"
            GetVirtualKeyCode = &HBA ' VK_OEM_1
        Case Else
            ' Add more special characters as needed
            GetVirtualKeyCode = 0 ' Unsupported character
    End Select
    
End Function

Private Function NeedsShift(ByVal char As String) As Boolean
'==================
'Special Characters
'==================

    ' Determine if Shift is needed for the character
    Select Case char
        Case "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "{", "}", "|", ":", """", "<", ">", "?"
            NeedsShift = True
        Case Else
            NeedsShift = False
    End Select
    
End Function

Private Sub Workbook_Open()
'=============
'Open Workbook
'=============

    'Frame1.Show
    Call UserForm_Initialize
    
End Sub

Private Sub CommandButton1_Click()
'=====
'Reset
'=====
    
    'Save workbook
    ThisWorkbook.Save
    
    'Clear Entry and list boxes
    Me.EnterBox1.Value = ""
    Me.ListBox3.Clear
    Me.ListBox4.Clear
    
    'Reset and clear Small & Large buttons
    Me.Small.Value = False
    Me.Large.Value = False
    
    'Set behaviors
    Cancel = True
    EnterBox1.EnterKeyBehavior = False
    EnterBox1.TabKeyBehavior = False
    
    Call UserForm_Initialize

End Sub

Private Sub CommandButton11_Click()
'==========
'Create Bin
'==========

    UserForm1.Show
    
End Sub

Private Sub CommandButton12_Click()
'=======================
'Backdoor to spreadsheet
'=======================

    Unload Me
    Application.WindowState = xlNormal

End Sub

Private Sub CommandButton13_Click()
'==========
'Print List
'==========
    
    'Declarations
    Dim listSheet As Worksheet
    Dim rowIndex As Integer
    Dim smallCounter As Integer
    Dim largeCounter As Integer
    
    'Initializations
    Set listSheet = Worksheets("List")
    rowIndex = 1
    smallCounter = 4
    largeCounter = 4
    
    'Save workbook
    ThisWorkbook.Save
    
    'No bin selected; Print message and exit sub
    If Me.ListBox2.List(0, 0) = "NS" Then
        MsgBox "No Bin is selected for printing"
        Exit Sub
    End If

    'Clears the worksheet before adding new items
    listSheet.Cells.ClearContents

    '------------------------------------------------------
    'Setup listSheet Sheet with default Headings and Values
    '------------------------------------------------------
    
    'Set Heading Values
    listSheet.Cells(1, 5).Value = "Bin: " & Me.ListBox2.List(0, 0)
    listSheet.Cells(3, 1).Value = "Small"
    listSheet.Cells(3, 2).Value = "Part"
    listSheet.Cells(3, 3).Value = "Date"
    listSheet.Cells(3, 9).Value = "Large"
    listSheet.Cells(3, 10).Value = "Part"
    listSheet.Cells(3, 11).Value = "Date"

    

    '***Bypass error caused by Do While loop when the list is exceeded***
    On Error GoTo errorBypass
    
    'Step through ListBox1 items and add to List worksheet
    Do While Me.ListBox1.List(rowIndex, 0) <> ""
        
        'Smalls (2 column groups on first page and one group on the second page)
        If Me.ListBox1.List(rowIndex, 2) = "Small" Then
            If smallCounter < 47 Then 'First Column
                listSheet.Cells(smallCounter, 1).Value = Me.ListBox1.List(rowIndex, 0)
                listSheet.Cells(smallCounter, 2).Value = Me.ListBox1.List(rowIndex, 1)
                listSheet.Cells(smallCounter, 3).Value = Me.ListBox1.List(rowIndex, 5)
                smallCounter = smallCounter + 1
            ElseIf smallCounter < 90 Then 'Second Column
                listSheet.Cells(3, 5).Value = "Small"
                listSheet.Cells(3, 6).Value = "Part"
                listSheet.Cells(3, 7).Value = "Date"
                listSheet.Cells(smallCounter - 43, 5).Value = Me.ListBox1.List(rowIndex, 0)
                listSheet.Cells(smallCounter - 43, 6).Value = Me.ListBox1.List(rowIndex, 1)
                listSheet.Cells(smallCounter - 43, 7).Value = Me.ListBox1.List(rowIndex, 5)
                smallCounter = smallCounter + 1
            Else 'Second Page first Column
                listSheet.Cells(47, 1).Value = "Small"
                listSheet.Cells(47, 2).Value = "Part"
                listSheet.Cells(47, 3).Value = "Date"
                listSheet.Cells(smallCounter - 42, 1).Value = Me.ListBox1.List(rowIndex, 0)
                listSheet.Cells(smallCounter - 42, 2).Value = Me.ListBox1.List(rowIndex, 1)
                listSheet.Cells(smallCounter - 42, 3).Value = Me.ListBox1.List(rowIndex, 5)
                smallCounter = smallCounter + 1
            End If
        Else 'Larges (One column group on the first page)
            listSheet.Cells(largeCounter, 9).Value = Me.ListBox1.List(rowIndex, 0)
            listSheet.Cells(largeCounter, 10).Value = Me.ListBox1.List(rowIndex, 1)
            listSheet.Cells(largeCounter, 11).Value = Me.ListBox1.List(rowIndex, 5)
            largeCounter = largeCounter + 1
        End If
        
        'Incriment to next row on ListBox1
        rowIndex = rowIndex + 1
    Loop

errorBypass:
    Resume a
a:
    
    'Final counts for Small and Large
    listSheet.Cells(1, 1).Value = "Small Count:"
    listSheet.Cells(1, 2).Value = smallCounter - 4
    listSheet.Cells(1, 9).Value = "Large Count:"
    listSheet.Cells(1, 10).Value = largeCounter - 4

    'Print List
    listSheet.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False '***Works!!!***
        

End Sub

Private Sub CommandButton14_Click()
'=========================
'Continue from Selected Pt
'=========================
    
    Dim binsSheet As Worksheet
    Set binsSheet = Worksheets("Bins")
    Dim selectedItem As Long
    Dim listedRow As Integer
    Dim myString As String
    
    'Bin is open
    If Me.ListBox2.List(0, 0) <> "NS" Then
        
        'Check if Tracking Station is open
        On Error GoTo errorBypass
        AppActivate "Tracking Station"
        Application.Wait Now() + TimeValue("00:00:01")
        
        'Tracking Station is open, proceed to next section
        GoTo skipError
errorBypass:
        MsgBox "Please make sure 'Tissue Discard' is open"
        Exit Sub
        
skipError:
    
        'Reset normal error function
        On Error GoTo 0
    
        'Start process from selected specimen in list
        For selectedItem = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
            If Me.ListBox1.Selected(selectedItem) = True Then '****it is selected***
            
                'Loop from selected specimen to upper boundary
                Do While selectedItem <= UBound(Me.ListBox1.List)
                    listedRow = Me.ListBox1.List(selectedItem, 4)
                    myString = binsSheet.Cells(listedRow, 2).Value '***Outputs the specimen's scan code (not truncated)***
                    Application.Wait Now() + TimeValue("00:00:01")
                    SendStringToCoPath myString
                    Call EnterButton
                
                    'Incriment to next item in list
                    selectedItem = selectedItem + 1
                Loop
            
                'Places excel back in focus
                AppActivate Application.Caption
                Call DeleteSection
                Exit Sub
            End If
        Next selectedItem
        GoTo none
        
    Else

        MsgBox "Please OPEN a specimen bin to resume Tissue Discard"
        Exit Sub
    End If

none:
    'No specimen is currently selected
    Application.Wait Now() + TimeValue("00:00:02")
    AppActivate Application.Caption
    MsgBox "Please select the next specimen to add to Specimen Discard"
    

End Sub

Private Sub CommandButton16_Click()
'======
'Delete
'======
    
    'Declarations
    Dim iRemove As VbMsgBoxResult
    Dim selectedItem As Long
    Dim selCol As Integer
    Dim selRow As Integer
    Dim listedRow As Integer
    
    'Initializations
    iRemove = MsgBox("Do you want to remove this/theese specimen permanently?", vbQuestion + vbYesNo, _
                     "Delete Specimen List?")
    
    If iRemove = vbYes Then
        
        'Check if user is in an active bin
        If Me.ListBox2.List(0, 0) <> "NS" Then
            
            'Set upper and lower bounds of list
            For selectedItem = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
                If Me.ListBox1.Selected(selectedItem) = True Then '****it is selected***
                    listedRow = Me.ListBox1.List(selectedItem, 4)
                    Rows(listedRow & ":" & listedRow).Delete Shift:=xlUp
                    Call Update 'keep from deleting multiple rows - Rework when time permits
                End If
            Next selectedItem
        Else
            MsgBox "Please scan or select and open a Bin to use this feature"
        End If
    Else
        'User does not want to delete
        MsgBox "Fine! I'll leave it alone then."
    End If
    
    Call Update
    Call Organize
    
End Sub

Private Sub CommandButton17_Click()
'=========
'Empty Bin
'=========
    
    Call DeleteSection
    
End Sub

Private Sub CommandButton18_Click()
'=====================
'Refresh Specimen List
'=====================
    
    Call Update
    Call Organize
    
End Sub

Private Sub CommandButton2_Click()
'==============
'Tissue Discard
'==============
    
    Dim binsSheet As Worksheet
    Set binsSheet = ThisWorkbook.Worksheets("Bins")
    Dim listedRow As Integer
    Dim selectedItem As Long
    Dim myString As String
    
    'Check if user is in an active bin
    If Me.ListBox2.List(0, 0) <> "NS" Then
        
        AppActivate "copath"
        Application.Wait Now() + TimeValue("00:00:01")
        SendStringToCoPath "spec" 'sets selection to specimen tracking
        Call EnterButton
        Application.Wait Now() + TimeValue("00:00:02")
        SendStringToCoPath "t" 'sets drop down to tissue discard
        Application.Wait Now() + TimeValue("00:00:02")
        Call EnterButton
        Application.Wait Now() + TimeValue("00:00:02")
        
        'Locate each item from list in binsSheet and input scan code into CoPath
        For selectedItem = 1 To UBound(Me.ListBox1.List)
            listedRow = Me.ListBox1.List(selectedItem, 4)
            myString = binsSheet.Cells(listedRow, 2).Value '***Outputs the specimen's scan code***
            Application.Wait Now + #12:00:01 AM#
            SendStringToCoPath myString
            Call EnterButton
        Next selectedItem
        
        AppActivate "discard tissue"
        
        Call DeleteSection
        
    Else
        MsgBox "Please scan or enter a specimen bin to begin Tissue Discard"
    End If
    
End Sub

Private Sub CommandButton4_Click()
'========
'Open Bin
'========

   Dim selectedItem As Long
    
    For selectedItem = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
        If Me.ListBox1.Selected(selectedItem) = True Then '****it is selected***
            Me.ListBox2.Clear
            Me.ListBox2.AddItem Me.ListBox1.List(selectedItem, 0)
            
            '***Updates ListBox3***
            With Me.ListBox3
                .Clear
                .AddItem "Start Bin"
                .List(0, 1) = Me.ListBox1.List(selectedItem, 0)
            End With
            Me.ListBox4.Clear
                        
        End If
        
    Next selectedItem
        
    Call Update
    Call Organize
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub CommandButton5_Click()
'===========
'Move to Bin
'===========
    
    'Declarations
    Dim selectedItem As Long
    Dim selCol As Integer
    Dim selRow As Integer
    
    'Check if user is in active bin
    If Me.ListBox2.List(0, 0) <> "NS" Then
        
        For selectedItem = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
            If Me.ListBox1.Selected(selectedItem) = True Then '****it is selected***
                UserForm2.Show
                Call Update
                Exit Sub
            End If
        Next selectedItem
    Else
        MsgBox "Please scan or select and open a Bin to use this feature"
    End If
End Sub

Private Sub CommandButton9_Click()
'====
'Exit
'====
    
    'Declarations
    Dim iExit As VbMsgBoxResult
    
    'Initializations
    iExit = MsgBox("Do you want to exit?", vbQuestion + vbYesNo, "Exit Search")

    If iExit = vbYes Then
        Unload Me
        ThisWorkbook.Save
        Application.Quit
    End If

End Sub

Private Sub EnterBox1_exit(ByVal Cancel As MSForms.ReturnBoolean)
'=========
'Entry Box
'=========
    
    'Declarations
    Dim userEntry As String
    Dim containerType As String
    Dim rowIndex As Integer
    Dim barCodeSheet As Worksheet
    Dim binsSheet As Worksheet
    Dim foundItem As Boolean
    Dim truncationPoint1 As Integer 'Truncation point 1 for scan value (full specimen number)
    Dim truncationPoint2 As Integer 'Truncation point 2 for scan value (full specimen number)
    Dim partNumber As Variant
    Dim moveSpecimen As VbMsgBoxResult
    
    'Initializations
    userEntry = UCase(Trim(EnterBox1.Value)) 'Trim added for Error management
    foundItem = False
    Set barCodeSheet = Worksheets("Barcode")
    Set binsSheet = ThisWorkbook.Worksheets("Bins")
    partNumber = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", _
               "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
               
    'Set behavior
    EnterBox1.EnterKeyBehavior = False
    EnterBox1.TabKeyBehavior = False
    
    '=========================================================================================================
    
    'Checks for userEntry discrepency where multiple entries made at once
    If userEntry Like "*;*;*;*" Then
        MsgBox ("Multiple case userEntry. Please try again")
        'Place cursor back in entry box
        Frame1.EnterBox1.SetFocus
        GoTo z
    End If
    
    '---------------------------------------------------------------------------------------------------------
    
    'Check for void entry
    If userEntry <> "" Then
    
        '==========================
        'Add bin number to ListBox2
        '==========================
        
        rowIndex = 1
        
        'Incriment through barCodeSheet and check values vs userEntry
        Do While barCodeSheet.Cells(rowIndex, 2).Value <> ""
            
            'Sets the Bin Identifier Box (ListBox2) to the userEntry if a match is found
            If barCodeSheet.Cells(rowIndex, 2).Value = userEntry Then
                
                '***Updates ListBox2***
                Me.ListBox2.Clear
                With Me.ListBox2
                    .AddItem barCodeSheet.Cells(rowIndex, 1)
                    Call Update
                    'GoTo z
                End With
                
                '***Updates ListBox3***
                With Me.ListBox3
                    .Clear
                    .AddItem "Start Bin"
                    .List(ListBox3.ListCount - 1, 1) = barCodeSheet.Cells(rowIndex, 1)
                End With
                
                '***Updates ListBox4***
                With Me.ListBox4
                    .Clear
                    '.AddItem "Start Bin"
                    '.List(0, 1) = barCodeSheet.Cells(rowIndex, 1)
                End With
                foundItem = True
                GoTo z
            End If
                
            rowIndex = rowIndex + 1
        Loop
        
        '---------------------------------------------------------------------------------------------------------
        
        '***********************************************************************
        'You are here because the search did not result in an active Bin Barcode
        '***********************************************************************
    
        '=======================================================
        'Search for specimen number and report Bin - No Bin Open
        '=======================================================
        
        'Value not found in barCode and no open Bin
        If foundItem = False And Me.ListBox2.List(0, 0) = "NS" Then
            rowIndex = 2
            
            'Sets Header for ListBox1
            Me.ListBox1.Clear
            Me.ListBox1.AddItem "Acession Number"
            Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Part"
            Me.ListBox1.List(ListBox1.ListCount - 1, 2) = "Bin"
            Me.ListBox1.List(ListBox1.ListCount - 1, 3) = "Container"
            
            'Incriment through binSheets to find matching values
            Do While binsSheet.Cells(rowIndex, 2).Value <> ""
                
                'Search string for userEntry
                If InStr(1, binsSheet.Cells(rowIndex, 2).Value, userEntry) <> 0 Then
                    
                    'Marks points in the Accession number to separate case from part
                    truncationPoint1 = InStr(1, binsSheet.Cells(rowIndex, 2), ";")
                    truncationPoint2 = InStr(truncationPoint1 + 1, binsSheet.Cells(rowIndex, 2), ";")
                    
                    'Adds each partial match to the running list
                    With Me.ListBox1
                        .AddItem Left(binsSheet.Cells(rowIndex, 2), truncationPoint1 - 1)
                        .List(ListBox1.ListCount - 1, 1) = binsSheet.Cells(rowIndex, 5)
                        .List(ListBox1.ListCount - 1, 2) = binsSheet.Cells(rowIndex, 1)
                        .List(ListBox1.ListCount - 1, 3) = binsSheet.Cells(rowIndex, 6)
                    End With
                    foundItem = True
                End If
                
                rowIndex = rowIndex + 1
            Loop
            
            'Message Prompt if Item not found - Done!
            If foundItem = False Then
                MsgBox "Item not found. Please scan a SPECIMEN BIN or enter valid SPECIMEN NUMBER"
            End If
        
        '=========================================================================================================
        
        '=====================================================
        'Add Specimen to Bins Sheet when Bin has been selected
        '=====================================================
        
        'Item not previously found and currently in active Bin
        ElseIf foundItem = False And Me.ListBox2.List(0, 0) <> "NS" Then
            
            rowIndex = 2
            
            'Small or Large specimen container selection
            If Small = True Then
                containerType = "Small"
            ElseIf Large = True Then
                containerType = "Large"
            Else
                MsgBox "Please select Small or Large"
                'foundItem = False
                GoTo z
            End If
            
            '3
            If InStr(1, userEntry, ";") = 0 Then
                MsgBox "Please scan a valid specimen"
                'foundItem = False
                GoTo z
            '3
            End If
            
            '---------------------------------------------------------------------------------------------------------
            
            '===================
            'Check for Duplicate
            '===================
            
            'Incriment through binsSheet for previously scanned value
            Do While binsSheet.Cells(rowIndex, 2).Value <> ""
                
                'userEntry is equal a previously scanned value
                If userEntry = binsSheet.Cells(rowIndex, 2).Value Then
                
                    'Prompt to move specimen
                    moveSpecimen = MsgBox("This specimen has already been added to BIN " _
                                           & binsSheet.Cells(rowIndex, 1) & _
                                           ". Do you want to remove this specimen and add it to the current bin", _
                                           vbQuestion + vbYesNo, "Exit Search")
                    
                    'Yes option selected
                    If moveSpecimen = vbYes Then
                        
                        'Change bin to current bin
                        binsSheet.Cells(rowIndex, 1).Value = Me.ListBox2.List(0, 0)
                        
                        'Change date to current date
                        binsSheet.Cells(rowIndex, 7).Value = Date
                        
                        'Change container type
                        If Small = True Then
                            containerType = "Small"
                        ElseIf Large = True Then
                            containerType = "Large"
                        End If
                        
                        'Set container type in binsSheet
                        binsSheet.Cells(rowIndex, 6).Value = containerType
                        
                        'Truncate userEntry for display
                        truncationPoint1 = InStr(1, binsSheet.Cells(rowIndex, 2), ";")
                        
                        'Add item to ListBox3 for reference
                        With Me.ListBox3
                            .AddItem Left(binsSheet.Cells(rowIndex, 2), truncationPoint1 - 1), 0
                            .List(0, 1) = partNumber(binsSheet.Cells(rowIndex, 5) - 1)
                            .List(0, 2) = ListBox3.ListCount - 1
                        End With
                        ListBox3.TopIndex = 0
                        
                        'Add item to ListBox3 for recently scanned value
                        With Me.ListBox4
                            .Clear
                            .AddItem Left(binsSheet.Cells(rowIndex, 2), truncationPoint1 - 1), 0
                            .List(0, 1) = partNumber(binsSheet.Cells(rowIndex, 5) - 1)
                        End With
                    
                    End If
                    
                    GoTo z
                Else
                    rowIndex = rowIndex + 1
                End If
            Loop
            
            '---------------------------------------------------------------------------------------------------------
            
            '==========================================
            'Add specimen entry to the end of Bins list
            '==========================================
            
            'Count rows and find first empty space in binsSheet
            rowIndex = binsSheet.Cells(Rows.count, 1).End(xlUp).Row + 1
            
            binsSheet.Cells(rowIndex, 1).Value = Me.ListBox2.List(0, 0) '***Bin***
            binsSheet.Cells(rowIndex, 2).Value = userEntry '***Serial Number***
            binsSheet.Cells(rowIndex, 3).Value = Mid(userEntry, 5, 2) '***Year***
                    
            truncationPoint1 = InStr(1, userEntry, "-")
            truncationPoint2 = InStr(1, userEntry, ";")
                    
            binsSheet.Cells(rowIndex, 4).Value = Mid(userEntry, truncationPoint1 + 1, _
                                                 truncationPoint2 - truncationPoint1 - 1) '***Case***
                    
            truncationPoint1 = InStr(1, userEntry, ";")
            truncationPoint2 = InStr(truncationPoint1 + 1, userEntry, ";")
                    
            binsSheet.Cells(rowIndex, 5).Value = Mid(userEntry, truncationPoint1 + 2, _
                                                 truncationPoint2 - truncationPoint1 - 2) '***Part***
            
            binsSheet.Cells(rowIndex, 6).Value = containerType '***Container***
            binsSheet.Cells(rowIndex, 7).Value = Date '***Date***
            
            '---------------------------------------------------------------------------------------------------------
            
            '==================
            'Add to Recent Scan
            '==================
            
            'Add userEntry to recent scan list
            With Me.ListBox3
                .AddItem Left(binsSheet.Cells(rowIndex, 2), truncationPoint1 - 1), 0
                .List(0, 1) = partNumber(binsSheet.Cells(rowIndex, 5) - 1)
                .List(0, 2) = ListBox3.ListCount - 1
            End With
            ListBox3.TopIndex = 0
            
            'Add userEntry to most recently scanned box
            With Me.ListBox4
                .Clear
                .AddItem Left(binsSheet.Cells(rowIndex, 2), truncationPoint1 - 1), 0
                .List(0, 1) = partNumber(binsSheet.Cells(rowIndex, 5) - 1)
            End With
            
        End If
        
        '---------------------------------------------------------------------------------------------------------
        '=========================================================================================================

z:
        
        EnterBox1.Value = ""
        Cancel = True
    
    'Allow user to exit EntryBox 1 if no entry
    Else
        
        Cancel = False
    End If
    
    '***This behavior might be redundant but needs testing***
    EnterBox1.EnterKeyBehavior = False
    EnterBox1.TabKeyBehavior = False
            
End Sub

Private Sub Large_Click()
'==========================
'Selection button for Large
'==========================
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()
'==================
'Current Bin number
'==================
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub ListBox3_Click()
'=================
'Current scan list
'=================
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub ListBox4_Click()
'================
'Most recent scan
'================
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub Small_Click()
'==========================
'Selection button for Small
'==========================
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub UserForm_Initialize()
'===============
'Initialize Form
'===============
    
    'Declarations
    Dim specimenRange As Integer
    Dim n As Integer
    Dim specimenCounter As Integer
    Dim binsSheet As Worksheet
    
    'Initializations
    Set binsSheet = ThisWorkbook.Worksheets("Bins")
    specimenRange = binsSheet.Cells(Rows.count, 1).End(xlUp).Row
    
    'Clear ListBox1 and add headers
    Me.ListBox1.Clear
    Me.ListBox1.AddItem "Active Bins"
    Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Count"
    
    'Clear ListBox2 and add NS (None Selected)
    Me.ListBox2.Clear
    Me.ListBox2.AddItem "NS"
    
    '---------------------------------------------------------------------------------------------------------
    
    'Loop through binsSheet and search for matching specimen
    For rowIndex = 2 To specimenRange
        specimenCounter = 0
        n = 1
        
        'Loop through previous bin numbers and find unique bins
        Do While rowIndex - n > 1
            
            'Bin matches previous bin number
            If binsSheet.Cells(rowIndex, 1).Value = binsSheet.Cells(rowIndex - n, 1) Then
                'Move up list until a unique bin number has been found
                GoTo z
            End If
            
            'Incriment until start of list is reached
            n = n + 1
        Loop
        
        '---------------------------------------------------------------------------------------------------------
        
        'Reset n to scan down list for specimen with matching Bin
        n = 0
        
        'Counts number of specimen in same bin
        Do While rowIndex + n <= specimenRange
            
            'Check for Bins matching current unique bin down the list and incriment counter
            If binsSheet.Cells(rowIndex, 1).Value = binsSheet.Cells(rowIndex + n, 1) Then
                specimenCounter = specimenCounter + 1
            End If
            
            'Incriment until end of list is reached
            n = n + 1
        Loop
        
        'Add unique bin number and count to the list
        With Me.ListBox1
            .AddItem binsSheet.Cells(rowIndex, 1)
            .List(ListBox1.ListCount - 1, 1) = specimenCounter
        End With
z:
    Next rowIndex
    
    'Set behaviors
    Frame1.EnterBox1.SetFocus
    EnterBox1.EnterKeyBehavior = False
    EnterBox1.TabKeyBehavior = False
    Cancel = True
    
End Sub

Sub DeleteSection()
'=================================
'Delete specimen from specimen bin
'=================================
    
    'Declarations
    Dim binsSheet As Worksheet
    Dim rowIndex As Integer
    Dim myString As String
    Dim myPath As String
    Dim thisBin As String
    Dim myDay, myMonth, myYear
    Dim iRemove As VbMsgBoxResult
    
    'Initializations
    Set binsSheet = ThisWorkbook.Worksheets("Bins")
    thisBin = Me.ListBox2.List(0, 0)
    myPath = Application.ThisWorkbook.Path & "\Discarded"
    myDay = Day(Date)
    myMonth = Month(Date)
    myYear = Year(Date)
    
    iRemove = MsgBox("Do you want to remove this bin from service and delete all specimen from this list?", _
                     vbQuestion + vbYesNo, "Delete Specimen List?")
    
    If iRemove = vbYes Then
        ThisWorkbook.SaveCopyAs Filename:=myPath & "\Discard Tissue " & _
                                            thisBin & " " & myYear & myMonth & myDay & ".xlsm"
        
        For rowIndex = binsSheet.Cells(Rows.count, 1).End(xlUp).Row To 2 Step -1
            If binsSheet.Cells(rowIndex, 1).Value = thisBin Then
                Rows(rowIndex & ":" & rowIndex).Delete Shift:=xlUp
            End If
        Next rowIndex
        
        Call UserForm_Initialize
            
    End If
z:

End Sub

Sub Update()
'=================================
'Updates ListBox1 after data entry
'=================================
    
    'Declarations
    Dim binsSheet As Worksheet
    Dim truncationPoint1 As Integer
    Dim truncationPoint2 As Integer
    Dim specimenCounter As Integer
    Dim rowIndex As Integer
    Dim Starter As Integer
    Dim Ender As Integer
    Dim currentBin As String
    Dim partNumber As Variant
    
    
    Set binsSheet = ThisWorkbook.Worksheets("Bins")
    currentBin = Me.ListBox2.List(0, 0)
    partNumber = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", _
                       "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    'Clear ListBox1
    Me.ListBox1.Clear
    
    'Add header to ListBox1
    Me.ListBox1.AddItem "Accession Number"
    Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Part"
    Me.ListBox1.List(ListBox1.ListCount - 1, 2) = "Container"
    Me.ListBox1.List(ListBox1.ListCount - 1, 3) = "Count"
    Me.ListBox1.List(ListBox1.ListCount - 1, 4) = "Row"
    Me.ListBox1.List(ListBox1.ListCount - 1, 5) = "Date"
    Me.ListBox1.List(ListBox1.ListCount - 1, 6) = "Year"
    Me.ListBox1.List(ListBox1.ListCount - 1, 7) = "Specimen"
    
    specimenCounter = 1
        
    For rowIndex = 2 To binsSheet.Cells(Rows.count, 1).End(xlUp).Row
    
        If binsSheet.Cells(rowIndex, 1).Value = currentBin Then
            
            truncationPoint1 = InStr(1, binsSheet.Cells(rowIndex, 2), ";")
            truncationPoint2 = InStr(truncationPoint1 + 1, binsSheet.Cells(rowIndex, 2), ";")
        
            With Me.ListBox1
                .AddItem Left(binsSheet.Cells(rowIndex, 2), truncationPoint1 - 1)
                .List(ListBox1.ListCount - 1, 1) = partNumber(binsSheet.Cells(rowIndex, 5) - 1)
                .List(ListBox1.ListCount - 1, 2) = binsSheet.Cells(rowIndex, 6)
                .List(ListBox1.ListCount - 1, 3) = specimenCounter
                .List(ListBox1.ListCount - 1, 4) = rowIndex
                .List(ListBox1.ListCount - 1, 5) = binsSheet.Cells(rowIndex, 7)
                .List(ListBox1.ListCount - 1, 6) = binsSheet.Cells(rowIndex, 3)
                .List(ListBox1.ListCount - 1, 7) = binsSheet.Cells(rowIndex, 4)
            End With
            specimenCounter = specimenCounter + 1
        End If
    Next rowIndex
    
End Sub

Sub Organize()
'=================
'Organize Listbox1
'=================

    Dim i As Long
    Dim j As Long
    Dim Temp As Variant
    Dim Y  As Integer
    With ListBox1
        For i = 1 To .ListCount - 1
            For j = i + 1 To .ListCount - 1
                If CLng(.List(i, 7)) >= CLng(.List(j, 7)) Then 'Evaluates as a number
                    For Y = 0 To .ColumnCount - 1
                        If Y = 3 Then
                            Y = 4 'Do not want to reorder count column
                        End If
                        Temp = .List(j, Y) '***Works when columns have strings***
                        .List(j, Y) = .List(i, Y)
                        .List(i, Y) = Temp
                    Next Y
                End If
            Next j
        Next i
    End With


End Sub
