VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frame1 
   Caption         =   "Discard Tissue"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15150
   OleObjectBlob   =   "Frame1 with Citrix Keyboard Map.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frame1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Const VK_RETURN = &HD
Private Const KEYEVENTF_KEYUP = &H2

Private Sub Workbook_Open()
Frame1.Show
End Sub

Private Sub CommandButton1_Click()
'*****
'Reset
'*****

Me.EnterBox1.Value = ""
Me.ListBox3.Clear
Me.ListBox4.Clear
Call Update
Call UserForm_Initialize

End Sub

Private Sub CommandButton10_Click()
    
    Call Update
    Frame1.EnterBox1.SetFocus
    
End Sub

Private Sub CommandButton11_Click()
    UserForm1.Show
End Sub

Private Sub CommandButton12_Click()
'***********************
'Backdoor to spreadsheet
'***********************

Unload Me

Application.WindowState = xlNormal

End Sub

Private Sub CommandButton13_Click()
'**********
'Print List
'**********

Call Update


If Me.ListBox2.List(0, 0) = "NS" Then
    MsgBox "No Bin is selected for printing"
    GoTo Bend
End If



'***Clear Sheet***
Worksheets("List").Cells.ClearContents

'*********
'New Start
'*********

Dim List As Worksheet
Set List = Worksheets("List")
Dim rVal As Integer
Dim SmVal As Integer
Dim LgVal As Integer

List.Cells(1, 5).Value = "Bin: " & Me.ListBox2.List(0, 0)
List.Cells(3, 1).Value = "Small"
List.Cells(3, 2).Value = "Part"
List.Cells(3, 3).Value = "Date"
List.Cells(3, 9).Value = "Large"
List.Cells(3, 10).Value = "Part"
List.Cells(3, 11).Value = "Date"

rVal = 1
SmVal = 4
LgVal = 4

'***Bypass error caused by Do While loop***
On Error GoTo err

Do While Me.ListBox1.List(rVal, 0) <> "" 'Throws error when list is exceeded
    
    If Me.ListBox1.List(rVal, 2) = "Small" Then
        If SmVal < 48 Then 'First Row
            List.Cells(SmVal, 1).Value = Me.ListBox1.List(rVal, 0)
            List.Cells(SmVal, 2).Value = Me.ListBox1.List(rVal, 1)
            List.Cells(SmVal, 3).Value = Me.ListBox1.List(rVal, 5)
            SmVal = SmVal + 1
        ElseIf SmVal < 92 Then 'Second Row
            List.Cells(3, 5).Value = "Small"
            List.Cells(3, 6).Value = "Part"
            List.Cells(3, 7).Value = "Date"
            List.Cells(SmVal - 44, 5).Value = Me.ListBox1.List(rVal, 0)
            List.Cells(SmVal - 44, 6).Value = Me.ListBox1.List(rVal, 1)
            List.Cells(SmVal - 44, 7).Value = Me.ListBox1.List(rVal, 5)
            SmVal = SmVal + 1
        Else 'Second Page first row
            List.Cells(48, 1).Value = "Small"
            List.Cells(48, 2).Value = "Part"
            List.Cells(48, 3).Value = "Date"
            List.Cells(SmVal - 43, 1).Value = Me.ListBox1.List(rVal, 0)
            List.Cells(SmVal - 43, 2).Value = Me.ListBox1.List(rVal, 1)
            List.Cells(SmVal - 43, 3).Value = Me.ListBox1.List(rVal, 5)
            SmVal = SmVal + 1
        End If
    Else
        List.Cells(LgVal, 9).Value = Me.ListBox1.List(rVal, 0)
        List.Cells(LgVal, 10).Value = Me.ListBox1.List(rVal, 1)
        List.Cells(LgVal, 11).Value = Me.ListBox1.List(rVal, 5)
        LgVal = LgVal + 1
    End If

    rVal = rVal + 1
Loop

err:
    Resume a
a:
    
List.Cells(1, 1).Value = "Small Count:"
List.Cells(1, 2).Value = SmVal - 4
List.Cells(1, 9).Value = "Large Count:"
List.Cells(1, 10).Value = LgVal - 4


'***Turn on when complete***
'***Print List***
Worksheets("List").PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False '***Works!!!***
        
Bend:

End Sub

Private Sub CommandButton14_Click()
'**************
'Tissue Discard
'**************
    
    Dim Bins As Worksheet
    Set Bins = Worksheets("Bins")
    Dim rVal As Integer
    rVal = 2
    Dim Counter As Integer
    Counter = 0
    Dim Starter As Integer
    Dim Ender As Integer
    Dim Mystring As String
    
    If Me.ListBox2.List(0, 0) <> "NS" Then
    
        
        '***Turn On***
        'AppActivate "Tracking Station" '***Good!***
        'Application.Wait Now + #12:00:01 AM#
        
        '***Delete when ready***
        ''SendKeys "spec"
        ''SendKeys "{enter}"
        ''Application.Wait Now + #12:00:02 AM#
        ''SendKeys "tissue discard"
        ''Application.Wait Now + #12:00:02 AM#
        ''SendKeys "{enter}"
        ''Application.Wait Now + #12:00:02 AM#
        
        
        
    Dim selItm As Long
    'Dim selItmAcc As Long
    'Dim selBin As String
    'Dim Bins As Worksheet
    'Set Bins = Worksheets("Bins")
    'Dim rVal As Integer
    
    For selItm = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
    'MsgBox Me.ListBox1.List(selItm, 4) '***Starts counting at 0****
    'MsgBox UBound(Me.ListBox1.List) '***Works***
        If Me.ListBox1.Selected(selItm) = True Then '****it is selected***
            
            
            
            
            
            
            
            
            
            
            'Counter = Counter + 1
            
            'selBin = Me.ListBox1.List(selItm, 0) '***Sets the New Bin number***
            
            'For selItmAcc = LBound(Frame1.ListBox1.List) To UBound(Frame1.ListBox1.List) '***Reference specimen selected on Frame1***
                'If Frame1.ListBox1.Selected(selItmAcc) = True Then
                
                '***************************************************
                
                'MsgBox selItm 'Properly lists row for selected item in UserForm2 ListBox1
                'MsgBox selItmAcc 'Properly lists row for selected item in Frame1 ListBox1
                'MsgBox Frame1.ListBox1.List(selItmAcc, 4) 'Properly lists row of spec
                
                'rVal = Frame1.ListBox1.List(selItmAcc, 4) '*********THIS!!!!!!***************
                    
                        '***Change bin to selected bin***
                        'Bins.Cells(rVal, 1).Value = selBin
                        '***Change date to current date***
                       'Bins.Cells(rVal, 7).Value = Date
                       
                '******************************************************
                'End If
            'Next selItmAcc
            GoTo a
            
        End If
        Counter = Counter + 1
        
    Next selItm
    GoTo none
        
        
a:
        'Do While Bins.Cells(rVal, 1).Value <> ""
            'If Bins.Cells(rVal, 1).Value = Me.ListBox2.List(0, 0) Then
                'Counter = Counter + 1
            
                
                ''***Turn on after testing***
                'Mystring = Bins.Cells(rVal, 2).Value
                'Application.Wait Now + #12:00:01 AM#
                'SendKeys Mystring
                'SendKeys "{enter}"
                
                ''***Turn off afteer testing***
                ''AppActivate "excel"
                ''MsgBox Mystring
                ''Application.Wait Now + #12:00:01 AM#
                
            
                'If Counter = 1 Then
                    'Starter = rVal
                    'Ender = rVal
                'Else
                    'Ender = rVal
                'End If
            'End If
            'rVal = rVal + 1
        'Loop
    
    'AppActivate "excel"
    
        'If Counter > 0 Then
            'Dim iRemove As VbMsgBoxResult
            'iRemove = MsgBox("Do you want to remove this bin from service and delete all specimen from this list?", vbQuestion + vbYesNo, "Delete Specimen List?")
            'If iRemove = vbYes Then
                
                ''***Turn on after testing**
                'Rows(Starter & ":" & Ender).Delete Shift:=xlUp
                ''Bins.Rows(Starter & ":" & Ender).Select
                
                'Me.EnterBox1.Value = ""
                'Me.ListBox3.Clear
                'Call Update
                'Call UserForm_Initialize
                
            'End If
        'End If
    'Else
none:
        'MsgBox "Please SELECT a specimen bin to resume Tissue Discard"
    End If
'none:
    'MsgBox "Please select the next specimen to add to Specimen Discard"
    
z:
End Sub



Private Sub CommandButton15_Click()



'AppActivate "notepad"
'SendKeys "This is a test"
''vbKeyReturn
'KeyAscii = 15  'Does not throw error, but did not work
'SendKeys "This is only a test"










'******************Example Code**********************************************

'Option Explicit
'Private Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Private Declare PtrSafe Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
'Private Const VK_RETURN = &HD
'Private Const KEYEVENTF_KEYUP = &H2

'AppActivate "copath"
        'Application.Wait Now + #12:00:01 AM#
        'SendKeys "spec"
        'SendKeys "{enter}"
        'Application.Wait Now + #12:00:02 AM#
        'SendKeys "tissue discard"
        'Application.Wait Now + #12:00:02 AM#
        'SendKeys "{enter}"
        'Application.Wait Now + #12:00:02 AM#

'Sub pleasework5()
    
    Dim keys As String
    Dim wsh As Object
    Dim mvk As Double
       
    Set wsh = CreateObject("WScript.Shell")
    mvk = MapVirtualKey(VK_RETURN, 0)
    'AppActivate "notepad"
    AppActivate "copath"
    
    Application.Wait Now() + TimeValue("00:00:01")
    'wsh.SendKeys ("2207062") 'Sends value
    '***CoPath is not cool with any of this***
    'wsh.SendKeys ("spec")
    'wsh.SendKeys ("s")
    'wsh.SendKeys ("p")
    'wsh.SendKeys ("e")
    'wsh.SendKeys ("c")
    'SendKeys "spec"
    
    '*****Now Acting as Return Button*****
    Application.Wait Now() + TimeValue("00:00:01")
    keybd_event VK_RETURN, mvk, 0, 0
    keybd_event VK_RETURN, 0, 0, 0
    keybd_event VK_RETURN, 0, KEYEVENTF_KEYUP, 0
    Application.Wait Now() + TimeValue("00:00:01")
    
    'wsh.SendKeys ("1")
    wsh.SendKeys ("tissue discard")
    
    Application.Wait Now() + TimeValue("00:00:01")
    keybd_event VK_RETURN, mvk, 0, 0
    keybd_event VK_RETURN, 0, 0, 0
    keybd_event VK_RETURN, mvk, KEYEVENTF_KEYUP, 0
    Application.Wait Now() + TimeValue("00:00:01")
    
    wsh.SendKeys ("ROSS23-39250;P1;KAI")
    
    Application.Wait Now() + TimeValue("00:00:02")
    keybd_event VK_RETURN, mvk, 0, 0
    keybd_event VK_RETURN, 0, 0, 0
    keybd_event VK_RETURN, mvk, KEYEVENTF_KEYUP, 0
    Application.Wait Now() + TimeValue("00:00:02")
    
    Set wsh = Nothing



End Sub

Private Sub CommandButton2_Click()

'**************
'Tissue Discard
'**************
    
    Dim Bins As Worksheet
    Set Bins = Worksheets("Bins")
    Dim rVal As Integer
    rVal = 2
    Dim Counter As Integer
    Counter = 0
    Dim Starter As Integer
    Dim Ender As Integer
    Dim Mystring As String
    
    If Me.ListBox2.List(0, 0) <> "NS" Then
    
        AppActivate "copath"
        Application.Wait Now + #12:00:01 AM#
        SendKeys "spec"
        SendKeys "{enter}"
        Application.Wait Now + #12:00:02 AM#
        SendKeys "tissue discard"
        Application.Wait Now + #12:00:02 AM#
        SendKeys "{enter}"
        Application.Wait Now + #12:00:02 AM#
    
        Do While Bins.Cells(rVal, 1).Value <> ""
            If Bins.Cells(rVal, 1).Value = Me.ListBox2.List(0, 0) Then
                Counter = Counter + 1
            
                
                '***Turn on after testing***
                Mystring = Bins.Cells(rVal, 2).Value
                Application.Wait Now + #12:00:01 AM#
                SendKeys Mystring
                'Application.Wait Now + #12:00:01 AM#
                SendKeys "{enter}"
                
                '***Turn off afteer testing***
                'AppActivate "excel"
                'MsgBox Mystring
                'Application.Wait Now + #12:00:01 AM#
                
            
                If Counter = 1 Then
                    Starter = rVal
                    Ender = rVal
                Else
                    Ender = rVal
                End If
            End If
            rVal = rVal + 1
        Loop
    
    AppActivate "excel"
    
        If Counter > 0 Then
            Dim iRemove As VbMsgBoxResult
            iRemove = MsgBox("Do you want to remove this bin from service and delete all specimen from this list?", vbQuestion + vbYesNo, "Delete Specimen List?")
            If iRemove = vbYes Then
                
                '***Turn on after testing**
                Rows(Starter & ":" & Ender).Delete Shift:=xlUp
                'Bins.Rows(Starter & ":" & Ender).Select
                
                Me.EnterBox1.Value = ""
                Me.ListBox3.Clear
                Call Update
                Call UserForm_Initialize
                
            End If
        End If
    Else
        MsgBox "Please scan or enter a specimen bin to begin Tissue Discard"
    End If
'z:
    
End Sub

Private Sub CommandButton4_Click()
'********
'Open Bin
'********

   Dim selItm As Long
    
    For selItm = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
        If Me.ListBox1.Selected(selItm) = True Then '****it is selected***
            Me.ListBox2.Clear
            Me.ListBox2.AddItem Me.ListBox1.List(selItm, 0)
            
            '***Updates ListBox3***
            With Me.ListBox3
                .Clear
                .AddItem "Start Bin"
                .List(0, 1) = Me.ListBox1.List(selItm, 0)
            End With
            Me.ListBox4.Clear
            
            
        End If
        
    Next selItm
    
        
    
    
    Call Update
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub CommandButton5_Click()
'***********
'Move to Bin
'***********
    
    
    '******************************
    'See if you can add multiselect
    '******************************
    
    If Me.ListBox2.List(0, 0) <> "NS" Then
        
        Dim selItm As Long
        Dim selCol As Integer
        Dim selRow As Integer
        'Dim sh As Worksheet
        'Set sh = ActiveSheet
    
        '***Must first be in active bin***
        '"Please Open or Scan an active bin and select a valid Specimen"
    
        For selItm = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
            If Me.ListBox1.Selected(selItm) = True Then '****it is selected***
                UserForm2.Show
                'selCol = Me.ListBox1.List(selItm, 2)
                'selRow = Me.ListBox1.List(selItm, 3)
                'sh.Cells(selRow, selCol).Delete Shift:=xlUp
            
                'If Cells(2, selCol).Value = 0 Then
                    'Columns(selCol).Delete Shift:=xlToLeft
                'End If
            'Else: MsgBox "No item is selected"
            Call Update
            GoTo z
            End If
        Next selItm
        'Call Update
    Else
        MsgBox "Please scan or select and open a Bin to use this feature"
    End If
    'Call UserForm_Initialize
z:
    
End Sub

Private Sub CommandButton9_Click()
'****
'Exit
'****

Dim iExit As VbMsgBoxResult

iExit = MsgBox("Do you want to exit?", vbQuestion + vbYesNo, "Exit Search")

If iExit = vbYes Then
    Unload Me
    ActiveWorkbook.Save
    Application.Quit
End If

End Sub

Private Sub CQY_Click()

End Sub

Private Sub EnterBox1_exit(ByVal Cancel As MSForms.ReturnBoolean)
'*********
'Entry Box
'*********

    '***Swap Container and count
    '***Needs to organize - Working well - verify
    
        
    '***Check everything for Ucase - Should be good
    '***Validate that all col + 1 are good and don't need to be changed to col + 2 - I think this is complete, but will need to continue to check
    '***Clean up colums so that only desired info is displayed*** - Do this when you have added all function; you might need column and row data for "Discard Tissue"
    '***Add function - "Discard Tissue"***
        '***Deleting list clears bin and saves receipt file (log) showing when specimen were removed form the workbook***
        


    Dim Entry As String
    Dim ConT As String
    Dim tDate As Date
    Entry = UCase(EnterBox1.Value)
    Dim col As Integer
    Dim rVal As Integer
    Dim bar As Worksheet
    Dim Bins As Worksheet
    Dim catch As Integer
    catch = 0
    Set bar = Worksheets("Barcode")
    Set Bins = Worksheets("Bins")
    Dim Pnt1 As Integer
    Dim Pnt2 As Integer
    EnterBox1.EnterKeyBehavior = False
    EnterBox1.TabKeyBehavior = False
    
    Dim PN As Variant
    PN = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    '1
    If Entry <> "" Then
    
        '**************************
        'Add bin number to ListBox2 - Done!
        '**************************
        
        rVal = 1
        
        Do While bar.Cells(rVal, 2).Value <> ""
            
            '*********************
            'Locate active Bin Row - Done!
            '*********************
            
            '2
            If bar.Cells(rVal, 2).Value = Entry Then
                
                '***Updates ListBox3***
                With Me.ListBox3
                    .Clear
                    .AddItem "Start Bin"
                    .List(ListBox3.ListCount - 1, 1) = bar.Cells(rVal, 1)
                End With
                With Me.ListBox4
                    .Clear
                    .AddItem "Start Bin"
                    .List(0, 1) = bar.Cells(rVal, 1)
                End With
                                
                Me.ListBox2.Clear
                With Me.ListBox2
                    .AddItem bar.Cells(rVal, 1)
                    GoTo a
                End With
            '2
            End If
                
            rVal = rVal + 1
        Loop
        
        '***Yow are here because the search did not result in an active Bin Barcode***
    
        '*****************************************
        'Search for specimen number and report Bin - Done!
        '*****************************************
        '2
        If bar.Cells(rVal, 2).Value = "" And Me.ListBox2.List(0, 0) = "NS" Then
            
            'col = 2
            rVal = 2
            
            Do While Bins.Cells(rVal, 2).Value <> ""
                '3
                If InStr(1, Bins.Cells(rVal, 2).Value, Entry) <> 0 Then '***Serches string for entry***
                    '***Allows script to skip clearing the list box in the loop - Entries are added after one another***
                    '4
                    If catch = 1 Then
                        GoTo cb
                    '4
                    End If
                         
                    '***Sets Header for ListBox1***
                    Me.ListBox1.Clear
                    Me.ListBox1.AddItem "Acession Number"
                    Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Part"
                    Me.ListBox1.List(ListBox1.ListCount - 1, 2) = "Bin"
                    Me.ListBox1.List(ListBox1.ListCount - 1, 3) = "Container"
cb:
                    '***Marks points in the Accession number to separate case from part***
                    Pnt1 = InStr(1, Bins.Cells(rVal, 2), ";")
                    Pnt2 = InStr(Pnt1 + 1, Bins.Cells(rVal, 2), ";")
                    '***Adds each partial match to the running list***
                    With Me.ListBox1
                        .AddItem Left(Bins.Cells(rVal, 2), Pnt1 - 1)
                        '.List(ListBox1.ListCount - 1, 1) = Mid(Bins.Cells(rVal, 2), Pnt1 + 1, Pnt2 - Pnt1 - 1)
                        .List(ListBox1.ListCount - 1, 1) = Bins.Cells(rVal, 5)
                        .List(ListBox1.ListCount - 1, 2) = Bins.Cells(rVal, 1)
                        .List(ListBox1.ListCount - 1, 3) = Bins.Cells(rVal, 6)
                    End With
                    '***Sets catch value to 1 so workbook progresses to end of Sub upon completing loop***
                    catch = 1
                '3
                End If
                rVal = rVal + 1
            Loop
            '3
            If catch = 1 Then
                GoTo z
            '3
            End If
            
            '********************************
            'Message Prompt if Item not found - Done!
            '********************************
            '3
            If Bins.Cells(rVal, 2).Value = "" Then
                MsgBox "Item not found. Please scan a SPECIMEN BIN or enter valid SPECIMEN NUMBER"
                GoTo z
            '3
            End If
        
        '*****************************************************
        'Add Specimen to Bins Sheet when Bin has been selected - Done!
        '*****************************************************
        '2
        ElseIf bar.Cells(rVal, 2).Value = "" And Me.ListBox2.List(0, 0) <> "NS" Then 'Double check that this is needed still?
            
            If Small = True Then
                ConT = "Small"
            ElseIf Large = True Then
                ConT = "Large"
            Else
                MsgBox "Please select Small or Large"
                GoTo z
            End If
            
            '3
            If InStr(1, Entry, ";") = 0 Then
                GoTo ms1
            '3
            End If
            
            '*******************
            'Check for Duplicate - Done!
            '*******************
            
            rVal = 2
            
            Do While Bins.Cells(rVal, 2).Value <> ""
                '3
                If Entry = Bins.Cells(rVal, 2).Value Then
    '********
    'Here!!!!
    '********
                    Dim DelMove As VbMsgBoxResult

                    DelMove = MsgBox("This specimen has already been added to BIN " & Bins.Cells(rVal, 1) & ". Do you want to remove this specimen and add it to the current bin", vbQuestion + vbYesNo, "Exit Search")

                    If DelMove = vbYes Then
                        
                        '***Change bin to selected bin***
                        Bins.Cells(rVal, 1).Value = Me.ListBox2.List(0, 0)
                        '***Change date to current date***
                        Bins.Cells(rVal, 7).Value = Date
                        
                        If Small = True Then
                            ConT = "Small"
                        ElseIf Large = True Then
                            ConT = "Large"
                        End If
                        
                        Bins.Cells(rVal, 6).Value = ConT
                        
                        Pnt1 = InStr(1, Bins.Cells(rVal, 2), ";")
                        
                        With Me.ListBox3
                            .AddItem Left(Bins.Cells(rVal, 2), Pnt1 - 1), 0
                            .List(0, 1) = PN(Bins.Cells(rVal, 5) - 1)
                            .List(0, 2) = ListBox3.ListCount - 1
                        End With
                        ListBox3.TopIndex = 0
                        
                        With Me.ListBox4
                            .Clear
                            .AddItem Left(Bins.Cells(rVal, 2), Pnt1 - 1), 0
                            .List(0, 1) = PN(Bins.Cells(rVal, 5) - 1)
                            '.List(0, 2) = ListBox3.ListCount - 1
                        End With
                        
                        'Call Update
                    
                    End If
                    
                    GoTo z
                Else
                    rVal = rVal + 1
                '3
                End If
            Loop
            '***If you are here you have an rVal for the first empty row***
            
            '******************************************
            'Add specimen entry to the end of Bins list - Done!
            '******************************************
            
            rVal = Bins.Cells(Rows.count, 1).End(xlUp).Row + 1
            
            Bins.Cells(rVal, 1).Value = Me.ListBox2.List(0, 0) '***Bin***
            Bins.Cells(rVal, 2).Value = Entry '***SN***
            Bins.Cells(rVal, 3).Value = Mid(Entry, 5, 2) '***Year***
                    
            Pnt1 = InStr(1, Entry, "-")
            Pnt2 = InStr(1, Entry, ";")
                    
            Bins.Cells(rVal, 4).Value = Mid(Entry, Pnt1 + 1, Pnt2 - Pnt1 - 1) '***Case***
                    
            Pnt1 = InStr(1, Entry, ";")
            Pnt2 = InStr(Pnt1 + 1, Entry, ";")
                    
            Bins.Cells(rVal, 5).Value = Mid(Entry, Pnt1 + 2, Pnt2 - Pnt1 - 2) '***Part***
            
            Bins.Cells(rVal, 6).Value = ConT '***Container***
            Bins.Cells(rVal, 7).Value = Date '***Date***
            
            
            
            
            '******************
            'Add to Recent Scan
            '******************
                        
            With Me.ListBox3
                .AddItem Left(Bins.Cells(rVal, 2), Pnt1 - 1), 0
                .List(0, 1) = PN(Bins.Cells(rVal, 5) - 1)
                .List(0, 2) = ListBox3.ListCount - 1
            End With
            ListBox3.TopIndex = 0
            
            With Me.ListBox4
                .Clear
                .AddItem Left(Bins.Cells(rVal, 2), Pnt1 - 1), 0
                .List(0, 1) = PN(Bins.Cells(rVal, 5) - 1)
                '.List(0, 2) = ListBox3.ListCount - 1
            End With
            
        '2
        End If
a:
        '***************************
        'Update ListBox1 after Entry
        '***************************
    
        'Call Update - Taking a lot of time as dataset grows
        
        GoTo z
        
ms1:
    MsgBox "Please scan a valid specimen"

z:
        
        EnterBox1.Value = ""
        Cancel = True
    '1
    Else
        
        Cancel = False
    '1
    End If
            
End Sub



Private Sub Large_Click()
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()

End Sub

Private Sub ListBox3_Click()
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub ListBox4_Click()
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub Small_Click()
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub UserForm_Initialize()
'***************
'Initialize Form - Done!
'***************
    
    Dim rVal As Integer
    Dim cntr As Integer
    Dim Bins As Worksheet
    Set Bins = Worksheets("Bins")


    Me.ListBox1.Clear
    Me.ListBox1.AddItem "Active Bins"
    Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Count"
    
    Me.ListBox2.Clear
    Me.ListBox2.AddItem "NS"
    
    For rVal = 2 To Bins.Cells(Rows.count, 1).End(xlUp).Row
        cntr = 1
        Do While Bins.Cells(rVal, 1).Value = Bins.Cells(rVal + 1, 1)
            cntr = cntr + 1
            rVal = rVal + 1
        Loop
        With Me.ListBox1
            .AddItem Bins.Cells(rVal, 1)
            .List(ListBox1.ListCount - 1, 1) = cntr
        End With
    Next rVal
    
    Frame1.EnterBox1.SetFocus
    
End Sub

Sub Update()
    
    '*********************************
    'Updates ListBox1 after data entry - Done
    '*********************************
        
    Dim Bins As Worksheet
    Dim Pnt1 As Integer
    Dim Pnt2 As Integer
    Dim cntr As Integer
    Dim rVal As Integer
    Dim Starter As Integer
    Dim Ender As Integer
    Dim curBin As String
    Set Bins = Worksheets("Bins")
    curBin = Me.ListBox2.List(0, 0)
    
    Dim PN As Variant
    PN = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    
    rVal = Bins.Cells(Rows.count, 1).End(xlUp).Row
    
    '***********
    'Sort by Bin - Done!
    '***********
    
    Bins.Sort.SortFields.Clear
    Bins.Sort.SortFields.Add Key:=Range("A2:A" & rVal), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Bins.Sort
        .SetRange Range("A2:G" & rVal)
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '*******************
    'Sort by year in Bin
    '*******************
    
    For rVal = 2 To Bins.Cells(Rows.count, 1).End(xlUp).Row
        Starter = rVal
        
        Do While Bins.Cells(Starter, 1) = Bins.Cells(rVal, 1)
            rVal = rVal + 1
        Loop
        
        '***Checks specimen year against different Bin***
        Do While Bins.Cells(Starter, 1) <> Bins.Cells(rVal - 1, 1)
            rVal = rVal - 1
        Loop
        
        '***********************
        If Starter <> rVal Then
            rVal = rVal - 1
        End If
        '***********************
    
        Ender = rVal
        
        Bins.Sort.SortFields.Clear
        Bins.Sort.SortFields.Add Key:=Range("C" & Starter & ":C" & Ender), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With Bins.Sort
            .SetRange Range("A" & Starter & ":G" & Ender)
            .header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Next rVal
    
    '*******************
    'Sort by case in year
    '*******************
    
    For rVal = 2 To Bins.Cells(Rows.count, 1).End(xlUp).Row
        Starter = rVal
        Do While Bins.Cells(Starter, 3) = Bins.Cells(rVal, 3)
            rVal = rVal + 1
            
        Loop
        
        '***Checks case against different Bin***
        Do While Bins.Cells(Starter, 1) <> Bins.Cells(rVal - 1, 1)
            rVal = rVal - 1
        Loop
        
        '***Checks case against different year***
        Do While Bins.Cells(Starter, 3) <> Bins.Cells(rVal - 1, 3)
            rVal = rVal - 1
        Loop
        
        '***********************
        If Starter <> rVal Then
            rVal = rVal - 1
        End If
        '***********************
        
        Ender = rVal
                        
        Bins.Sort.SortFields.Clear
        Bins.Sort.SortFields.Add Key:=Range("D" & Starter & ":D" & Ender), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With Bins.Sort
            .SetRange Range("A" & Starter & ":G" & Ender)
            .header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Next rVal
    
    '*******************
    'Sort by part in case
    '*******************
    
    For rVal = 2 To Bins.Cells(Rows.count, 1).End(xlUp).Row
        
        Starter = rVal
        Do While Bins.Cells(Starter, 4) = Bins.Cells(rVal, 4)
            rVal = rVal + 1
        Loop
        
        '***Checks part against different Bin***
        Do While Bins.Cells(Starter, 1) <> Bins.Cells(rVal - 1, 1)
            rVal = rVal - 1
        Loop
        
        '***Checks part against different year***
        Do While Bins.Cells(Starter, 3) <> Bins.Cells(rVal - 1, 3)
            rVal = rVal - 1
        Loop
        
        '***Checks part against different case***
        Do While Bins.Cells(Starter, 4) <> Bins.Cells(rVal - 1, 4)
            rVal = rVal - 1
        Loop
        
        '***********************
        If Starter <> rVal Then
            rVal = rVal - 1
        End If
        '***********************
        
        Ender = rVal
        
        Bins.Sort.SortFields.Clear
        Bins.Sort.SortFields.Add Key:=Range("E" & Starter & ":E" & Ender), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With Bins.Sort
            .SetRange Range("A" & Starter & ":G" & Ender)
            .header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Next rVal
        
    '*****
    '*****
    '*****
    '*****
    '*****
    
    rVal = 2
    
    Me.ListBox1.Clear
    
    Me.ListBox1.AddItem "Accession Number"
    Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Part"
    Me.ListBox1.List(ListBox1.ListCount - 1, 2) = "Container"
    Me.ListBox1.List(ListBox1.ListCount - 1, 3) = "Count"
    Me.ListBox1.List(ListBox1.ListCount - 1, 4) = "Row"
    Me.ListBox1.List(ListBox1.ListCount - 1, 5) = "Date"
    
    cntr = 1
        
    For rVal = 2 To Bins.Cells(Rows.count, 1).End(xlUp).Row
    
        If Bins.Cells(rVal, 1).Value = curBin Then
            
            Pnt1 = InStr(1, Bins.Cells(rVal, 2), ";")
            Pnt2 = InStr(Pnt1 + 1, Bins.Cells(rVal, 2), ";")
        
            With Me.ListBox1
                .AddItem Left(Bins.Cells(rVal, 2), Pnt1 - 1)
                .List(ListBox1.ListCount - 1, 1) = PN(Bins.Cells(rVal, 5) - 1)
                .List(ListBox1.ListCount - 1, 2) = Bins.Cells(rVal, 6)
                .List(ListBox1.ListCount - 1, 3) = cntr
                .List(ListBox1.ListCount - 1, 4) = rVal
                .List(ListBox1.ListCount - 1, 5) = Bins.Cells(rVal, 7)
            End With
            cntr = cntr + 1
        End If
    Next rVal
End Sub
