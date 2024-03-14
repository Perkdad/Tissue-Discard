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
Private Sub Workbook_Open()
Frame1.Show
Call UserForm_Initialize
End Sub

'Move to Bin Multiselect
Private Sub CommandButton1_Click()
'*****
'Reset
'*****

ActiveWorkbook.Save
Me.EnterBox1.Value = ""
Me.ListBox3.Clear
Me.ListBox4.Clear
Me.Small.Value = False
Me.Large.Value = False
Call UserForm_Initialize

End Sub

Private Sub CommandButton11_Click()
'**********
'Create Bin
'**********

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

    ActiveWorkbook.Save
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
'*************************
'Continue from Selected Pt
'*************************
    
    Dim Bins As Worksheet
    Set Bins = Worksheets("Bins")
    Dim selItm As Long
    Dim lisRow As Integer
    Dim Mystring As String
    
    If Me.ListBox2.List(0, 0) <> "NS" Then
            
        '***Turn On***
        On Error GoTo err
        AppActivate "Tracking Station" '***Good!***
        Application.Wait Now + #12:00:01 AM#
        
        GoTo SkpErr
err:
    MsgBox "Please make sure 'Tissue Discard' is open"
    Resume z
        
SkpErr:
    On Error GoTo 0
    
    For selItm = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
        If Me.ListBox1.Selected(selItm) = True Then '****it is selected***
            Do While selItm <= UBound(Me.ListBox1.List)
                lisRow = Me.ListBox1.List(selItm, 4)
                
                Mystring = Bins.Cells(lisRow, 2).Value '***Outputs the specimen's scan code***
                Application.Wait Now + #12:00:01 AM#
                SendKeys Mystring
                SendKeys "{enter}"
            
                selItm = selItm + 1
            Loop
                AppActivate "discard tissue"
                Call DeleteSection
            GoTo z
        End If
    Next selItm
    GoTo none
        
a:
        
    Else

        MsgBox "Please OPEN a specimen bin to resume Tissue Discard"
        GoTo z
    End If

none:
    Application.Wait Now + #12:00:02 AM#
    AppActivate "discard tissue"
    MsgBox "Please select the next specimen to add to Specimen Discard"
    
z:

End Sub

Private Sub CommandButton16_Click()
'*******
'Delete
'******

    Dim iRemove As VbMsgBoxResult
    iRemove = MsgBox("Do you want to remove this/theese specimen permanently?", vbQuestion + vbYesNo, "Delete Specimen List?")
    
    If iRemove = vbYes Then

        If Me.ListBox2.List(0, 0) <> "NS" Then
        
            Dim selItm As Long
            Dim selCol As Integer
            Dim selRow As Integer
            Dim lisRow As Integer
        
            '***Must first be in active bin***
    
            For selItm = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
                If Me.ListBox1.Selected(selItm) = True Then '****it is selected***
                    lisRow = Me.ListBox1.List(selItm, 4)
                    Rows(lisRow & ":" & lisRow).Delete Shift:=xlUp
                    'UserForm2.Show
                    'GoTo z
                End If
            Next selItm
        Else
            MsgBox "Please scan or select and open a Bin to use this feature"
        End If
    Else
        MsgBox "Fine! I'll leave it alone then."
    End If
    
    Call Update
    Call Organize


End Sub

Private Sub CommandButton2_Click()
'**************
'Tissue Discard
'**************
    
    Dim Bins As Worksheet
    Set Bins = Worksheets("Bins")
    Dim lisRow As Integer
    Dim selItm As Long
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
        
        For selItm = 1 To UBound(Me.ListBox1.List)
            lisRow = Me.ListBox1.List(selItm, 4)
            Mystring = Bins.Cells(lisRow, 2).Value '***Outputs the specimen's scan code***
            Application.Wait Now + #12:00:01 AM#
            SendKeys Mystring
            SendKeys "{enter}"
        Next selItm
        
        AppActivate "discard tissue"
        
        Call DeleteSection
        
    Else
        MsgBox "Please scan or enter a specimen bin to begin Tissue Discard"
    End If
z:
    
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
    Call Organize
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
        
        '***Must first be in active bin***
    
        For selItm = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
            If Me.ListBox1.Selected(selItm) = True Then '****it is selected***
                UserForm2.Show
                Call Update
                GoTo z
            End If
        Next selItm
    Else
        MsgBox "Please scan or select and open a Bin to use this feature"
    End If
z:
    'Call Update
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

Private Sub EnterBox1_exit(ByVal Cancel As MSForms.ReturnBoolean)
'*********
'Entry Box
'*********
    
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
    'EnterBox1.EnterKeyBehavior = False 'Moved to the end of the sub
    'EnterBox1.TabKeyBehavior = False
    
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
                
                '***Updates ListBox4***
                With Me.ListBox4
                    .Clear
                    .AddItem "Start Bin"
                    .List(0, 1) = bar.Cells(rVal, 1)
                End With
                                
                '***Updates ListBox2***
                Me.ListBox2.Clear
                With Me.ListBox2
                    .AddItem bar.Cells(rVal, 1)
                    Call Update
                    GoTo z
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
                
                    '************************************************************
                    'Removes Specimen from Another bin and Adds to the Active Bin
                    '************************************************************
                    
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
                        
                        Call Update
                        Call Organize
                    
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
            'Add to Recent Scan - Done!
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
            End With
            
        '2
        End If
a:
        '***************************
        'Update ListBox1 after Entry
        '***************************
    
        Call Update
        Call Organize
        
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
    
    EnterBox1.EnterKeyBehavior = False
    EnterBox1.TabKeyBehavior = False
            
End Sub

Private Sub Large_Click()
    Frame1.EnterBox1.SetFocus
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox2_Click()
    Frame1.EnterBox1.SetFocus
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
    
    Dim stRange As Integer
    Dim n As Integer
    Dim cntr As Integer 'Counter
    Dim Bins As Worksheet
    Set Bins = Worksheets("Bins")
    stRange = Bins.Cells(Rows.count, 1).End(xlUp).Row

    Me.ListBox1.Clear
    Me.ListBox1.AddItem "Active Bins"
    Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Count"
    
    Me.ListBox2.Clear
    Me.ListBox2.AddItem "NS"
    
    '***Sets range***
    For rVal = 2 To stRange
        cntr = 0
        n = 1
        '***Searches for the same bin number up list, if match, bin is skipped***
        Do While rVal - n > 1
            If Bins.Cells(rVal, 1).Value = Bins.Cells(rVal - n, 1) Then
                GoTo z
            End If
            n = n + 1
        Loop
        
        '***Reste integer to scan down list***
        n = 0
        '***Counts number of specimen in same bin***
        Do While rVal + n <= stRange
            If Bins.Cells(rVal, 1).Value = Bins.Cells(rVal + n, 1) Then
                cntr = cntr + 1
            End If
            n = n + 1
        Loop
        '***Adds Items to ListBox1***
        With Me.ListBox1
            .AddItem Bins.Cells(rVal, 1)
            .List(ListBox1.ListCount - 1, 1) = cntr
        End With
z:
    Next rVal
    
    Frame1.EnterBox1.SetFocus
    EnterBox1.EnterKeyBehavior = False
    EnterBox1.TabKeyBehavior = False
    
End Sub

Sub DeleteSection()

    '*********************************
    'Delete specimen from specimen bin
    '*********************************
    
    Dim Bins As Worksheet
    Set Bins = Worksheets("Bins")
    Dim rVal As Integer
    Dim Mystring As String
    Dim myPath As String
    myPath = Application.ThisWorkbook.Path & "\Discarded"
    Dim myDay, myMonth, myYear
    myDay = Day(Date)
    myMonth = Month(Date)
    myYear = Year(Date)
    
    Dim iRemove As VbMsgBoxResult
    iRemove = MsgBox("Do you want to remove this bin from service and delete all specimen from this list?", vbQuestion + vbYesNo, "Delete Specimen List?")
    
    If iRemove = vbYes Then
        ActiveWorkbook.SaveCopyAs Filename:= _
            myPath & "\Discard Tissue " & Me.ListBox2.List(0, 0) & " " & myYear & myMonth & myDay & ".xlsm"
        
        For rVal = Bins.Cells(Rows.count, 1).End(xlUp).Row To 2 Step -1
            If Bins.Cells(rVal, 1).Value = Me.ListBox2.List(0, 0) Then
                Rows(rVal & ":" & rVal).Delete Shift:=xlUp
            End If
        Next rVal
        
        Call UserForm_Initialize
            
    End If
z:

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
    
    Me.ListBox1.Clear
    
    Me.ListBox1.AddItem "Accession Number"
    Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Part"
    Me.ListBox1.List(ListBox1.ListCount - 1, 2) = "Container"
    Me.ListBox1.List(ListBox1.ListCount - 1, 3) = "Count"
    Me.ListBox1.List(ListBox1.ListCount - 1, 4) = "Row"
    Me.ListBox1.List(ListBox1.ListCount - 1, 5) = "Date"
    Me.ListBox1.List(ListBox1.ListCount - 1, 6) = "Year"
    Me.ListBox1.List(ListBox1.ListCount - 1, 7) = "Specimen"
    
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
                .List(ListBox1.ListCount - 1, 6) = Bins.Cells(rVal, 3)
                .List(ListBox1.ListCount - 1, 7) = Bins.Cells(rVal, 4)
            End With
            cntr = cntr + 1
        End If
    Next rVal
    
End Sub

Sub Organize()
'*****************
'Organize Listbox1
'*****************

Dim i As Long
Dim j As Long
Dim Temp As Variant
Dim Y  As Integer
With ListBox1
    For i = 1 To .ListCount - 1
        For j = i + 1 To .ListCount - 1
            'If .List(i, 7) >= .List(j, 7) Then '***Original that worked - Evaluates as a string***
            If CLng(.List(i, 7)) >= CLng(.List(j, 7)) Then '***Evaluates as a number***
                For Y = 0 To .ColumnCount - 1
                If Y = 3 Then
                    Y = 4 '***Do not want to reorder count column***
                End If
                Temp = .List(j, Y) '***Works when columns have strings***
                'Temp = CLng(.List(j, Y)) '***Results in error because other rows have strings***
                .List(j, Y) = .List(i, Y)
                .List(i, Y) = Temp
                Next Y
            End If
        Next j
    Next i
End With

'******************
'Save for reference
'******************

'***Works very well but only moves 1 row***

'Dim i As Long
'Dim j As Long
'Dim Temp As Variant
'With ListBox1
    'For i = 1 To .ListCount - 1
        'For j = i + 1 To .ListCount - 1
            'If .List(i) > .List(j) Then
                'Temp = .List(j)
                '.List(j) = .List(i)
                '.List(i) = Temp
            'End If
        'Next j
    'Next i
'End With

'*********************************************************************

'For i = LBound(LbList, 1) To UBound(LbList, 1) - 1
    'For j = i + 1 To UBound(LbList, 1)
        'If IsNumeric(LbList(i, Column)) And IsNumeric(LbList(j, Column)) Then
            'If CDbl(LbList(i, Column)) > CDbl(LbList(j, Column)) Then
                'For c = 0 To ListBox1.ColumnCount - 1
                    'sTemp = LbList(i, c)
                    'LbList(i, c) = LbList(j, c)
                    'LbList(j, c) = sTemp
                'Next c
            'End If
        'Else
            'If StrComp(LbList(i, Column), LbList(j, Column), vbTextCompare) = 1 Then
                'For c = 0 To ListBox1.ColumnCount - 1
                    'sTemp = LbList(i, c)
                    'LbList(i, c) = LbList(j, c)
                    'LbList(j, c) = sTemp
                'Next c
            'End If
        'End If
    'Next j
'Next i



End Sub
