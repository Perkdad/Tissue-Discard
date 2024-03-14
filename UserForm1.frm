VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Bin Editor"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4905
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
        
    Dim bar As Worksheet
    Dim Bins As Worksheet
    Set bar = Worksheets("Barcode")
    Set Bins = Worksheets("Bins")
    Dim cBin As String
    Dim cBar As String
    cBin = UCase(TextBox1.Value)
    cBar = UCase(TextBox2.Value)
    Dim rVal As Integer
    Dim answer As Integer
    
    '**********************************
    'Checks for a bin number in textbox - Done!
    '**********************************
    
    If cBin = "" Then
        MsgBox "Please Enter a Bin Number to proceed"
        TextBox2.Value = ""
        GoTo z
    End If
    
    '****************************
    'Check for a valid Bin Number - Done!
    '****************************
    
    If Not UCase(cBin) Like "[A-Z]##" Then
        MsgBox "You must enter a Bin number consisting of a Letter and 2 numerical Digits"
        TextBox1.Value = ""
        TextBox2.Value = ""
        GoTo z
    End If
    
    '************************************************
    'Checks to see if the Bin is active with specimen - Done!
    '************************************************
    
    rVal = 1
    Do While Bins.Cells(rVal, 1) <> ""
        If UCase(Bins.Cells(rVal, 1).Value) = UCase(cBin) Then
            GoTo m1
        End If
        rVal = rVal + 1
    Loop
    
    '***************************************
    'Checks if the barcode is already in use - Done!
    '***************************************
        
    rVal = 1
    Do While bar.Cells(rVal, 2) <> ""
        If UCase(bar.Cells(rVal, 2).Value) = UCase(cBar) Then
            GoTo m2 '***assigned to another bin***
        End If
        rVal = rVal + 1
    Loop
    
    rVal = 1
    
    Do While Bins.Cells(rVal, 2) <> ""
        If UCase(Bins.Cells(rVal, 2).Value) = UCase(cBar) Then
            GoTo m4
        End If
        rVal = rVal + 1
    Loop
    
    '********************************
    'If bin and bar have been entered - Done!
    '********************************
           
    If cBin <> "" And cBar <> "" Then
        
        '************************
        'Checks for valid barcode - Done!
        '************************
        
        If InStr(cBar, ";") = 0 Then 'double check this works
            GoTo m3
        End If
        
        'ROSS23-15305;P1;KAI
        
        '*************************************************
        'Checks if the Bin is currently assigned a barcode - Done!
        '*************************************************
        
        rVal = 2
        Do While bar.Cells(rVal, 1) <> ""
            If UCase(bar.Cells(rVal, 1).Value) = UCase(cBin) Then
                answer = MsgBox("This Bin currently has a Barcode. Would you like to update the Barcode?", vbYesNo)
                If answer = vbNo Then
                    TextBox1.Value = ""
                    TextBox2.Value = ""
                    GoTo z
                Else
                    GoTo a
                End If
            End If
            rVal = rVal + 1
        Loop
        
a:
        
        '******************************************************************
        'Searches for bin in list else appends bin to list and adds barcode - Done!
        '******************************************************************
        
        If bar.Cells(rVal, 1) <> "" Then
            bar.Cells(rVal, 2).Value = UCase(cBar)
        Else
            bar.Cells(rVal, 2).Value = UCase(cBar)
            bar.Cells(rVal, 1).Value = UCase(cBin)
        End If
            TextBox1.Value = ""
            TextBox2.Value = ""
        GoTo z
        
    End If
        
    '***********************
    'Remove bin from service - Done!
    '***********************
    
    If cBin <> "" And cBar = "" Then
    
        rVal = 2
        Do While bar.Cells(rVal, 1) <> ""
            If bar.Cells(rVal, 1) = UCase(cBin) Then
                answer = MsgBox("Are you sure you would like to remove this Bin from service and dissociate the Barcode?", vbYesNo)
                If answer = vbYes Then
                    bar.Range("A" & rVal & ":B" & rVal).Delete Shift:=xlUp
                    TextBox1.Value = ""
                    TextBox2.Value = ""
                    GoTo z
                End If
                GoTo z
            End If
            
            rVal = rVal + 1
            
        Loop
        
        MsgBox "You have not entered an active Bin number"
        TextBox1.Value = ""
        TextBox2.Value = ""
    End If
    
m1:
    MsgBox "This Bin is active and may not be edited at this time"
    TextBox1.Value = ""
    TextBox2.Value = ""
    GoTo z
m2:
    MsgBox "This Barcode is already assigned to another bin"
    TextBox2.Value = ""
    GoTo z
m3:
    MsgBox "Please enter a valid barcode"
    TextBox2.Value = ""
    GoTo z
m4:
    MsgBox "This is an active Barcode for a specimen in Bin " & Bins.Cells(rVal, 1)
    TextBox1.Value = ""
    TextBox2.Value = ""
    GoTo z
    
z:

Call UserForm_Initialize

End Sub


Private Sub ListBox1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()
    
    'Application.Visible = False
    'Organize bin numbers
    
    UserForm1.TextBox1.SetFocus
    
    Dim bar As Worksheet
    Dim Bins As Worksheet
    Set bar = Worksheets("Barcode")
    Set Bins = Worksheets("Bins")
    Dim rVal As Integer
    Dim roCnt As Integer
    rVal = bar.Cells(Rows.count, 1).End(xlUp).Row
    
    bar.Sort.SortFields.Clear
    bar.Sort.SortFields.Add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With bar.Sort
        .SetRange Range("A1:B" & rVal)
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Me.ListBox1.Clear
    
    Me.ListBox1.AddItem "Bin"
    Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Barcode"
    
    For roCnt = 2 To rVal
    
        With Me.ListBox1
            .AddItem bar.Cells(roCnt, 1)
            .List(ListBox1.ListCount - 1, 1) = bar.Cells(roCnt, 2)
            
        End With
    Next roCnt
End Sub
