VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Bin Editor"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4920
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************
'Program:   Discard Tissue Workbook
'Form:      UserForm1 Bin Editor
'Author:    Aaron Perkins
'Date:      10/28/2025
'Version:   1.4.2
'**********************************

Private Sub CommandButton1_Click()
'======
'Update
'======
    
    'Declarations
    Dim barCodeSheet As Worksheet
    Dim binsSheet As Worksheet
    Dim userEntryBin As String
    Dim userBarCode As String
    Dim rowIndex As Integer
    Dim updateMessage As Integer
    
    'Initialization
    Set barCodeSheet = ThisWorkbook.Worksheets("Barcode")
    Set binsSheet = ThisWorkbook.Worksheets("Bins")
    
    userEntryBin = UCase(TextBox1.Value)
    userBarCode = UCase(TextBox2.Value)
    
    '=========================================================================================================
    
    '=================================
    'Check for a bin number in textbox
    '=================================
    
    If userEntryBin = "" Then
        MsgBox "Please Enter a Bin Number to proceed"
        TextBox2.Value = ""
        Call UserForm_Initialize
        Exit Sub
    End If
    
    '============================
    'Check for a valid Bin Number
    '============================
    
    If Not UCase(userEntryBin) Like "[A-Z]##" Then
        MsgBox "You must enter a Bin number consisting of a Letter and 2 numerical Digits"
        TextBox1.Value = ""
        TextBox2.Value = ""
        Call UserForm_Initialize
        Exit Sub
    End If
    
    '===============================================
    'Check to see if the Bin is active with specimen
    '===============================================
    
    rowIndex = 1
    
    'Loop through rows until empty row found
    Do While binsSheet.Cells(rowIndex, 1) <> ""
        If UCase(binsSheet.Cells(rowIndex, 1).Value) = UCase(userEntryBin) Then
            MsgBox "This Bin is active and may not be edited at this time"
            TextBox1.Value = ""
            TextBox2.Value = ""
            Call UserForm_Initialize
            Exit Sub
        End If
        rowIndex = rowIndex + 1
    Loop
    
    '---------------------------------------------------------------------------------------------------------
    
    '=======================================
    'Checks if the barcode is already in use
    '=======================================
    
    '--------------------------------
    'The barcode is assigned to a bin
    '--------------------------------
    
    rowIndex = 1
    
    'Loop through rows until empty row found
    Do While barCodeSheet.Cells(rowIndex, 2) <> ""
        If UCase(barCodeSheet.Cells(rowIndex, 2).Value) = UCase(userBarCode) Then
            MsgBox "This Barcode is already assigned to another bin"
            TextBox2.Value = ""
            Call UserForm_Initialize
            Exit Sub
        End If
        rowIndex = rowIndex + 1
    Loop
    
    '--------------------------------------------
    'The barcode is currently a specimen in a bin
    '--------------------------------------------
    
    rowIndex = 1
    
    'Loop through rows until empty row found
    Do While binsSheet.Cells(rowIndex, 2) <> ""
        
        'The barcode matches a value in the binsSheet
        If UCase(binsSheet.Cells(rowIndex, 2).Value) = UCase(userBarCode) Then
            MsgBox "This is an active Barcode for a specimen in Bin " & binsSheet.Cells(rowIndex, 1)
            TextBox1.Value = ""
            TextBox2.Value = ""
            Call UserForm_Initialize
            Exit Sub
        End If
        rowIndex = rowIndex + 1
    Loop
    
    '---------------------------------------------------------------------------------------------------------
    
    '=====================
    'Bin and Barcode entry
    '=====================
    
    'Bin and barcode values entered
    If userEntryBin <> "" And userBarCode <> "" Then
        
        '-----------------------
        'Check for valid barcode
        '-----------------------
        
        'Search for ";" in Ross barcode to validate
        If InStr(userBarCode, ";") = 0 Then
            MsgBox "Please enter a valid barcode"
            TextBox2.Value = ""
            Call UserForm_Initialize
            Exit Sub
        End If
        
        
        '------------------------------------------------
        'Check if the Bin is currently assigned a barcode
        '------------------------------------------------
        
        rowIndex = 2
        
        'Loop through rows until empty row found
        Do While barCodeSheet.Cells(rowIndex, 1) <> ""
            
            'User entry matches a value in the barCodeSheet
            If UCase(barCodeSheet.Cells(rowIndex, 1).Value) = UCase(userEntryBin) Then
                
                'Message prompt to update Barcode
                updateMessage = MsgBox("This Bin currently has a Barcode. Would you like to update the Barcode?", vbYesNo)
                
                'No response does not update and exits sub
                If updateMessage = vbNo Then
                    TextBox1.Value = ""
                    TextBox2.Value = ""
                    Call UserForm_Initialize
                    Exit Sub
                Else
                    GoTo loopExit
                End If
            End If
            rowIndex = rowIndex + 1
        Loop
        
loopExit:
        
        '---------------------------------------------------------------
        'Update bin barcode if in list or append bin and barcode to list
        '---------------------------------------------------------------
        
        'update
        If barCodeSheet.Cells(rowIndex, 1) <> "" Then
            barCodeSheet.Cells(rowIndex, 2).Value = UCase(userBarCode)
        
        'append
        Else
            barCodeSheet.Cells(rowIndex, 2).Value = UCase(userBarCode)
            barCodeSheet.Cells(rowIndex, 1).Value = UCase(userEntryBin)
        End If
        
        'Reset text boxes
        TextBox1.Value = ""
        TextBox2.Value = ""
        Call UserForm_Initialize
        Exit Sub
        
    End If
    
    '---------------------------------------------------------------------------------------------------------
        
    '=======================
    'Remove bin from service
    '=======================
    
    'Bin number entered with no barcode
    If userEntryBin <> "" And userBarCode = "" Then
        
        rowIndex = 2
        
        'Loop through rows until empty row found
        Do While barCodeSheet.Cells(rowIndex, 1) <> ""
            
            'Bin entry matches bin in barCodeSheet
            If barCodeSheet.Cells(rowIndex, 1) = UCase(userEntryBin) Then
                
                'Message prompt to remove bin from service
                updateMessage = MsgBox("Are you sure you would like to remove this Bin from service and dissociate the Barcode?", vbYesNo)
                
                'Yes response deletes bin and details out of the barCodeSheet
                If updateMessage = vbYes Then
                    barCodeSheet.Range("A" & rowIndex & ":B" & rowIndex).Delete Shift:=xlUp
                    TextBox1.Value = ""
                    TextBox2.Value = ""
                    Call UserForm_Initialize
                    Exit Sub
                End If
                
                'No response exits sub
                Call UserForm_Initialize
                Exit Sub
            End If
            
            rowIndex = rowIndex + 1
            
        Loop
        
        'Unrecognized Bin entry
        MsgBox "You have not entered an active Bin number"
        TextBox1.Value = ""
        TextBox2.Value = ""
        Call UserForm_Initialize
    End If
    
End Sub


Private Sub ListBox1_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Initialize()

'===================
'Initialize UserForm
'===================
    
    'Declarations
    Dim barCodeSheet As Worksheet
    Dim binsSheet As Worksheet
    Dim rowIndex As Integer
    Dim rowCount As Integer
    
    'Initializations
    Set barCodeSheet = ThisWorkbook.Worksheets("Barcode")
    Set binsSheet = ThisWorkbook.Worksheets("Bins")
    rowIndex = barCodeSheet.Cells(Rows.count, 1).End(xlUp).Row
    
    '=========================================================================================================
    
    'Set focus on TextBox1 (Bin Number)
    UserForm1.TextBox1.SetFocus
    
    'Sort barCodesheet
    barCodeSheet.Sort.SortFields.Clear
    barCodeSheet.Sort.SortFields.Add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With barCodeSheet.Sort
        .SetRange Range("A1:B" & rowIndex)
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Clear ListBox1 of any previous data
    Me.ListBox1.Clear
    
    'Add headers to ListBox1
    Me.ListBox1.AddItem "Bin"
    Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Barcode"
    
    'Add items to ListBox1
    For rowCount = 2 To rowIndex
        With Me.ListBox1
            .AddItem barCodeSheet.Cells(rowCount, 1)
            .List(ListBox1.ListCount - 1, 1) = barCodeSheet.Cells(rowCount, 2)
            
        End With
    Next rowCount
End Sub
