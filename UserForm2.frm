VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Select New Specimen Bin"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3870
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================
'Program:   Discard Tissue Workbook
'Form:      UserForm2 Select New Specimen Bin
'Author:    Aaron Perkins
'Date:      10/28/2025
'Version:   1.4.2
'============================================

Private Sub CommandButton1_Click()
'=================
'Update Bin Number
'=================
    
    'Declarations
    Dim selectedItem As Integer
    Dim selectedItemFrame1 As Integer
    Dim selectedBin As String
    Dim binsSheet As Worksheet
    Dim rowIndex As Integer
    
    'Initializations
    Set binsSheet = ThisWorkbook.Worksheets("Bins")
    
    '=========================================================================================================
    
    
    For selectedItem = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
        If Me.ListBox1.Selected(selectedItem) = True Then 'it is selected
            
            'Sets the New Bin number to the selected item in the list box
            selectedBin = Me.ListBox1.List(selectedItem, 0)
            
            'Loop through items in Frame1 until selected item is found and then update
            For selectedItemFrame1 = LBound(Frame1.ListBox1.List) To UBound(Frame1.ListBox1.List)
                If Frame1.ListBox1.Selected(selectedItemFrame1) = True Then 'it is selected
                    rowIndex = Frame1.ListBox1.List(selectedItemFrame1, 4)
                    
                    'Change bin to selected bin
                    binsSheet.Cells(rowIndex, 1).Value = selectedBin
                    
                    'Change date to current date
                    binsSheet.Cells(rowIndex, 7).Value = Date
                    
                End If
            Next selectedItemFrame1
        End If
    Next selectedItem
    
    'Close this UserForm2
    Unload Me
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
'===============
'Initialize Form
'===============
    
    'Declarations
    Dim barCodeSheet As Worksheet
    Dim rowIndex As Integer
    Dim rowCount As Integer
    
    'Initializations
    Set barCodeSheet = Worksheets("Barcode")
    rowIndex = barCodeSheet.Cells(Rows.count, 1).End(xlUp).Row
    
    '=========================================================================================================
    
    'Clear all data from ListBox1
    Me.ListBox1.Clear
    
    'Add header
    Me.ListBox1.AddItem "Bin"
    
    'Loop through barCodeSheet and add items to ListBox1
    For rowCount = 2 To rowIndex
        With Me.ListBox1
            .AddItem barCodeSheet.Cells(rowCount, 1)
            
        End With
    Next rowCount
    
End Sub
