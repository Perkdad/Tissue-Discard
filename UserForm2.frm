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
Private Sub CommandButton1_Click()
'***********
'Update
'***********
    
    
    '******************************
    'See if you can add multiselect
    '******************************
        
    Dim selItm As Long
    Dim selItmAcc As Long
    Dim selBin As String
    Dim Bins As Worksheet
    Set Bins = Worksheets("Bins")
    Dim rVal As Integer
    
    For selItm = LBound(Me.ListBox1.List) To UBound(Me.ListBox1.List)
        If Me.ListBox1.Selected(selItm) = True Then '****it is selected***
            
            selBin = Me.ListBox1.List(selItm, 0) '***Sets the New Bin number***
            
            For selItmAcc = LBound(Frame1.ListBox1.List) To UBound(Frame1.ListBox1.List) '***Reference specimen selected on Frame1***
                If Frame1.ListBox1.Selected(selItmAcc) = True Then
                
                '***************************************************
                                
                rVal = Frame1.ListBox1.List(selItmAcc, 4)
                    
                        '***Change bin to selected bin***
                        Bins.Cells(rVal, 1).Value = selBin
                        '***Change date to current date***
                       Bins.Cells(rVal, 7).Value = Date
                       
                '******************************************************
                End If
            Next selItmAcc
            
        End If
    Next selItm
    Unload Me
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
'***************
'Initialize Form - Done!
'***************
    
    'UserForm1.TextBox1.SetFocus
    
    Dim bar As Worksheet
    'Dim Bins As Worksheet
    Set bar = Worksheets("Barcode")
    'Set Bins = Worksheets("Bins")
    Dim rVal As Integer
    Dim roCnt As Integer
    rVal = bar.Cells(Rows.count, 1).End(xlUp).Row
    
    'bar.Sort.SortFields.Clear
    'bar.Sort.SortFields.Add Key:=Range("B1"), _
        'SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    'With bar.Sort
        '.SetRange Range("A1:B" & rVal)
        '.header = xlYes
        '.MatchCase = False
        '.Orientation = xlTopToBottom
        '.SortMethod = xlPinYin
        '.Apply
    'End With
    
    Me.ListBox1.Clear
    
    Me.ListBox1.AddItem "Bin"
    'Me.ListBox1.List(ListBox1.ListCount - 1, 1) = "Barcode"
    
    For roCnt = 2 To rVal
    
        With Me.ListBox1
            .AddItem bar.Cells(roCnt, 1)
            '.List(ListBox1.ListCount - 1, 1) = bar.Cells(roCnt, 2)
            
        End With
    Next roCnt
    
End Sub
