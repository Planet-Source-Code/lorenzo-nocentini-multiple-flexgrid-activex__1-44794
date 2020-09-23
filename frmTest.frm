VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Test"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3480
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":11F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":158C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdConnectMF 
      Caption         =   "Load data in to Multiple Flexgrid #0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   3135
   End
   Begin Test.MFlex MFlex1 
      Height          =   5640
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   9948
      Titles          =   "Title 0;Title 1;Title 2;Title 3;Title 4;Title 5;"
   End
   Begin VB.Label Label1 
      Caption         =   "Select cell with left mouse click and to change use Mouse Rigth Clik"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   480
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   9405
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As New ADODB.Connection      ' Database Connection
Dim rspecas As New ADODB.Recordset  ' Database RecordSet

' Load a database in the MULTIPLE Flexgrid (ActiveX) in the first grid
Private Sub cmdConnectMF_Click()
    On Error Resume Next
    Dim llocal As String: llocal = App.Path & "\Test.mdb"
    db.Close
    On Error GoTo 0
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + llocal + ";Mode=ReadWrite;Persist Security Info=False"
    
    rspecas.CursorType = adOpenKeyset
    rspecas.CursorLocation = adUseClient
    rspecas.LockType = adLockPessimistic
    rspecas.ActiveConnection = db
    rspecas.Index = "PrimaryKey"
    rspecas.Source = "SELECT * FROM COMUNI"
    rspecas.Open
    'MsgBox rspecas.RecordCount
    rspecas.MoveFirst
        
    Dim lTimer As Long
    
    Screen.MousePointer = vbHourglass
    
    MFlex1.Grids(0).Refresh
    lTimer = Timer
    
    MFlex1.Grids(0).Visible = False
    MFlex1.Grids(0).Rows = 1
    MFlex1.Grids(0).Rows = rspecas.RecordCount + 1
    MFlex1.Grids(0).Cols = rspecas.Fields.Count - 1
    MFlex1.Grids(0).Row = 1
    MFlex1.Grids(0).Col = 0
    MFlex1.Grids(0).RowSel = MFlex1.Grids(0).Rows - 1
    MFlex1.Grids(0).ColSel = MFlex1.Grids(0).Cols - 1
    MFlex1.Grids(0).Clip = rspecas.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    
    MFlex1.Grids(0).Row = 1
    MFlex1.Grids(0).Visible = True
    
     ' Write flexgrid titles
    MFlex1.TextMatrix "ID", 0, 0, 0
    MFlex1.TextMatrix "Code", 1, 0, 0
    MFlex1.TextMatrix "Name", 2, 0, 0
    MFlex1.TextMatrix "CAP", 3, 0, 0
    MFlex1.TextMatrix "Province", 4, 0, 0
    MFlex1.TextMatrix "Belfiore", 5, 0, 0
    MFlex1.TextMatrix "Region", 6, 0, 0
    MFlex1.TextMatrix "Asl", 7, 0, 0
    MFlex1.TextMatrix "District", 8, 0, 0
    
    ' Update the grid #0 Title
    MFlex1.PutTitle 0, "Database Values"
    
    ' Set a default heigth for grid #0
    MFlex1.FlexHeight 0, 3000
    
    Label1.Visible = True
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Execution time: " & Timer - lTimer & " sec." & vbCr & "of " & MFlex1.Grids(0).Rows - 1 & " record"
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
    
    ' Define the picturebox for the icons
    MFlex1.ilist = ImageList1
    
    ' Write categories titles
    MFlex1.PutTitle 0, "Status"
    MFlex1.PutTitle 1, "Options"
    MFlex1.PutTitle 2, "Mode"
    MFlex1.PutTitle 3, "Remote"
    MFlex1.PutTitle 4, "Battery"
    MFlex1.PutTitle 5, "Calibration"
        
    ' Initialize flexgrids number of rows/columns
    For i = 0 To 5
        ' In this example I set all flexgrids with 3 columns and 1 row
        MFlex1.InitializeGrid 9, 1, i
    Next i
    
    ' Initialize each column width and alignement
    For i = 0 To 5
        ' Colum 0: width = 900, alignement = center-center
        MFlex1.InitSingleCol i, 0, 600, flexAlignCenterCenter
        
        ' Colum 1: width = 900, alignement = center-center
        MFlex1.InitSingleCol i, 1, 900, flexAlignCenterCenter
        
        ' Colum 2: width = 2400, alignement = center-center
        MFlex1.InitSingleCol i, 2, 2400, flexAlignCenterCenter
        
        ' Colum 3: width = 900, alignement = center-center
        MFlex1.InitSingleCol i, 3, 900, flexAlignCenterCenter
        
        ' Colum 4: width = 900, alignement = center-center
        MFlex1.InitSingleCol i, 4, 900, flexAlignCenterCenter
        
        ' Colum 5: width = 900, alignement = center-center
        MFlex1.InitSingleCol i, 5, 900, flexAlignCenterCenter
        
        ' Colum 6: width = 900, alignement = center-center
        MFlex1.InitSingleCol i, 6, 900, flexAlignCenterCenter
        
        ' Colum 7: width = 900, alignement = center-center
        MFlex1.InitSingleCol i, 7, 900, flexAlignCenterCenter
        
        ' Colum 8: width = 900, alignement = center-center
        MFlex1.InitSingleCol i, 8, 600, flexAlignCenterCenter
    Next i
    
    ' Write flexgrid titles
    For i = 0 To 5
        MFlex1.TextMatrix "Value", 0, 0, i
        MFlex1.TextMatrix "Range", 1, 0, i
        MFlex1.TextMatrix "Description", 2, 0, i
    Next i
    
    ' Add some fake values into the flexgrids
    Randomize                               ' Initialize random numbers
    For j = 0 To MFlex1.TotalFlexes - 1     ' For all flexgrids
        For i = 0 To Int((10 * Rnd) + 1)    ' Add random numbers
            AddRowInFlex Int((50 * Rnd) + 1), "1.50", "Description " & i & " " & j, j
        Next i
    Next j
    
End Sub

' Manually add a row
Sub AddRowInFlex(ByVal value As Variant, ByVal Range As String, ByVal Description As String, Index As Integer)
Dim CurrRow As Integer
       
    ' Add a new row
    MFlex1.AddRow Index
    
    ' Get the last row (the one I added)
    CurrRow = MFlex1.Rows(Index) - 1
    
    ' Put my values
    MFlex1.TextMatrix CStr(value), 0, CurrRow, Index
    MFlex1.TextMatrix CStr(Range), 1, CurrRow, Index
    MFlex1.TextMatrix CStr(Description), 2, CurrRow, Index

End Sub

' Here I just resize the MFlex according to the form size (if you want)
Private Sub Form_Resize()
Dim i As Integer
    On Error Resume Next
    'MFlex1.Width = Me.Width - 45
    'MFlex1.Height = Me.Height - 600
    'For i = 0 To 5
        ' I enlarge the column "Description" according to the form width
        'MFlex1.InitSingleCol i, 2, Me.Width - 2185, flexAlignCenterCenter
    'Next i
End Sub

Private Sub MFlex1_FlexMouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Dim strOld As String
        Dim strNew As String
        Dim a As Variant
    
        If MFlex1.Grids(Index).Col = 0 Then Exit Sub    ' Primary Key cannot be modified
    
        strOld = MFlex1.Grids(Index).TextMatrix(MFlex1.Grids(Index).Row, MFlex1.Grids(Index).Col)
        strNew = InputBox("Enter New Value", "New Value", strOld)
    
        If DBModify(strNew, MFlex1.Grids(Index).Row, MFlex1.Grids(Index).Col) Then
            MFlex1.Grids(Index).TextMatrix(MFlex1.Grids(Index).Row, MFlex1.Grids(Index).Col) = strNew
        End If
    End If
End Sub

Public Function DBModify(oque As String, r As Long, c As Long) As Boolean
    On Error GoTo fim    ' the "fim" is just to control the error
   
    rspecas.MoveFirst
    rspecas.Move r - 1
    If rspecas.EOF Or rspecas.BOF Then GoTo fim
    rspecas.Fields(c) = oque
    rspecas.Update
    DBModify = True
    Exit Function
   
fim:
   DBModify = False
End Function

