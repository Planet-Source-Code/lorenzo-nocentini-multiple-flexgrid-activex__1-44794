VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl MFlex 
   AccessKeys      =   "FlexNum"
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   ScaleHeight     =   9180
   ScaleWidth      =   7215
   ToolboxBitmap   =   "MFlex.ctx":0000
   Begin VB.PictureBox pctBack 
      BackColor       =   &H00D05C28&
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   100
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6420
      TabIndex        =   9
      Top             =   7320
      Visible         =   0   'False
      Width           =   6420
      Begin VB.Label lblCInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   100
         Left            =   840
         TabIndex        =   10
         Top             =   120
         Width           =   45
      End
      Begin VB.Image imgAdd 
         Height          =   285
         Index           =   100
         Left            =   6060
         Tag             =   "0"
         Top             =   45
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgAdd_HI 
         Height          =   285
         Index           =   100
         Left            =   6060
         Tag             =   "0"
         Top             =   45
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgHide_HI 
         Height          =   285
         Index           =   100
         Left            =   6060
         Tag             =   "0"
         Top             =   45
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Image imgPicTitle 
         Height          =   240
         Index           =   100
         Left            =   240
         Top             =   67
         Width           =   240
      End
      Begin VB.Line lnBack 
         BorderColor     =   &H00FFFFFF&
         Index           =   100
         Visible         =   0   'False
         X1              =   0
         X2              =   6465
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Image imgHide 
         Height          =   285
         Index           =   100
         Left            =   6060
         Top             =   45
         Width           =   285
      End
   End
   Begin VB.PictureBox pctHand 
      Height          =   345
      Left            =   6570
      Picture         =   "MFlex.ctx":0312
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctAdd_Hi 
      Height          =   345
      Left            =   6570
      Picture         =   "MFlex.ctx":0464
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctAdd_Press 
      Height          =   345
      Left            =   6570
      Picture         =   "MFlex.ctx":0922
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctAdd 
      Height          =   345
      Left            =   6570
      Picture         =   "MFlex.ctx":0DE0
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctHide_Hi 
      Height          =   345
      Left            =   6570
      Picture         =   "MFlex.ctx":129E
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctHide_Press 
      Height          =   345
      Left            =   6570
      Picture         =   "MFlex.ctx":175C
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox pctHide 
      Height          =   345
      Left            =   6570
      Picture         =   "MFlex.ctx":1C1A
      ScaleHeight     =   285
      ScaleWidth      =   270
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.VScrollBar vsbCInfo 
      CausesValidation=   0   'False
      Height          =   3640
      LargeChange     =   1800
      Left            =   6430
      SmallChange     =   300
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Value           =   15
      Width           =   255
   End
   Begin VB.Frame fraBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6465
   End
   Begin MSFlexGridLib.MSFlexGrid fxgCInfo 
      Height          =   615
      Index           =   100
      Left            =   0
      TabIndex        =   11
      Tag             =   "0"
      Top             =   7680
      Visible         =   0   'False
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1085
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   100
      BackColor       =   16774384
      ForeColor       =   0
      BackColorFixed  =   15244408
      ForeColorFixed  =   16777215
      BackColorSel    =   16764603
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLinesFixed  =   0
      ScrollBars      =   2
      AllowUserResizing=   2
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   9
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   8
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   7
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   6
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   5
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   4
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   3
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   2
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   1
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Index           =   0
      Left            =   6600
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "MFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ********************************************************************************
' ********************************************************************************
' **                                                                            **
' **   MULTIPLE FLEXGRID ACTIVEX                                                **
' **                                                                            **
' **   - Made by Nocentini Lorenzo                                              **
' **   - Many Thanks to João Chamorra for help and suggestions                  **
' **                                                                            **
' ********************************************************************************
' ********************************************************************************
'
' How to use the "Multiple Flexgrid"
'
'
' PROPERTIES:
'
' - TotalFlexes       Total number of Tables. The tables start from index 0,
'                     so if you put 6 tables they will go from 0 to 5
'
' - Titles            It represents the titles of the categories which
'                     contain the tables. They must be separated by a ";"
'                     < NOTE: if you know a better way to handle these
'                     strings in the ActiveX please tell me >
'
' - Rows              Total number of rows of a specified flexgrid
'
' - bcolor            Background color
'
' - Grids             Manually control the flexgrids inside the activex
'
'
' METHODS:
'
' - InitializeGrid    Used to set the number of colums and rows of a single
'                     flexgrid
'
' - InitSingleCol     Used to define each column widht and alignement of
'                     a single flexgrid
'
' - TextMatrix        Similar to the MSFlexGrid TextMatrix method (with one
'                     more parameter to define the index of the flexgrid)
'                     Use this to put text everywhere in the tables
'
' - GetText           Get text from Flexgrid
'
' - PutTitle          Set a single category title
'
' - AddRow            Add a row in a specified table
'
' - RefreshDB         Update the grids with database values
'
'
' EVENTS:
'
' - FlexMouseDown     Mouse down on a flexgrid

Option Explicit

Private TotFlex As Integer
Dim FlexTitles(10) As String
Dim imlista As ImageList

Event FlexMouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Function Row(Index As Integer) As Integer
    If Index < TotFlex Then Row = fxgCInfo(Index).Row
End Function

Function Rows(Index As Integer) As Integer
    If Index < TotFlex Then Rows = fxgCInfo(Index).Rows
End Function

Function Col(Index As Integer) As Integer
    If Index < TotFlex Then Col = fxgCInfo(Index).Col
End Function

' Set colums width and text alignement
Public Sub InitSingleCol(Index As Integer, Col, Width, Alignement As AlignmentSettings)
    On Error Resume Next
    If Index < TotFlex Then
        fxgCInfo(Index).Col = Col
        fxgCInfo(Index).ColAlignment(Col) = Alignement
        fxgCInfo(Index).ColWidth(Col) = Width
    End If
End Sub

' Set a flexgrid height
Public Sub FlexHeight(Index As Integer, Height As Long)
    On Error Resume Next
    If Index < TotFlex Then
        fxgCInfo(Index).Height = Height
    End If
End Sub

' Initialize the Flexgrids
Public Sub InitializeGrid(Cols As Integer, Rows As Integer, Index)
    If Index < TotFlex Then
        fxgCInfo(Index).Cols = Cols
        fxgCInfo(Index).Rows = Rows
    End If
End Sub

' Write text into the flexgrid
Public Sub TextMatrix(Text As String, Col As Integer, Row As Integer, Index As Integer)
    On Error Resume Next
    If Index < TotFlex Then fxgCInfo(Index).TextMatrix(Row, Col) = Text
End Sub

' Read text from the flexgrid
Public Function GetText(Col As Integer, Row As Integer, Index As Integer)
    If Index < TotFlex Then GetText = fxgCInfo(Index).TextMatrix(Row, Col)
End Function

Private Sub fxgCInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent FlexMouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub fxgCInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Call the same routine to hide the highlighted images
    UserControl_MouseMove 0, 0, 0, 0
End Sub

' Shows the table
Private Sub imgAdd_HI_Click(Index As Integer)
    imgAdd_HI(Index).Tag = 0
    imgHide_HI(Index).Tag = 1
    imgAdd(Index).Visible = False
    imgAdd_HI(Index).Visible = False
    imgHide(Index).Visible = True
    imgHide_HI(Index).Visible = True
    imgHide_HI(Index).ZOrder
    fxgCInfo(Index).Visible = True      ' Shows the flexgrid
    fxgCInfo(Index).Tag = 1             ' Remember that the flexgrid is shown
    lnBack(Index).Visible = False       ' Hide the divisory line
    ArrangeFlexes                       ' Re-Arrange the flexgrids
End Sub

Private Sub imgAdd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When the mouse is over an image, I show the highlight one
    ' If the Tag property is "1" then it means that the highlight picture
    ' is already visible, so there is no need to show it again (so it does not flicker)
    If imgAdd_HI(Index).Tag = 0 Then
        imgAdd_HI(Index).Tag = 1
        imgAdd_HI(Index).Visible = True
        lblCInfo(Index).ForeColor = &HFFE3D9
        imgAdd_HI(Index).ZOrder
    End If
End Sub

' Hides the table
Private Sub imgHide_HI_Click(Index As Integer)
    imgAdd_HI(Index).Tag = 1
    imgHide_HI(Index).Tag = 0
    imgHide(Index).Visible = False
    imgHide_HI(Index).Visible = False
    imgAdd(Index).Visible = True
    imgAdd_HI(Index).Visible = True
    imgAdd_HI(Index).ZOrder
    fxgCInfo(Index).Visible = False     ' Hides the flexgrid
    fxgCInfo(Index).Tag = 0             ' Remember that the flexgrid is hidden
    lnBack(Index).Visible = True        ' Show a divisor line
    ArrangeFlexes                       ' Re-Arrange the flexgrids
End Sub

Private Sub imgHide_HI_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When an image is pressed I show the proper image
    imgHide_HI(Index).Picture = pctHide_Press.Picture
End Sub

Private Sub imgHide_HI_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' The mouse button is up: remove the pressed image
    imgHide_HI(Index).Picture = pctHide_Hi.Picture
End Sub

Private Sub imgAdd_HI_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When an image is pressed I show the proper image
    imgAdd_HI(Index).Picture = pctAdd_Press.Picture
End Sub

Private Sub imgAdd_HI_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' The mouse button is up: remove the pressed image
    imgAdd_HI(Index).Picture = pctAdd_Hi.Picture
End Sub

Private Sub imgHide_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When the mouse is over an image, I show the highlight one
    ' If the Tag property is "1" then it means that the highlight picture
    ' is already visible, so there is no need to show it again (so it does not flicker)
    If imgHide_HI(Index).Tag = 0 Then
        imgHide_HI(Index).Tag = 1
        imgHide_HI(Index).Visible = True
        lblCInfo(Index).ForeColor = &HFFE3D9
        imgHide_HI(Index).ZOrder
    End If
End Sub

Private Sub lblCInfo_Click(Index As Integer)
    ' When the user clicks on the label it happens the same thing when he
    ' clicks on the hide/show images
    If fxgCInfo(Index).Tag = 1 Then
        imgHide_HI_Click Index
    Else
        imgAdd_HI_Click Index
    End If
End Sub

Private Sub lblCInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' When the user moves the mouse on the label it happens the same thing when he
    ' moves the mouse on the hide/show images
    If fxgCInfo(Index).Tag = 1 Then
        imgHide_MouseMove Index, 0, 0, 0, 0
    Else
        imgAdd_MouseMove Index, 0, 0, 0, 0
    End If
End Sub

Private Sub pctBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Call the same routine to hide the highlighted images
    UserControl_MouseMove 0, 0, 0, 0
End Sub

' This function changes the flexgrids position when a flexgrid is explosed
' or reduced, or when the form size is changed
' It also set the scrollbar min. and max. values
Private Sub ArrangeFlexes()
Dim i As Integer
Dim j As Integer
Dim TotTop
    TotTop = 0
    For i = 0 To TotFlex - 1
        pctBack(i).Top = TotTop
        fxgCInfo(i).Top = pctBack(i).Top + pctBack(i).Height
        TotTop = TotTop + pctBack(i).Height
        If fxgCInfo(i).Tag = 1 Then
            TotTop = TotTop + fxgCInfo(i).Height
        End If
    Next i
    fraBack.Height = TotTop
    ' If the height of the flexgrids is higher than the form height then
    ' the scrollbar will be enabled
    If TotTop > UserControl.Height Then
        vsbCInfo.Visible = True
        vsbCInfo.Max = TotTop - UserControl.Height
        vsbCInfo.value = 0
    Else
        vsbCInfo.Visible = False
        fraBack.Top = 0
    End If
End Sub

Private Sub UserControl_Initialize()
Dim i As Integer
    On Error Resume Next
    For i = 0 To TotFlex - 1
        ' Creates controls dynamically (Number is user defined)
        Load pctBack(i)
        Load imgPicTitle(i)
        Load lblCInfo(i)
        Load imgHide(i)
        Load imgAdd(i)
        Load imgHide_HI(i)
        Load imgAdd_HI(i)
        Load fxgCInfo(i)
        Load lnBack(i)
        
        ' Controls Container is back frame
        Set fxgCInfo(i).Container = fraBack
        Set pctBack(i).Container = fraBack
        
        ' Default position (will be adjusted later in "ArrangeFlexes"
        fxgCInfo(i).Left = 0
        pctBack(i).Left = 0
        pctBack(i).Top = 0
        
        ' Make visible
        pctBack(i).Visible = True
        fxgCInfo(i).Visible = True
        lblCInfo(i).Visible = True
        imgPicTitle(i).Visible = True
        
        ' Set Container control (pctBack)
        Set imgPicTitle(i).Container = pctBack(i)
        Set lblCInfo(i).Container = pctBack(i)
        Set imgHide(i).Container = pctBack(i)
        Set imgAdd(i).Container = pctBack(i)
        Set imgHide_HI(i).Container = pctBack(i)
        Set imgAdd_HI(i).Container = pctBack(i)
        Set lnBack(i).Container = pctBack(i)
        
        ' Set positions inside control (pctBack)
        imgPicTitle(i).Left = 240
        imgPicTitle(i).Top = 67
        lblCInfo(i).Left = 840
        lblCInfo(i).Top = 90
        imgHide(i).Left = 6060
        imgHide(i).Top = 45
        imgAdd(i).Left = 6060
        imgAdd(i).Top = 45
        imgHide_HI(i).Left = 6060
        imgHide_HI(i).Top = 45
        imgAdd_HI(i).Left = 6060
        imgAdd_HI(i).Top = 45
        lnBack(i).X1 = 0
        lnBack(i).X2 = pctBack(i).Width
        lnBack(i).Y1 = 360
        lnBack(i).Y2 = 360
    
        ' Load command pictures
        imgHide(i).Picture = pctHide.Picture
        imgAdd(i).Picture = pctAdd.Picture
        imgHide_HI(i).Picture = pctHide_Hi.Picture
        imgAdd_HI(i).Picture = pctAdd_Hi.Picture
        imgHide_HI(i).MouseIcon = pctHand.Picture
        imgHide_HI(i).MousePointer = 99
        imgAdd_HI(i).MouseIcon = pctHand.Picture
        imgAdd_HI(i).MousePointer = 99
        lblCInfo(i).MouseIcon = pctHand.Picture
        lblCInfo(i).MousePointer = 99
        If fxgCInfo(i).Tag = 0 Then
            imgHide(i).Visible = False
            imgAdd(i).Visible = True
            lnBack(i).Visible = True
        Else
            imgHide(i).Visible = True
            imgAdd(i).Visible = False
            lnBack(i).Visible = False
        End If
        lblCInfo(i).Caption = "Title " & i
    Next i
End Sub

Private Sub UserControl_InitProperties()
Dim i As Integer
    TotFlex = 6
    UserControl_Initialize
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Index
    ' Hide the hilighted images
    For Index = 0 To TotFlex - 1
        ' If the Tag property is "0" then it means that the highlight picture
        ' is not visible, so there is no need to hide it again
        If imgAdd_HI(Index).Tag = 1 Then
            imgAdd_HI(Index).Tag = 0
            imgAdd_HI(Index).Visible = False
            lblCInfo(Index).ForeColor = &HFFFFFF
        End If
        If imgHide_HI(Index).Tag = 1 Then
            imgHide_HI(Index).Tag = 0
            imgHide_HI(Index).Visible = False
            lblCInfo(Index).ForeColor = &HFFFFFF
        End If
    Next Index
End Sub

Private Sub UserControl_Resize()
Dim i As Integer

    On Error Resume Next
    
    ' Change Height/Width of the controls according to the form size
    vsbCInfo.Height = UserControl.Height
    vsbCInfo.Left = UserControl.Width - 380
    fraBack.Width = UserControl.Width - 345
    For i = 0 To TotFlex - 1
        fxgCInfo(i).Width = UserControl.Width - 360
        'fxgCInfo(i).ColWidth(2) = fxgCInfo(i).Width - 2000
        pctBack(i).Width = UserControl.Width - 360
        lnBack(i).X2 = pctBack(i).Width
        imgHide(i).Left = pctBack(i).Width - 360
        imgHide_HI(i).Left = pctBack(i).Width - 360
        imgAdd(i).Left = pctBack(i).Width - 360
        imgAdd_HI(i).Left = pctBack(i).Width - 360
        imgHide(6).Left = pctBack(i).Width - 360
    Next i
    
    ArrangeFlexes
    
End Sub

Public Sub PutTitle(Index As Integer, Text As String)
    If Index < TotFlex Then lblCInfo(Index).Caption = Text
End Sub

Private Sub vsbCInfo_Change()
    ' All the flexgrids are placed on a frame, so to scroll them i just change
    ' the frame height
    fraBack.Top = -vsbCInfo.value
End Sub

Private Sub vsbCInfo_Scroll()
    vsbCInfo_Change
End Sub

Public Sub AddRow(Index As Integer)
       
    ' Set the proper height (so it does not automatically show the flexgrid srollbar)
    fxgCInfo(Index).Height = (240 * (fxgCInfo(Index).Rows + 1)) + 15
    
    ' Creates a new row
    fxgCInfo(Index).Rows = fxgCInfo(Index).Rows + 1
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim dummy As String
Dim i As Integer
    On Error Resume Next
    TotFlex = PropBag.ReadProperty("TotalFlexes", 6)
    UserControl_Initialize
    UserControl_Resize
    dummy = PropBag.ReadProperty("Titles", "")
    AnalyzeTitles dummy
    For i = 0 To TotFlex - 1
        imgPicTitle(i).Picture = PropBag.ReadProperty("Pic" & i, vbEmpty)
        imgTitle(i).Picture = PropBag.ReadProperty("Pic" & i, vbEmpty)
    Next i
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim dummy As String
Dim i As Integer
    On Error Resume Next
    dummy = ""
    For i = 0 To TotFlex - 1
        dummy = dummy & lblCInfo(i).Caption & ";"
    Next i
    Call PropBag.WriteProperty("TotalFlexes", TotFlex, 6)
    Call PropBag.WriteProperty("Titles", dummy, "")
End Sub

Sub SetTitles()
Dim i As Integer
    For i = 0 To TotFlex - 1
        lblCInfo(i).Caption = FlexTitles(i)
    Next i
End Sub

Private Sub AnalyzeTitles(ByVal New_Value As String)
Dim SPos As Integer
Dim NextPos As Integer
Dim i As Integer
Dim dummy As String
    On Error Resume Next
    If Right(New_Value, 1) <> ";" And New_Value <> "" Then New_Value = New_Value & ";"
    NextPos = 1
    i = 0
    For i = 0 To 5
        SPos = InStr(New_Value, ";")
        If SPos > 0 Then
            dummy = Left(New_Value, SPos - 1)
        Else
            dummy = New_Value
        End If
        
        FlexTitles(i) = dummy
        New_Value = Right(New_Value, Len(New_Value) - SPos)
    Next i
    SetTitles
End Sub

Public Property Get TotalFlexes() As Integer
Attribute TotalFlexes.VB_Description = "Number of Tables"
    TotalFlexes = TotFlex
End Property

Public Property Let TotalFlexes(ByVal New_Value As Integer)
    If New_Value < 1 Or New_Value > 10 Then Exit Property
    TotFlex = New_Value
    PropertyChanged "TotalFlexes"
    UserControl_Initialize
    UserControl_Resize
End Property

Public Property Get Titles() As String
Attribute Titles.VB_Description = "Enter each table title separated by "";"""
Dim i As Integer
    On Error Resume Next
    Titles = ""
    For i = 0 To TotFlex - 1
        Titles = Titles & lblCInfo(i).Caption & ";"
    Next i
End Property

Public Property Let Titles(ByVal New_Value As String)
    On Error Resume Next
    AnalyzeTitles New_Value
    PropertyChanged "Titles"
End Property

Public Property Get Grids(Index) As MSFlexGrid
    ' This way a can manually control any flex grid adding data as I wish
    Set Grids = fxgCInfo(Index)
End Property

Public Property Let bcolor(ByVal cor As Long)
    ' To control the back color of the mflex
    UserControl.BackColor = cor
End Property

Public Property Get labels(Index) As Label
    ' To control the labels….forecolor etc…
    ' For example imaging that you want to put a different color in the label of an flex grid …
    Set labels = lblCInfo(Index)
End Property

' Associate an imagelist with the title icons
Public Property Let ilist(ByRef elista As ImageList)
Dim i As Integer
    On Error Resume Next
    Set imlista = elista
    For i = 0 To TotFlex - 1
        imgPicTitle(i).Picture = imlista.ListImages(i + 1).Picture
        If Err.Number > 0 Then
            imgPicTitle(i).Picture = LoadPicture()
            Err.Clear
        End If
    Next i
End Property

