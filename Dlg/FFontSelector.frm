VERSION 5.00
Begin VB.Form FFontSelector 
   Caption         =   "Select a Font"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   7455
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox LBFonts 
      Height          =   3885
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CheckBox ChkPitchVariable 
      Caption         =   "Variable"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Value           =   1  'Aktiviert
      Width           =   975
   End
   Begin VB.CheckBox ChkPitchFixed 
      Caption         =   "Fixed"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Value           =   1  'Aktiviert
      Width           =   855
   End
   Begin VB.ComboBox CmbQuality 
      Height          =   345
      Left            =   5760
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   3000
      Width           =   1815
   End
   Begin VB.ComboBox CmbFamily 
      Height          =   345
      Left            =   5760
      TabIndex        =   28
      Text            =   "Combo1"
      Top             =   3720
      Width           =   1815
   End
   Begin VB.ComboBox CmbPitch 
      Height          =   345
      Left            =   5760
      TabIndex        =   26
      Text            =   "Combo1"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox TxtOrient 
      Alignment       =   2  'Zentriert
      Height          =   345
      Left            =   5760
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox TxtEscape 
      Alignment       =   2  'Zentriert
      Height          =   345
      Left            =   5760
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   840
      Width           =   1815
   End
   Begin VB.PictureBox PBColor 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Ausgefüllt
      Height          =   975
      Left            =   6600
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   33
      Top             =   4200
      Width           =   975
   End
   Begin VB.ComboBox CmbWidth 
      Height          =   345
      Left            =   5760
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox CmbPointSize 
      Height          =   345
      Left            =   5760
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox TxtFontName 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3135
   End
   Begin VB.CheckBox ChkStrikeout 
      Caption         =   "Durchgestrichen"
      Height          =   255
      Left            =   3360
      TabIndex        =   31
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CheckBox ChkUnderline 
      Caption         =   "Unterstrichen"
      Height          =   255
      Left            =   3360
      TabIndex        =   30
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CheckBox ChkItalic 
      Caption         =   "Italic, Kursiv"
      Height          =   255
      Left            =   3360
      TabIndex        =   29
      Top             =   4200
      Width           =   1455
   End
   Begin VB.ComboBox CmbClipPrec 
      Height          =   345
      Left            =   5760
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   2640
      Width           =   1815
   End
   Begin VB.ComboBox CmbOutPrec 
      Height          =   345
      Left            =   5760
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.ComboBox CmbCharSet 
      Height          =   345
      Left            =   5760
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.ComboBox CmbWeight 
      Height          =   345
      Left            =   5760
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   36
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   35
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label LblFontsCount 
      Caption         =   "XXX"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Pitch:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin VB.Label LblFontQuality 
      Caption         =   "Schrift-Qualität:"
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label LblFontFamily 
      Caption         =   "Schriftart Familie:"
      Height          =   375
      Left            =   3360
      TabIndex        =   27
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label LblFontPitch 
      Caption         =   "Font Pitch:"
      Height          =   375
      Left            =   3360
      TabIndex        =   25
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label LblOrient 
      Caption         =   "Orientierung (1/10 °):"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label LblEscapement 
      Caption         =   "Escapement (1/10 °):"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label LblColor 
      Caption         =   "Farbe:"
      Height          =   375
      Left            =   5760
      TabIndex        =   32
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label LblWidth 
      Caption         =   "Schriftbreite:"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Größe, Höhe:"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Name der Schriftart:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Clipping-Präzision:"
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Ausgabe-Präzision:"
      Height          =   375
      Left            =   3360
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label LblCharSet 
      Caption         =   "Zeichensatz:"
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label LblWeight 
      Caption         =   "Gewichtung, Schriftschnitt:"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label LblFontSample 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Quick Brown Fox Jumps Over The Lazy Dog. Franz Jagt im komplett verwahrlosten Taxi quer durch Bayern."
      Height          =   1575
      Left            =   120
      TabIndex        =   34
      Top             =   5280
      Width           =   7455
   End
End
Attribute VB_Name = "FFontSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Return As VbMsgBoxResult
Private m_Obj As Font
Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare Function EnumFontFamiliesExW Lib "gdi32" (ByVal hDC As Long, ByRef lpLogFont As Any, ByVal lpEnumFontProc As Long, ByVal lParam As Collection, Optional ByVal dw As Long = 0) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Private FontShared      As Font 'all shared functions on all then Enums
Private m_FontsAll      As Collection
Private m_FontsFixed    As Collection
Private m_FontsVariable As Collection
Private m_bInitializing As Boolean

Private Sub Form_Load()
    'm_bInitializing = True
    Set FontShared = New Font
    m_Return = VbMsgBoxResult.vbCancel
    FillCombos
    LblFontSample.Caption = "Quick Brown Fox Jumps Over The Lazy Dog." & vbCrLf & _
                            "Franz Jagt im komplett verwahrlosten Taxi quer durch Bayern." & vbCrLf & _
                            "AaBbCcDdEe FfGgHhIiJj KkLlMmNnOoPp QqRrSsTtUu VvWwXxYyZz" & vbCrLf & _
                            "ÄäÖöÜüß 0123456789 +-*/ !""§$%&/()=?`´**+~#'"
    Set m_FontsAll = MFonts.FontsAll
    Set m_FontsFixed = MFonts.FontsFixed
    Set m_FontsVariable = MFonts.FontsVariable
    If m_FontsAll Is Nothing Then Exit Sub
    If m_FontsAll.Count = 0 Then Exit Sub
    LBFonts_Fill m_FontsAll
    'm_bInitializing = False
End Sub

Private Sub Form_Resize()
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim L As Single: L = 0 'LblFontSample.Left
    Dim t As Single: t = LblFontSample.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - t - 2 * brdr - BtnOK.Height
    If W > 0 And H > 0 Then LblFontSample.Move L, t, W, H
    t = t + H + brdr
    W = BtnOK.Width: H = BtnOK.Height
    L = Me.ScaleWidth / 2 - brdr - W
    If W > 0 And H > 0 Then BtnOK.Move L, t, W, H
    L = Me.ScaleWidth / 2 + brdr
    If W > 0 And H > 0 Then BtnCancel.Move L, t, W, H
End Sub

Sub LBFonts_Fill(FontList As Collection)
    LBFonts.Clear
    Dim fnt As Font
    For Each fnt In FontList
        LBFonts.AddItem fnt.name
    Next
    LblFontsCount.Caption = FontList.Count
End Sub

Sub UpdateFontsView()
    If ChkPitchFixed.Value = CheckBoxConstants.vbChecked And ChkPitchVariable.Value = CheckBoxConstants.vbUnchecked Then
        LBFonts_Fill m_FontsFixed
    ElseIf ChkPitchFixed.Value = CheckBoxConstants.vbUnchecked And ChkPitchVariable.Value = CheckBoxConstants.vbChecked Then
        LBFonts_Fill m_FontsVariable
    ElseIf ChkPitchFixed.Value = CheckBoxConstants.vbChecked And ChkPitchVariable.Value = CheckBoxConstants.vbChecked Then
        LBFonts_Fill m_FontsAll
    End If
End Sub

Function LBFonts_Search(ByVal name As String) As Long
    LBFonts_Search = -1
    If LBFonts.ListCount = 0 Then Exit Function
    Dim i As Long
    For i = 0 To LBFonts.ListCount - 1
        If LBFonts.List(i) = name Then
            LBFonts_Search = i
            Exit Function
        End If
    Next
End Function

Private Sub ChkPitchFixed_Click()
    UpdateFontsView
End Sub

Private Sub ChkPitchVariable_Click()
    UpdateFontsView
End Sub

Sub FillCombos()
    Dim i As Long
    For i = 5 To 100: Me.CmbPointSize.AddItem i: Next
    For i = 0 To 99:  Me.CmbWidth.AddItem i:  Next
    
    FontShared.EFontWeight_ToCBLB Me.CmbWeight
    FontShared.ECharSet_ToCBLB Me.CmbCharSet
    
    FontShared.EOutPrecision_ToCBLB Me.CmbOutPrec
    FontShared.EClipPrecision_ToCBLB Me.CmbClipPrec
    FontShared.EFontQuality_ToCBLB Me.CmbQuality
    FontShared.EFontPitch_ToCBLB Me.CmbPitch
    FontShared.EFontFamily_ToCBLB Me.CmbFamily
End Sub

Public Function ShowDialog(Obj_inout As Font, Owner As Form) As VbMsgBoxResult
    Set m_Obj = Obj_inout.Clone
    UpdateView m_Obj
    Me.Show vbModal, Owner
    If m_Return = VbMsgBoxResult.vbCancel Then Exit Function
    Obj_inout.NewC m_Obj
End Function

Private Sub BtnOK_Click()
    m_Return = VbMsgBoxResult.vbOK
    UpdateData m_Obj
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Public Sub UpdateView(obj As Object)
    'Set m_Obj = obj
    With m_Obj
        Me.TxtFontName.Text = .name
        Me.TxtEscape.Text = .Escapement
        Me.TxtOrient.Text = .Orientation
        Me.ChkItalic.Value = IIf(.Italic, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
        Me.ChkStrikeout.Value = IIf(.Strikethrough, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
        Me.ChkUnderline = IIf(.Underline, CheckBoxConstants.vbChecked, CheckBoxConstants.vbUnchecked)
        Me.CmbCharSet.Text = FontShared.ECharSet_ToStr(.Charset)
        Me.CmbClipPrec.Text = FontShared.EClipPrecision_ToStr(.ClipPrecision)
        Me.CmbOutPrec.Text = FontShared.EOutPrecision_ToStr(.OutPrecision)
        Me.CmbPitch.Text = FontShared.EFontPitch_ToStr(.Pitch)
        Me.CmbFamily.Text = FontShared.EFontFamily_ToStr(.Family)
        Me.CmbPointSize.Text = CStr(.Size)
        Me.CmbWidth.Text = CStr(.Width)
        Me.CmbWeight.Text = FontShared.EFontWeight_ToStr(.Weight)
        Me.CmbQuality.Text = FontShared.EFontQuality_ToStr(.Quality)
        'Me.CmbPitch.Text = FontShared.EFontPitch_ToStr(.Pitch)
        Me.PBColor.BackColor = .Color
    End With
    'den namen suchen und in der Listbox anzeigen
    Dim nam As String: nam = m_Obj.name
    Dim i As Long: i = LBFonts_Search(nam)
    If i > 0 Then LBFonts.ListIndex = i
    UpdateFontView
End Sub

Public Sub UpdateFontView()
    Set Me.LblFontSample.Font = m_Obj.ToStdFont
    Me.LblFontSample.ForeColor = m_Obj.Color
End Sub

Public Function UpdateData(obj As Object) As Boolean
    Set m_Obj = obj
    m_Obj.SetParams name:=Me.TxtFontName.Text, _
                    PointSize:=CSng(Me.CmbPointSize.Text), _
                    Width:=CLng(Me.CmbWidth.Text), _
                    Escapement:=CLng(Me.TxtEscape.Text), _
                    Orientation:=CLng(Me.TxtOrient.Text), _
                    Weight:=FontShared.EFontWeight_Parse(Me.CmbWeight.Text), _
                    Color:=CLng(Me.PBColor.BackColor), _
                    bItalic:=Me.ChkItalic.Value = CheckBoxConstants.vbChecked, _
                    bUnderline:=Me.ChkUnderline.Value = CheckBoxConstants.vbChecked, _
                    bStrikeout:=Me.ChkStrikeout.Value = CheckBoxConstants.vbChecked, _
                    Charset:=FontShared.ECharSet_Parse(Me.CmbCharSet.Text), _
                    OutPrec:=FontShared.EOutPrecision_Parse(Me.CmbOutPrec.Text), _
                    ClipPrec:=FontShared.EClipPrecision_Parse(Me.CmbClipPrec.Text), _
                    Quality:=FontShared.EFontQuality_Parse(Me.CmbQuality.Text), _
                    Pitch:=FontShared.EFontPitch_Parse(Me.CmbPitch.Text), _
                    Family:=FontShared.EFontFamily_Parse(Me.CmbFamily.Text)
    UpdateFontView
    
'Name, PointSize, Width, Escapement, Orientation, Weight, Color, bItalic, bUnderline, bStrikeOut, CharSet, OutPrec, ClipPrec, Quality, Pitch, Family
End Function

Private Sub LBFonts_Click()
    Dim li As Long: li = LBFonts.ListIndex
    TxtFontName.Text = LBFonts.List(li)
    TxtFontName_LostFocus
End Sub

Private Sub PBColor_DblClick()
    Dim ColorDlg As ColorDialog: Set ColorDlg = New ColorDialog
    If ColorDlg.ShowDialog(Me) = VbMsgBoxResult.vbCancel Then Exit Sub
    PBColor.BackColor = ColorDlg.Color
    m_Obj.SetParams Color:=PBColor.BackColor
    UpdateFontView
End Sub

Private Sub TxtFontName_LostFocus()
    Dim fnam As String: fnam = TxtFontName.Text
    Dim fnt As Font: Set fnt = AllFonts_Item(fnam)
    If fnt Is Nothing Then Exit Sub
    Dim sf As StdFont: Set sf = m_Obj.ToStdFont
    m_Obj.NewC fnt
    m_Obj.SetParams PointSize:=sf.Size, Weight:=sf.Weight, Color:=PBColor.BackColor, bItalic:=sf.Italic, bUnderline:=sf.Underline, bStrikeout:=sf.Strikethrough, Charset:=sf.Charset
    m_Obj.Bold = sf.Bold
    'm_Obj.SetParams Name:=Me.TxtFontName.Text
    UpdateView m_Obj
    UpdateFontView
End Sub

Private Property Get AllFonts_Item(ByVal name As String) As Font
Try: On Error GoTo Catch
    Set AllFonts_Item = m_FontsAll.Item(name)
Catch:
End Property

Private Sub CmbPointSize_Click()
    m_Obj.SetParams PointSize:=CSng(CmbPointSize.Text)
    UpdateFontView
End Sub

Private Sub CmbWeight_Click()
    m_Obj.SetParams Weight:=FontShared.EFontWeight_Parse(CmbWeight.Text)
    UpdateFontView
End Sub

Private Sub CmbCharSet_Click()
    m_Obj.SetParams Charset:=FontShared.ECharSet_Parse(CmbCharSet.Text)
    UpdateFontView
End Sub

Private Sub CmbOutPrec_Click()
    m_Obj.SetParams OutPrec:=FontShared.EOutPrecision_Parse(CmbOutPrec.Text)
    UpdateFontView
End Sub

Private Sub CmbClipPrec_Click()
    m_Obj.SetParams ClipPrec:=FontShared.EClipPrecision_Parse(CmbClipPrec.Text)
    UpdateFontView
End Sub

Private Sub CmbQuality_Click()
    m_Obj.SetParams Quality:=FontShared.EFontQuality_Parse(CmbQuality.Text)
    UpdateFontView
End Sub

Private Sub CmbPitch_Click()
    m_Obj.SetParams Pitch:=FontShared.EFontPitch_Parse(CmbPitch.Text)
    UpdateFontView
End Sub

Private Sub CmbFamily_Click()
    m_Obj.SetParams Family:=FontShared.EFontFamily_Parse(CmbFamily.Text)
    UpdateFontView
End Sub


Private Sub ChkItalic_Click()
    m_Obj.SetParams bItalic:=ChkItalic.Value = CheckBoxConstants.vbChecked
    UpdateFontView
End Sub

Private Sub ChkUnderline_Click()
    m_Obj.SetParams bUnderline:=ChkUnderline.Value = CheckBoxConstants.vbChecked
    UpdateFontView
End Sub

Private Sub ChkStrikeout_Click()
    m_Obj.SetParams bStrikeout:=ChkStrikeout.Value = CheckBoxConstants.vbChecked
    UpdateFontView
End Sub

