VERSION 5.00
Begin VB.Form frmPreview 
   Caption         =   "Preview"
   ClientHeight    =   5715
   ClientLeft      =   2220
   ClientTop       =   2265
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.VScrollBar VScroll1 
      Height          =   5445
      Left            =   6300
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   285
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5430
      Width           =   6315
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6315
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Click Here To Change Font Size"
      Top             =   5460
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2685
      Left            =   0
      ScaleHeight     =   2685
      ScaleWidth      =   5325
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   30
      Width           =   5325
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lCurrentSize As Long
Dim FS(0 To 9) As Long
Private sPrintThis() As String
Private LineDisplay As Long
Public charWidth As Long

Public Property Let PrintText(New_PrintText() As String)
    sPrintThis = New_PrintText
End Property

Private Sub Form_Load()
    lCurrentSize = 3
    FS(0) = 5
    FS(1) = 6
    FS(2) = 8
    FS(3) = 9
    FS(4) = 10
    FS(5) = 11
    FS(6) = 12
    FS(7) = 14
    FS(8) = 16
    FS(9) = 18
End Sub

Private Sub Form_Resize()
    Dim i As Long
    Dim onLast As Boolean

    HScroll1.Enabled = False
    VScroll1.Enabled = False
    
    If Me.Width < 1200 Then Me.Width = 1200
    If Me.Height < 1200 Then Me.Height = 1200
    VScroll1.Move Me.ScaleWidth - VScroll1.Width, 0, VScroll1.Width, Me.ScaleHeight - HScroll1.Height
    HScroll1.Move 0, Me.ScaleHeight - HScroll1.Height, Me.ScaleWidth - VScroll1.Width
    Picture1.Move 0, 0, IIf(Picture1.Width < VScroll1.Left, VScroll1.Left, Picture1.Width), HScroll1.Top
    Picture2.Move HScroll1.Width, VScroll1.Height
    LineDisplay = HScroll1.Top / Picture1.TextHeight("M")

    If LineDisplay < UBound(sPrintThis) Then
        With VScroll1
            onLast = (.Value = .Max)
            .Max = UBound(sPrintThis) \ (LineDisplay - 3)
            If .Max < 10 Then
                .LargeChange = 5
            ElseIf .Max < 100 Then
                .LargeChange = 10
            ElseIf .Max < 300 Then
                .LargeChange = 100
            ElseIf .Max < 1000 Then
                .LargeChange = 250
            Else
                .LargeChange = 500
            End If
            If .Value > .Max Or onLast Then .Value = .Max
            .Enabled = True
        End With
    End If
    Picture1.Cls
    For i = 1 To LineDisplay
        If (i + VScroll1.Value * (LineDisplay - 3)) Mod 2 <> 0 Then Picture1.Line (Picture1.ScaleWidth, Picture1.CurrentY + Picture1.TextHeight("M"))-(Picture1.CurrentX, Picture1.CurrentY), Val(Picture1.Tag), BF
        If (i + VScroll1.Value * (LineDisplay - 3)) <= UBound(sPrintThis) Then
            If Picture1.Width < Picture1.TextWidth(sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))) + Picture1.TextWidth("M") Then Picture1.Width = Picture1.TextWidth(sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))) + Picture1.TextWidth("M")
            If Left(sPrintThis((i + VScroll1.Value * (LineDisplay - 3))), 10) = "<PAGELINE>" Then
                Picture1.Print String(10, Mid(sPrintThis((i + VScroll1.Value * (LineDisplay - 3))), 11, 1)) & Right(sPrintThis((i + VScroll1.Value * (LineDisplay - 3))), Len(sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))) - 10)
            Else
                Picture1.Print sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))
            End If
        Else
            Picture1.Print
        End If
    Next i
    If Picture1.Width > VScroll1.Left Then
        With HScroll1
            onLast = (.Value = .Max)
            .Max = (Picture1.Width - VScroll1.Left) / Screen.TwipsPerPixelX / 10
            If .Max < 10 Then
                .LargeChange = 5
            ElseIf .Max < 100 Then
                .LargeChange = 10
            ElseIf .Max < 300 Then
                .LargeChange = 100
            ElseIf .Max < 1000 Then
                .LargeChange = 250
            Else
                .LargeChange = 500
            End If
            If .Value > .Max Or onLast Then .Value = .Max
            If .Max > 0 Then .Enabled = True
        End With
    End If
End Sub

Private Sub HScroll1_Change()
    Dim i As Long

    If Not HScroll1.Enabled Then Exit Sub
    Picture1.Cls
    For i = 1 To LineDisplay
        If ((i + VScroll1.Value * (LineDisplay - 3))) Mod 2 <> 0 Then Picture1.Line (Picture1.ScaleWidth, Picture1.CurrentY + Picture1.TextHeight("M"))-(Picture1.CurrentX, Picture1.CurrentY), Val(Picture1.Tag), BF
        If (i + VScroll1.Value * (LineDisplay - 3)) <= UBound(sPrintThis) Then
            If Picture1.Width < Picture1.TextWidth(sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))) + Picture1.TextWidth("M") Then Picture1.Width = Picture1.TextWidth(sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))) + Picture1.TextWidth("M")
            If Left(sPrintThis((i + VScroll1.Value * (LineDisplay - 3))), 10) = "<PAGELINE>" Then
                Picture1.Print String(10, Mid(sPrintThis((i + VScroll1.Value * (LineDisplay - 3))), 11, 1)) & Right(sPrintThis((i + VScroll1.Value * (LineDisplay - 3))), Len(sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))) - 10)
            Else
                Picture1.Print sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))
            End If
        Else
            Picture1.Print
        End If
    Next i
    Picture1.Move -(HScroll1.Value * 10 * Screen.TwipsPerPixelX)
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If lCurrentSize > 0 Then lCurrentSize = lCurrentSize - 1
        If Picture1.FontSize <> FS(lCurrentSize) Then
            Picture1.FontSize = FS(lCurrentSize)
            Picture1.FontBold = False
            Form_Resize
        End If
    ElseIf Button = vbRightButton Then
        If lCurrentSize < 9 Then lCurrentSize = lCurrentSize + 1
        If Picture1.FontSize <> FS(lCurrentSize) Then
            Picture1.FontSize = FS(lCurrentSize)
            Picture1.FontBold = False
            Form_Resize
        End If
    End If
    Picture2.Cls
    Picture2.CurrentX = 10
    Picture2.CurrentY = 10
    Picture2.Print Format(Picture1.FontSize, "##")
End Sub

Private Sub VScroll1_Change()
    Dim i As Long

    If Not VScroll1.Enabled Then Exit Sub
    Picture1.Cls
    For i = 1 To LineDisplay
        If ((i + VScroll1.Value * (LineDisplay - 3))) Mod 2 <> 0 Then Picture1.Line (Picture1.ScaleWidth, Picture1.CurrentY + Picture1.TextHeight("M"))-(Picture1.CurrentX, Picture1.CurrentY), Val(Picture1.Tag), BF
        If (i + VScroll1.Value * (LineDisplay - 3)) <= UBound(sPrintThis) Then
            If Picture1.Width < Picture1.TextWidth(sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))) + Picture1.TextWidth("M") Then Picture1.Width = Picture1.TextWidth(sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))) + Picture1.TextWidth("M")
            If Left(sPrintThis((i + VScroll1.Value * (LineDisplay - 3))), 10) = "<PAGELINE>" Then
                Picture1.Print String(10, Mid(sPrintThis((i + VScroll1.Value * (LineDisplay - 3))), 11, 1)) & Right(sPrintThis((i + VScroll1.Value * (LineDisplay - 3))), Len(sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))) - 10)
            Else
                Picture1.Print sPrintThis((i + VScroll1.Value * (LineDisplay - 3)))
            End If
        Else
            Picture1.Print
        End If
    Next i
End Sub

