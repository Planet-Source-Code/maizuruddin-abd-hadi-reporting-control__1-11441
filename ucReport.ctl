VERSION 5.00
Begin VB.UserControl ucReport 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ucReport.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ucReport.ctx":066A
End
Attribute VB_Name = "ucReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Printing Orientation Potrait or Landscape
Public Enum PaperPrintOrientation
    ppoPRORLandscape = vbPRORLandscape
    ppoPRORPortrait = vbPRORPortrait
End Enum

'Default property values
Private Const m_def_AllowSkipPageOnLastRow = False
Private Const m_def_BottomMargin = 3
Private Const m_def_FooterLength = 6
Private Const m_def_LeftMargin = 5
Private Const m_def_PaperOrientation = 1
Private Const m_def_PreviewEvenBackColor = &H80000005
Private Const m_def_PreviewForeColor = &H0&
Private Const m_def_PreviewOddBackColor = &HDDFDDB
Private Const m_def_PrintFontSize = 9
Private Const m_def_RightMargin = 5
Private Const m_def_TopMargin = 3

'Property value holder
Private m_AllowSkipPageOnLastRow As Boolean
Private m_BottomMargin As Long
Private m_FooterLength As Long
Private m_LeftMargin As Long
Private m_PaperOrientation As PaperPrintOrientation
Private m_PageLength As Long
Private m_PageWidth As Long
Private m_PreviewOddBackColor As OLE_COLOR
Private m_PreviewEvenBackColor As OLE_COLOR
Private m_PreviewForeColor As OLE_COLOR
Private m_PrintFontSize As Single
Private m_RightMargin As Long
Private m_TopMargin As Long

Private m_PageNum As Long

'Control variables
Dim blnStartReport As Boolean
Dim blnHasGrouping As Boolean
Dim GroupingID() As String
Dim PreviousValue() As String
Dim RaisedEvents() As Boolean
Dim lLineCount As Long
Dim blnJustStarted As Boolean
Dim sPrintThis() As String
Dim blnFinishRaised As Boolean
Dim blnRaised As Boolean
Dim blnPageFooter As Boolean
Dim blnFlushed As Boolean
Dim blnPreview As Boolean
Dim blnFirstPage As Boolean
Dim blnLastPage As Boolean

Dim m_PrintTo As Printer

'Events decalaration
'Events for reporting
Event FirstPageHeader()
Event PageHeader()
Event BeforeGroupOf(ByVal GroupID As String, ByVal GroupValue As String)
Event OnEveryRow()
Event AfterGroupOf(ByVal GroupID As String, ByVal GroupValue As String)
Event OnLastRow()
Event PageFooter()
'Events for printing
Event StartPrinting()
Event Printing(ByVal Percentage As Single, Cancel As Boolean)
Event FinishPrinting()
Event AbortPrinting()

'Property : AllowSkipPageOnLastRow
' On the last row of data to be printed, do we
' allow the SkipPage command to be used or not.
Public Property Get AllowSkipPageOnLastRow() As Boolean
    AllowSkipPageOnLastRow = m_AllowSkipPageOnLastRow
End Property
Public Property Let AllowSkipPageOnLastRow(ByVal New_AllowSkipPageOnLastRow As Boolean)
    m_AllowSkipPageOnLastRow = New_AllowSkipPageOnLastRow
End Property

'Property : BottomMargin
' To declare number of lines as the report
' bottom margin
Public Property Get BottomMargin() As Long
    BottomMargin = m_BottomMargin
End Property
Public Property Let BottomMargin(ByVal New_BottomMargin As Long)
    If blnPreview Then Exit Property
    m_BottomMargin = New_BottomMargin
    PropertyChanged "BottomMargin"
End Property

'Method     : CancelReport
'Parameters : none
' To reinitialize variables indicating no report
' was produced
Public Sub CancelReport()
    If Not blnPreview Then Exit Sub
    
    ReDim sPrintThis(0)
    blnStartReport = False
    blnFlushed = False
    blnPreview = False
End Sub

'Method     : FinishReport
'Parameters : none
' The last method to be called during reporting
Public Sub FinishReport()
    Dim i As Long

    'Only if reporting has been started with
    'the StartReport do we allow this method
    'to proceed
    If Not blnStartReport Then Exit Sub
    
    'To indicate that this will be the last page
    'printed
    blnLastPage = True
    
    'Check to see whether grouping was declared
    If blnHasGrouping Then
        'If grouping was declared, raised the
        'AfterGroup event for every grouping
        For i = UBound(GroupingID) To LBound(GroupingID) Step -1
            RaiseEvent AfterGroupOf(GroupingID(i), PreviousValue(i))
        Next i
    End If
    
    'Raise the OnLastRow event
    RaiseEvent OnLastRow
    
    'Check whether printing has reached,
    'bottom of report
    blnFinishRaised = True
    If lLineCount < m_PageLength - m_FooterLength - m_BottomMargin Then
        'if reached bottom of report raise the
        'PageFooter event
        lLineCount = m_PageLength - m_FooterLength - m_BottomMargin
        ReDim Preserve sPrintThis(1 To m_PageNum * m_PageLength)
        blnPageFooter = True
        RaiseEvent PageFooter
    End If
    
    'Reinitialize variable to indicate reporting
    'was done
    m_PageNum = 0
    blnPageFooter = False
    blnStartReport = False
    blnLastPage = False
    blnFirstPage = False
    blnFlushed = True
    blnPreview = True
End Sub

'Property : FooterLength
' To declare number of lines as the report
' trailer
Public Property Get FooterLength() As Long
    FooterLength = m_FooterLength
End Property
Public Property Let FooterLength(ByVal New_FooterLength As Long)
    If blnPreview Then Exit Property
    m_FooterLength = New_FooterLength
    PropertyChanged "FooterLength"
End Property

'Property : LeftMargin
' To declare number of lines as the report
' left margin
Public Property Get LeftMargin() As Long
    LeftMargin = m_LeftMargin
End Property
Public Property Let LeftMargin(ByVal New_LeftMargin As Long)
    If blnPreview Then Exit Property
    m_LeftMargin = New_LeftMargin
    PropertyChanged "LeftMargin"
End Property

'Property (READONLY): PageLength
' To get the number of lines of the report
Public Property Get PageLength() As Long
    PageLength = m_PageLength
End Property

'Property (READONLY): PageNum
' To get current page number, during the
' reporting process only
Public Property Get PageNum() As Long
    'If reporting has not been started the
    'PageNum property will yield the value of 0
    If Not blnStartReport Then Exit Property
    PageNum = m_PageNum
End Property

'Property (READONLY): PageWidth
' To get the report printable width
Public Property Get PageWidth() As Long
    PageWidth = m_PageWidth
End Property

'Property : PaperOrientation
' To declare the report printing orientation
Public Property Get PaperOrientation() As PaperPrintOrientation
    PaperOrientation = m_PaperOrientation
End Property
Public Property Let PaperOrientation(ByVal New_PaperOrientation As PaperPrintOrientation)
    If blnPreview Then Exit Property
    m_PaperOrientation = New_PaperOrientation
    PropertyChanged "PaperOrientation"
End Property

'Method     : Preview
'Parameters : none
' this method is used to previewed the report
' produced
Public Sub Preview()
    Dim frmTemp As Form
    Dim i As Long
    Dim k As Long
    
    'Check whether reporting process is complete
    If Not blnFinishRaised Then Exit Sub
    If Not blnPreview Then Exit Sub
    
    On Error GoTo PreviewError
    i = LBound(sPrintThis)
    On Error GoTo 0
    UserControl.Extender.Parent.MousePointer = vbHourglass
    Set frmTemp = New frmPreview
    Load frmTemp
    With frmTemp
        .Picture1.BackColor = m_PreviewEvenBackColor
        .Picture1.ForeColor = m_PreviewForeColor
        .Picture1.Tag = m_PreviewOddBackColor
        .Picture1.FontSize = 9
        .Picture1.FontBold = False
        .PrintText = sPrintThis
        .Picture2.Print Format(.Picture1.FontSize, "##")
        UserControl.Extender.Parent.MousePointer = vbDefault
        .Show vbModal
    End With
    Unload frmTemp
    Set frmTemp = Nothing
PreviewError:
    On Error GoTo 0
    Set frmTemp = Nothing
End Sub

'Property : PreviewEvenBackColor
' To declare the background color for every even
' lines of the report during preview
Public Property Get PreviewEvenBackColor() As OLE_COLOR
    PreviewEvenBackColor = m_PreviewEvenBackColor
End Property
Public Property Let PreviewEvenBackColor(ByVal New_PreviewEvenBackColor As OLE_COLOR)
    m_PreviewEvenBackColor = New_PreviewEvenBackColor
    PropertyChanged "PreviewEvenBackColor"
End Property

'Property : PreviewForeColor
' To declare the foreground color during preview
Public Property Get PreviewForeColor() As OLE_COLOR
    PreviewForeColor = m_PreviewForeColor
End Property
Public Property Let PreviewForeColor(ByVal New_PreviewForeColor As OLE_COLOR)
    m_PreviewForeColor = New_PreviewForeColor
    PropertyChanged "PreviewForeColor"
End Property

'Property : PreviewOddBackColor
' To declare the background color for every odd
' lines of the report during preview
Public Property Get PreviewOddBackColor() As OLE_COLOR
    PreviewOddBackColor = m_PreviewOddBackColor
End Property
Public Property Let PreviewOddBackColor(ByVal New_PreviewOddBackColor As OLE_COLOR)
    m_PreviewOddBackColor = New_PreviewOddBackColor
    PropertyChanged "PreviewOddBackColor"
End Property

'Method     : SaveToFile
'Parameters : FileName
' this method is used to save the report
' produced to a file specified by the FileName
Public Sub SaveToFile(ByVal FileName As String)
    Dim i As Long
    Dim l As Long
    Dim OutFile As String
    Dim blnProceed As Boolean

    'Check for FileName validity
    If Len(Trim(FileName)) < 1 Then Exit Sub
    
    'Check whether reporting process is complete
    If Not blnFinishRaised Then Exit Sub
    If Not blnPreview Then Exit Sub
    
    blnProceed = True
    'Check whether the specified file exist or not
    OutFile = Dir(FileName)
    'If exist, display a message box inquiring
    'whether the file can be overwrite or not
    If OutFile <> "" Then blnProceed = (MsgBox("File " & OutFile & " exist." & vbCrLf & "Do you wish to override ?", vbQuestion + vbYesNo, "Override ?") = vbYes)

    If blnProceed Then
ReOpenFile:
        l = FreeFile
        On Error GoTo ErrorOpenFile
        Open FileName For Output Access Write Lock Write As #l
        On Error GoTo 0
        For i = LBound(sPrintThis) To UBound(sPrintThis)
            If Left(sPrintThis(i), 10) = "<PAGELINE>" Then
                Print #l, Space(m_LeftMargin) & String(10, Mid(sPrintThis(i), 11, 1)) & Right(sPrintThis(i), Len(sPrintThis(i)) - 10)
            Else
                Print #l, Space(m_LeftMargin) & sPrintThis(i)
            End If
        Next i
        Close #l
        On Error GoTo 0
    End If
    Exit Sub
ErrorOpenFile:
    On Error Resume Next
'    Close #l
    If MsgBox("Could not open " & OutFile & "." & vbCrLf & "Do you wish to retry ?", vbExclamation + vbOKCancel) = vbOK Then
        GoTo ReOpenFile
    Else
        On Error GoTo 0
        Exit Sub
    End If
End Sub

'Method     : SendToPrinter
'Parameters : none
' this method is used to print the report
' produced
Public Sub SendToPrinter()
    Dim i As Long
    Dim k As Single
    Dim Cancel As Boolean

    'Check whether reporting process is complete
    If Not blnFinishRaised Then Exit Sub
    If Not blnPreview Then Exit Sub
    
    'raise the StartPrinting event
    RaiseEvent StartPrinting
    
    For i = LBound(sPrintThis) To UBound(sPrintThis)
        'raise the Printing event
        RaiseEvent Printing(((i + 1) / (UBound(sPrintThis) - LBound(sPrintThis) + 1)), Cancel)
        'Check to determine whether user has opted to cancel printing
        If Cancel Then Exit For
        If Left(sPrintThis(i), 10) = "<PAGELINE>" Then
            k = m_PrintTo.CurrentY
            If Mid(sPrintThis(i), 11, 1) = "-" Then
                m_PrintTo.Line (m_PrintTo.ScaleWidth - m_PrintTo.TextWidth("M") * m_RightMargin, k + m_PrintTo.TextHeight("M") \ 2)-(m_PrintTo.TextWidth("M") * m_LeftMargin, k + m_PrintTo.TextHeight("M") \ 2)
            ElseIf Mid(sPrintThis(i), 11, 1) = "=" Then
                m_PrintTo.Line (m_PrintTo.ScaleWidth - m_PrintTo.TextWidth("M") * m_RightMargin, k + m_PrintTo.TextHeight("M") \ 2 - 2)-(m_PrintTo.TextWidth("M") * m_LeftMargin, k + m_PrintTo.TextHeight("M") \ 2 - 2)
                m_PrintTo.Line (m_PrintTo.ScaleWidth - m_PrintTo.TextWidth("M") * m_RightMargin, k + m_PrintTo.TextHeight("M") \ 2 + 2)-(m_PrintTo.TextWidth("M") * m_LeftMargin, k + m_PrintTo.TextHeight("M") \ 2 + 2)
            Else
                m_PrintTo.CurrentX = m_PrintTo.TextWidth("M") * m_LeftMargin
                m_PrintTo.Print String(10, Mid(sPrintThis(i), 11, 1)) & Right(sPrintThis(i), Len(sPrintThis(i)) - 10)
            End If
            m_PrintTo.CurrentY = k + m_PrintTo.TextHeight("M")
        Else
            m_PrintTo.CurrentX = m_PrintTo.TextWidth("M") * m_LeftMargin
            m_PrintTo.Print sPrintThis(i)
        End If
        'Check whether we need to call NewPage
        'method of the printer object
        If i Mod m_PageLength = 0 Then m_PrintTo.NewPage
    Next i
    'Check whether printing was cancelled
    If Cancel Then
        m_PrintTo.KillDoc
        'raise event AbortPrinting to indicate
        'printing was cancelled
        RaiseEvent AbortPrinting
    Else
        m_PrintTo.EndDoc
        'raise event FinishPrinting to indicate
        'printing was completed
        RaiseEvent FinishPrinting
    End If
    blnFlushed = False
End Sub

'Property : PrintFontSize
' To declare the font size to used during printing
Public Property Get PrintFontSize() As Single
    PrintFontSize = m_PrintFontSize
End Property
Public Property Let PrintFontSize(ByVal New_PrintFontSize As Single)
    If blnPreview Then Exit Property
    m_PrintFontSize = New_PrintFontSize
    PropertyChanged "PrintFontSize"
End Property

'Method     : PrintLine
'Parameters : LineChar
' this method is to generate a horizontal line
' across the reporting page, the line produced
' is defined by the Linechar
Public Sub PrintLine(Optional ByVal LineChar As String = "-")
    'Only if reporting has been started with
    'the StartReport do we allow this method
    'to proceed
    If Not blnStartReport Then Exit Sub
    
    'Check whether method was called during the
    'PageFooter event
    If blnPageFooter Then If lLineCount + 1 > m_PageLength Then Exit Sub

    'Check whether the LineChar is not blank
    If Len(Trim(LineChar)) < 1 Then
        'If the LineChar is blank then it is a
        'just like skipping a line, so we called
        'then SkipLine method instead
        SkipLine
        Exit Sub
    End If
    
    'Check whether method was called during
    'FinishReport method
    If Not blnFinishRaised Then
        If lLineCount + 1 >= m_PageLength - m_FooterLength - m_BottomMargin Then
            'Check whether the PageFooter
            'event was previously raised
            If Not blnRaised Then
                'Raise the PageFooter event for
                'the current page
                blnRaised = True
                lLineCount = m_PageLength - m_FooterLength - m_BottomMargin
                blnPageFooter = True
                ReDim Preserve sPrintThis(1 To m_PageNum * m_PageLength)
                RaiseEvent PageFooter
                blnPageFooter = False
                
                'Raise the PageHeader event for
                'the next page
                lLineCount = 1 + m_TopMargin
                m_PageNum = m_PageNum + 1
                ReDim Preserve sPrintThis(1 To (m_PageNum - 1) * m_PageLength + lLineCount)
                RaiseEvent PageHeader
                If blnFirstPage Then blnFirstPage = False
                blnRaised = False
            End If
        End If
    End If
    lLineCount = lLineCount + 1
    ReDim Preserve sPrintThis(1 To (m_PageNum - 1) * m_PageLength + lLineCount)
    sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount) = "<PAGELINE>" & String(m_PageWidth - 10, LineChar)
End Sub

'Method     : PrintOutput
'Parameters : NextLine, Output
' this method is used to write to the report
' based upon passed parameter, the NextLine is
' to indicate that whether printing should
' proceed to next line or on the current line.
' The Output parameter must be passed as follows
'    [column position], [value to be printed],
'    [column position], [value to be printed],
'    ...
' [column position] - is to indicate the the
'   starting position to write
' [value to be printed] - is the value to be
'   written to the report
Public Sub PrintOutput(ByVal NextLine As Boolean, ParamArray Output() As Variant)
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim z() As Variant
    
    If LBound(Output) = UBound(Output) And IsArray(Output(LBound(Output))) Then
        z = Output(LBound(Output))
    Else
        z = Output
    End If
    
    'Only if reporting has been started with
    'the StartReport do we allow this method
    'to proceed
    If Not blnStartReport Then Exit Sub
    
    'Check whether method was called during the
    'PageFooter event
    If blnPageFooter Then If lLineCount + IIf(NextLine, 1, 0) > m_PageLength Then Exit Sub
    
    'Check whether method was called during
    'FinishReport method
    If Not blnFinishRaised Then
        If lLineCount + IIf(NextLine, 1, 0) >= m_PageLength - m_FooterLength - m_BottomMargin Then
            'Check whether the PageFooter
            'event was previously raised
            If Not blnRaised Then
                'Raise the PageFooter event for
                'the current page
                blnRaised = True
                lLineCount = m_PageLength - m_FooterLength - m_BottomMargin
                blnPageFooter = True
                ReDim Preserve sPrintThis(1 To m_PageNum * m_PageLength)
                RaiseEvent PageFooter
                blnPageFooter = False
                
                'Raise the PageHeader event for
                'the next page
                lLineCount = 1 + m_TopMargin
                m_PageNum = m_PageNum + 1
                ReDim Preserve sPrintThis(1 To (m_PageNum - 1) * m_PageLength + lLineCount)
                RaiseEvent PageHeader
                If blnFirstPage Then blnFirstPage = False
                blnRaised = False
            End If
        End If
    End If
    If NextLine Then
        lLineCount = lLineCount + 1
        ReDim Preserve sPrintThis(1 To (m_PageNum - 1) * m_PageLength + lLineCount)
    End If
    For i = LBound(z) To UBound(z) Step 2
        If z(i + 1) = "%PAGENUM%" Then
            z(i + 1) = m_PageNum
        ElseIf z(i + 1) = "%DATE%" Then
            z(i + 1) = Format(Now, "Short Date")
        ElseIf z(i + 1) = "%TIME%" Then
            z(i + 1) = Format(Now, "hh:mm:ss")
        ElseIf z(i + 1) = "%DATETIME%" Then
            z(i + 1) = Format(Now, "Short Date") & " " & Format(Now, "hh:mm:ss")
        End If
        If z(i) >= 1 And z(i) <= m_PageWidth Then
            If Len(sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount)) < z(i) + Len(z(i + 1)) Then sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount) = sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount) & Space((z(i) + Len(z(i + 1))) - (Len(sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount))))
            Mid(sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount), z(i), Len(z(i + 1))) = z(i + 1)
        Else
            If Len(sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount)) <= m_LeftMargin Then sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount) = Space(m_LeftMargin - Len(sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount)) + 1)
            sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount) = sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount) & z(i + 1)
        End If
        If Len(sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount)) > m_PageWidth Then sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount) = Left(sPrintThis((m_PageNum - 1) * m_PageLength + lLineCount), m_PageWidth)
    Next i
End Sub

'Method     : ReportOutput
'Parameters : GroupValue, GroupDelimiter
' this method is used during the report process
' after calling the StartReport before
' FinishReport, this method usually called
' multiple times especially if there was
' grouping defined at StartReport
Public Sub ReportOutput(Optional ByVal GroupValue As String = "", Optional GroupDelimiter As String = ",")
    Dim i As Long
    Dim j As Long
    Dim z() As String

    'Only if reporting has been started with
    'the StartReport do we allow this method
    'to proceed
    If Not blnStartReport Then Exit Sub
    
    'Check to see whether grouping was declared
    If blnHasGrouping Then
        For j = i To UBound(GroupingID)
            RaisedEvents(j) = False
        Next j
        z = Split(GroupValue, GroupDelimiter)
        For i = LBound(GroupingID) To UBound(GroupingID)
            If i < LBound(z) Or i > UBound(z) Then Exit For
            'Check whether events should be
            'raised for a particular group
            If PreviousValue(i) <> z(i) Then
                For j = i To UBound(GroupingID)
                    RaisedEvents(j) = True
                Next j
                Exit For
            End If
        Next i
        'Check whether the report has just been
        'started with the the StartReport method
        If Not blnJustStarted Then
            For i = UBound(GroupingID) To LBound(GroupingID) Step -1
                'Check whether the AfterGroup
                'event should be raised or not
                If RaisedEvents(i) Then RaiseEvent AfterGroupOf(GroupingID(i), PreviousValue(i))
            Next i
        End If
        For i = LBound(GroupingID) To UBound(GroupingID)
            PreviousValue(i) = z(i)
            'Check whether the BeforeGroup event
            'should be raised or not
            If RaisedEvents(i) Then RaiseEvent BeforeGroupOf(GroupingID(i), z(i))
        Next i
        'Change the variable status of
        'JustStarted to false
        If blnJustStarted Then blnJustStarted = False
    End If
    'raise event OnEveryRow
    RaiseEvent OnEveryRow
    'Change the variable status of FirstPage to
    'false
    If blnFirstPage Then blnFirstPage = False
End Sub

'Property : RightMargin
' To declare number of lines as the report
' right margin
Public Property Get RightMargin() As Long
    RightMargin = m_RightMargin
End Property
Public Property Let RightMargin(ByVal New_RightMargin As Long)
    If blnPreview Then Exit Property
    m_RightMargin = New_RightMargin
    PropertyChanged "RightMargin"
End Property

'Method     : SkipLine
'Parameters : NumberOfLines
' this method is used to print number of blank
' lines depending upon the passed parameter
' value of NumberOfLines
Public Sub SkipLine(Optional ByVal NumberOfLines As Long = 1)
    'Only if reporting has been started with
    'the StartReport do we allow this method
    'to proceed
    If Not blnStartReport Then Exit Sub
    
    'Check to ensure valid value of
    'NumberOfLines has been passed
    If NumberOfLines < 1 Or NumberOfLines > m_PageLength Then Exit Sub
    
    'Check whether method was called during the
    'PageFooter event
    If blnPageFooter Then If lLineCount + NumberOfLines > m_PageLength Then Exit Sub
    
    'Check whether method was called during
    'FinishReport method
    If Not blnFinishRaised Then
        If lLineCount + NumberOfLines >= m_PageLength - m_FooterLength - m_BottomMargin Then
            'Check whether the PageFooter
            'event was previously raised
            If Not blnRaised Then
                'Raise the PageFooter event for
                'the current page
                blnRaised = True
                lLineCount = m_PageLength - m_FooterLength - m_BottomMargin
                blnPageFooter = True
                ReDim Preserve sPrintThis(1 To m_PageNum * m_PageLength)
                RaiseEvent PageFooter
                blnPageFooter = False
                
                'Raise the PageHeader event for
                'the next page
                lLineCount = 1 + m_TopMargin
                m_PageNum = m_PageNum + 1
                ReDim Preserve sPrintThis(1 To (m_PageNum - 1) * m_PageLength + lLineCount)
                RaiseEvent PageHeader
                If blnFirstPage Then blnFirstPage = False
                blnRaised = False
            End If
        End If
    End If
    lLineCount = lLineCount + NumberOfLines
    ReDim Preserve sPrintThis(1 To (m_PageNum - 1) * m_PageLength + lLineCount)
End Sub

'Method     : SkipPage
'Parameters : none
' this method is used to skip directly to next
' page eventhou printing has not reached the
' bottom of reporting area
Public Sub SkipPage()
    'Only if reporting has been started with
    'the StartReport do we allow this method
    'to proceed
    If Not blnStartReport Then Exit Sub
    
    'Check whether report is on the last page
    'user has allowed skip page on the last row
    If blnLastPage And Not m_AllowSkipPageOnLastRow Then Exit Sub
    
    'Check whether method was called during
    'StartReport method
    If blnFirstPage Then Exit Sub
    
    'Check whether method was called during
    'FinishReport method
    If blnFinishRaised Then Exit Sub
    
    'Check whether method was called during the
    'PageFooter event
    If blnPageFooter Then Exit Sub
    
    'Check whether printing is currently at
    'page header
    If lLineCount = 1 + m_TopMargin Then Exit Sub
    
    'Check whether the PageFooter
    'event was previously raised
    If Not blnRaised Then
        'Raise the PageFooter event for the
        'current page
        lLineCount = m_PageLength - m_FooterLength - m_BottomMargin
        blnRaised = True
        blnPageFooter = True
        ReDim Preserve sPrintThis(1 To m_PageNum * m_PageLength)
        RaiseEvent PageFooter
        blnPageFooter = False
        blnRaised = False
        
        'Raise the PageHeader event for the
        'next page
        lLineCount = 1 + m_TopMargin
        m_PageNum = m_PageNum + 1
        ReDim Preserve sPrintThis(1 To (m_PageNum - 1) * m_PageLength + lLineCount)
        RaiseEvent PageHeader
    End If
End Sub

'Method     : StartReport
'Parameters : PrintTo, GroupBy, GroupDelimiter
' this method is used start the reporting
' process, it required the selected printer
' object. The groupby, if nescessary, to invoke
' the BeforeGroup and AfterGroup event
Public Sub StartReport(PrintTo As Object, Optional ByVal GroupBy As String = "", Optional GroupDelimiter As String = ",")
    Dim i As Long
    Dim j As Long
    Dim blnRecheck As Boolean

    'Check to see whether previous reporting
    'exist and if exist called the CancelReport
    'method
    If blnPreview Then CancelReport
    
    'Initialize the printer object
    Set m_PrintTo = PrintTo
    m_PrintTo.ScaleMode = 3
    m_PrintTo.Font.Name = "Courier New"
    m_PrintTo.Font.Size = m_PrintFontSize
    m_PrintTo.Orientation = m_PaperOrientation
    
    'Calculate the page length & width
    m_PageLength = Int(m_PrintTo.ScaleHeight / m_PrintTo.TextHeight("M"))
    m_PageWidth = Int(m_PrintTo.ScaleWidth / m_PrintTo.TextWidth("M")) - m_LeftMargin - m_RightMargin
    
    'Initialize starting report variable
    blnStartReport = True
    blnHasGrouping = False
    blnFinishRaised = False
    blnJustStarted = True
    blnRaised = False
    blnLastPage = False
    blnFirstPage = True
    lLineCount = 1 + m_TopMargin
    m_PageNum = 1
    ReDim sPrintThis(1 To lLineCount)
    
    'Check whether the grouping exists
    If Len(Trim(GroupBy)) > 0 Then
        GroupingID = Split(GroupBy, GroupDelimiter)
CheckBlankGrouping:
        blnHasGrouping = True
        blnRecheck = False
        For i = LBound(GroupingID) To UBound(GroupingID)
            If Len(Trim(GroupingID(i))) < 1 Then
                If i + 1 <= UBound(GroupingID) Then
                    GroupingID(i) = GroupingID(i + 1)
                    For j = i + 1 To UBound(GroupingID) - 1
                        GroupingID(j) = GroupingID(j + 1)
                    Next j
                    GroupingID(UBound(GroupingID)) = ""
                Else
                    If LBound(GroupingID) = UBound(GroupingID) Then
                        blnHasGrouping = False
                    Else
                        ReDim Preserve GroupingID(LBound(GroupingID) To UBound(GroupingID) - 1)
                        blnRecheck = True
                    End If
                    Exit For
                End If
            End If
        Next i
        'Ensure no NULL grouping was passed
        If blnRecheck Then GoTo CheckBlankGrouping
    End If
    If blnHasGrouping Then
        ReDim PreviousValue(LBound(GroupingID) To UBound(GroupingID))
        ReDim RaisedEvents(LBound(GroupingID) To UBound(GroupingID))
    End If
    
    'Start the reporting process
    'raise the FirstPageHeader event
    RaiseEvent FirstPageHeader
    'raise the PageHeader event
    RaiseEvent PageHeader
End Sub

'Property : TopMargin
' To declare number of lines as the report
' top margin
Public Property Get TopMargin() As Long
    TopMargin = m_TopMargin
End Property
Public Property Let TopMargin(ByVal New_TopMargin As Long)
    If blnPreview Then Exit Property
    m_TopMargin = New_TopMargin
    PropertyChanged "TopMargin"
End Property

Private Sub UserControl_InitProperties()
    'Initialize property holder when first created
    m_AllowSkipPageOnLastRow = m_def_AllowSkipPageOnLastRow
    m_BottomMargin = m_def_BottomMargin
    m_FooterLength = m_def_FooterLength
    m_LeftMargin = m_def_LeftMargin
    m_PaperOrientation = m_def_PaperOrientation
    m_PreviewEvenBackColor = m_def_PreviewEvenBackColor
    m_PreviewForeColor = m_def_PreviewForeColor
    m_PreviewOddBackColor = m_def_PreviewOddBackColor
    m_PrintFontSize = m_def_PrintFontSize
    m_RightMargin = m_def_RightMargin
    m_TopMargin = m_def_TopMargin
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Initialize property holder for created controls
    m_AllowSkipPageOnLastRow = PropBag.ReadProperty("AllowSkipPageOnLastRow", m_def_AllowSkipPageOnLastRow)
    m_BottomMargin = PropBag.ReadProperty("BottomMargin", m_def_BottomMargin)
    m_FooterLength = PropBag.ReadProperty("FooterLength", m_def_FooterLength)
    m_LeftMargin = PropBag.ReadProperty("LeftMargin", m_def_LeftMargin)
    m_PaperOrientation = PropBag.ReadProperty("PaperOrientation", m_def_PaperOrientation)
    m_PreviewEvenBackColor = PropBag.ReadProperty("PreviewEvenBackColor", m_def_PreviewEvenBackColor)
    m_PreviewForeColor = PropBag.ReadProperty("PreviewForeColor", m_def_PreviewForeColor)
    m_PreviewOddBackColor = PropBag.ReadProperty("PreviewOddBackColor", m_def_PreviewOddBackColor)
    m_PrintFontSize = PropBag.ReadProperty("PrintFontSize", m_def_PrintFontSize)
    m_RightMargin = PropBag.ReadProperty("RightMargin", m_def_RightMargin)
    m_TopMargin = PropBag.ReadProperty("TopMargin", m_def_TopMargin)
End Sub

Private Sub UserControl_Resize()
    'Resize control to a specific size since
    'this control is invisible at run time
    Static blnInSub As Boolean
    
    If blnInSub Then Exit Sub
    blnInSub = True
    With UserControl
        .Width = 375
        .Height = 345
    End With
    blnInSub = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Save changed property valued
    Call PropBag.WriteProperty("AllowSkipPageOnLastRow", m_AllowSkipPageOnLastRow, m_def_AllowSkipPageOnLastRow)
    Call PropBag.WriteProperty("BottomMargin", m_BottomMargin, m_def_BottomMargin)
    Call PropBag.WriteProperty("FooterLength", m_FooterLength, m_def_FooterLength)
    Call PropBag.WriteProperty("LeftMargin", m_LeftMargin, m_def_LeftMargin)
    Call PropBag.WriteProperty("PaperOrientation", m_PaperOrientation, m_def_PaperOrientation)
    Call PropBag.WriteProperty("PreviewOddBackColor", m_PreviewOddBackColor, m_def_PreviewOddBackColor)
    Call PropBag.WriteProperty("PreviewEvenBackColor", m_PreviewEvenBackColor, m_def_PreviewEvenBackColor)
    Call PropBag.WriteProperty("PreviewForeColor", m_PreviewForeColor, m_def_PreviewForeColor)
    Call PropBag.WriteProperty("PrintFontSize", m_PrintFontSize, m_def_PrintFontSize)
    Call PropBag.WriteProperty("RightMargin", m_RightMargin, m_def_RightMargin)
    Call PropBag.WriteProperty("TopMargin", m_TopMargin, m_def_TopMargin)
End Sub

