Attribute VB_Name = "PasteImageFit"
' Word Paste Image Fit — VBA companion (Phase B path).
' Current MVP paste path is Python + COM. This module mirrors the fit rules
' for users who prefer a .dotm / macro-only install later.

Option Explicit

Public Const CM_TO_POINTS As Double = 28.3464567

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Public Type FitStyle
    MaxWidthMode As String          ' percent_of_content | fixed_cm
    MaxWidthValue As Double
    MaxHeightMode As String         ' percent_of_page | fixed_cm
    MaxHeightValue As Double
    AllowUpscale As Boolean
    Align As String                 ' left | center | right
    FitTableCell As Boolean
End Type

Public Function DefaultNotesStyle() As FitStyle
    Dim s As FitStyle
    s.MaxWidthMode = "percent_of_content"
    s.MaxWidthValue = 92
    s.MaxHeightMode = "percent_of_page"
    s.MaxHeightValue = 75
    s.AllowUpscale = False
    s.Align = "center"
    s.FitTableCell = True
    DefaultNotesStyle = s
End Function

Public Sub SmartPasteImage()
    Dim s As FitStyle
    s = DefaultNotesStyle()
    Call PasteAndFit(s)
End Sub

Public Sub PasteAndFit(ByRef style As FitStyle)
    Dim rng As Range
    Dim ils As InlineShape
    Dim contentW As Double, contentH As Double, pageH As Double
    Dim cellW As Double, cellH As Double
    Dim hasCell As Boolean
    Dim maxW As Double, maxH As Double
    Dim scaleW As Double, scaleH As Double, scale As Double

    Set rng = Selection.Range
    contentW = ActiveDocument.PageSetup.PageWidth - ActiveDocument.PageSetup.LeftMargin - ActiveDocument.PageSetup.RightMargin
    contentH = ActiveDocument.PageSetup.PageHeight - ActiveDocument.PageSetup.TopMargin - ActiveDocument.PageSetup.BottomMargin
    pageH = ActiveDocument.PageSetup.PageHeight

    hasCell = False
    If Selection.Information(wdWithInTable) Then
        hasCell = True
        cellW = Selection.Cells(1).Width - 8
        On Error Resume Next
        cellH = Selection.Cells(1).Height - 8
        On Error GoTo 0
    End If

    rng.Collapse wdCollapseStart
    rng.Paste

    If rng.InlineShapes.Count < 1 Then
        MsgBox "No inline picture detected after paste. Set Insert/paste pictures as: In line with text.", vbExclamation
        Exit Sub
    End If

    Set ils = rng.InlineShapes(1)
    ils.LockAspectRatio = msoTrue

    If style.MaxWidthMode = "fixed_cm" Then
        maxW = style.MaxWidthValue * CM_TO_POINTS
    Else
        If style.FitTableCell And hasCell Then
            maxW = cellW * (style.MaxWidthValue / 100#)
        Else
            maxW = contentW * (style.MaxWidthValue / 100#)
        End If
    End If

    If style.FitTableCell And hasCell Then
        maxW = MinD(maxW, cellW)
    Else
        maxW = MinD(maxW, contentW)
    End If

    If style.MaxHeightMode = "fixed_cm" Then
        maxH = style.MaxHeightValue * CM_TO_POINTS
    Else
        maxH = pageH * (style.MaxHeightValue / 100#)
    End If
    maxH = MinD(maxH, contentH)
    If style.FitTableCell And hasCell And cellH > 0 Then
        maxH = MinD(maxH, cellH)
    End If

    scaleW = maxW / ils.Width
    scaleH = maxH / ils.Height
    scale = MinD(scaleW, scaleH)
    If Not style.AllowUpscale Then
        If scale > 1 Then scale = 1
    End If

    ils.Width = ils.Width * scale

    Select Case LCase$(style.Align)
        Case "left": Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        Case "right": Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Case Else: Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End Select
End Sub
