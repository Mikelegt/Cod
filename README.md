' -----------------
'| Why-Why Wizard  |
'| ©J.G.Harrington |
'|   v1.16 2008    |
' -----------------
'v1.02 - bugs corrected
'v1.03 - RPN added
'v1.04 - multiple language capability
'v1.05 - Polish added
'v1.06 - bug in font size changing fixed
'v1.07 - latin american spanish added - October 04
'v1.08 - cell selection/rightmouse disabled in help screen, Italian added Dec 04
'v1.08a - bug fixed - rpn table opening when not wanted Apr05
'v1.09 - arabic, with help from Waleed Shaban
'v1.10 - romanian, with help from Daniel Gombos Jun 05
'v1.11 - 2 blank sheets added request from Andrzej Mularski Targowek
'v1.12 - German what/why/root cause changed to was/warum/Ursache in popup menu - request Burkhard Buecken
'v1.13 - Portuguese & Dutch added, help from Marc Winkelman
'v1.14 - Portuguese-BRA added, help from Rogerio Macorin, bug fixed with centering date/rpn table Aug06
'v1.15 - Multiple countermeasures functionalty (reapplied from an old source)
'v1.16 - bugs fixed for office 2007

Option Explicit
Public Const _
    AdminPass As String = "x-x"

Public strType As String, iWhys As Integer, boEnableDelete As Boolean, _
    boWhat As Boolean, boRootCause As Boolean, iHeight As Integer, _
    iWidth As Integer, iFont As Integer, iCount1 As Integer, iCount2 As Integer, _
    boEliminate As Boolean
Dim rgWhy As Range, shp As Shape, rgLang As Range, iZoomArabic As Integer

Sub ResetPopups()

    Set rgLang = Sheets("Help").Range("languages").Columns(Sheets("Help").Range("lang_setting"))
    On Error Resume Next
    Application.CommandBars("Why Popup").Delete
    Application.CommandBars("Root Popup").Delete
    Application.CommandBars("Counter Popup").Delete
    Application.CommandBars("Reset Popup").Delete
    On Error GoTo 0
    With CommandBars.Add(Name:="Why Popup", Position:=msoBarPopup)
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "AddWhy"
            .Caption = rgLang.Cells(18)
            .FaceId = 3975
            If boEliminate = True Then .Enabled = False
        End With
        If boWhat = True Then
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "AddWhat"
                .Caption = rgLang.Cells(19)
                .FaceId = 2553
            End With
        End If
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "DeleteWhy"
            .Caption = rgLang.Cells(20)
            .FaceId = 358
            If boEnableDelete = False Then .Enabled = False
        End With
        If boWhat = False Then
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "AssignRootCause"
                .Caption = rgLang.Cells(21)
                If ActiveSheet.Range("align") = "Left2Right" Then
                    .FaceId = 2526
                Else
                    .FaceId = 2527
                End If
                If boEliminate = True Then .Enabled = False
                If boEnableDelete = False Or boRootCause = False Then .Enabled = False
            End With
        End If
        If boEliminate = False And boWhat = False Then
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "Eliminate"
                .Caption = rgLang.Cells(22)
                .FaceId = 1088
                If boEnableDelete = False Or boRootCause = False Then .Enabled = False
            End With
        ElseIf boEliminate = True And boWhat = False Then
            With .Controls.Add(Type:=msoControlButton)
                .OnAction = "UnEliminate"
                .Caption = rgLang.Cells(23)
                .FaceId = 1087
                If boEnableDelete = False Or boRootCause = False Then .Enabled = False
            End With
        End If
    End With
    With CommandBars.Add(Name:="Root Popup", Position:=msoBarPopup)
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "AssignCountermeasure"
            .Caption = rgLang.Cells(24)
            .FaceId = 328
            'If ActiveCell.Offset(1, 2) = "counter" Then .Enabled = False
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "UnassignRootCause"
            .Caption = rgLang.Cells(25)
            If ActiveSheet.Range("align") = "Left2Right" Then
                .FaceId = 2527
            Else
                .FaceId = 2526
            End If
            If ActiveCell.Offset(1, 2) = "counter" Then .Enabled = False
        End With
    End With
    With CommandBars.Add(Name:="Counter Popup", Position:=msoBarPopup)
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "DeleteCountermeasure"
            .Caption = rgLang.Cells(26)
            .FaceId = 358
        End With
    End With
    With CommandBars.Add(Name:="Reset Popup", Position:=msoBarPopup)
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "ShowViewControls"
            .Caption = rgLang.Cells(27)
            If fmView.Visible = True Then .Enabled = False
            .FaceId = 2089
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "Reset"
            .Caption = rgLang.Cells(28)
            .FaceId = 18
        End With
    End With

End Sub
Sub AddWhat()
    UnprotectMe
    Application.ScreenUpdating = False
    
'frig zoom to 100% to solve Bill Gates arabic bug!
    iZoomArabic = ActiveWindow.Zoom
    ActiveWindow.Zoom = 100
    
    Range("B65536").End(xlUp).Select
    Selection.Range("A1:A2").EntireRow.Insert
    With Selection.Range("A1:A2").EntireRow
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Cells(1, 14 + 2 * Range("why_count")).Range("A1:A2").Borders(xlEdgeRight).Weight = xlMedium
        .Locked = True
        .Font.Italic = False
    End With
    With ActiveCell
        .Locked = False
        .FormulaHidden = False
        .RowHeight = Range("height")
        .Font.Bold = False
        .Font.Size = Range("font")
        If ActiveSheet.Range("align") = "Left2Right" Then
            .HorizontalAlignment = xlLeft
        Else
            .HorizontalAlignment = xlRight
        End If
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.ColorIndex = 0
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    With ActiveCell.Offset(1, 0)
        .RowHeight = 5
        .Value = "what------00"
        .Font.Bold = False
        .Locked = True
        .FormulaHidden = True
        .Font.ColorIndex = 2
    End With
    ProtectMe
    ActiveWindow.Zoom = iZoomArabic
    AutoFit Range("autofit")
End Sub

Sub AddWhy()
Dim strStartShape As String, strEndShape As String, iDropCount As Integer, _
    strConnector As String, lgColWidth As Long, iConn1 As Integer, iConn2 As Integer
    UnprotectMe
    Application.ScreenUpdating = False
    
 'frig zoom to 100% to solve Bill Gates arabic bug!
    iZoomArabic = ActiveWindow.Zoom
    ActiveWindow.Zoom = 100

    iWidth = Range("width")
    iHeight = Range("height")
    lgColWidth = ActiveCell.Width
    iFont = Range("font")
    Set rgWhy = ActiveCell
    Range("last_conn") = Range("last_conn") + 1
    ActiveCell.Offset(1, 0) = Left(ActiveCell.Offset(1, 0), 8) & _
        "-" & Format(iWhys + 1, "00")
    ActiveCell.Offset(0, 2).Select
    If Cells(5, ActiveCell.Column) = Sheets("Help").Range("languages"). _
        Columns(Sheets("Help").Range("lang_setting")).Cells(8, 1) Then 'add another why column
        ActiveCell.Offset(0, -1).EntireColumn.Insert
        ActiveCell.Offset(0, -1).EntireColumn.Insert
        ActiveCell.ColumnWidth = iWidth
        ActiveCell.Offset(0, -1).ColumnWidth = 4
        Range("why_count") = Range("why_count") + 1
        Cells(5, ActiveCell.Column) = Sheets("Help").Range("languages"). _
            Columns(Sheets("Help").Range("lang_setting")).Cells(7, 1) & " " & Range("why_count")
       ' AutoFit Range("autofit")
    End If
    If iWhys > 0 Then 'find the last why
        iDropCount = 1
        Do Until iDropCount = iWhys
            ActiveCell.Offset(2, 0).Select
            If IsEmpty(ActiveCell.Offset(1, 0)) = False Then iDropCount = iDropCount + 1
        Loop
        ActiveCell.Offset(2, 0).Select
        
        'find next below or left full cell
        Do Until IsEmpty(ActiveCell.Offset(1, 0)) = False Or _
            ActiveCell.Offset(1, 0).End(xlToLeft).Column <> 1 Or _
            Cells(ActiveCell.Offset(1, 0).Row, 2).End(xlDown).Row = 65536
            ActiveCell.Offset(2, 0).Select
        Loop
        
        
        
        Selection.Range("A1:A2").EntireRow.Insert

        With Selection.Range("A1:A2").EntireRow
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Cells(1, 14 + 2 * Range("why_count")).Range("A1:A2").Borders(xlEdgeRight).Weight = xlMedium
            .Locked = True
            .Font.Italic = False
        End With
        Selection.RowHeight = iHeight
        Selection.Offset(1, 0).RowHeight = 5
    End If
    
    strStartShape = ActiveSheet.Shapes.AddShape(msoShapeFlowchartConnector, _
        rgWhy.Left + lgColWidth, rgWhy.Top + iHeight / 2, 1, 1).Name
    ActiveSheet.Shapes(strStartShape).Name = "start-" & Format(Range("last_conn"), "000")
    strStartShape = "start-" & Format(Range("last_conn"), "000")
    strEndShape = ActiveSheet.Shapes.AddShape(msoShapeFlowchartConnector, _
        ActiveCell.Left, ActiveCell.Top + iHeight / 2, 1, 1).Name
    ActiveSheet.Shapes(strEndShape).Name = "end-" & Format(Range("last_conn"), "000")
    strEndShape = "end-" & Format(Range("last_conn"), "000")
    ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 453#, 55.5, 4.5, 53.25) _
        .Select
    strConnector = "conn-" & Format(Range("last_conn"), "000")
    Selection.Name = strConnector
    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    Selection.ShapeRange.Flip msoFlipHorizontal
    Selection.ShapeRange.Flip msoFlipVertical
    If Sheets("Why-Why").Range("align") = "Left2Right" Then
        iConn1 = 7
        iConn2 = 3
    ElseIf Sheets("Why-Why").Range("align") = "Right2Left" Then
        iConn1 = 3
        iConn2 = 7
    End If
    Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes( _
        strStartShape), iConn1
    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes( _
        strEndShape), iConn2
    ActiveCell.Select
    With ActiveCell
        .Locked = False
        .FormulaHidden = False
        .Font.Bold = False
        .Font.Size = iFont
        If ActiveSheet.Range("align") = "Left2Right" Then
            .HorizontalAlignment = xlLeft
        Else
            .HorizontalAlignment = xlRight
        End If
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.ColorIndex = 0
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    With ActiveCell.Offset(1, 0)
        .Value = strConnector & "-00"
        .Font.Bold = False
        .Locked = True
        .FormulaHidden = True
        .Font.ColorIndex = 2
    End With
    CheckLastWhy
    ProtectMe
    ActiveWindow.Zoom = iZoomArabic
    AutoFit Range("autofit")
    
End Sub

Sub DeleteWhy()
Dim rgCounter As Range, iTreeCount As Integer, iBottom As Integer
    UnprotectMe
    Application.ScreenUpdating = False

    If Left(ActiveCell.Offset(1, 0), 5) = "conn-" Then 'delete why
    'check if its the top of a list
        Set rgWhy = ActiveCell
        If Val(Right(ActiveCell.Offset(1, -2), 2)) > 1 Then 'it is at the top of a list of whys
        'establish where the bottom of the tree is
            If Val(Right(ActiveCell.Offset(1, -2), 2)) = 2 Then
                iBottom = Range("B65536").End(xlUp).Row
                rgWhy.Offset(3, -2).Select
Start1:
                Do Until IsEmpty(ActiveCell) = False Or ActiveCell.Row > iBottom
                    ActiveCell.Offset(2, 0).Select
                Loop
                If ActiveCell.Row > iBottom Then
                    If ActiveCell.Column = 2 Then
                        iTreeCount = iBottom - 1
                        GoTo End1
                    End If
                Else
                    iTreeCount = ActiveCell.Row - 2
                    GoTo End1
                End If
                Cells(rgWhy.Row + 3, ActiveCell.Column - 2).Select
                GoTo Start1
End1:
            Else
                iTreeCount = ActiveCell.Row + 5
                Do Until Cells(iTreeCount, ActiveCell.Column) <> ""
                    iTreeCount = iTreeCount + 2
                Loop
                iTreeCount = iTreeCount - 2
            End If
            rgWhy.Select
        'delete top connector
            ActiveSheet.Shapes("conn-" & Mid(ActiveCell.Offset(1, 0), 6, 3)).Delete
            ActiveSheet.Shapes("start-" & Mid(ActiveCell.Offset(1, 0), 6, 3)).Delete
            ActiveSheet.Shapes("end-" & Mid(ActiveCell.Offset(1, 0), 6, 3)).Delete
        
       ' move the next line up
            Range(ActiveCell.Offset(2, 0), Cells(iTreeCount, 9 + 2 * Range("why_count"))).Select
            Selection.Cut
            rgWhy.Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
        'reconnect the connector
            ActiveSheet.Shapes("conn-" & Mid(ActiveCell.Offset(1, 0), 6, 3)).Select
            Selection.ShapeRange.ConnectorFormat.EndDisconnect
            Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet _
                .Shapes("end-" & Mid(ActiveCell.Offset(1, 0), 6, 3)), 3
        'reduce the why counter by 1
            Set rgCounter = ActiveCell.Offset(1, -2)
            rgCounter = Left(rgCounter, 9) & _
                Format((Val(Right(rgCounter, 2)) - 1), "00")
        'delete the 2 rows at the bottom of the tree
            Cells(iTreeCount - 1, ActiveCell.Column).Select
            ActiveCell.Range("A1:A2").EntireRow.Delete
        Else 'it's not at the top
        'delete cell, connector & borders
            ActiveSheet.Shapes("conn-" & Mid(ActiveCell.Offset(1, 0), 6, 3)).Delete
            ActiveSheet.Shapes("start-" & Mid(ActiveCell.Offset(1, 0), 6, 3)).Delete
            ActiveSheet.Shapes("end-" & Mid(ActiveCell.Offset(1, 0), 6, 3)).Delete
            
            With ActiveCell
                .Borders(xlEdgeLeft).LineStyle = xlNone
                .Borders(xlEdgeTop).LineStyle = xlNone
                .Borders(xlEdgeBottom).LineStyle = xlNone
                .Borders(xlEdgeRight).LineStyle = xlNone
                .Borders(xlDiagonalDown).LineStyle = xlNone
                .Borders(xlDiagonalUp).LineStyle = xlNone
                .ClearContents
                .Locked = True
            End With
            ActiveCell.Offset(1, 0).ClearContents
            ActiveCell.Offset(1, 0).Font.Italic = False
        'reduce the why counter by 1
            Set rgCounter = ActiveCell.Offset(1, -2)
            If IsEmpty(rgCounter) = True Then _
                Set rgCounter = ActiveCell.Offset(1, -2).End(xlUp)
            
            rgCounter = Left(rgCounter, 9) & _
                Format((Val(Right(rgCounter, 2)) - 1), "00")
        
        'check to left see if rows need deleting
        
            If ActiveCell.Offset(1, 0).End(xlToLeft).Column = 1 Then
                ActiveCell.Range("A1:A2").EntireRow.Delete
                AutoFit Range("autofit")
            End If
         'check if columns need deleting
            If ActiveCell.End(xlUp).Row = 5 And ActiveCell.End(xlDown).Row = 65536 Then
                ActiveCell.Offset(0, -1).Range("A1:B1").EntireColumn.Delete
                AutoFit Range("autofit")
                Range("why_count") = Range("why_count") - 1
            End If
        End If
    ElseIf Left(ActiveCell.Offset(1, 0), 5) = "what-" Then 'delete what
        ActiveCell.Range("A1:A2").EntireRow.Delete
        AutoFit Range("autofit")
    End If
    CheckLastWhy
    ProtectMe
End Sub
Sub AssignRootCause()
Dim rgRootCause As Range, strEndShape As String, iWhyNum As Integer, strWhyConn As String, _
    iConn1 As Integer, iConn2 As Integer
    UnprotectMe
    Application.ScreenUpdating = False
    
 'frig zoom to 100% to solve Bill Gates arabic bug!
    iZoomArabic = ActiveWindow.Zoom
    ActiveWindow.Zoom = 100

    Set rgWhy = ActiveCell
    Set rgRootCause = Cells(ActiveCell.Row, 4 + 2 * Range("why_count"))
    With rgRootCause
        .Value = rgWhy.Value
        .Locked = False
        .FormulaHidden = False
        .Font.Bold = False
        .Font.Size = Range("font")
        If ActiveSheet.Range("align") = "Left2Right" Then
            .HorizontalAlignment = xlLeft
        Else
            .HorizontalAlignment = xlRight
        End If
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Font.ColorIndex = 0
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    With rgWhy
        .ClearContents
        .ClearFormats
    End With
    strWhyConn = Left(rgWhy.Offset(1, 0), 8)
'add a new arrow
    iWhyNum = Val(Mid(rgWhy.Offset(1, 0), 6, 3))
    Range("last_conn") = Range("last_conn") + 1
    strEndShape = ActiveSheet.Shapes.AddShape(msoShapeFlowchartConnector, _
        rgRootCause.Left, rgRootCause.Top + Range("height") / 2, 1, 1).Name
    ActiveSheet.Shapes(strEndShape).Name = "end-" & Format(Range("last_conn"), "000")
    
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 172.5, 204#, 156.75, _
        2.25).Select
    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    Selection.ShapeRange.Flip msoFlipHorizontal
    Selection.ShapeRange.Flip msoFlipVertical
    If Sheets("Why-Why").Range("align") = "Left2Right" Then
        iConn1 = 7
        iConn2 = 3
    Else
        iConn1 = 3
        iConn2 = 7
    End If
    Selection.ShapeRange.ConnectorFormat.BeginConnect _
        ActiveSheet.Shapes("end-" & Format(iWhyNum, "000")), iConn1
    Selection.ShapeRange.ConnectorFormat.EndConnect _
        ActiveSheet.Shapes("end-" & Format(Range("last_conn"), "000")), iConn2
    Selection.Name = "conn-" & Format(Range("last_conn"), "000")
    ActiveSheet.Shapes(strWhyConn).Select
    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadNone

    rgRootCause.Select
    With rgRootCause.Offset(1, 0)
        .Value = "root-" & Format(Range("last_conn"), "000")
        .Font.Bold = False
        .Locked = True
        .FormulaHidden = True
        .Font.ColorIndex = 2
    End With
    CheckLastWhy
    ActiveWindow.Zoom = iZoomArabic
    ProtectMe
End Sub
Sub CheckLastWhy()
Dim boNoLastWhy As Boolean, rgLastWhy As Range
    boNoLastWhy = True
    For iCount1 = 7 To Range("B65536").End(xlUp).Row - 1 Step 2
        Set rgLastWhy = Cells(iCount1, 2 + 2 * Range("why_count"))
        
        If IsEmpty(rgLastWhy) = False And IsEmpty(rgLastWhy.Offset(0, 2)) _
            = True Then boNoLastWhy = False
    Next iCount1
    If boNoLastWhy = True Then 'hide last why column
        With Cells(5, 2 + 2 * Range("why_count"))
            .Font.ColorIndex = 2
            .Range("A1:B1").EntireColumn.ColumnWidth = 0.5
        End With
    Else
        With Cells(5, 2 + 2 * Range("why_count"))
            .Font.ColorIndex = 0
            .EntireColumn.ColumnWidth = Range("width")
            .Range("B1").EntireColumn.ColumnWidth = 4
        End With
    End If
    AutoFit Range("autofit")
End Sub
Sub Eliminate()
    UnprotectMe
    With ActiveCell.Borders(xlDiagonalDown)
        .LineStyle = xlContinuous
        .ColorIndex = 3
    End With
    With ActiveCell.Borders(xlDiagonalUp)
        .LineStyle = xlContinuous
        .ColorIndex = 3
    End With
    ActiveCell.Offset(1, 0).Font.Italic = True
    ProtectMe
End Sub
Sub UnEliminate()
    UnprotectMe
    ActiveCell.Borders(xlDiagonalDown).LineStyle = xlNone
    ActiveCell.Borders(xlDiagonalUp).LineStyle = xlNone
    ActiveCell.Offset(1, 0).Font.Italic = False
    ProtectMe
End Sub
Sub Reset()
Dim X As Integer
    If MsgBox(Sheets("Help").Range("languages").Cells(29, Sheets("Help").Range("lang_setting")), _
        vbCritical + vbOKCancel, "Why-Why Wizard") = vbCancel Then Exit Sub
    UnprotectMe
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Cells(6, 2).Select
    Do Until ActiveCell = "bottom"
        ActiveCell.EntireRow.Delete
    Loop
    Cells(5, 3).Select
    Do Until ActiveCell.Offset(0, 1) = Sheets("Help").Range("languages").Columns(Sheets("Help").Range("lang_setting")).Cells(8, 1)
        ActiveCell.EntireColumn.Delete
    Loop

    For Each shp In ActiveSheet.Shapes
        If shp.Name <> "helpbox" Then shp.Delete
    Next shp
    Range("last_conn") = 0
    Range("why_count") = 0
    Range("autofit") = False
    Range("rpn") = False
    Range("B3") = ""
    AddWhat
    UnprotectMe
    SetHeight 30
    SetWidth 25
    SetFont 10
    ActiveWindow.Zoom = 100
    Range("B3").Select
    Unload fmView
    UnprotectMe
    Columns("E:M").Hidden = True
    ProtectMe
    Application.EnableEvents = True
End Sub

Sub UnassignRootCause()

Dim rgRootCause As Range, strEndShape As String, strConn As String
    UnprotectMe
    Application.ScreenUpdating = False
    Set rgRootCause = ActiveCell
    Set rgWhy = ActiveCell.Offset(1, 0).End(xlToLeft).Offset(-1, 0)
    With rgWhy
        .Value = rgRootCause.Value
        .Locked = False
        .Font.Bold = False
        .Font.Size = Range("font")
        If ActiveSheet.Range("align") = "Left2Right" Then
            .HorizontalAlignment = xlLeft
        Else
            .HorizontalAlignment = xlRight
        End If
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    With rgRootCause
        .ClearContents
        .ClearFormats
    End With
    strConn = Right(rgRootCause.Offset(1, 0), 3)
    ActiveSheet.Shapes("conn-" & strConn).Delete
    ActiveSheet.Shapes("end-" & strConn).Delete
    
    rgRootCause.Offset(1, 0).ClearContents
    strConn = Mid(rgWhy.Offset(1, 0), 6, 3)
    ActiveSheet.Shapes("conn-" & strConn).Select
    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    rgWhy.Select
    CheckLastWhy
    ProtectMe
End Sub
Sub AssignCountermeasure()
    UnprotectMe
    If Range("rpn") = False Then
        Cells(1, 5 + Range("why_count") * 2).Range("A1:D1").EntireColumn.Hidden = False
    Else
        Cells(1, 5 + Range("why_count") * 2).Range("A1:I1").EntireColumn.Hidden = False
    End If



    'ADDED CODE *****************************
    Dim numCountMeasures As Integer
    
    ActiveCell.Offset(0, 2).Select
    Do Until IsEmpty(ActiveCell.Offset(1, 0)) Or (IsEmpty(ActiveCell.Offset(1, -2)) = False And numCountMeasures > 0)

        ActiveCell.Offset(2, 0).Select
        numCountMeasures = numCountMeasures + 1
    Loop
    ActiveCell.Offset(0, -2).Select

    
    'find next below or left full cell
    Do Until IsEmpty(ActiveCell.Offset(1, 0)) = False Or _
            ActiveCell.Offset(1, 0).End(xlToLeft).Column <> 1 Or _
            Cells(ActiveCell.Offset(1, 0).Row, 2).End(xlDown).Row = 65536
            ActiveCell.Offset(2, 0).Select
    Loop
    
    
    If numCountMeasures > 0 Then
    
        Selection.Range("A1:A2").EntireRow.Insert
        AutoFit Range("autofit")
        Selection.RowHeight = iHeight
        Selection.Offset(1, 0).RowHeight = 5
    End If
    
    '*****************************************

    With ActiveCell.Range("C1:E1")
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlHairline
        If ActiveSheet.Range("align") = "Left2Right" Then
            .HorizontalAlignment = xlLeft
        Else
            .HorizontalAlignment = xlRight
        End If
        .VerticalAlignment = xlCenter
        .Font.Size = Range("font")
        .Font.Bold = False
        .Font.ColorIndex = 0
        .Locked = False
        .FormulaHidden = False
        .Range("B1:C1").HorizontalAlignment = xlCenter
        .Range("C1").NumberFormat = "d-mmm-yy"
        With .Range("A2")
            .Value = "counter"
            .Font.ColorIndex = 2
            .Locked = True
            .Font.Bold = False
            .FormulaHidden = True
        End With
        .Select
        .Range("A1").Select
    End With
    With ActiveCell.Range("E1:H1")
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlHairline
        If ActiveSheet.Range("align") = "Left2Right" Then
            .HorizontalAlignment = xlLeft
        Else
            .HorizontalAlignment = xlRight
        End If
        .VerticalAlignment = xlCenter
        .Font.Size = Range("font")
        .Font.Bold = False
        .Font.ColorIndex = 0
        .NumberFormat = "0"
        .Locked = False
        .FormulaHidden = False
        .HorizontalAlignment = xlCenter
        With .Range("D1")
            .FormulaR1C1 = "=RC[-3]*RC[-2]*RC[-1]"
            .Font.Bold = True
            .Locked = True
        End With
    End With
    ProtectMe
    AutoFit Range("autofit")
End Sub
Sub DeleteCountermeasure()
    UnprotectMe
    
    With ActiveCell.Range("A1:H1")
        .ClearContents
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Locked = True
        .Range("A2").ClearContents
    End With
    If Cells(5, 6 + Range("why_count") * 2).End(xlDown).Row = 65536 Then
        If Range("rpn") = False Then
            Cells(1, 5 + Range("why_count") * 2).Range("A1:D1").EntireColumn.Hidden = True
        Else
            Cells(1, 5 + Range("why_count") * 2).Range("A1:I1").EntireColumn.Hidden = True
        End If
        AutoFit Range("autofit")
    End If
    ActiveCell.Offset(0, -2).Select
    
    If IsEmpty(ActiveCell.Offset(1, 0)) Then
        Selection.Range("A1:A2").EntireRow.Delete
        AutoFit Range("autofit")
    End If
    
    ProtectMe
End Sub
Sub ShowViewControls()
    fmView.Show 0
End Sub
Sub SetWidth(iCurrentWidth)
    UnprotectMe
    For iCount1 = 2 To 6 + 2 * Range("why_count") Step 2
        If Cells(1, iCount1).EntireColumn.ColumnWidth > 1 Then _
            Cells(1, iCount1).EntireColumn.ColumnWidth = iCurrentWidth
    Next iCount1
    Range("width") = iCurrentWidth
    ProtectMe
    AutoFit Range("autofit")
End Sub

Sub SetHeight(iCurrentHeight)
    UnprotectMe
    For iCount1 = 6 To Range("B65536").End(xlUp).Row - 1 Step 2
        Cells(iCount1, 1).EntireRow.RowHeight = iCurrentHeight
    Next iCount1
    Range("height") = iCurrentHeight
    For Each shp In ActiveSheet.Shapes
        If shp.AutoShapeType = msoShapeFlowchartConnector Then _
            shp.Top = shp.TopLeftCell.Top + iCurrentHeight / 2
    Next shp
    ProtectMe
    AutoFit Range("autofit")
End Sub
Sub SetFont(iCurrentFont)
    UnprotectMe
    For iCount2 = 6 To Range("B65536").End(xlUp).Row - 1 Step 2
        For iCount1 = 2 To 6 + 2 * Range("why_count") Step 2
            Cells(iCount2, iCount1).Font.Size = iCurrentFont
        Next iCount1
        Cells(iCount2, iCount1 - 1).Font.Size = iCurrentFont
        Cells(iCount2, iCount1).Font.Size = iCurrentFont
        Cells(iCount2, iCount1 + 2).Range("A1:D1").Font.Size = iCurrentFont
    Next iCount2
    Range("B3").Font.Size = iCurrentFont
    Range("font") = iCurrentFont
    If Cells(1, 6 + 2 * Range("why_count")).EntireColumn.Hidden = False Then _
        Cells(1, 7 + 2 * Range("why_count")).Range("A1:B1").EntireColumn.AutoFit
    If Cells(1, 10 + 2 * Range("why_count")).EntireColumn.Hidden = False Then _
        Cells(1, 10 + 2 * Range("why_count")).Range("A1:D1").EntireColumn.AutoFit
    ProtectMe
End Sub
Sub ProtectMe()
    ActiveSheet.Protect AdminPass
    Application.EnableEvents = True
End Sub
Sub UnprotectMe()
    Application.EnableEvents = False
    ActiveSheet.Unprotect AdminPass
End Sub
Sub AutoFit(boYes As Boolean)
Dim iZoom As Integer, rgTarget As Range
    If boYes = False Then Exit Sub
    Set rgTarget = ActiveCell
    Range("Print_Area").Select
    ActiveWindow.Zoom = True
    Range("A1").Select
    iZoom = ActiveWindow.Zoom
    ActiveWindow.Zoom = iZoom - iZoom Mod 5
    fmView.lblZoom = ActiveWindow.Zoom
    Range("A1").Select
    rgTarget.Select
End Sub
Sub HelpGo()
Dim obSheet As Worksheet
    Application.ScreenUpdating = False
    ThisWorkbook.Unprotect AdminPass
    Sheets("Help").Visible = True
    For Each obSheet In ThisWorkbook.Worksheets
        If obSheet.Name <> "Help" Then obSheet.Visible = False
    Next obSheet
    Sheets("Help").Select
    Range("A1").Select
    ThisWorkbook.Protect AdminPass
End Sub
Sub HelpBack()
Dim obSheet As Worksheet
    Application.ScreenUpdating = False
    ThisWorkbook.Unprotect AdminPass
    For Each obSheet In ThisWorkbook.Worksheets
        If obSheet.Name <> "Help" Then obSheet.Visible = True
    Next obSheet
    Sheets("Help").Visible = False
    Sheets("Why-Why").Select
    ThisWorkbook.Protect AdminPass
End Sub
Sub SetLanguage()
Dim rgLanguage As Range, iLanguage As Integer
    Application.ScreenUpdating = False
    Set rgWhy = ActiveCell
'set language ranges
    Application.EnableEvents = False
    ThisWorkbook.Unprotect AdminPass
    Sheets("Help").Visible = True
    Sheets("Help").Select
    Sheets("Help").Unprotect AdminPass
    iLanguage = fmView.cmbLanguage.ListIndex + 1
    Sheets("Help").Range("lang_setting") = iLanguage
    iCount1 = Sheets("Help").Range("languages").Rows.Count
    Set rgLanguage = Sheets("Help").Range("languages") _
        .Range(Cells(1, iLanguage), Cells(iCount1, iLanguage))
'update help sheet
    ActiveSheet.Shapes("back").Select
    Selection.Characters.Text = rgLanguage(45)
    With Sheets("Help")
        .Cells(5, 4) = rgLanguage(30)
        .Cells(7, 4) = rgLanguage(31)
        .Cells(8, 4) = rgLanguage(32)
        .Cells(9, 4) = rgLanguage(33)
        .Cells(10, 4) = rgLanguage(34)
        .Cells(11, 4) = rgLanguage(35)
        .Cells(12, 4) = rgLanguage(36)
        .Cells(13, 4) = rgLanguage(37)
        .Cells(15, 4) = rgLanguage(38)
        .Cells(16, 4) = rgLanguage(39)
        .Cells(17, 4) = rgLanguage(40)
        .Cells(18, 4) = rgLanguage(41)
        .Cells(19, 4) = rgLanguage(42)
        .Cells(20, 4) = rgLanguage(43)
        .Cells(21, 4) = rgLanguage(44)
        .Rows("7:21").Rows.AutoFit
        .Range("A1").Select
    End With

'update why-why sheet
    Sheets("Why-Why").Select
    ActiveSheet.Unprotect AdminPass
    ActiveSheet.Shapes("helpbox").Select
    Selection.Characters.Text = rgLanguage(5)
    With Sheets("Why-Why")
        .Unprotect AdminPass
        .Range("A1") = rgLanguage(2)
        .Range("B2") = rgLanguage(4)
        .Range("B5") = rgLanguage(6)
        For iCount1 = 1 To .Range("why_count")
            .Cells(5, 2 + iCount1 * 2) = rgLanguage(7) & " " & Format(iCount1, "0")
        Next iCount1
        With .Cells(1, 4 + .Range("why_count") * 2)
            .Cells(5, 1) = rgLanguage(8)
            .Cells(5, 3) = rgLanguage(9)
            .Cells(5, 4) = rgLanguage(10)
            .Cells(5, 5) = rgLanguage(11)
            .Cells(5, 7) = rgLanguage(12)
            .Cells(5, 8) = rgLanguage(13)
            .Cells(5, 9) = rgLanguage(14)
            .Cells(5, 10) = rgLanguage(15)
            .Cells(3, 3) = rgLanguage(16)
            .Cells(3, 7) = rgLanguage(17)
        End With
        If Cells(1, 6 + 2 * Range("why_count")).EntireColumn.Hidden = False Then _
            Cells(1, 7 + 2 * Range("why_count")).Range("A1:B1").EntireColumn.AutoFit
        If fmView.chRPN = True Then _
            Cells(1, 7 + 2 * Range("why_count")).Range("D1:G1").EntireColumn.AutoFit
    End With
    rgWhy.Select
    
    'portion of code added to handle right to left languages like arabic
    With Sheets("Why-Why")
        If rgLanguage(57) = "Left2Right" Then
            If ActiveSheet.DisplayRightToLeft = True Then
                ActiveSheet.DisplayRightToLeft = False
            End If
            Sheets("Sheet1").DisplayRightToLeft = False
            Sheets("Sheet2").DisplayRightToLeft = False
            .Range("align") = "Left2Right"
            .Range("Print_Area").HorizontalAlignment = xlLeft
            .Cells(1, 7 + 2 * Range("why_count")).Range("A1:G1").EntireColumn.HorizontalAlignment = xlCenter
            .Rows("3:5").HorizontalAlignment = xlCenter
            .Range("B3").HorizontalAlignment = xlLeft
        Else
            If ActiveSheet.DisplayRightToLeft = False Then
                ActiveSheet.DisplayRightToLeft = True
            End If
            Sheets("Sheet1").DisplayRightToLeft = True
            Sheets("Sheet2").DisplayRightToLeft = True
            .Range("align") = "Right2Left"
            .Range("Print_Area").HorizontalAlignment = xlRight
            .Cells(1, 7 + 2 * Range("why_count")).Range("A1:G1").EntireColumn.HorizontalAlignment = xlCenter
            .Rows("3:5").HorizontalAlignment = xlCenter
            .Range("B3").HorizontalAlignment = xlRight
        End If
    End With
    
    Sheets("Help").Protect AdminPass
    Sheets("Help").Visible = False
    Sheets("Why-Why").Protect AdminPass
    ThisWorkbook.Protect AdminPass
    Application.EnableEvents = True
    Application.ScreenUpdating = False
End Sub
