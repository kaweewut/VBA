Attribute VB_Name = "Module3"
Sub ToolSet()
Attribute ToolSet.VB_ProcData.VB_Invoke_Func = "l\n14"
Dim wb As Workbook
Dim ws As Worksheet
Dim a1, a2, a3, a4, a5, a6, a7, a8, a9, a10, a11, a12, a13, a14, a15 As Range
Dim b(1 To 15) As Variant
Dim Tooldata(1 To 400, 1 To 15) As Variant
Dim i, j, k, n, z, clTab As Integer
Dim MC As String
Dim Minval, Maxval As Double
Dim actWBName, tempWBName As String
Dim ToolSet As VbMsgBoxResult 'create tool set
Dim Job(1 To 4) As Variant ' for mold number,part name
Dim Folderpath As String
Dim pos As Integer
'Dim ExpireDate As Date
        
'        ExpireDate = #5/24/2025#

'        If Date > ExpireDate Then
'            MsgBox "Out of license", vbCritical, "Expire"
'            Exit Sub
'        End If

    k = 0
    z = 0
    Minval = 9999
    Maxval = -9999
    For Each wb In Application.Workbooks
        If wb.Name = "GIFU_ProcessSheet.xls" Then
            wb.Close Savechanges:=False
            Exit For
        End If
        If wb.Sheets(1).Cells(2, 1) = "Cutting time" Then
            If Minval > Val(wb.Sheets(1).Cells(3, 21).Value) Then
                Minval = Val(wb.Sheets(1).Cells(3, 21).Value)
            End If
            If Maxval < Val(wb.Sheets(1).Cells(3, 21).Value) Then
                Maxval = Val(wb.Sheets(1).Cells(3, 21).Value)
            End If
            Job(1) = wb.Sheets(1).Cells(9, 1)
            Job(2) = wb.Sheets(1).Cells(9, 5)
            Job(3) = wb.Sheets(1).Cells(9, 11)
            Job(4) = wb.Sheets(1).Cells(13, 5)
        End If
    Next wb
    For Each wb In Application.Workbooks ' keep data
            z = z + 1
            MC = wb.Sheets(1).Cells(9, 11)
            i = 0
        For Each ws In wb.Sheets
            If ws.Range("H1") = "H Geometry" Then ' 1 tool form
                Do While i < 400
                    If Tooldata(ws.Cells(3, 5) + i, 1) = "" Or (Tooldata(ws.Cells(3, 5) + i, 2) = ws.Cells(3, 17) And Tooldata(ws.Cells(3, 5) + i, 3) = ws.Cells(6, 17) _
                    And Tooldata(ws.Cells(3, 5) + i, 6) = ws.Cells(29, 15) _
                    And Tooldata(ws.Cells(3, 5) + i, 7) = ws.Cells(39, 15) And Tooldata(ws.Cells(3, 5) + i, 8) = ws.Cells(49, 15)) Then
                        Tooldata(ws.Cells(3, 5) + i, 1) = ws.Cells(3, 5)
                        Tooldata(ws.Cells(3, 5) + i, 2) = ws.Cells(3, 17)
                        Tooldata(ws.Cells(3, 5) + i, 3) = ws.Cells(6, 17)
                        Tooldata(ws.Cells(3, 5) + i, 4) = ws.Cells(11, 17)
                        If Tooldata(ws.Cells(3, 5) + i, 5) > 0 And (Tooldata(ws.Cells(3, 5) + i, 5) <= ws.Cells(23, 15) Or Tooldata(ws.Cells(3, 5) + i, 5) >= ws.Cells(23, 15)) Then
                            k = 1
                        End If
                        If ws.Cells(23, 15) <= 0 Then
                            ws.Cells(23, 15) = 1
                            ws.Cells(53, 15) = 1
                        End If
                        If Tooldata(ws.Cells(3, 5) + i, 5) < ws.Cells(23, 15) Then 'AL
                            Tooldata(ws.Cells(3, 5) + i, 5) = ws.Cells(23, 15)
                            Tooldata(ws.Cells(3, 5) + i, 10) = ws.Cells(59, 18)
                            Tooldata(ws.Cells(3, 5) + i, 11) = ws.Cells(53, 15)
                            Tooldata(ws.Cells(3, 5) + i, 12) = ws.Cells(43, 15)
                            Tooldata(ws.Cells(3, 5) + i, 13) = ws.Cells(63, 15)
                            Tooldata(ws.Cells(3, 5) + i, 14) = ws.Cells(33, 15)
                            Tooldata(ws.Cells(3, 5) + i, 15) = ws.Cells(23, 15)
                        End If
                        Tooldata(ws.Cells(3, 5) + i, 6) = ws.Cells(29, 15)
                        Tooldata(ws.Cells(3, 5) + i, 7) = ws.Cells(39, 15)
                        Tooldata(ws.Cells(3, 5) + i, 8) = ws.Cells(49, 15)
                        If ws.Cells(3, 30) = "G41" Or ws.Cells(3, 30) = "G42" Then 'for 2.5D re-machining
                            ws.Cells(11, 6) = ws.Cells(3, 5)
                        End If
                        If ws.Cells(11, 6) <> "" And (MC = "M852-G90" Or MC = "M852-G91" Or MC = "M852-M00") Then 'D Wear
                            Tooldata(ws.Cells(3, 5) + i, 9) = ws.Cells(3, 5) + 40
                        ElseIf ws.Cells(11, 6) <> "" And (MC = "MCD-G90" Or MC = "MCD-G91" Or MC = "MCD-M00") Then
                            Tooldata(ws.Cells(3, 5) + i, 9) = ws.Cells(3, 5) + 60
                        ElseIf ws.Cells(11, 6) <> "" Then
                            Tooldata(ws.Cells(3, 5) + i, 9) = ws.Cells(11, 6)
                        End If
                        Exit Do
                    End If

                    i = i + 100

                Loop
            ElseIf ws.Range("F1") = "H Geometry" Then ' 10 tool form
                'i = 0
                For j = 0 To 9
                Set a1 = ws.Cells(3 + 6 * j, 4)
                Set a2 = ws.Cells(3 + 6 * j, 16)
                Set a3 = ws.Cells(6 + 6 * j, 16)
                Set a4 = ws.Cells(3 + 6 * j, 18)
                Set a5 = ws.Cells(3 + 6 * j, 28)
                Set a6 = ws.Cells(3 + 6 * j, 41)
                Set a7 = ws.Cells(3 + 6 * j, 37)
                Set a8 = ws.Cells(6 + 6 * j, 18)
                Set a9 = ws.Cells(3 + 6 * j, 11)
                Set a10 = ws.Cells(5 + 6 * j, 34)
                Set a11 = ws.Cells(3 + 6 * j, 35)
                Set a12 = ws.Cells(3 + 6 * j, 39)
                Set a13 = ws.Cells(3 + 6 * j, 30)
                Set a14 = ws.Cells(3 + 6 * j, 43)
                Set a15 = ws.Cells(3 + 6 * j, 28)

                b(1) = a1.Value
                b(2) = a2.Value
                b(3) = a3.Value
                b(4) = a4.Value
                b(5) = a5.Value
                b(6) = a6.Value
                b(7) = a7.Value
                b(8) = a8.Value
                b(9) = a9.Value
                b(10) = a10.Value
                b(11) = a11.Value
                b(12) = a12.Value
                b(13) = a13.Value
                b(14) = a14.Value
                b(15) = a15.Value
                
                If b(9) >= 1 And (MC = "M852-G90" Or MC = "M852-G91" Or MC = "M852-M00") Then
                    b(9) = b(1) + 40
                ElseIf b(9) >= 1 And (MC = "MCD-G90" Or MC = "MCD-G91" Or MC = "MCD-M00") Then
                    b(9) = b(1) + 60
                ElseIf b(9) >= 1 Then
                    b(9) = b(1)
                End If
                    If b(1) <> "" Then
                    Do While i < 400
                        If Tooldata(b(1) + i, 1) = "" Or (Tooldata(b(1) + i, 2) = b(2) And Tooldata(b(1) + i, 3) = b(3) _
                        And Tooldata(b(1) + i, 6) = b(6) And Tooldata(b(1) + i, 7) = b(7) _
                        And Tooldata(b(1) + i, 8) = b(8)) Then
                            Tooldata(b(1) + i, 1) = b(1)
                            Tooldata(b(1) + i, 2) = b(2)
                            Tooldata(b(1) + i, 3) = b(3)
                            Tooldata(b(1) + i, 4) = b(4)
                            If Tooldata(b(1) + i, 5) > 0 And (Tooldata(b(1) + i, 5) <= b(5) Or Tooldata(b(1) + i, 5) >= b(5)) Then
                                k = 1
                            End If
                            If Tooldata(b(1) + i, 5) < b(5) Then 'AL
                                Tooldata(b(1) + i, 5) = b(5)
                                Tooldata(b(1) + i, 10) = b(10)
                                Tooldata(b(1) + i, 11) = b(11)
                                Tooldata(b(1) + i, 12) = b(12)
                                Tooldata(b(1) + i, 13) = b(13)
                                Tooldata(b(1) + i, 14) = b(14)
                                Tooldata(b(1) + i, 15) = b(5)
                            End If
                            Tooldata(b(1) + i, 6) = b(6)
                            Tooldata(b(1) + i, 7) = b(7)
                            Tooldata(b(1) + i, 8) = b(8)
                            If b(9) <> "" Then 'D Wear
                                Tooldata(b(1) + i, 9) = b(9)
                            End If
                            Exit Do
                        End If

                            i = i + 100

                    Loop
                    'i = 0
                    End If
                Next j
            End If
        Next ws
    Next wb
Dim ExpireDate As Date
        
        ExpireDate = #5/24/2025#
        
        If Date > ExpireDate Then
            MsgBox "Out of license", vbCritical, "Expire"
            Exit Sub
        End If
    
    For Each wb In Application.Workbooks 'ระบายสี
        i = 0
        For Each ws In wb.Sheets
            If ws.Range("H1") = "H Geometry" Then ' 1 tool form
                Do While i < 400
                    If Tooldata(ws.Cells(3, 5) + i, 2) = ws.Cells(3, 17) And Tooldata(ws.Cells(3, 5) + i, 3) = ws.Cells(6, 17) _
                    And Tooldata(ws.Cells(3, 5) + i, 6) = ws.Cells(29, 15) _
                    And Tooldata(ws.Cells(3, 5) + i, 7) = ws.Cells(39, 15) And Tooldata(ws.Cells(3, 5) + i, 8) = ws.Cells(49, 15) Then
                        If ws.Cells(23, 15) < Tooldata(ws.Cells(3, 5) + i, 5) Then 'AL less than
                            If Tooldata(ws.Cells(3, 5) + i, 9) >= 1 Then
                                ws.Cells(11, 6) = Tooldata(ws.Cells(3, 5) + i, 9) 'D Wear add
                            End If
                                ws.Range("L21:T66").Interior.ColorIndex = 15
                                ws.Tab.ColorIndex = 15
                                Exit Do
                        ElseIf ws.Cells(23, 15) = Tooldata(ws.Cells(3, 5) + i, 5) Then 'Al Max
                            If Tooldata(ws.Cells(3, 5) + i, 9) >= 1 Then
                                ws.Cells(11, 6) = Tooldata(ws.Cells(3, 5) + i, 9) 'D Wear add
                            End If
                                Tooldata(ws.Cells(3, 5) + i, 5) = 1000
                                Exit Do
                        End If
                    End If
                i = i + 100
                Loop
            ElseIf ws.Range("F1") = "H Geometry" Then ' 10 tool form
                'i = 0
                n = 0
                clTab = 0
                For j = 0 To 9
                    Set a1 = ws.Cells(3 + 6 * j, 4)
                    Set a2 = ws.Cells(3 + 6 * j, 16)
                    Set a3 = ws.Cells(6 + 6 * j, 16)
                    Set a4 = ws.Cells(3 + 6 * j, 18)
                    Set a5 = ws.Cells(3 + 6 * j, 28)
                    Set a6 = ws.Cells(3 + 6 * j, 41)
                    Set a7 = ws.Cells(3 + 6 * j, 37)
                    Set a8 = ws.Cells(6 + 6 * j, 18)
                    Set a9 = ws.Cells(3 + 6 * j, 11)
                    b(1) = a1.Value
                    b(2) = a2.Value
                    b(3) = a3.Value
                    b(4) = a4.Value
                    b(5) = a5.Value
                    b(6) = a6.Value
                    b(7) = a7.Value
                    b(8) = a8.Value
                    b(9) = a9.Value
                    If b(1) <> "" Then
                    Do While i < 400
                        If Tooldata(b(1) + i, 2) = b(2) And Tooldata(b(1) + i, 3) = b(3) _
                        And Tooldata(b(1) + i, 6) = b(6) And Tooldata(b(1) + i, 7) = b(7) And Tooldata(b(1) + i, 8) = b(8) Then
                                If ws.Cells(3 + 6 * j, 28) < Tooldata(b(1) + i, 5) Then 'AL less than
                                    ws.Cells(3 + 6 * j, 11) = Tooldata(b(1) + i, 9) ' D Wear add
                                    clTab = clTab + 1
                                    If j = 0 Then
                                        ws.Range("A3:AR8").Interior.ColorIndex = 15
                                    ElseIf j = 1 Then
                                        ws.Range("A9:AR14").Interior.ColorIndex = 15
                                    ElseIf j = 2 Then
                                        ws.Range("A15:AR20").Interior.ColorIndex = 15
                                    ElseIf j = 3 Then
                                        ws.Range("A21:AR26").Interior.ColorIndex = 15
                                    ElseIf j = 4 Then
                                        ws.Range("A27:AR32").Interior.ColorIndex = 15
                                    ElseIf j = 5 Then
                                        ws.Range("A33:AR38").Interior.ColorIndex = 15
                                    ElseIf j = 6 Then
                                        ws.Range("A39:AR44").Interior.ColorIndex = 15
                                    ElseIf j = 7 Then
                                        ws.Range("A45:AR50").Interior.ColorIndex = 15
                                    ElseIf j = 8 Then
                                        ws.Range("A51:AR56").Interior.ColorIndex = 15
                                    ElseIf j = 9 Then
                                        ws.Range("A57:AR62").Interior.ColorIndex = 15
                                    End If
                                    n = n + 1
                                    Exit Do
                                ElseIf ws.Cells(3 + 6 * j, 28) = Tooldata(b(1) + i, 5) Then 'AL Max
                                    ws.Cells(3 + 6 * j, 11) = Tooldata(b(1) + i, 9) ' D Wear add
                                    Tooldata(b(1) + i, 5) = 1000
                                    clTab = clTab + 1
                                    Exit Do
                                End If
                        End If
                            i = i + 100
                    Loop
                    'i = 0
                    End If
                Next j
                    If n = clTab Then
                        ws.Tab.ColorIndex = 15
                    End If
            End If
        Next ws
        If k = 1 Then
            wb.Sheets(1).Cells(47, 15) = wb.Sheets(1).Cells(54, 15)
        Else
            wb.Sheets(1).Cells(47, 15) = wb.Sheets(1).Cells(55, 15)
        End If
        If wb.Sheets(1).Cells(9, 11) = "A100-MX" Or wb.Sheets(1).Cells(9, 11) = "KBT-MX" Or wb.Sheets(1).Cells(9, 11) = "HMC10-MX" Then
            wb.Sheets(1).Cells(46, 15) = wb.Sheets(1).Cells(56, 15)
            wb.Sheets(1).Cells(46, 15) = Replace(wb.Sheets(1).Cells(46, 15).Value, "aaa", wb.Sheets(1).Cells(5, 7))
            wb.Sheets(1).Cells(46, 15) = Replace(wb.Sheets(1).Cells(46, 15).Value, "bbb", wb.Sheets(1).Cells(5, 9))
            wb.Sheets(1).Cells(43, 15) = wb.Sheets(1).Cells(59, 15)
        End If
        If wb.Sheets(1).Cells(9, 11) = "A100-G91S" Or wb.Sheets(1).Cells(9, 11) = "HMC10-G91S" Or wb.Sheets(1).Cells(9, 11) = "KBT-G91S" _
        Or wb.Sheets(1).Cells(9, 11) = "A100-G90S" Or wb.Sheets(1).Cells(9, 11) = "HMC10-G90S" Or wb.Sheets(1).Cells(9, 11) = "KBT-G90S" _
        Or wb.Sheets(1).Cells(9, 11) = "A100-M00S" Or wb.Sheets(1).Cells(9, 11) = "HMC10-M00S" Or wb.Sheets(1).Cells(9, 11) = "KBT-M00S" Then
            wb.Sheets(1).Cells(46, 15) = wb.Sheets(1).Cells(56, 15)
            wb.Sheets(1).Cells(46, 15) = Replace(wb.Sheets(1).Cells(46, 15).Value, "aaa", wb.Sheets(1).Cells(5, 7))
            wb.Sheets(1).Cells(46, 15) = Replace(wb.Sheets(1).Cells(46, 15).Value, "bbb", wb.Sheets(1).Cells(5, 9))

        End If
        If wb.Sheets(1).Cells(9, 11) = "A100(X)" Then
            wb.Sheets(1).Cells(43, 15) = wb.Sheets(1).Cells(59, 15)
        End If
        
        If k = 1 And z > 2 Then
            wb.Sheets(1).Cells(45, 15) = wb.Sheets(1).Cells(58, 15)
            wb.Sheets(1).Cells(45, 15) = Replace(wb.Sheets(1).Cells(45, 15).Value, "xxx", Minval)
            wb.Sheets(1).Cells(45, 15) = Replace(wb.Sheets(1).Cells(45, 15).Value, "yyy", Maxval)
        End If
        Application.DisplayAlerts = False
        wb.Save
    Next wb
    
    If k = 1 Then
        ToolSet = MsgBox("Do you want create tool set ?", vbYesNo, "Tool set create")
    End If
    If ToolSet = vbYes Then

        If InStr(ActiveWorkbook.Path, "NCP") > 0 Then
            If z > 2 Then
                pos = InStr(ActiveWorkbook.Path, "NCP")
                Folderpath = Left(ActiveWorkbook.Path, pos + Len("NCP") - 1)
                FileCopy "D:\TOOLS\Option\Gifu\ProcessSheet\Excel Tool\Cover.xlsx", Folderpath & "\Tool Set_" & Job(1) & "_" & Job(4) & "_Order" & Minval & "-" & Maxval & ".xlsx"
                Workbooks.Open Filename:=Folderpath & "\Tool Set_" & Job(1) & "_" & Job(4) & "_Order" & Minval & "-" & Maxval & ".xlsx"
            Else
                Folderpath = ActiveWorkbook.Path
                FileCopy "D:\TOOLS\Option\Gifu\ProcessSheet\Excel Tool\Cover.xlsx", Folderpath & "\Tool Set_" & Job(1) & "_" & Job(4) & "_Order" & Minval & ".xlsx"
                Workbooks.Open Filename:=Folderpath & "\Tool Set_" & Job(1) & "_" & Job(4) & "_Order" & Minval & ".xlsx"
            End If
        Else
            Folderpath = "D:"
            If z > 2 Then
                FileCopy "D:\TOOLS\Option\Gifu\ProcessSheet\Excel Tool\Cover.xlsx", Folderpath & "\Tool Set_" & Job(1) & "_" & Job(4) & "_Order" & Minval & "-" & Maxval & ".xlsx"
                Workbooks.Open Filename:=Folderpath & "\Tool Set_" & Job(1) & "_" & Job(4) & "_Order" & Minval & "-" & Maxval & ".xlsx"
            Else
                FileCopy "D:\TOOLS\Option\Gifu\ProcessSheet\Excel Tool\Cover.xlsx", Folderpath & "\Tool Set_" & Job(1) & "_" & Job(4) & "_Order" & Minval & ".xlsx"
                Workbooks.Open Filename:=Folderpath & "\Tool Set_" & Job(1) & "_" & Job(4) & "_Order" & Minval & ".xlsx"
            End If
        End If
        actWBName = ActiveWorkbook.Name
        Workbooks.Open Filename:="D:\TOOLS\Option\Gifu\ProcessSheet\Excel Tool\ToolSet.xlsx"
        tempWBName = ActiveWorkbook.Name
        Workbooks(actWBName).Activate
        i = 1
        j = 0

        Do While i < 400
            Do While j <= 9
                If j = 0 Or i = 1 Or i = 101 Or i = 201 Or i = 301 Then
                    If Tooldata(i, 1) <> "" Then
                        j = 0
                        Workbooks(tempWBName).Sheets("Sheet1").Copy After:=Workbooks(actWBName).Sheets(Sheets.Count)
                        Workbooks(actWBName).Activate
                        Set ws = Workbooks(actWBName).Sheets(Sheets.Count)
                        If i = 1 Then
                            ws.Cells(2, 1) = "Set1"
                        ElseIf i = 101 Then
                            ws.Cells(2, 1) = "Set2"
                        ElseIf i = 201 Then
                            ws.Cells(2, 1) = "Set3"
                        ElseIf i = 301 Then
                            ws.Cells(2, 1) = "Set4"
                        End If
                    End If
                End If
                If Tooldata(i, 1) = "" Then
                    i = i + 1
                Else
                    ws.Cells(3 + j * 6, 4) = Tooldata(i, 1)
                    If Tooldata(i, 2) >= 90 Then
                        ws.Cells(3 + j * 6, 16) = Tooldata(i, 2)
                    ElseIf Tooldata(i, 2) > 0 Then
                        ws.Cells(3 + j * 6, 16) = "R" & Tooldata(i, 2)
                    End If
                    ws.Cells(6 + j * 6, 16) = Tooldata(i, 3)
                    ws.Cells(3 + j * 6, 18) = Tooldata(i, 4)
                    ws.Cells(3 + j * 6, 41) = Tooldata(i, 6)
                    ws.Cells(3 + j * 6, 37) = Tooldata(i, 7)
                    ws.Cells(6 + j * 6, 18) = Tooldata(i, 8)
                    ws.Cells(3 + j * 6, 11) = Tooldata(i, 9)
                    ws.Cells(5 + j * 6, 34) = Tooldata(i, 10)
                    ws.Cells(3 + j * 6, 35) = Tooldata(i, 11)
                    ws.Cells(3 + j * 6, 39) = Tooldata(i, 12)
                    ws.Cells(3 + j * 6, 30) = Tooldata(i, 13)
                    ws.Cells(3 + j * 6, 43) = Tooldata(i, 14)
                    ws.Cells(3 + j * 6, 28) = Tooldata(i, 15)
                    j = j + 1
                    i = i + 1
                End If
          
                If i > 400 Then
                    GoTo mask
                End If
                If j = 10 Then
                    j = 0
                End If
            Loop
mask: Loop
        If z > 2 Then
            Workbooks(actWBName).Sheets(1).Cells(31, 15) = "For set tool in order " & Minval & "-" & Maxval
            Workbooks(actWBName).Sheets(1).Cells(3, 21) = Minval & "-" & Maxval
        Else
            Workbooks(actWBName).Sheets(1).Cells(31, 15) = "For set tool in order " & Minval
            Workbooks(actWBName).Sheets(1).Cells(3, 21) = Minval
        End If
        Workbooks(actWBName).Sheets(1).Cells(9, 1) = Job(1)
        Workbooks(actWBName).Sheets(1).Cells(9, 5) = Job(2)
        Workbooks(actWBName).Sheets(1).Cells(9, 11) = Job(3)
        Workbooks(actWBName).Sheets(1).Cells(13, 5) = Job(4)
        Workbooks(tempWBName).Activate
        ActiveWorkbook.Close
        Workbooks(actWBName).Activate
        ActiveWorkbook.Save
        MsgBox "Tool set created in " & ActiveWorkbook.Path & "\" & ActiveWorkbook.Name, vbInformation
    End If
End Sub
