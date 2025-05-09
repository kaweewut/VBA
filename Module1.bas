Attribute VB_Name = "Module1"
Sub McChange()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim fileText As String
    Dim newText As String
    Dim CellA As Range
    Dim CellB As Range
    Dim valueA As String
    Dim valueB As String
    Dim i, MC As String
    Dim Angle, j, k, z, pos, AngleCode, HMC10Code, MCin, MCout As String 'for Program Angle
    Dim FileProps As Office.DocumentProperties
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set folder = fso.getfolder(ThisWorkbook.Path & "\PG") 'Open file in same workbook adress
        MC = "AAAA"
        MCin = "BBBB"
        MCout = "CCCC"
        
        For Each file In folder.Files
            k = 0
            z = 0
            Open file For Input As #1 'MC master Auto Check
                Do While Not EOF(1)
                    Line Input #1, textline
                    If InStr(1, textline, "(A100)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("A2")
                        Exit Do
                    ElseIf InStr(1, textline, "(GF8)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("B2")
                        Exit Do
                    ElseIf InStr(1, textline, "(KBT)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("C2")
                        Exit Do
                    ElseIf InStr(1, textline, "(M611)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("D2")
                        Exit Do
                    ElseIf InStr(1, textline, "(M852)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("E2")
                        Exit Do
                    ElseIf InStr(1, textline, "(M860)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("F2")
                        Exit Do
                    ElseIf InStr(1, textline, "(M1270)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("G2")
                        Exit Do
                    ElseIf InStr(1, textline, "(M2110)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("H2")
                        Exit Do
                    ElseIf InStr(1, textline, "(MCVA2)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("I2")
                        Exit Do
                    ElseIf InStr(1, textline, "(MCD)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("J2")
                        Exit Do
                    ElseIf InStr(1, textline, "(MCR)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("K2")
                        Exit Do
                    ElseIf InStr(1, textline, "(MCV)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("L2")
                        Exit Do
                    ElseIf InStr(1, textline, "(HMC10)", vbTextCompare) > 0 Then
                        Set CellA = Sheets("CODE").Range("M2")
                        Exit Do
                    End If
                Loop
            Close #1
            
        Dim ExpireDate As Date
            
        Set FileProps = ThisWorkbook.CustomDocumentProperties

        On Error Resume Next
        ExpireDate = FileProps("ExpireDate").Value
        On Error GoTo 0
        
        If ExpireDate = 0 Then
            MsgBox "ExpireDate value not found in properties", vbCritical, "Value no found"
            Exit Sub
        End If
        
        If Date > ExpireDate Then
            MsgBox "Out of license!", vbCritical, "Expire"
            Exit Sub
        End If
                
                Set CellB = Sheets("CODE").Range(Sheets("Transform").Cells(4, 5))
                 valueA = CellA.Value
                 valueB = CellB.Value
                 
                 If valueA = valueB Then 'Check program duplicate
                    MsgBox "Select duplicate machines", vbCritical
                    GoTo lastline
                    'Exit Sub
                End If
                
                If valueB = "M1270" Then
                    MC = valueB
                ElseIf valueB = "A100" Then
                    MCout = valueB
                    MCin = valueA
                ElseIf valueB = "KBT" Then
                    MCout = valueB
                    MCin = valueA
                ElseIf valueB = "HMC10" Then
                    MCout = valueB
                    MCin = valueA
                End If
                
                If valueB = "A100" Then
                    z = 1
                ElseIf valueB = "KBT" Then
                    z = 1
                ElseIf valueB = "HMC10" Then
                    z = 1
                End If
                
                
        If valueA = "M852" And valueB = "MCD" Then 'Problem clear in D wear
                Set CellA = CellA.Offset(31, 0)
                Set CellB = CellB.Offset(31, 0)
                valueA = CellA.Value
                valueB = CellB.Value
                i = -1
            ElseIf valueB = "MCR" Then 'Clear M201(X-END3)
                Set CellA = CellA.Offset(31, 0)
                Set CellB = CellB.Offset(31, 0)
                valueA = CellA.Value
                valueB = CellB.Value
                i = -1
            Else
                i = 1
        End If
                
                If valueA = "A100" Then 'Find Angle
                    k = 0
                    j = 1
                    Open file For Input As #1 'Angle Check
                    Do While Not EOF(1)
                    Line Input #1, textline
                        If InStr(1, textline, "G65P9000", vbTextCompare) > 0 Then
                            pos = InStr(textline, "A") + 1
                            Angle = Mid(textline, pos, Len(textline) - pos + 1)
                            Angle = Trim(Angle)
                            AngleCode = textline
                            k = 1
                            
                            Exit Do

                        End If
                        j = j + 1
                        If j >= 40 Then
                            Exit Do
                        End If
                    Loop
                    Close #1
                End If
                If valueA = "KBT" Then 'Find Angle
                    k = 0
                    j = 1
                    Open file For Input As #1 'Angle Check
                    Do While Not EOF(1)
                    Line Input #1, textline
                        If InStr(1, textline, "G111", vbTextCompare) > 0 Then
                            pos = InStr(textline, "A") + 1
                            Angle = Mid(textline, pos, Len(textline) - pos + 1)
                            Angle = Left(Angle, InStr(Angle, "B") - 1)
                            Angle = Trim(Angle)
                            AngleCode = textline
                            k = 1
                            
                            Exit Do

                        End If
                        j = j + 1
                        If j >= 40 Then
                            Exit Do
                        End If
                    Loop
                    Close #1
                End If
                If valueA = "HMC10" Then 'Find Angle
                    k = 0
                    j = 1
                    Open file For Input As #1 'Angle Check
                    Do While Not EOF(1)
                    Line Input #1, textline
                        If InStr(1, textline, "M217", vbTextCompare) > 0 Then
                            pos = InStr(textline, "B") + 1
                            Angle = Mid(textline, pos, Len(textline) - pos + 1)
                            Angle = Left(Angle, InStr(Angle, "S") - 1)
                            Angle = Trim(Angle)
                            AngleCode = textline
                            k = 1
                            
                            Exit Do

                        End If
                        j = j + 1
                        If j >= 40 Then
                            Exit Do
                        End If
                    Loop
                    Close #1
                End If
                If valueA = "HMC10" Then 'Find Angle
                    Open file For Input As #1 'Angle Check
                    Do While Not EOF(1)
                    Line Input #1, textline
                        If InStr(1, textline, "G54.1", vbTextCompare) > 0 Then
                            HMC10Code = textline
                            
                            Exit Do

                        End If
                    Loop
                    Close #1
                End If
                
                If k = 1 And z = 0 Then 'Machine not rotate angle

                        MsgBox file.Name & " have angle", vbCritical
                        GoTo lastline
                        'Exit Sub
                End If
                
            
                 fileText = getFileText(file.Path)
                If CellA <> CellB Then
                    newText = Replace(fileText, valueA, valueB)
                Else
                    newText = fileText
                End If
                 Set CellA = CellA.Offset(i, 0)
                 Set CellB = CellB.Offset(i, 0)
                 valueA = CellA.Value
                 valueB = CellB.Value
            Do While CellA <> "" And CellB <> ""
                    If valueA <> valueB Then
                        newText = Replace(newText, valueA, valueB)
                    Else
                    End If
                        Set CellA = CellA.Offset(i, 0)
                        Set CellB = CellB.Offset(i, 0)
                        valueA = CellA.Value
                        valueB = CellB.Value
            Loop
                If MCin = "HMC10" And k = 1 Then 'Master HMC10
                    'newText = Replace(newText, "(M25", "(MA25")
                    'newText = Replace(newText, "M25", "")
                    'newText = Replace(newText, "(MA25", "(M25")
                    'newText = Replace(newText, HMC10Code, "")
                    'newText = Replace(newText, "M24", "")
                    newText = Replace(newText, "" & vbCrLf & "M25" & vbCrLf & HMC10Code & vbCrLf & "M24" & vbCrLf & "", "")
                    If MCout = "A100" Then 'Case A100
                        newText = Replace(newText, AngleCode, "G65P9000 W54. C1. A" & Angle)
                    Else 'Case KBT
                        newText = Replace(newText, AngleCode, "G111 A" & Angle & " B58. C54.")
                    End If
                
                End If
                
                If MCout = "HMC10" And k = 1 Then 'Change to HMC10
                    

                        newText = Replace(newText, AngleCode, "M217 T54 B" & Angle & " S101" & vbCrLf & "" & vbCrLf & "M25" & vbCrLf & "G90G00G54.1P1B0" & vbCrLf & "M24")

                End If
                
                If k = 1 And MCin = "A100" And MCout = "KBT" Then
                    
                    newText = Replace(newText, AngleCode, "G111 A" & Angle & " B58. C54.")
                
                ElseIf k = 1 And MCin = "KBT" And MCout = "A100" Then
                    newText = Replace(newText, AngleCode, "G65P9000 W54. C1. A" & Angle)
                
                End If
                
                

                If MC = "M1270" Then 'M1270 name to M5070 problem
                    newText = Replace(newText, "M5070", "M1270")
                End If
                    setFileText file.Path, newText
lastline:        Next
        
        MsgBox "Transform is complete", vbInformation
        

End Sub

Function getFileText(filePath)
    Dim fso As Object
    Dim file As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set file = fso.OpenTextFile(filePath, 1)
        getFileText = file.ReadAll
        file.Close
End Function
Sub setFileText(filePath, text)
    Dim fso As Object
    Dim file As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set file = fso.OpenTextFile(filePath, 2, True)
        file.Write text
        file.Close
End Sub

