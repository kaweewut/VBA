Attribute VB_Name = "Module1"
Sub Rotate()
    Dim strLine As String
    Dim fileName As String
    Dim folderPath As String
    Dim OldName, NewName As String
    Dim angle As Long
    Dim i, j As Integer
    Dim posX, posY, posI, posJ, valX, valY, valI, valJ As String
    Dim ExpireDate As Date
    Dim FileProps As Office.DocumentProperties
    
        'ExpireDate = #6/24/2025#
        Set FileProps = ThisWorkbook.CustomDocumentProperties
        
        On Error Resume Next
            ExpireDate = FileProps("ExpireDate").Value
        On Error GoTo 0
        
        If ExpireDate = 0 Then
            MsgBox "ExpireDate value not found in properties", vbCritical, "Value no found"
            Exit Sub
        End If
        
        If Date > ExpireDate Then
            MsgBox "Out of license", vbCritical, "Expire"
            Exit Sub
        End If

        folderPath = ThisWorkbook.Path & "\PG\"
        fileName = Dir(folderPath & "*.*")
        angle = Sheets("TransForm").Cells(2, 5)
        
        If angle = 90 Then
            i = 1
            j = -1
        ElseIf angle = -90 Then
            i = -1
            j = 1
        Else
            i = -1
            j = -1
        End If
            
        Do While fileName <> ""
            
            OldName = folderPath & fileName
            Open OldName For Input As #1
            NewName = Left(OldName, InStrRev(OldName, ".") - 1) & "_R" & angle & "." & Right(OldName, Len(OldName) - InStrRev(OldName, "."))
            Open NewName For Output As #2
            
            Do Until EOF(1)
                Line Input #1, strLine
                If InStr(strLine, "(") > 0 Or InStr(strLine, "G90G00G54") > 0 Or InStr(strLine, "G90G00X0.Y0.") > 0 Or InStr(strLine, "G43Z") > 0 Then
                    GoTo lastline
                ElseIf InStr(strLine, "X") = 0 And InStr(strLine, "Y") = 0 And InStr(strLine, "I") = 0 And InStr(strLine, "J") = 0 Then
                    GoTo lastline
                End If
                
                        posX = InStr(strLine, "X") + 1
                        valX = Mid(strLine, posX, Len(strLine) - posX + 1)
                        valX = Val(valX)
                        posY = InStr(strLine, "Y") + 1
                        valY = Mid(strLine, posY, Len(strLine) - posY + 1)
                        valY = Val(valY)
                        posI = InStr(strLine, "I") + 1
                        valI = Mid(strLine, posI, Len(strLine) - posI + 1)
                        valI = Val(valI)
                        posJ = InStr(strLine, "J") + 1
                        valJ = Mid(strLine, posJ, Len(strLine) - posJ + 1)
                        valJ = Val(valJ)
                       
                If angle = 90 Or angle = -90 Then
                        If InStr(strLine, "X") > 0 And InStr(strLine, "Y") > 0 Then
                            strLine = Replace(strLine, "X" & Format(valX, "0.###"), "X" & Format(j * valY, "0.###"))
                            strLine = Replace(strLine, "Y" & Format(valY, "0.###"), "Y" & Format(i * valX, "0.###"))
                        ElseIf InStr(strLine, "X") > 0 Then
                            strLine = Replace(strLine, "X" & Format(valX, "0.###"), "Y" & Format(i * valX, "0.###"))
                        ElseIf InStr(strLine, "Y") > 0 Then
                            strLine = Replace(strLine, "Y" & Format(valY, "0.###"), "X" & Format(j * valY, "0.###"))
                        End If
                        If InStr(strLine, "I") > 0 And InStr(strLine, "J") > 0 Then
                            strLine = Replace(strLine, "I" & Format(valI, "0.###"), "I" & Format(j * valJ, "0.###"))
                            strLine = Replace(strLine, "J" & Format(valJ, "0.###"), "J" & Format(i * valI, "0.###"))
                        ElseIf InStr(strLine, "I") > 0 Then
                            strLine = Replace(strLine, "I" & Format(valI, "0.###"), "J" & Format(i * valI, "0.###"))
                        ElseIf InStr(strLine, "J") > 0 Then
                            strLine = Replace(strLine, "J" & Format(valJ, "0.###"), "I" & Format(j * valJ, "0.###"))
                        End If
                Else
                        strLine = Replace(strLine, "X" & Format(valX, "0.###"), "X" & Format(i * valX, "0.###"))
                        strLine = Replace(strLine, "Y" & Format(valY, "0.###"), "Y" & Format(j * valY, "0.###"))
                        strLine = Replace(strLine, "I" & Format(valI, "0.###"), "I" & Format(i * valI, "0.###"))
                        strLine = Replace(strLine, "J" & Format(valJ, "0.###"), "J" & Format(j * valJ, "0.###"))
                End If

lastline:       Print #2, strLine
            Loop
        
            Close #1
            Close #2
            
            fileName = Dir()
        Loop

        MsgBox "Rotate data, Angle = " & angle & " is complete.", vbInformation
        
End Sub
