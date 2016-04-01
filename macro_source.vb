Sub A_Button_Click()
    Sheets("Game").Unprotect
    If Range("G6").Value > 13 Then
        If Range("G6").Value < 21 Then
            If Range("G5").Value > 5 Then
                If Range("G5").Value < 17 Then
                    Range("G7") = Range("G8")
                    Range("G5").Value = 17
                    Range("G6").Value = 11
                End If
            End If
            If Range("G5").Value > 23 Then
                If Range("G5").Value < 29 Then
                    Range("G7") = Range("G9")
                    Range("G5").Value = 17
                    Range("G6").Value = 11
                End If
            End If
        End If
    End If
    
    update
End Sub

Sub IO_Button_click()
    Sheets("Game").Unprotect
    Dim current_status As Integer
    If Range("G4").Value <> "1" Then
        Range("G4").Value = "1"
        current_status = 1
    Else
        Range("G4").Value = "0"
        current_status = 0
    End If
    
    If current_status = 1 Then
        update
    Else
        Range("AM6", "BV31").Interior.Color = RGB(0, 0, 0)
        Range("G5").Value = 17
        Range("G6").Value = 11
        Range("G7").Value = 0
        Range("G8").Value = 0
        Range("G9").Value = 0
        Range("G10").Value = ""
    End If
    
    
End Sub

Sub update()
    If Range("G4").Value = 0 Then
        Exit Sub
    End If
    Range("AM6", "BV31").Interior.Color = RGB(255, 255, 255)
    Range("AU8", "BO13").Interior.Color = RGB(182, 182, 182)
    Range("AU22", "AX25").Interior.Color = RGB(30, 230, 170)
    Range("BL22", "BO25").Interior.Color = RGB(30, 230, 170)
    
    Dim playerX As Integer
    Dim playerY As Integer
    playerY = Range("G5").Value
    playerX = Range("G6").Value
    Range("AM6").Offset(playerX, playerY).Interior.Color = RGB(50, 180, 210)
    Range("AM6").Offset(playerX + 1, playerY).Interior.Color = RGB(50, 180, 210)
    Range("AM6").Offset(playerX, playerY + 1).Interior.Color = RGB(50, 180, 210)
    Range("AM6").Offset(playerX + 1, playerY + 1).Interior.Color = RGB(50, 180, 210)
    Range("AU8").Value = Range("CB4").Offset(Range("G7").Value, 0)
    Range("G8").Value = Range("CB4").Offset(Range("G7").Value, 1)
    Range("G9").Value = Range("CB4").Offset(Range("G7").Value, 2)
    Sheets("Game").Protect
End Sub

Sub Up_Button_Click()
    Sheets("Game").Unprotect

    If Range("G4").Value = 0 Then
        Exit Sub
    End If
    
    If Range("G6").Value < 2 Then
        Exit Sub
    End If
    
    Range("G6").Value = Range("G6").Value - 2
    update
End Sub

Sub Down_Button_Click()
    Sheets("Game").Unprotect

    If Range("G4").Value = 0 Then
        Exit Sub
    End If
    
    If Range("G6").Value > 25 Then
        Exit Sub
    End If
    
    Range("G6").Value = Range("G6").Value + 2
    update
End Sub

Sub Left_Button_Click()
    Sheets("Game").Unprotect

    If Range("G4").Value = 0 Then
        Exit Sub
    End If
    
    
    If Range("G5").Value < 2 Then
        Exit Sub
    End If
    
    Range("G5").Value = Range("G5").Value - 2
    update
    
End Sub


Sub Right_Button_Click()
    Sheets("Game").Unprotect

    If Range("G4").Value = 0 Then
        Exit Sub
    End If
    
    If Range("G5").Value > 34 Then
        Exit Sub
    End If
    
    Range("G5").Value = Range("G5").Value + 2
    update
End Sub
