Attribute VB_Name = "DiggerMOD"
Public Sub Head_To(Index As Integer, X As Integer, Y As Integer)
    SetPixel Stru.hDC, Ant(Index).AntX, Ant(Index).AntY, vbBlack
    
    If Ant(Index).AntX <= X Then
        Ant(Index).AntX = Ant(Index).AntX + Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    If Ant(Index).AntX >= X Then
        Ant(Index).AntX = Ant(Index).AntX - Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    If Ant(Index).AntY <= Y Then
       Ant(Index).AntY = Ant(Index).AntY + Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    If Ant(Index).AntY >= Y Then
        Ant(Index).AntY = Ant(Index).AntY - Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    
    'If Ant(Index).AntX >= (Stru.XTarget - 2) And Ant(Index).AntX <= (Stru.XTarget + 2) And Ant(Index).AntY >= (Stru.YTarget - 2) And Ant(Index).AntY <= (Stru.YTarget + 2) Then
    '    For III = 1 To 500
    '        If Ant(III).AntTask = "DIGGER" Then
    '            Ant(III).AntStopTime = 100
    '            Ant(III).AtDig = 1
    '            Ant(III).MovingToTarget = 0
    '            Choose_Digger_Target
    '        End If
    '    Next III
    'End If
    
    SetPixel Stru.hDC, Ant(Index).AntX, Ant(Index).AntY, RGB(0, 255, 255)
End Sub

Public Sub Choose_Digger_Target()
    Randomize
    Stru.XTarget = Int(Rnd * 600)
    Stru.YTarget = Int(Rnd * 400)
End Sub
