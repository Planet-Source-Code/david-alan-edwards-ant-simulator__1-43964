Attribute VB_Name = "SoldierMOD"
Public Sub Head_To2(Index As Integer, X As Integer, Y As Integer)
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
    
    If Ant(Index).AntX >= (Enemy(0).X - 2) And Ant(Index).AntX <= (Enemy(0).X + 2) And Ant(Index).AntY >= (Enemy(0).Y - 2) And Ant(Index).AntY <= (Enemy(0).Y + 2) Then
        Ant(Index).AntStopTime = 0
        Enemy(0).Health = Enemy(0).Health - 1
    End If
    
    SetPixel Stru.hDC, Ant(Index).AntX, Ant(Index).AntY, RGB(255, 255, 0)
End Sub

