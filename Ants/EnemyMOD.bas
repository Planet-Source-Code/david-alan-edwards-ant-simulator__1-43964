Attribute VB_Name = "EnemyMOD"
Public Sub Spawn_Enemy(Index As Integer)
    Randomize
    rndpos = Int(Rnd * 4)
    If rndpos = 0 Then
        Enemy(Index).X = 10 'Int(Rnd * AntFRM.ColonyFRA.Width / Screen.TwipsPerPixelX)
        Enemy(Index).Y = 10 'Int(Rnd * AntFRM.ColonyFRA.Height / Screen.TwipsPerPixelY)
    ElseIf rndpos = 1 Then
        Enemy(Index).X = AntFRM.ColonyFRA.Width / Screen.TwipsPerPixelX - 10
        Enemy(Index).Y = AntFRM.ColonyFRA.Height / Screen.TwipsPerPixelY - 10
    ElseIf rndpos = 2 Then
        Enemy(Index).X = 10 'Int(Rnd * AntFRM.ColonyFRA.Width / Screen.TwipsPerPixelX)
        Enemy(Index).Y = AntFRM.ColonyFRA.Height / Screen.TwipsPerPixelY - 10
    ElseIf rndpos = 3 Then
        Enemy(Index).X = AntFRM.ColonyFRA.Width / Screen.TwipsPerPixelX - 10
        Enemy(Index).Y = 10 'Int(Rnd * AntFRM.ColonyFRA.Height / Screen.TwipsPerPixelY)
    End If
    Enemy(Index).Health = 10000
    Enemy(Index).Speed = 1
    Enemy(Index).Stop = 100
    Find_Ant Index
End Sub

Public Sub Find_Ant(Index As Integer)
     Randomize
oiawjfwa:
     Enemy(Index).Target = Int(Rnd * 499 + 1)
     If Ant(Enemy(Index).Target).AntName = "DEAD" Then GoTo oiawjfwa:
End Sub

Public Sub Head_To3(Index As Integer, X As Integer, Y As Integer)
    SetPixel Stru.hDC, Enemy(Index).X, Enemy(Index).Y, vbBlack
    SetPixel Stru.hDC, Enemy(Index).X + 1, Enemy(Index).Y, vbBlack
    SetPixel Stru.hDC, Enemy(Index).X - 1, Enemy(Index).Y, vbBlack
    SetPixel Stru.hDC, Enemy(Index).X + 1, Enemy(Index).Y + 1, vbBlack
    SetPixel Stru.hDC, Enemy(Index).X - 1, Enemy(Index).Y + 1, vbBlack
    SetPixel Stru.hDC, Enemy(Index).X + 1, Enemy(Index).Y - 1, vbBlack
    SetPixel Stru.hDC, Enemy(Index).X - 1, Enemy(Index).Y - 1, vbBlack
    SetPixel Stru.hDC, Enemy(Index).X, Enemy(Index).Y + 1, vbBlack
    SetPixel Stru.hDC, Enemy(Index).X, Enemy(Index).Y - 1, vbBlack
    
    If Enemy(Index).X <= X Then
        Enemy(Index).X = Enemy(Index).X + Enemy(Index).Speed ' + Rnd * 2 - 1
    End If
    If Enemy(Index).X >= X Then
        Enemy(Index).X = Enemy(Index).X - Enemy(Index).Speed ' + Rnd * 2 - 1
    End If
    If Enemy(Index).Y <= Y Then
       Enemy(Index).Y = Enemy(Index).Y + Enemy(Index).Speed ' + Rnd * 2 - 1
    End If
    If Enemy(Index).Y >= Y Then
        Enemy(Index).Y = Enemy(Index).Y - Enemy(Index).Speed ' + Rnd * 2 - 1
    End If
    
    If Enemy(0).X >= (Ant(Enemy(0).Target).AntX - 2) And Enemy(0).X <= (Ant(Enemy(0).Target).AntX + 2) And Enemy(0).Y >= (Ant(Enemy(0).Target).AntY - 2) And Enemy(0).Y <= (Ant(Enemy(0).Target).AntY + 2) Then
        Ant(Enemy(0).Target).AntX = Ant(0).AntX
        Ant(Enemy(0).Target).AntY = Ant(0).AntY
        Ant(Enemy(0).Target).AntName = "DEAD"
        'Enemy(Index).Stop = 100
        'Enemy(Index).Stop2 = 100
        Stru.Pop = Stru.Pop - 1
        Ant(Enemy(0).Target).AntIsDead = 1
        Find_Ant Index
    End If
    
    SetPixel Stru.hDC, Enemy(Index).X, Enemy(Index).Y, RGB(255, 0, 0)
    SetPixel Stru.hDC, Enemy(Index).X + 1, Enemy(Index).Y, RGB(255, 0, 0)
    SetPixel Stru.hDC, Enemy(Index).X - 1, Enemy(Index).Y, RGB(255, 0, 0)
    SetPixel Stru.hDC, Enemy(Index).X + 1, Enemy(Index).Y + 1, RGB(255, 0, 0)
    SetPixel Stru.hDC, Enemy(Index).X - 1, Enemy(Index).Y + 1, RGB(255, 0, 0)
    SetPixel Stru.hDC, Enemy(Index).X + 1, Enemy(Index).Y - 1, RGB(255, 0, 0)
    SetPixel Stru.hDC, Enemy(Index).X - 1, Enemy(Index).Y - 1, RGB(255, 0, 0)
    SetPixel Stru.hDC, Enemy(Index).X, Enemy(Index).Y + 1, RGB(255, 0, 0)
    SetPixel Stru.hDC, Enemy(Index).X, Enemy(Index).Y - 1, RGB(255, 0, 0)
End Sub

