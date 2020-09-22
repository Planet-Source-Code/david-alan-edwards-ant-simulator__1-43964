Attribute VB_Name = "FoodMOD"
Public Sub Locate_Nearest_Food(Index As Integer)
    
End Sub

Public Sub Head_To_Food(Index As Integer)
    SetPixel Stru.hDC, Ant(Index).AntX, Ant(Index).AntY, vbBlack
    If Ant(Index).AntX <= Food.SourceX Then
        Ant(Index).AntX = Ant(Index).AntX + Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    If Ant(Index).AntX >= Food.SourceX Then
        Ant(Index).AntX = Ant(Index).AntX - Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    If Ant(Index).AntY <= Food.SourceY Then
       Ant(Index).AntY = Ant(Index).AntY + Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    If Ant(Index).AntY >= Food.SourceY Then
        Ant(Index).AntY = Ant(Index).AntY - Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    
    SetPixel Stru.hDC, Ant(Index).AntX, Ant(Index).AntY, RGB(255, 255, 255)
    
    If Ant(Index).AntX >= (Food.SourceX - 2) And Ant(Index).AntX <= (Food.SourceX + 2) And Ant(Index).AntY >= (Food.SourceY - 2) And Ant(Index).AntY <= (Food.SourceY + 2) Then
        If Food.Quantity >= 0 Then
            Ant(Index).AntTask = "RETURNING FOOD"
            Ant(Index).AntHolding = 1
            Ant(Index).AntContents = 12000
            Ant(Index).AntStopTime = 200
            Food.Quantity = Food.Quantity - 12000
        Else
            Food.Quantity = -1
            Ant(Index).AntTask = "RETURNING FOOD"
            Kill_Food
            Food.Quantity = 1000000
            Food.SourceX = Int(Rnd * AntFRM.ColonyFRA.Width / Screen.TwipsPerPixelX)
            Food.SourceY = Int(Rnd * AntFRM.ColonyFRA.Height / Screen.TwipsPerPixelY)
            Create_Food
        End If
    End If
End Sub

Public Sub Head_To_Queen(Index As Integer)
    SetPixel Stru.hDC, Ant(Index).AntX, Ant(Index).AntY, vbBlack
    If Ant(Index).AntX <= Ant(0).AntX Then
        Ant(Index).AntX = Ant(Index).AntX + Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    If Ant(Index).AntX >= Ant(0).AntX Then
        Ant(Index).AntX = Ant(Index).AntX - Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    If Ant(Index).AntY <= Ant(0).AntY Then
       Ant(Index).AntY = Ant(Index).AntY + Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    If Ant(Index).AntY >= Ant(0).AntY Then
        Ant(Index).AntY = Ant(Index).AntY - Ant(Index).AntSpeed + Rnd * 2 - 1
    End If
    
    If Ant(Index).AntX >= (Ant(0).AntX - 2) And Ant(Index).AntX <= (Ant(0).AntX + 2) And Ant(Index).AntY >= (Ant(0).AntY - 2) And Ant(Index).AntY <= (Ant(0).AntY + 2) Then
        Ant(Index).AntTask = "FOOD"
        Stru.FoodSupply = Stru.FoodSupply + Ant(Index).AntContents
        Ant(Index).AntContents = 0
        Ant(Index).AntStopTime = 200
        Ant(Index).AntHolding = 0
    End If
    
    SetPixel Stru.hDC, Ant(Index).AntX, Ant(Index).AntY, RGB(255, 255, 255)
End Sub

Public Sub Create_Food()
    SetPixel Stru.hDC, Food.SourceX, Food.SourceY, vbGreen
    SetPixel Stru.hDC, Food.SourceX - 1, Food.SourceY, vbGreen
    SetPixel Stru.hDC, Food.SourceX + 1, Food.SourceY, vbGreen
    SetPixel Stru.hDC, Food.SourceX - 1, Food.SourceY + 1, vbGreen
    SetPixel Stru.hDC, Food.SourceX + 1, Food.SourceY + 1, vbGreen
    SetPixel Stru.hDC, Food.SourceX - 1, Food.SourceY - 1, vbGreen
    SetPixel Stru.hDC, Food.SourceX + 1, Food.SourceY - 1, vbGreen
    SetPixel Stru.hDC, Food.SourceX, Food.SourceY + 1, vbGreen
    SetPixel Stru.hDC, Food.SourceX, Food.SourceY - 1, vbGreen
End Sub

Public Sub Kill_Food()
    SetPixel Stru.hDC, Food.SourceX, Food.SourceY, vbBlack
    SetPixel Stru.hDC, Food.SourceX - 1, Food.SourceY, vbBlack
    SetPixel Stru.hDC, Food.SourceX + 1, Food.SourceY, vbBlack
    SetPixel Stru.hDC, Food.SourceX - 1, Food.SourceY + 1, vbBlack
    SetPixel Stru.hDC, Food.SourceX + 1, Food.SourceY + 1, vbBlack
    SetPixel Stru.hDC, Food.SourceX - 1, Food.SourceY - 1, vbBlack
    SetPixel Stru.hDC, Food.SourceX + 1, Food.SourceY - 1, vbBlack
    SetPixel Stru.hDC, Food.SourceX, Food.SourceY + 1, vbBlack
    SetPixel Stru.hDC, Food.SourceX, Food.SourceY - 1, vbBlack
End Sub
