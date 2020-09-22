Attribute VB_Name = "QueenMOD"
Public Sub Create_Queen()
    SetPixel Stru.hDC, (Ant(0).AntX), (Ant(0).AntY), vbYellow
    SetPixel Stru.hDC, (Ant(0).AntX - 1), (Ant(0).AntY), vbYellow
    SetPixel Stru.hDC, (Ant(0).AntX + 1), (Ant(0).AntY), vbYellow
    SetPixel Stru.hDC, (Ant(0).AntX - 1), (Ant(0).AntY - 1), vbYellow
    SetPixel Stru.hDC, (Ant(0).AntX + 1), (Ant(0).AntY - 1), vbYellow
    SetPixel Stru.hDC, (Ant(0).AntX - 1), (Ant(0).AntY + 1), vbYellow
    SetPixel Stru.hDC, (Ant(0).AntX + 1), (Ant(0).AntY + 1), vbYellow
    SetPixel Stru.hDC, (Ant(0).AntX), (Ant(0).AntY - 1), vbYellow
    SetPixel Stru.hDC, (Ant(0).AntX), (Ant(0).AntY + 1), vbYellow
End Sub
