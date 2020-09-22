VERSION 5.00
Begin VB.Form AntFRM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ant Coloniser"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox HTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7440
      TabIndex        =   15
      Text            =   "0"
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox SC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Text            =   "0"
      Top             =   7770
      Width           =   495
   End
   Begin VB.TextBox DC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6240
      TabIndex        =   10
      Text            =   "0"
      Top             =   7770
      Width           =   495
   End
   Begin VB.TextBox FC 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6840
      TabIndex        =   9
      Text            =   "0"
      Top             =   7770
      Width           =   495
   End
   Begin VB.TextBox PopTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4320
      TabIndex        =   7
      Text            =   "0"
      Top             =   7770
      Width           =   975
   End
   Begin VB.TextBox Food2TXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Text            =   "0"
      Top             =   7770
      Width           =   975
   End
   Begin VB.TextBox FoodTXT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "0"
      Top             =   7770
      Width           =   975
   End
   Begin VB.CommandButton QuitBUT 
      Caption         =   "Quit"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Frame ColonyFRA 
      BackColor       =   &H00014294&
      BorderStyle     =   0  'None
      Caption         =   "The Colony"
      ForeColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      Begin VB.Timer EventsTIM 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   240
         Top             =   360
      End
   End
   Begin VB.CommandButton GoBUT 
      Caption         =   "Start Simulation"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Enemy Health"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7440
      TabIndex        =   16
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "S"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   14
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "D"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "F"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Population"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Available Food"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Food Supply"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   7560
      Width           =   975
   End
End
Attribute VB_Name = "AntFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ColonyFRA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim I As Integer
    If Button = 1 Then
        For I = 1 To 500
            Ant(I).AntX = X / Screen.TwipsPerPixelX
            Ant(I).AntY = Y / Screen.TwipsPerPixelY
        Next I
    ElseIf Button = 2 Then
        Food.Quantity = 1000000
        Food.SourceX = X / Screen.TwipsPerPixelX
        Food.SourceY = Y / Screen.TwipsPerPixelY
    End If
End Sub

Private Sub EventsTIM_Timer()
    Create_Queen
    Create_Food
    Dim I As Integer
    For I = 0 To 500
    If Ant(I).AntIsDead = 0 Then
        If Ant(I).AntStopTime <= 0 Then
            Ant(I).AntTimer = Ant(I).AntTimer - 1
            If Ant(I).AntTask = "FOOD" Then
                If Ant(I).AntFoodS = 0 Then
                    If Stru.FoodSupply >= 500000 Then
                        Randomize
                        CCC = Int(Rnd * 5)
                        If CCC = 2 Or CCC = 3 Then
                            Ant(I).AntTask = "DIGGER"
                            Ant(I).AntFoodS = 0
                        ElseIf CCC = 1 Or CCC = 0 Then
                            Ant(I).AntTask = "SOLDIER"
                            Ant(I).AntFoodS = 0
                        ElseIf CCC = 4 Then
                            Ant(I).AntTask = "FOOD"
                            Ant(I).AntFoodS = 1
                        End If
                    Else
                        Ant(I).AntTask = "FOOD"
                        Ant(I).AntFoodS = 0
                    End If
                End If
                Locate_Nearest_Food I
                    Head_To_Food I
            ElseIf Ant(I).AntTask = "RETURNING FOOD" Then
                Head_To_Queen I
            ElseIf Ant(I).AntTask = "DIGGER" Then
                If Stru.FoodSupply >= 500000 Then
                    Ant(I).AntTask = "DIGGER"
                    Ant(I).AntFoodS = 0
                Else
                    Ant(I).AntTask = "FOOD"
                    Ant(I).AntFoodS = 0
                End If
                    Randomize
                    Head_To I, Int(Rnd * 600), Int(Rnd * 600)
            ElseIf Ant(I).AntTask = "SOLDIER" Then
                If Stru.FoodSupply >= 500000 Then
                    Ant(I).AntTask = "SOLDIER"
                    Ant(I).AntFoodS = 0
                Else
                    Ant(I).AntTask = "FOOD"
                    Ant(I).AntFoodS = 0
                End If
                    Randomize
                    Head_To2 I, Enemy(0).X, Enemy(0).Y
            ElseIf Ant(I).AntTask = "CHILDREN" Then
                If Stru.Pop < 500 Then
                    Stru.Pop = Stru.Pop + 1
                    Randomize
                    AntRND = Int(Rnd * 2)
                    If AntRND = 0 Then
                        Ant(Stru.Pop).AntFunction = "SOLDER"
                        Ant(Stru.Pop).AntFoodS = 0
                    ElseIf AntRND = 1 Then
                        Ant(Stru.Pop).AntFunction = "WORKER"
                        Ant(Stru.Pop).AntFoodS = 0
                    End If
                    Ant(Stru.Pop).AntHolding = 0
                    Ant(Stru.Pop).AntName = "ALIVE"
                    Ant(Stru.Pop).AntTask = "FOOD"
                    Ant(Stru.Pop).AntTimer = 1000
                    Ant(Stru.Pop).AntX = 300
                    Ant(Stru.Pop).AntY = 300
                    Ant(Stru.Pop).AntIsDead = 0
                    Randomize
                    Randomize
                    Randomize
                    Ant(Stru.Pop).AntSpeed = Rnd * 2 + 0.5
                    Ant(Stru.Pop).AntStopTime = Int(Rnd * 400)
                    Ant(0).AntStopTime = 40
                    Ant(0).AntAge = 1000000
                Else
                    For II = 1 To 500
                        If Ant(I).AntName = "DEAD" Then
                            Ant(I).AntName = "ALIVE"
                            Exit For
                        End If
                    Next II
                End If
            End If
            
            If Ant(I).AntTimer <= 0 Then
                Ant(I).AntTimer = 100
            End If
            
            If Stru.FoodSupply <= 0 Then
                Stru.FoodTimer = Stru.FoodTimer - 1
                If Stru.FoodTimer <= 0 Then
                    Events_Stop
                End If
            End If
            Stru.FoodSupply = Stru.FoodSupply - 1.1
        Else
            Ant(I).AntStopTime = Ant(I).AntStopTime - 1
        End If
next1:
        Ant(I).AntAge = Ant(I).AntAge + 1
    End If
    
    If Ant(I).AntName = "ALIVE" Then
        If Ant(I).AntTask = "SOLDIER" Then
            SCV = SCV + 1
        ElseIf Ant(I).AntTask = "FOOD" Or Ant(I).AntTask = "RETURNING FOOD" Then
            FCV = FCV + 1
        ElseIf Ant(I).AntTask = "DIGGER" Then
            DCV = DCV + 1
        End If
    End If
    
    If TimerV >= 10 Then
        If I < 1 Then
            If Enemy(I).Health <= 0 Then
                Spawn_Enemy (I)
            Else
                Head_To3 I, Ant(Enemy(I).Target).AntX, Ant(Enemy(I).Target).AntY
            End If
        End If
    End If
    Next I
    
    DC.Text = DCV
    SC.Text = SCV
    FC.Text = FCV
    
    Food2TXT.Text = Food.Quantity
    PopTXT.Text = DCV + SCV + FCV
    FoodTXT.Text = Stru.FoodSupply
    Ant(0).AntIsDead = 0
    Ant(0).AntAge = 10000
    Ant(0).AntFunction = "QUEEN"
    Ant(0).AntHolding = 1
    Ant(0).AntName = "The Queen"
    Ant(0).AntTask = "CHILDREN"
    Ant(0).AntTimer = 1
    Ant(0).AntX = 300
    Ant(0).AntY = 300
    Ant(0).AntFoodS = 0
    Ant(0).AntIsDead = 0
    TimerV = TimerV + 1
    Me.Caption = "Ant Coloniser: " & TimerV
    HTXT.Text = Enemy(0).Health
    Enemy(Index).Stop = Enemy(Index).Stop - 1
    DCV = 0
    SCV = 0
    FCV = 0
End Sub

Private Sub Form_Load()
    Choose_Digger_Target
    ColonyFRA.BackColor = &H14294
    Stru.Pop = 100
    Me.Show
    Stru.hDC = GetDC(ColonyFRA.hwnd)
    Load_Ants
    Events_Start
End Sub

Private Sub Load_Ants()
    Dim I As Integer
    Dim AntRND As Integer
    
    Stru.FoodSupply = 2000000
    Stru.FoodTimer = 10000
    Stru.TheClock = 0
    Stru.Deads = 0
    
    FoodTXT.Text = Stru.FoodSupply
    
    Food.SourceX = 400
    Food.SourceY = 100
    Food.Quantity = 1000000
    
    Ant(0).AntAge = 1
    Ant(0).AntFunction = "QUEEN"
    Ant(0).AntHolding = 1
    Ant(0).AntName = "The Queen"
    Ant(0).AntTask = "CHILDREN"
    Ant(0).AntTimer = 100
    Ant(0).AntX = 300
    Ant(0).AntY = 300
    Ant(0).AntFoodS = 0
    Ant(0).AntIsDead = 0
    
    For I = 1 To 100
        Ant(I).AntAge = 0
        Randomize
        AntRND = Int(Rnd * 2)
        If AntRND = 0 Then
            Ant(I).AntFunction = "SOLDER"
            Ant(I).AntFoodS = 0
        ElseIf AntRND = 1 Then
            Ant(I).AntFunction = "WORKER"
            Ant(I).AntFoodS = 0
        End If
        Ant(I).AntHolding = 0
        Ant(I).AntName = "ALIVE"
        Ant(I).AntTask = "FOOD"
        Ant(I).AntTimer = 1000
        Ant(I).AntX = 300
        Ant(I).AntY = 300
        Randomize
        Randomize
        Randomize
        Ant(I).AntSpeed = Rnd * 2 + 0.5
        Ant(I).AntStopTime = Int(Rnd * 400)
    Next I
    For I = 101 To 500
        Ant(I).AntHolding = 0
        Ant(I).AntName = "DEAD"
        Ant(I).AntIsDead = 1
    Next I
End Sub

Private Sub Events_Start()
    Stru.Status = 1
    GoBUT.Caption = "Stop Simulation"
    EventsTIM.Enabled = True
End Sub

Private Sub Events_Stop()
    Stru.Status = 0
    GoBUT.Caption = "Start Simulation"
    EventsTIM.Enabled = False
End Sub

Private Sub GoBUT_Click()
    If GoBUT.Caption = "Start Simulation" Then
        Events_Start
    ElseIf GoBUT.Caption = "Stop Simulation" Then
        Events_Stop
    End If
End Sub

Private Sub QuitBUT_Click()
    Stru.Status = 0
    Unload Me
End Sub
