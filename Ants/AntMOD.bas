Attribute VB_Name = "AntMOD"
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Type AntTYPE
    AntName As String
    AntAge As Long
    AntFunction As String
    AntTimer As Long
    AntHolding As Integer
    AntContents As Long
    AntX As Integer
    AntY As Integer
    AntTask As String
    AntFoodS As Integer
    AntSpeed As Single
    AntStopTime As Long
    AntIsDead As Integer
    AtDig As Integer
    MovingToTarget As Integer
    XTarget As Integer
    YTarget As Integer
End Type

Private Type StructureTYPE
    TheClock As Long
    Status As Integer
    FoodSupply As Single
    FoodTimer As Long
    hDC As Long
    Pop As Long
    Deads As Long
    XTarget As Integer
    YTarget As Integer
End Type

Private Type FoodTYPE
    SourceX As Integer
    SourceY As Integer
    Quantity As Long
End Type

Private Type EnemyTYPE
    X As Integer
    Y As Integer
    Target As Integer
    Health As Integer
    Speed As Integer
    Stop As Integer
    Stop2 As Integer
End Type

Public TimerV As Double

Public Ant(500) As AntTYPE
Public Enemy(5) As EnemyTYPE
Public Stru As StructureTYPE
Public Food As FoodTYPE
