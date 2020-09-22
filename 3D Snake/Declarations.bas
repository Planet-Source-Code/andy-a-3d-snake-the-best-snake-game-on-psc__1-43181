Attribute VB_Name = "Declarations"
Public IsoEngine As New IsoEngine
Public Cam As New IECamera
Public Helper As New IEHelper
Public Music As New IEMusic
Public Sound As New IESound

'Map Variables
Public Map() As Integer
Public MapWidth As Integer
Public MapHeight As Integer

'Tile Data Variables
Public TexNames() As String
Public TexColors() As Long
Public TexTypes() As Byte
Public NumTex As Integer

'Status Variables
Public Level As Integer
Public Score As Long
Public Lives As Integer
Public FruitsLeft As Integer

'Fruit Variables
Public Type Fruit
    X As Single
    Y As Single
    FruitType As Byte
End Type
Public CurrentFruit As Fruit

'Snake Variables
Public Grow As Integer
Public CrashCounter As Integer
Public Enum Orientation
    TL = 0
    TR = 1
    BR = 2
    BL = 3
End Enum
Public Type Segment
    X As Single
    Y As Single
    Orientation As Orientation
End Type
Public Segs() As Segment
Public Direction() As D3DVECTOR2 'This is an array so it can store
                                 'a queue of directions (eg. the user
                                 'quickly presses up-left and then up-right)

'Timers
Public MoveTimer As Single

'Options
Public SnakeSpeed As Single
Public MusicVolume As Single
Public SoundVolume As Single
Public FruitsPerLevel As Integer
Public StartAtLevel As Integer

'Counters
Public i As Integer
Public j As Integer

'Other Variables
Public StopGame As Boolean
Public Pause As Boolean
