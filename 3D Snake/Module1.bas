Attribute VB_Name = "Module1"
Public Sub LoadIsoEngine()
    'Set up IsoEngine
    IsoEngine.SetAppPath App.Path & "\"
    IsoEngine.SetLogFilename "Log.txt"
    
        'Set resolution to 800x600x24 or 800x600x16
        If IsoEngine.SetDisplayMode(800, 600, 32) = False Then
            IsoEngine.SetDisplayMode 800, 600, 16
        End If
        
    IsoEngine.Initialize frmMain.picGame.hWnd, frmMain.picGame.Width, frmMain.picGame.Height, , D3DFMT_D32
    IsoEngine.SetTileSize 128, 64
    
    'Set up terrain
    IsoEngine.LoadTextureInfo "Tile Image Data.tex", TexNames, TexColors, TexTypes, NumTex
    IsoEngine.LoadTileTextures "Graphics\Tile ", ".bmp"
    
    'Load textures
    IsoEngine.LoadTexture "Graphics\Fruit 1.bmp", "f1", Black
    IsoEngine.LoadTexture "Graphics\Fruit 2.bmp", "f2", Black
    IsoEngine.LoadTexture "Graphics\Fruit 3.bmp", "f3", Black
    IsoEngine.LoadTexture "Graphics\Head.bmp", "head", Black '32
    IsoEngine.LoadTexture "Graphics\Body.bmp", "body", Black 'x
    IsoEngine.LoadTexture "Graphics\Tail.bmp", "tail", Black '24
    
    'Load fonts
    IsoEngine.CreateFont "Status", "Arial", 10, True, False, False
    
    'Load music
    Music.Initialize frmMain.picGame.hWnd
    Music.Playlist_AddItem "Sounds\Game.mid"
    
    'Initialize sound engine
    Sound.Initialize frmMain.picGame.hWnd
    
    'Set up camera
    Cam.AutoCheckCamPos = True
    Cam.ShowBorder = False
    Cam.WrapWorld = False
End Sub

Public Sub LoadLevel(Optional ByVal RestartLevel As Boolean = False)
    If Not RestartLevel Then
        'Loads the map of the current level
        IsoEngine.LoadMapFromFile Level & ".map", Map, MapWidth, MapHeight
    End If
    
    'Resizes the snake to 2 segments (head and tail)
    ReDim Segs(2)
    
    'Sets the startup position and orientation of the snake
    Select Case Level
        Case 0
            Segs(1).X = 16.5 * 64
            Segs(1).Y = 24 * 32
            Segs(1).Orientation = TR
        Case 1, 2, 5
            Segs(1).X = 392
            Segs(1).Y = 820
            Segs(1).Orientation = TR
        Case 3
            Segs(1).X = 576
            Segs(1).Y = 820
            Segs(1).Orientation = TR
        Case 4
            Segs(1).X = 276
            Segs(1).Y = 896
            Segs(1).Orientation = TR
        Case 6
            Segs(1).X = 8 * 64
            Segs(1).Y = 28 * 32
            Segs(1).Orientation = TR
        Case 7
            Segs(1).X = 12 * 64
            Segs(1).Y = 26 * 32
            Segs(1).Orientation = TR
        Case 8
            Segs(1).X = 12 * 64
            Segs(1).Y = 26 * 32
            Segs(1).Orientation = TL
        Case 9
            Segs(1).X = 3 * 64
            Segs(1).Y = 7 * 32
            Segs(1).Orientation = TR
        Case 10
            Segs(1).X = 5 * 64
            Segs(1).Y = 15 * 32
            Segs(1).Orientation = TR
    End Select
    
    'Automatically places the tail of the snake based
    'on the position and orientation of the head
    ReDim Direction(1)
    Select Case Segs(1).Orientation
        Case TL
            Segs(2).Orientation = BR
            Segs(2).X = Segs(1).X + 24
            Segs(2).Y = Segs(1).Y + 12
            Direction(1).X = -2
            Direction(1).Y = -1
        Case TR
            Segs(2).Orientation = BL
            Segs(2).X = Segs(1).X - 24
            Segs(2).Y = Segs(1).Y + 12
            Direction(1).X = 2
            Direction(1).Y = -1
        Case BL
            Segs(2).Orientation = TR
            Segs(2).X = Segs(1).X + 24
            Segs(2).Y = Segs(1).Y - 12
            Direction(1).X = -2
            Direction(1).Y = 1
        Case BR
            Segs(2).Orientation = TL
            Segs(2).X = Segs(1).X - 24
            Segs(2).Y = Segs(1).Y - 12
            Direction(1).X = 2
            Direction(1).Y = 1
    End Select
    
    'Resets some variables
    Grow = 2
    CrashCounter = -1
    
    If Not RestartLevel Then
        'Gives the user a chance to get ready for the next level
        Render
        MsgBox "Press <space> to start level " & Level & ".", vbInformation, "Level " & Level - 1 & " Complete"
        frmMain.picGame.SetFocus
        Music.Volume = MusicVolume
        Music.Play True
    End If
End Sub

Sub PlaceFruit()
    Randomize Timer
    Do
        i = Int(Rnd(1) * 32) + 1
        j = Int(Rnd(1) * 32) + 1
        If Map(i, j) = 2 And ((i + 1) * 128 <> CurrentFruit.X And (j + 1) * 64 <> CurrentFruit.Y) Then
            CurrentFruit.X = (i) * 64 - 64
            CurrentFruit.Y = (j) * 32 - 64
            CurrentFruit.FruitType = Int(Rnd(1) * 3) + 1
            Exit Do
        End If
    Loop
End Sub

Public Sub ResetGame()
    ReDim Segs(0)
    Lives = 5
    Score = 0
    Level = StartAtLevel
    FruitsLeft = FruitsPerLevel
    LoadLevel
    PlaceFruit
End Sub

Public Sub RenderLoop()
    MoveTimer = Timer
    On Error Resume Next
    If StopGame Then IsoEngine.RenderToScreen
    StopGame = False
    frmMain.picGame.SetFocus
    Do
        DoEvents
        If Timer - MoveTimer >= SnakeSpeed And Pause = False Then
            MoveSnake
            MoveTimer = Timer
        End If
        Render
    Loop Until StopGame = True
    Music.Stop_
    Sound.Stop_
End Sub

Sub Render()
    IsoEngine.Clear
        Cam.CenterAtPixel Segs(1).X, Segs(1).Y
        IsoEngine.DrawMap
        IsoEngine.Sprite_Begin
            For i = 1 To UBound(Segs)
                If i = 1 Then
                    IsoEngine.DrawSprite GetTex("head"), Pnt(Segs(i).X + Orientation2XOffset(Segs(i).Orientation), Segs(i).Y + Orientation2YOffset(Segs(i).Orientation)), Orientation2Scale(Segs(i).Orientation), 0, White, True
                ElseIf i = UBound(Segs) Then
                    IsoEngine.DrawSprite GetTex("tail"), Pnt(Segs(i).X + Orientation2XOffset(Segs(i).Orientation), Segs(i).Y + Orientation2YOffset(Segs(i).Orientation)), Orientation2Scale(Segs(i).Orientation), 0, White, True
                Else
                    IsoEngine.DrawSprite GetTex("body"), Pnt(Segs(i).X + Orientation2XOffset(Segs(i).Orientation), Segs(i).Y + Orientation2YOffset(Segs(i).Orientation)), Orientation2Scale(Segs(i).Orientation), 0, White, True
                End If
            Next
        IsoEngine.Sprite_End
        IsoEngine.DrawTexture GetTex("f" & CurrentFruit.FruitType), RECT(CurrentFruit.X, CurrentFruit.Y, 0, 0), White, , , , True
        IsoEngine.DrawBox RECT(678, 0, 100, 75), red + TransMedFlag
        IsoEngine.DrawText 680, 2, "Score: " & Score & vbCrLf & "Lives Left: " & Lives & vbCrLf & "Fruits Left: " & FruitsLeft & vbCrLf & "Level: " & Level, GetFont("Status"), White
        If Pause Then IsoEngine.DrawBox RECT(0, 0, frmMain.picGame.Width, frmMain.picGame.Height), DarkGrey + TransHvyFlag
        On Error Resume Next
    IsoEngine.RenderToScreen
End Sub

Function Orientation2XOffset(Orientation As Orientation) As Integer
    Select Case Orientation
        Case TR: Orientation2XOffset = 48
        Case BR: Orientation2XOffset = 64
        Case TL: Orientation2XOffset = -8
        Case Else: Orientation2XOffset = 0
    End Select
End Function

Function Orientation2YOffset(Orientation As Orientation) As Integer
    Select Case Orientation
        Case BL, BR: Orientation2YOffset = 24
        Case TL: Orientation2YOffset = -12
        Case Else: Orientation2YOffset = 0
    End Select
End Function

Sub MoveSnake()
    'If the snake has no direction, do not move it
    If UBound(Direction) = 0 Then Exit Sub
    
    'Changes the snake's direction if needed
    If UBound(Direction) > 1 Then
        For i = 1 To UBound(Direction) - 1
            Direction(i) = Direction(i + 1)
        Next
        ReDim Preserve Direction(UBound(Direction) - 1)
    End If
       
    'Grows the snake if needed
    If Grow > 0 Then
        Grow = Grow - 1
        ReDim Preserve Segs(UBound(Segs) + 1)
    End If
    
    'Moves the head of the snake
    Dim TmpHead As Segment
    TmpHead = Segs(1)
    Segs(1).Orientation = Direction2Orientation(Direction(1))
    Segs(1).X = Segs(1).X + Direction(1).X * 12
    Segs(1).Y = Segs(1).Y + Direction(1).Y * 12
    
    'Moves the body and tail of the snake
    For i = UBound(Segs) To 3 Step -1
        Segs(i) = Segs(i - 1)
    Next
    Select Case True
        Case Segs(UBound(Segs) - 1).X < Segs(UBound(Segs)).X And Segs(UBound(Segs) - 1).Y < Segs(UBound(Segs)).Y
            Segs(UBound(Segs)).Orientation = BR
        Case Segs(UBound(Segs) - 1).X < Segs(UBound(Segs)).X And Segs(UBound(Segs) - 1).Y > Segs(UBound(Segs)).Y
            Segs(UBound(Segs)).Orientation = TR
        Case Segs(UBound(Segs) - 1).X > Segs(UBound(Segs)).X And Segs(UBound(Segs) - 1).Y > Segs(UBound(Segs)).Y
            Segs(UBound(Segs)).Orientation = TL
        Case Segs(UBound(Segs) - 1).X > Segs(UBound(Segs)).X And Segs(UBound(Segs) - 1).Y < Segs(UBound(Segs)).Y
            Segs(UBound(Segs)).Orientation = BL
    End Select
    Segs(2).X = TmpHead.X
    Segs(2).Y = TmpHead.Y
    
    'Checks to see if the snake can eat a fruit
    If Collide_Box2Box(Segs(1).X, Segs(1).Y, Segs(1).X + 48, Segs(1).Y + 24, CurrentFruit.X, CurrentFruit.Y, 64, 64) Then
        Score = Score + 100 * (FruitsPerLevel + 1 - FruitsLeft)
        Grow = (11 - FruitsLeft) * 2
        FruitsLeft = FruitsLeft - 1
        Sound.Play App.Path & "\Sounds\GetFruit.wav", SoundVolume, , 11000
        If FruitsLeft = 0 Then
            FruitsLeft = FruitsPerLevel
            Level = Level + 1
            If Level = 11 Then
                Music.Stop_
                MsgBox "Congratulations!  You have beaten all 11 levels!  For a more challenging game, try increasing the speed of the snake.", vbExclamation, "You Win!"
                frmMain.btnStop_Click
                Exit Sub
            End If
            LoadLevel
            Exit Sub
        End If
        PlaceFruit
    End If
    
    'Checks to see if the snake crashes into itself
    If CrashCounter > -1 Then
        CrashCounter = CrashCounter + 1
        If CrashCounter = 15 Then CrashCounter = -1
    End If
    For i = 2 To UBound(Segs)
        If Segs(1).X = Segs(i).X And Segs(1).Y = Segs(i).Y And CrashCounter = -1 Then
            CrashCounter = 0
            Lives = Lives - 1
            Sound.Play App.Path & "\Sounds\HitSnake.wav", SoundVolume, , 11000
            If Lives = 0 Then
                Music.Stop_
                MsgBox "You crashed into yourself!" & vbCrLf & vbCrLf & "YOU LOSE!", vbCritical, "G A M E   O V E R"
                frmMain.btnStop_Click
                Exit Sub
            Else
                Music.Stop_
                MsgBox "You crashed into yourself!  Press <space> to keep playing!", vbCritical, "You Lose A Life"
                Music.Play True
                ReDim Direction(0)
                frmMain.picGame.SetFocus
                Exit Sub
            End If
        End If
    Next
    
    'Checks to see if the snake falls into the lava
    Dim TmpSnakeX(3) As Single, TmpSnakeY(3) As Single
    'There is an extra +32 when inputting the coordinates
    'of the snake because the tiles are 32 high
    TmpSnakeX(0) = Segs(1).X + 24 - Cam.X: TmpSnakeY(0) = Segs(1).Y + 6 + 32 - Cam.Y 'Top corner of head
    TmpSnakeX(1) = Segs(1).X + 24 - Cam.X: TmpSnakeY(1) = Segs(1).Y + 18 + 32 - Cam.Y 'Bottom corner
    TmpSnakeX(2) = Segs(1).X + 12 - Cam.X: TmpSnakeY(2) = Segs(1).Y + 12 + 32 - Cam.Y 'Left corner
    TmpSnakeX(3) = Segs(1).X + 36 - Cam.X: TmpSnakeY(3) = Segs(1).Y + 12 + 32 - Cam.Y 'Right corner
    For i = 0 To 3
        PixelToTile TmpSnakeX(i), TmpSnakeY(i)
    Next
    If Map(TmpSnakeX(0), TmpSnakeY(0)) = 1 Or _
       Map(TmpSnakeX(1), TmpSnakeY(1)) = 1 Or _
       Map(TmpSnakeX(2), TmpSnakeY(2)) = 1 Or _
       Map(TmpSnakeX(3), TmpSnakeY(3)) = 1 Then
        Lives = Lives - 1
        Sound.Play App.Path & "\Sounds\Fall.wav", SoundVolume, , 11000
        If Lives = 0 Then
            Music.Stop_
            MsgBox "You fell into the lava!" & vbCrLf & vbCrLf & "YOU LOSE!", vbCritical, "G A M E   O V E R"
            frmMain.btnStop_Click
            Exit Sub
        Else
            Music.Stop_
            MsgBox "You fell into the lava!  Press <space> to keep playing!", vbCritical, "You Lose A Life"
            Dim TmpSnakeLength As Integer, TmpGrow As Integer
            TmpSnakeLength = UBound(Segs) - 2
            TmpGrow = Grow
            LoadLevel
            frmMain.picGame.SetFocus
            Grow = TmpSnakeLength + TmpGrow
            Exit Sub
        End If
    End If
End Sub

Function Direction2Orientation(Direction As D3DVECTOR2) As Orientation
    If Direction.X = -2 And Direction.Y = -1 Then
        Direction2Orientation = TL
    ElseIf Direction.X = 2 And Direction.Y = -1 Then
        Direction2Orientation = TR
    ElseIf Direction.X = 2 And Direction.Y = 1 Then
        Direction2Orientation = BR
    ElseIf Direction.X = -2 And Direction.Y = 1 Then
        Direction2Orientation = BL
    End If
End Function

Sub UnloadGame()
    IsoEngine.Unload
    Sound.Unload
    Music.Unload
    Set Sound = Nothing
    Set Helper = Nothing
    Set Cam = Nothing
    Set IsoEngine = Nothing
End Sub

Function Orientation2Scale(ByVal Orientation As Orientation) As D3DVECTOR2
    Select Case Orientation
        Case TL
            Orientation2Scale.X = 1
            Orientation2Scale.Y = 1
        Case TR
            Orientation2Scale.X = -1
            Orientation2Scale.Y = 1
        Case BR
            Orientation2Scale.X = -1
            Orientation2Scale.Y = -1
        Case BL
            Orientation2Scale.X = 1
            Orientation2Scale.Y = -1
    End Select
End Function
