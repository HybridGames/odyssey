Attribute VB_Name = "modProcessEvents"
'This Mod holds the processing code that is used by modProcess - ProcessString
'This way the code that handles individual work is isolated from other processing code
'The order of the Subs was based on the order the code was listed in the ProcessString sub

'Error Logging On
Sub ErrorLoggingOn(Data As String)
    Dim BanDate As Long

    If Len(Data) >= 1 Then
        Select Case Asc(Mid$(Data, 1, 1))
        Case 0    'Custom Message
            If Len(Data) >= 2 Then
                MsgBox Mid$(Data, 2), vbOKOnly + vbExclamation, TitleString
            End If
        Case 1    'Invalid User/Pass
            MsgBox "Invalid user name/password!", vbOKOnly + vbExclamation, TitleString
        Case 2    'Account already in use
            MsgBox "Someone is already using that account!", vbOKOnly + vbExclamation, TitleString
        Case 3    'Banned
            If Len(Data) >= 5 Then
                BanDate = Asc(Mid$(Data, 2, 1)) * 16777216 + Asc(Mid$(Data, 3, 1)) * 65536 + Asc(Mid$(Data, 4, 1)) * 256& + Asc(Mid$(Data, 5, 1))
                If Len(Data) > 5 Then
                    MsgBox "You are banned from The Odyssey Online Classic until " + CStr(CDate(BanDate)) + " (" + Mid$(Data, 6) + ")!", vbOKOnly, TitleString
                Else
                    MsgBox "You are banned from The Odyssey Online Classic until " + CStr(CDate(BanDate)) + "!", vbOKOnly, TitleString
                End If
                CloseClientSocket 3
            End If
        Case 4    'Server Full
            MsgBox "The server is full, please try again in a few minutes!", vbOKOnly + vbExclamation, TitleString
        Case 5    'Multiple Login
            MsgBox "You may not log in multiple times from the same computer!", vbOKOnly + vbExclamation, TitleString
        End Select
    End If
    CloseClientSocket 4
End Sub

'Error Creating a New Account
Sub ErrorCreatingNewAccount(Data As String)
    If Len(Data) >= 1 Then
        Select Case Asc(Mid$(Data, 1, 1))
        Case 0    'Custom Message
            If Len(Data) >= 2 Then
                MsgBox Mid$(Data, 2), vbOKOnly + vbExclamation, TitleString
            End If
        Case 1    'User name already in use
            MsgBox "That user name is already in use.  Please try another.", vbOKOnly + vbExclamation, TitleString
        End Select
    End If
    CloseClientSocket 2
End Sub

'Account Created
Sub AccountCreated(Data As String)
    CloseClientSocket 0
    MsgBox "Your account has been created successfully!  Please write down your user name and password somewhere safe so that you do not loose them.  Click Login to connect to the game server.", vbOKOnly + vbExclamation, TitleString
End Sub

'Logged On, Includes Character Data
Sub LoggedOn(Data As String)
    Dim A As Long

    If frmWait_Loaded = True Then Unload frmWait
    If frmLogin_Loaded = True Then Unload frmLogin
    If frmNewCharacter_Loaded = True Then Unload frmNewCharacter
    CWalkStep = 0
    
    If Len(Data) >= 10 Then
        With Character
            .name = vbNullString
            .Class = Asc(Mid$(Data, 1, 1))
            .Gender = Asc(Mid$(Data, 2, 1))
            .Sprite = Asc(Mid$(Data, 3, 1)) * 256 + Asc(Mid$(Data, 4, 1))
            .Level = Asc(Mid$(Data, 5, 1))
            .status = Asc(Mid$(Data, 6, 1))
            .Guild = Asc(Mid$(Data, 7, 1))
            .GuildRank = Asc(Mid$(Data, 8, 1))
            .Access = Asc(Mid$(Data, 9, 1))
            .index = Asc(Mid$(Data, 10, 1))
            .Experience = Asc(Mid$(Data, 11, 1)) * 16777216 + Asc(Mid$(Data, 12, 1)) * 65536 + Asc(Mid$(Data, 13, 1)) * 256& + Asc(Mid$(Data, 14, 1))
    
            Data = Mid$(Data, 15)
            A = InStr(Data, vbNullChar)
            If A > 1 Then
                .name = Mid$(Data, 1, A - 1)
                If A < Len(Data) Then
                    .Desc = Mid$(Data, A + 1)
                End If
            End If
        End With
        
        SetMap 0
        
        For A = 1 To MaxUsers
            Guild(A).name = vbNullString
            With Player(A)
                .Sprite = 0
                .Map = 0
            End With
        Next A
        
        With Character
            For A = 1 To MaxInvObjects
                With .Inv(A)
                    .Object = 0
                    .EquippedNum = 0
                    .value = 0
                    .ItemPrefix = 0
                    .ItemSuffix = 0
                End With
            Next A
        End With
        
        frmWait.Show
        frmWait.lblStatus = "Receiving Game Data ..."
        frmWait.btnCancel.Visible = True
        frmWait.Refresh
        SendSocket Chr$(7) + Chr$(1)    'I wanna play
    Else
        Character.Class = 0
        frmNewCharacter.Show
        frmWait.Hide
    End If
End Sub

'Password Changed
Sub PasswordChanged(Data As String)
    If frmWait_Loaded = True Then Unload frmWait
End Sub

'Sets the Message of the Day text
Sub SetMotd(Data As String)
    MOTDText = Data
End Sub

'When another player joins the game
Sub PlayerJoinedGame(Data As String)
    Dim PlayerIndex As Long

    If Len(Data) >= 7 Then
        PlayerIndex = Asc(Mid$(Data, 1, 1))
        With Player(PlayerIndex)
            .Ignore = False
            .IsDead = False
            .Sprite = Asc(Mid$(Data, 2, 1)) * 256 + Asc(Mid$(Data, 3, 1))
            .status = Asc(Mid$(Data, 4, 1))
            .Guild = Asc(Mid$(Data, 5, 1))
            .MaxHP = Asc(Mid$(Data, 6, 1))
            .name = Mid$(Data, 7)
            If CMap > 0 Then
                If Not .status = 25 Then
                    If .status = 2 Then
                        PrintChat "All hail " + .name + ", a new adventurer in this land!", 3
                    Else
                        PrintChat .name + " has joined the game!", 3
                    End If
                End If
            End If
            UpdatePlayerColor PlayerIndex
        End With
    End If
End Sub

'Player Left Game
Sub PlayerLeftGame(Data As String)
    Dim PlayerIndex As Long

    If Len(St) = 1 Then
        PlayerIndex = Asc(Mid$(St, 1, 1))
        If PlayerIndex >= 1 Then
            With Player(PlayerIndex)
                If Not .status = 25 Then
                    PrintChat .name + " has left the game!", 3
                End If
                PlayerLeftMap PlayerIndex
                .Sprite = 0
                .IsDead = False
            End With
        End If
    End If
End Sub

'Player Joined Map
Sub PlayerJoinedMap(Data As String)
    Dim PlayerIndex As Long

    If Len(St) = 7 Then
        PlayerIndex = Asc(Mid$(St, 1, 1))
        With Player(PlayerIndex)
            .Map = CMap
            .X = Asc(Mid$(St, 2, 1))
            .Y = Asc(Mid$(St, 3, 1))
            .D = Asc(Mid$(St, 4, 1))
            .Sprite = Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1))
            .status = Asc(Mid$(St, 7, 1))
            .XO = .X * 32
            .YO = .Y * 32
            AddPlayerLight PlayerIndex
            .PlayerIndex = 0
        End With
    End If
End Sub

'Player Left Map
Sub PlayerLeftMap(Data As String)
    Dim PlayerIndex As Long

    If Len(St) = 1 Then
        PlayerIndex = Asc(Mid$(St, 1, 1))
        If PlayerIndex >= 1 Then
            PlayerLeftMap PlayerIndex
            RemovePlayerLight PlayerIndex
        End If
    End If
End Sub

'Player Moved
Sub PlayerMoved(Data As String)
    Dim PlayerIndex As Long
    
    If Len(St) = 5 Then
        PlayerIndex = Asc(Mid$(St, 1, 1))
        If PlayerIndex > 0 And PlayerIndex < MaxUsers Then
            With Player(PlayerIndex)
                If .X * 32 = .XO And .Y * 32 = .YO Then
                    .X = Asc(Mid$(St, 2, 1))
                    .Y = Asc(Mid$(St, 3, 1))
                Else
                    .XO = .X * 32
                    .YO = .Y * 32
                    .X = Asc(Mid$(St, 2, 1))
                    .Y = Asc(Mid$(St, 3, 1))
                End If
                .D = Asc(Mid$(St, 4, 1))
                .WalkStep = Asc(Mid$(St, 5, 1))
                .IsDead = False
            End With
        End If
    End If
End Sub

'Player Say
Sub Say(Data As String)
    Dim PlayerIndex As Long

    If Len(St) >= 2 Then
        PlayerIndex = Asc(Mid$(St, 1, 1))
        If Player(PlayerIndex).Ignore = False Then
            PrintChat Player(PlayerIndex).name + " says, " + Chr$(34) + Mid$(St, 2) + Chr$(34), 7
        End If
    End If
End Sub

'When this player joins a map
Sub JoinMap(Data As String)
    Dim A As Long
    Dim PlayerCount As Long

    If Len(St) = 13 Then
        If MapEdit = True Then CloseMapEdit
        'Destroy Projectiles
        For A = 1 To MaxProjectiles
            DestroyEffect A
        Next A
        
        For A = 1 To MaxFloatText
            ClearFloatText A
        Next A
        
        For A = 1 To MaxUsers
            Player(A).HP = 0
        Next A
        
        'If Map = 0 then we're just logging in
        If CMap = 0 Then
            Tick = timeGetTime
            Character.LastMove = Tick
            St1 = vbNullString
            PlayerCount = 0
            For A = 1 To MaxUsers
                With Player(A)
                    If .Sprite > 0 And A <> Character.index And Not .status = 25 Then
                        PlayerCount = PlayerCount + 1
                        St1 = St1 + ", " + .name
                    End If
                End With
            Next A
            Character.IsDead = False
            SetHP GetMaxHP
            SetEnergy GetMaxEnergy
            SetMana GetMaxMana
            
            LoadOptions
            Load frmMain
            
            If PlayerCount > 0 Then
                St1 = Mid$(St1, 2)
                PrintChat "Welcome to the Odyssey Online Classic!", 15
                PrintChat "There are " + CStr(PlayerCount) + " other players online:", 15
                PrintChat St1, 15
            Else
                PrintChat "Welcome to the Odyssey Online Classic!", 15
                PrintChat "There are no other users currently online.", 15
            End If
            
            PrintChat MOTDText, 11
        End If
        
        SetMap Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
        CX = Asc(Mid$(St, 3, 1))
        CY = Asc(Mid$(St, 4, 1))
        CDir = Asc(Mid$(St, 5, 1))
        CXO = CX * 32
        CYO = CY * 32
        
        For A = 0 To MaxMonsters
            Map.Monster(A).Monster = 0
        Next A
        
        For A = 0 To MaxMapObjects
            Map.Object(A).Object = 0
            Map.Object(A).ItemPrefix = 0
            Map.Object(A).ItemSuffix = 0
        Next A
        
        For A = 0 To 9
            Map.Door(A).Att = 0
        Next A
        
        For A = 1 To MaxUsers
            Player(A).Map = 0
        Next A
        
        ClearLighting
        Freeze = True
        LoadMapFromCache CMap
        If RequestedMap = False Then
            If Map.Version <> Asc(Mid$(St, 6, 1)) * 16777216 + Asc(Mid$(St, 7, 1)) * 65536 + Asc(Mid$(St, 8, 1)) * 256& + Asc(Mid$(St, 9, 1)) Or CheckSum(MapData) <> Asc(Mid$(St, 10, 1)) * 16777216 + Asc(Mid$(St, 11, 1)) * 65536 + Asc(Mid$(St, 12, 1)) * 256& + Asc(Mid$(St, 13, 1)) Then
                RequestedMap = True
                SendSocket Chr$(45)
            Else
                ShowMap
            End If
        End If
    End If
End Sub

'Error on creating a new character
Sub ErrorCreatingCharacter(Data As String)
    If frmWait_Loaded = True Then
        Unload frmWait
    End If
    
    frmNewCharacter.Show
    MsgBox "That name is already in use, please try another!", vbOKOnly + vbExclamation, TitleString
End Sub

'New Map Object Sent
Sub NewMapObject(Data As String)
    Dim MapObject As Long
    Dim Rnd As Long
    
    If Len(St) = 11 Then
        MapObject = Asc(Mid$(St, 1, 1))
        If MapObject <= MaxMapObjects Then
            With Map.Object(MapObject)
                .Object = GetInt(Mid$(St, 2, 2))
                .X = Asc(Mid$(St, 4, 1))
                .Y = Asc(Mid$(St, 5, 1))
                .ItemPrefix = Asc(Mid$(St, 6, 1))
                .ItemSuffix = Asc(Mid$(St, 7, 1))
                Rnd = Int(Rnd * 9)
                .XOffset = Rnd
                Rnd = Int(Rnd * 9)
                .YOffset = Rnd
                .value = Asc(Mid$(St, 8, 1)) * 16777216 + Asc(Mid$(St, 9, 1)) * 65536 + Asc(Mid$(St, 10, 1)) * 256& + Asc(Mid$(St, 11, 1))
                .PickedUp = 0
                RedrawMapTile CLng(.X), CLng(.Y)
            End With
        End If
    End If
End Sub

'Erases a Map Object
Sub EraseMapObject(Data As String)
    Dim MapObject As Long

    If Len(St) = 1 Then
        MapObject = Asc(Mid$(St, 1, 1))
        If MapObject <= MaxMapObjects Then
            With Map.Object(MapObject)
                .Object = 0
                .ItemPrefix = 0
                .ItemSuffix = 0
                .value = 0
                .PickedUp = 0
                RedrawMapTile CLng(.X), CLng(.Y)
            End With
        End If
    End If
End Sub

'Displays Messages
Sub Message(Data As String)
    If Len(Data) >= 1 Then
        Select Case Asc(Mid$(Data, 1, 1))
        Case 0    'Custom
            If Len(Data) > 2 Then
                PrintChat Mid$(Data, 2), 7
            End If
        Case 1    'Inv full
            PrintChat "Your inventory is full!", 7
        Case 2    'Map full
            PrintChat "There is too much already on the ground here to drop that.", 7
        Case 3    'No such object
            PrintChat "No such object.", 7
        Case 4    'No such player
            PrintChat "No such player.", 7
        Case 5    'No such monster
            PrintChat "No such monster.", 7
        Case 6    'Player is too far away
            PrintChat "Player is too far away.", 7
        Case 7    'Monster is too far away
            PrintChat "Monster is too far away.", 7
        Case 8    'You cannot use that
            PrintChat "You cannot use that object.", 7
        Case 9    'Friendly Zone - can't attack
            PrintChat "This is a friendly area, you cannot attack here!", 7
        Case 10    'Cannot attack immortal
            PrintChat "You may not attack an immortal!", 7
        Case 11    'You are an immortal
            PrintChat "Immortals may not attack other players!", 7
        Case 12    'Can't attack monsters here
            PrintChat "You cannot attack these monsters!", 7
        Case 13    'Ban list full
            PrintChat "The ban list is full!", 7
        Case 14    'Not invited to join
            PrintChat "You have not been invited to join any guild.", 7
        Case 15    'Not enough cash
            PrintChat "You do not have enough gold to do that!", 7
        Case 16    'Guild name in use
            PrintChat "That name is already used either by another player or guild.  Please try another.", 7
        Case 17    'Guild full
            PrintChat "That guild is full!", 7
        Case 18    'too many guilds
            PrintChat "Too many guilds already exist.  You may join another guild or try again later.", 7
        Case 19    'cannot attack player -- he is not in guild
            PrintChat "That player is not in a guild -- you may not attack non-guild players.", 7
        Case 20    'cannot attack player -- you are not in guild
            PrintChat "You must be a member of a guild to attack other players.", 7
        Case 21    'not in a hall
            PrintChat "You are not in a guild hall!", 7
        Case 22    'hall already owned
            PrintChat "This hall is already owned by another guild.", 7
        Case 23    'already have hall
            PrintChat "Your guild already owns a hall.  You must move out of your old hall before you may purchase a new one.", 7
        Case 24    'don't have enough money to buy hall
            PrintChat "Your guild does not have enough money in its bank account to buy this hall.  Type /guild hallinfo for the price information of this hall.", 7
        Case 25    'do not own a guild hall
            PrintChat "Your guild does not own a hall.", 7
        Case 26    'need 5 members
            PrintChat "You must have atleast 5 members in your guild before you may do that.", 7
        Case 27    'Can't afford that
            PrintChat "You do not have the items required to purchase that!", 7
        Case 28    'Not in a bank
            PrintChat "You are not in a bank!", 7
        Case 29    'too far away
            PrintChat "That player is too far away to hit!", 7
        Case 30    'must be Level # to join guild
            PrintChat "You must be at least Level " + CStr(World.GuildJoinLevel) + " to join a guild!", 7
        Case 32    'Must be in a smithy shop
            PrintChat "You are not in a blacksmithy shop!", 7
        Case 33    'Do not have enough money
            PrintChat "You do not have enough money to repair this object!", 7
        Case 34    'Do not have specified object
            PrintChat "You do not have the object to be repaired!", 7
        Case 35    'Not high enough stats
            PrintChat "You do not meet the requirements to use this item!", 7
        Case 36    'Ammo
            PrintChat "This ammo must be used with a projectile weapon!", 7
        Case 37    'Ammo
            PrintChat "This ammo is not used with this weapon!", 7
        Case 40    'Can't use
            PrintChat "Your class cannot use this object!", 7
        Case 41    'Squelched or Jailed
            PrintChat "Nobody hears your pitiful cries.", 7
        Case 42    'New Day
            PrintChat "A new day dawns in the land of Odyssey!", 7
        Case 43    'No Guild MOTD
            PrintChat "Your guild does not currently have a message of the day.  If you are a lord or a founder, you can set this by typing /guild motd <message>.", 7
        Case 44    'Bank is Full
            PrintChat "Your bank is full!", 7
        Case 45    'E-mail is already saved
            PrintChat "Your account already has an e-mail address saved!", 7
        Case 46    'E-mail updated
            PrintChat "Your e-mail address has been stored!", 7
        Case 47    'E-mail needed
            PrintChat "You have not entered an email address for your account.  Type /email to enter it.  This is the only way to recover your account if it is lost or stolen!", 7
        Case 48    'Server is being backed up
            PrintChat "Server is being backed up.", 7
        End Select
    End If
End Sub

'Updates character's inventory
Sub NewInventoryObject(Data As String)
    Dim InvIndex As Long

    If Len(Data) = 9 Then
        InvIndex = Asc(Mid$(Data, 1, 1))
        If InvIndex >= 1 And InvIndex <= 20 Then
            With Character.Inv(InvIndex)
                .Object = GetInt(Mid$(Data, 2, 2))
                .value = Asc(Mid$(Data, 4, 1)) * 16777216 + Asc(Mid$(Data, 5, 1)) * 65536 + Asc(Mid$(Data, 6, 1)) * 256& + Asc(Mid$(Data, 7, 1))
                .ItemPrefix = Asc(Mid$(Data, 8, 1))
                .ItemSuffix = Asc(Mid$(Data, 9, 1))
            End With
            If frmMain.picRepair.Visible = True Then
                DisplayRepair
            ElseIf frmMain.picSellObject.Visible = True Then
                DisplaySell
            End If
            RefreshInventory
        End If
    End If
End Sub

'Removes an item from inventory
Sub EraseInventoryObject(Data As String)
    Dim InvIndex As Long

    If Len(Data) = 1 Then
        InvIndex = Asc(Mid$(Data, 1, 1))
        If InvIndex >= 1 And InvIndex <= 20 Then
            With Character.Inv(InvIndex)
                If .EquippedNum > 0 And .Object > 0 Then
                    If Object(.Object).Type = 10 Then
                        Character.Projectile = False
                    ElseIf Object(.Object).Type = 11 Then
                        Character.Ammo = 0
                    End If
                End If
                .Object = 0
                .ItemPrefix = 0
                .ItemSuffix = 0
                .value = 0
                .EquippedNum = 0
            End With
            If frmMain.picRepair.Visible = True Then
                DisplayRepair
            ElseIf frmMain.picSellObject.Visible = True Then
                DisplaySell
            End If
            RefreshInventory
        End If
    End If
End Sub

'Use Object
Sub UseObject(Data As String)
    Dim InvIndex As Long

    If Len(Data) >= 1 Then
        InvIndex = Asc(Mid$(Data, 1, 1))
        If InvIndex >= 1 And InvIndex <= 20 Then
            If Character.Inv(InvIndex).Object > 0 Then
                Select Case Object(Character.Inv(InvIndex).Object).Type
                Case 2, 3, 4    'Armor pieces
                    Character.EquippedObject(Object(Character.Inv(InvIndex).Object).Type).Object = Character.Inv(InvIndex).Object
                    Character.EquippedObject(Object(Character.Inv(InvIndex).Object).Type).value = Character.Inv(InvIndex).value
                    Character.EquippedObject(Object(Character.Inv(InvIndex).Object).Type).ItemPrefix = Character.Inv(InvIndex).ItemPrefix
                    Character.EquippedObject(Object(Character.Inv(InvIndex).Object).Type).ItemSuffix = Character.Inv(InvIndex).ItemSuffix
                    Character.Inv(InvIndex).Object = 0
                    Character.Inv(InvIndex).value = 0
                    Character.Inv(InvIndex).ItemPrefix = 0
                    Character.Inv(InvIndex).ItemSuffix = 0
                Case 8    'Ring
                    Character.EquippedObject(5).Object = Character.Inv(InvIndex).Object
                    Character.EquippedObject(5).value = Character.Inv(InvIndex).value
                    Character.EquippedObject(5).ItemPrefix = Character.Inv(InvIndex).ItemPrefix
                    Character.EquippedObject(5).ItemSuffix = Character.Inv(InvIndex).ItemSuffix
                    Character.Inv(InvIndex).Object = 0
                    Character.Inv(InvIndex).value = 0
                    Character.Inv(InvIndex).ItemPrefix = 0
                    Character.Inv(InvIndex).ItemSuffix = 0
                Case 1    'Weapons
                    Character.EquippedObject(1).Object = Character.Inv(InvIndex).Object
                    Character.EquippedObject(1).value = Character.Inv(InvIndex).value
                    Character.EquippedObject(1).ItemPrefix = Character.Inv(InvIndex).ItemPrefix
                    Character.EquippedObject(1).ItemSuffix = Character.Inv(InvIndex).ItemSuffix
                    Character.Inv(InvIndex).Object = 0
                    Character.Inv(InvIndex).value = 0
                    Character.Inv(InvIndex).ItemPrefix = 0
                    Character.Inv(InvIndex).ItemSuffix = 0
                    Character.Projectile = False
                Case 10    'Projectile Weapon
                    Character.EquippedObject(1).Object = Character.Inv(InvIndex).Object
                    Character.EquippedObject(1).value = Character.Inv(InvIndex).value
                    Character.EquippedObject(1).ItemPrefix = Character.Inv(InvIndex).ItemPrefix
                    Character.EquippedObject(1).ItemSuffix = Character.Inv(InvIndex).ItemSuffix
                    Character.Inv(InvIndex).Object = 0
                    Character.Inv(InvIndex).value = 0
                    Character.Inv(InvIndex).ItemPrefix = 0
                    Character.Inv(InvIndex).ItemSuffix = 0
                    Character.Projectile = True
                Case 11    'Ammo (Stays in inventory)
                    Character.Ammo = InvIndex
                    Character.Inv(InvIndex).EquippedNum = InvIndex
                End Select
                RefreshInventory
            End If
        End If
    End If
End Sub

'Stop Using, unequip
Sub StopUsingObject(Data As String)
    Dim InvIndex As Long

    If Len(Data) = 1 Then
        InvIndex = Asc(Mid$(Data, 1, 1))
        If InvIndex >= 1 And InvIndex <= 20 Then
            If Character.Inv(InvIndex).Object > 0 Then
                Character.Inv(InvIndex).EquippedNum = False
                If Object(Character.Inv(InvIndex).Object).Type = 10 Then Character.Projectile = False
                If Object(Character.Inv(InvIndex).Object).Type = 11 Then Character.Ammo = 0
                RefreshInventory
            End If
        Else
            If InvIndex >= 21 Then    'Ok
                InvIndex = InvIndex - 20
                If Character.EquippedObject(InvIndex).Object > 0 Then
                    If Object(Character.EquippedObject(InvIndex).Object).Type = 10 Then Character.Projectile = False
                    Character.EquippedObject(InvIndex).Object = 0
                    Character.EquippedObject(InvIndex).value = 0
                    Character.EquippedObject(InvIndex).ItemPrefix = 0
                    Character.EquippedObject(InvIndex).ItemSuffix = 0
                    RefreshInventory
                End If
            End If
        End If
    End If
End Sub

'This Player has joined the game
Sub JoinedGame(Data As String)
    Dim A As Long

    If frmWait_Loaded = True Then
        frmWait.lblStatus = "Loading Game ..."
        frmWait.lblStatus.Refresh
    End If
    For A = 1 To MaxUsers
        Player(A).Map = 0
    Next A
    With Character
        For A = 1 To MaxInvObjects
            With .Inv(A)
                .Object = 0
                .EquippedNum = 0
                .value = 0
                .ItemPrefix = 0
                .ItemSuffix = 0
            End With
        Next A
        For A = 1 To 5
            .EquippedObject(A).Object = 0
            .EquippedObject(A).value = 0
            .EquippedObject(A).ItemPrefix = 0
            .EquippedObject(A).ItemSuffix = 0
        Next A
    End With
    Load frmMain
    frmMain.WindowState = 0
    SetHP GetMaxHP
    SetMana GetMaxMana
    SetEnergy GetMaxEnergy
    DrawStats
    RedrawMap = True
    If Character.Level = 1 Then frmMain.picHelp.Visible = True
    blnPlaying = True
    ResetTimers
End Sub

'Tell message /t
Sub Tell(Data As String)
    Dim PlayerIndex As Long

    If Len(Data) >= 2 Then
        PlayerIndex = Asc(Mid$(Data, 1, 1))
        If PlayerIndex >= 1 Then
            With Player(PlayerIndex)
                If .Ignore = False Then
                    PrintChat .name + " tells you, " + Chr$(34) + Mid$(Data, 2) + Chr$(34), 10
                End If
            End With
        End If
    End If
End Sub

'Broadcast message
Sub Broadcast(Data As String)
    Dim PlayerIndex As Long

    If Len(Data) >= 2 And options.Broadcasts = True Then
        PlayerIndex = Asc(Mid$(Data, 1, 1))
        If PlayerIndex >= 1 Then
            If Player(PlayerIndex).Ignore = False Then
                PrintChat Player(PlayerIndex).name + ": " + Mid$(Data, 2), 13
            End If
        End If
    End If
End Sub

'Emote
Sub Emote(Data As String)
    Dim PlayerIndex As Long

    If Len(Data) >= 2 Then
        PlayerIndex = Asc(Mid$(Data, 1, 1))
        If Player(PlayerIndex).Ignore = False Then
            PrintChat Player(PlayerIndex).name + " " + Mid$(Data, 2), 11
        End If
    End If
End Sub

'Yell Message
Sub Yell(Data As String)
    Dim PlayerIndex As Long
    
    If Len(Data) >= 2 Then
        PlayerIndex = Asc(Mid$(Data, 1, 1))
        If Player(PlayerIndex).Ignore = False Then
            PrintChat Player(PlayerIndex).name + " yells, " + Chr$(34) + Mid$(Data, 2) + Chr$(34), 7
        End If
    End If
End Sub

'Server wide message
Sub ServerMessage(Data As String)
    If Len(Data) > 0 Then
        PrintChat "Server Message: " + Data, 9
    End If
End Sub
