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
