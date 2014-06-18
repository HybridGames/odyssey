Attribute VB_Name = "modProcessEvents"
'Client Creating a new Account
Sub NewAccount(Data As String, Player As PlayerData)
    Dim ByteCount As Long
    Dim NameLength As Long
    Dim UserName As String
    Dim Password As String

    With Player
        If .ClientVer = CurrentClientVer Then
            ByteCount = InStr(1, Data, Chr$(0))
            If ByteCount > 1 And ByteCount < Len(Data) Then
                UserName = Trim$(Mid$(Data, 1, ByteCount - 1))
                NameLength = Len(UserName)
                If NameLength >= 3 And NameLength <= 15 And ValidName(UserName) Then
                    UserRS.Index = "User"
                    UserRS.Seek "=", UserName

                    '@todo this doesn't seem necessary
                    'E = 0
                    'UserName = UCase$(UserName)
                    'For F = 1 To MaxUsers
                    '    If F <> Index Then
                    '        If UserName = UCase$(Player(F).User) Then
                    '            E = 1
                    '            Exit For
                    '        End If
                    '    End If
                    'Next F

                    If UserRS.NoMatch = True And GuildNum(UserName) = 0 Then 'And E = 0 Then
                        UserRS.AddNew
                        UserRS!User = UserName
                        .User = UserName
                        Password = Trim$(UCase$(Mid$(Data, ByteCount + 1)))
                        If Len(Password) > 15 Then
                            UserRS!Password = Left$(Password, 15)
                        Else
                            UserRS!Password = Password
                        End If
                        UserRS.Update
                        UserRS.Seek "=", .User
                        .Bookmark = UserRS.Bookmark
                        .Access = 0
                        .Class = 0
                        SavePlayerData Index
                        SendSocket Index, Chr$(2)    'New account created!
                        AddSocketQue Index
                    Else
                        SendSocket Index, Chr$(1) + Chr$(1)    'User Already Exists
                        AddSocketQue Index
                    End If
                Else
                    Hacker Index, "A.79"
                End If
            Else
                AddSocketQue Index
            End If
        Else
            SendSocket Index, Chr$(116)
            AddSocketQue Index
        End If
    End With
End Sub
