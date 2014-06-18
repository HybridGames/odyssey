Attribute VB_Name = "modProcessEvents"
'Client Creating a new Account
Sub NewAccount(Data As String, Player As PlayerData)
    Dim ByteCount As Long

    With Player
        If .ClientVer = CurrentClientVer Then
            ByteCount = InStr(1, Data, Chr$(0))
            If ByteCount > 1 And ByteCount < Len(Data) Then
                St1 = Trim$(Mid$(Data, 1, ByteCount - 1))
                B = Len(St1)
                If B >= 3 And B <= 15 And ValidName(St1) Then
                    UserRS.Index = "User"
                    UserRS.Seek "=", St1

                    'E = 0
                    'St1 = UCase$(St1)
                    'For F = 1 To MaxUsers
                    '    If F <> Index Then
                    '        If St1 = UCase$(Player(F).User) Then
                    '            E = 1
                    '            Exit For
                    '        End If
                    '    End If
                    'Next F

                    If UserRS.NoMatch = True And GuildNum(St1) = 0 Then 'And E = 0 Then
                        UserRS.AddNew
                        UserRS!User = St1
                        .User = St1
                        St1 = Trim$(UCase$(Mid$(Data, ByteCount + 1)))
                        If Len(St1) > 15 Then
                            UserRS!Password = Left$(St1, 15)
                        Else
                            UserRS!Password = St1
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
