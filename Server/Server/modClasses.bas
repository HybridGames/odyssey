Attribute VB_Name = "modClasses"
Option Explicit

Sub CreateClassData()
    With Class(1)    'Knight
        .StartHP = 25
        .StartEnergy = 20
        .StartMana = 10
        .MaxHP = 155
        .MaxEnergy = 70
        .MaxMana = 20
    End With
    With Class(2)    'Rogue
        .StartHP = 25
        .StartEnergy = 30
        .StartMana = 10
        .MaxHP = 115
        .MaxEnergy = 100
        .MaxMana = 20
    End With
End Sub
