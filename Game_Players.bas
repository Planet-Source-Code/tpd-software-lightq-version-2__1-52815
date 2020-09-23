Attribute VB_Name = "Game_Players"
' ######################################
' ##
' ##  LightQ player functions
' ##

Type Players
  Name      As String * 16
  Level     As Long
  Time      As Long
End Type

' level indexes for players
Public CurLevel() As Long
'
' adds a new player to the game
'
Sub DoAddNewPlayer(Name As String)

    Dim CurPlayers() As Players
    Dim NumPlayers   As Long
    
    Open App.Path & "\players.dat" For Binary As #1
    
    Get #1, , NumPlayers
    
    If NumPlayers > 0 Then
       ReDim CurPlayers(NumPlayers - 1)
       Get #1, , CurPlayers()
       
       For i = 0 To UBound(CurPlayers)
         If Trim(CurPlayers(i).Name) = Name Then
            ShowError "The player " & Name & " already exist."
            Close #1
            Exit Sub
         End If
       Next i
       
       ReDim Preserve CurPlayers(UBound(CurPlayers) + 1)
    Else
       ReDim CurPlayers(0)
    End If
    
    With CurPlayers(UBound(CurPlayers))
         .Level = 1
         .Name = Trim(UCase(Name))
         .Time = GetTickCount()
    End With
    
    NumPlayers = NumPlayers + 1
    
    Seek #1, 1
    Put #1, , NumPlayers
    Put #1, , CurPlayers()
    
    Close #1
    
    DoLoadPlayers

End Sub

'
' removes a player from the game
'
Sub DoRemovePlayer(Name As String)

    If DoQuestion("Are you sure you want to remove player " & Name, "REMOVE PLAYER", vbYesNo) = vbNo Then Exit Sub
    
    Dim CurPlayers() As Players
    Dim NumPlayers   As Long
    
    Open App.Path & "\players.dat" For Binary As #1
      Get #1, , NumPlayers
      ReDim CurPlayers(NumPlayers - 1)
      Get #1, , CurPlayers()
    Close #1
    Kill App.Path & "\players.dat"
    NumPlayers = NumPlayers - 1
    Open App.Path & "\players.dat" For Binary As #1
      Put #1, , NumPlayers
      For i = 0 To UBound(CurPlayers)
         If Trim(CurPlayers(i).Name) <> Name Then
            Put #1, , CurPlayers(i)
         End If
      Next i
    Close #1
    
    DoLoadPlayers

End Sub

'
' updates player status level index
'
Sub DoUpdatePlayerLevel(Name As String)

    Dim CurPlayers() As Players
    Dim NumPlayers   As Long
    Dim LevelSolved  As Long
    Dim cLevel       As Long
    
    EngineSelect = True
   
    With frmScreen
       LevelSolved = Val(.Levels.List(.Levels.ListIndex))
    End With
  
    'update the current level for the current player (after solving a level)
    Open App.Path & "\players.dat" For Binary As #1
      Get #1, , NumPlayers
      ReDim CurPlayers(NumPlayers - 1)
      Get #1, , CurPlayers()
    
      For i = 0 To UBound(CurPlayers)
         If Trim(CurPlayers(i).Name) = Name Then
            
            With frmScreen
            
               If CurPlayers(i).Level = LevelSolved Then
                  CurPlayers(i).Level = CurPlayers(i).Level + 1
                  CurLevel(i) = CurPlayers(i).Level
                  
                  If CurPlayers(i).Level > .Levels.ListCount Then
                     .Levels.Tag = "ALL_DONE"
                  End If
                  
                  cLevel = CurLevel(i)
               Else
                  If LevelSolved + 1 < .Levels.ListCount Then
                     cLevel = LevelSolved + 1
                  Else
                     .Levels.Tag = "ALL_DONE"
                  End If
               End If
               
               If cLevel <= .Levels.ListCount Then
                  .Levels.ListIndex = cLevel - 1
               End If
            
            End With
            
            Exit For
         End If
      
      Next i

      Seek #1, 5
      Put #1, , CurPlayers()

    Close #1
    
    'update player status on screen
    frmScreen.UpdatePlayerStats
    EngineSelect = False
    
End Sub

'
' loads all players
'
Sub DoLoadPlayers()
    
    Dim CurPlayers() As Players
    Dim NumPlayers   As Long
    
    'load all players and populate player list
    Open App.Path & "\players.dat" For Binary As #1
       Get #1, , NumPlayers
       If NumPlayers > 0 Then
          ReDim CurPlayers(NumPlayers - 1)
          ReDim CurLevel(NumPlayers - 1)
          Get #1, , CurPlayers()
       Else
          ReDim CurPlayers(0)
          ReDim CurLevel(0)
       End If
    Close #1
    
    With frmScreen.lstPlayers
       .Clear
       For i = 0 To NumPlayers - 1
          CurLevel(i) = CurPlayers(i).Level
          .AddItem Trim(CurPlayers(i).Name)
       Next i
    End With
     
End Sub

