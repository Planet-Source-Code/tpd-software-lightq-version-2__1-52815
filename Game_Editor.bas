Attribute VB_Name = "Game_Editor"
' ######################################
' ##
' ##  LightQ Game Editor functions
' ##

Private Type EditX
   Object As String
   is_rot As Boolean
End Type

Private Type BinData
   version   As Long          ' version
   user      As String * 16   ' user name
   Name      As String * 128  ' level name
   passw(15) As Byte          ' MD5 Hash
End Type

Private Type UserPriv
   user      As String * 16
   passw(15) As Byte
End Type

Public CurrentLevel As Long

Private BinHdr As BinData
Private UsrPrv As UserPriv

' This is a state variable that make skip the "_click" event when the engine alters a value
Public EngineSelect As Boolean

' The X and Y Index of the selected edit object
Public oYSelected As Long
Public oXSelected As Long
Public YSelected  As Long
Public XSelected  As Long

' Flag (KeyCode) if we want to tell the CORE to shift the matrix
Public DoShiftOnMatrix As Integer

' holds all editable objects
Public EditObjects()  As EditX

Sub InitEditObjects()

    Dim Num As Long
    Dim St  As String
    
    Open App.Path & "\data\objects\edit.map" For Input As #1
    
        Input #1, Num
        
        ReDim EditObjects(Num - 1)
        
        For i = 0 To Num - 1
           Line Input #1, St
           
           With EditObjects(i)
              .Object = Trim(Split(St, ",")(0))
              .is_rot = IIf(Val(Split(St, ",")(1)), True, False)
           End With
        Next i
    
    Close #1
    
    ' select object window
    With frmScreen.eObjects
      .Width = SP_W * 10 + 16
      .Height = SP_H * 6 + 16
      .Move 100, frmScreen.ScaleHeight - .Height - SP_H
      .Visible = False
    End With
    
    With frmScreen
      .eObjects_TL.x1 = 0
      .eObjects_TL.X2 = .eObjects.ScaleWidth - 1
      .eObjects_TL.y1 = 0
      .eObjects_TL.Y2 = 0
    
      .eObjects_BL.x1 = 0
      .eObjects_BL.X2 = .eObjects.ScaleWidth - 1
      .eObjects_BL.y1 = .eObjects.ScaleHeight - 1
      .eObjects_BL.Y2 = .eObjects.ScaleHeight - 1
    
      .eObjects_LL.x1 = 0
      .eObjects_LL.X2 = 0
      .eObjects_LL.y1 = 0
      .eObjects_LL.Y2 = .eObjects.ScaleHeight - 1
    
      .eObjects_RL.x1 = .eObjects.ScaleWidth - 1
      .eObjects_RL.X2 = .eObjects.ScaleWidth - 1
      .eObjects_RL.y1 = 0
      .eObjects_RL.Y2 = .eObjects.ScaleHeight - 1
    End With
    
    With frmScreen.scroll_layer
      .Width = SP_W * 10
      .Height = SP_H * 6
      .Move 8, 8
    End With
    
    ' select object props
    With frmScreen.edit_obj
      .ScaleMode = vbPixels
      .AutoRedraw = True
      .ClipControls = True
      .Move 0, 0, frmScreen.scroll_layer.ScaleWidth, 6 * SP_H * Screen.TwipsPerPixelY
    End With
    
    YSelected = -1

End Sub

Sub ShowEditObjects()

    Dim Ani As Animations
  
    Dim eX As Long
    Dim eY As Long
    Dim i  As Long
  
    eX = 0
    eY = 0
    For i = 0 To UBound(EditObjects)
       Ani = GetAnimationInfo(EditObjects(i).Object)
       BitBlt frmScreen.edit_obj.hdc, eX * SP_W, eY * SP_H, SP_W, SP_H, hSprite(H_SPRITES), Ani.x_start * SP_W, Ani.y_start * SP_H, vbSrcCopy
       eX = eX + 1
       If eX > 9 Then
          eX = 0
          eY = eY + 1
       End If
    Next i

    With FontEx
       .Name = "Arial"
       .Size = 7
       .Bold = False
       .Italic = False
       .Strikethrough = False
       .UnderLine = False
       .Charset = 0
       .Colour = vbWhite
    End With
    TextEx.Draw FontEx, SP_H, 0, 0, SP_W, frmScreen.edit_obj.hdc, "CLEAR", F_CENTER Or F_SINGLELINE Or F_VCENTER

    frmScreen.eObjects.Visible = True

End Sub

Sub ShowEditSettings()
  
    If Not InField(CurX, CurY) Then Exit Sub
    
    If GetPropertiesFromObject Then
       frmScreen.eSettings.Visible = True
    End If

End Sub

Function GetPropertiesFromObject() As Boolean

    Dim oName     As String
    Dim oCol      As Byte
    Dim oDir      As Byte
    
    EngineSelect = True
    
    GetObject gMatrix(CurX, CurY).sprite, oName, oDir, oCol

    If Len(oName) = 0 Then Exit Function

    GetPropertiesFromObject = True
    
    With frmScreen
    
        For i = 1 To 7
            If oCol = i Then .obj_col(i).Value = True
        Next i

        .obj_props(0).Enabled = False
        .obj_props(1).Enabled = False
        .obj_props(0).Value = vbUnchecked
        .obj_props(1).Value = vbUnchecked
               
        Select Case oName
        Case "LR"
          .obj_props(1).Value = vbChecked
          .lay_col.Visible = True
          EngineSelect = False
          Exit Function
        Case "LS"
          .obj_props(1).Value = vbUnchecked
          .lay_col.Visible = True
          EngineSelect = False
          Exit Function
        End Select
        
        .lay_col.Visible = False
          
        ' object may be moved in game?
        If ObjectIsMovable(oName) Then
          .obj_props(0).Enabled = True
          .obj_props(0).Value = IIf(eMatrix(CurX, CurY).is_mov = True, vbChecked, vbUnchecked)
        End If
        
        ' object may be rotated in game?
        If GetMaxDir(oName) > 0 Then
          .obj_props(1).Enabled = True
          .obj_props(1).Value = IIf(eMatrix(CurX, CurY).is_rot = True, vbChecked, vbUnchecked)
        End If
          
    End With
    
    EngineSelect = False
    
End Function

Sub ApplyColor(Index As Integer)

  Dim oName As String
  Dim oDir  As Byte
  Dim oCol  As Byte
  
  ' apply new color value in edit/settings mode
  With gMatrix(CurX, CurY)
    GetObject .sprite, oName, oDir, oCol
    .sprite = oName & "_" & oDir
    If oCol > 0 Then
       .sprite = .sprite & ":" & Index
    End If
  End With
  
  CalcBeams

End Sub

Sub ApplySettings()

  With eMatrix(CurX, CurY)
     .is_mov = IIf(frmScreen.obj_props(0).Value = vbChecked, True, False)
     .is_rot = IIf(frmScreen.obj_props(1).Value = vbChecked, True, False)
  End With

End Sub

Sub FixEditMatrix()
  
  Dim oName     As String
  Dim oCol      As Byte
  Dim oDir      As Byte
    
  With eMatrix(CurX, CurY)
     
     GetObject gMatrix(CurX, CurY).sprite, oName, oDir, oCol
   
     Select Case oName
     Case "LS"
       .is_rot = False
       .is_mov = False
     Case "LR"
       .is_rot = True
       .is_mov = False
     Case Else
       If GetMaxDir(oName) = 0 Then
         .is_rot = False
       Else
         .is_rot = EditObjects(SelectedObjectIndex()).is_rot
       End If
       .is_mov = ObjectIsMovable(oName)
     End Select
  
  End With
  
End Sub

Sub SaveLevel(Packname As String, Password As String)

  Dim NumMovable As Long
  Dim LevLine    As String
  Dim oName      As String
  Dim oCol       As Byte
  Dim oDir       As Byte
  Dim nIndex     As String
  Dim Object     As String * 6
  Dim tMapName   As String
  Dim Md         As Byte
  Dim pwmd5()    As Byte
  
  On Error GoTo SaveError
  
  For Y = REAL_Y1 To REAL_Y2
    For X = REAL_X1 To REAL_X2
      If Len(gMatrix(X, Y).sprite) And eMatrix(X, Y).is_mov Then
         NumMovable = NumMovable + 1
      End If
    Next X
  Next Y
 
  If NumMovable > 0 Then
    DoQuestion "Your level has " & NumMovable & " inventory item(s)" & vbCrLf & "Be sure to leave enough room at the right site of your level" & vbCrLf & _
           "The level will need " & NumMovable \ (REAL_Y2 - REAL_Y1 + 2) + 2 & " additional colums!", , vbOKOnly
  End If
  
  If CurrentLevel = 0 Then
     nIndex = GetNextLevelIndex(Packname)
  Else
     nIndex = Right("00" & Trim(Str(CurrentLevel)), 3)
  End If
  CurrentLevel = nIndex

  tMapName = DoInput("NAME YOUR LEVEL [ NR " & nIndex & " ]", MapName)

  If Len(tMapName) = 0 Then Exit Sub

  MapName = UCase(tMapName)

  Open App.Path & "\data\levels\" & Packname & "\level_" & nIndex & ".lev" For Output As #1
  Open App.Path & "\data\levels\" & Packname & "\level_" & nIndex & ".bin" For Binary As #2
  
    BinHdr.Name = Trim(UCase(MapName))
    BinHdr.version = 1000
    pwmd5() = MD5(Password)
    For i = 0 To UBound(BinHdr.passw)
       BinHdr.passw(i) = pwmd5(i + 1)
    Next i
        
    Print #1, Trim(UCase(BinHdr.Name))
    
    Seek #2, Len(BinHdr) + 1
    
    For Y = REAL_Y1 To REAL_Y2
      LevLine = vbNullString
      For X = REAL_X1 To CL_C - 2
         With gMatrix(X, Y)
            Object = .sprite
            Put #2, , Object
            Put #2, , eMatrix(X, Y).is_mov
            Put #2, , eMatrix(X, Y).is_rot
            Put #2, , .is_rot = True
                       
            If X > REAL_X1 Then
               LevLine = LevLine & "."
            End If
            LevLine = LevLine & IIf(eMatrix(X, Y).is_mov = True, Space(6), _
                                    Left(Trim(.sprite) & IIf(eMatrix(X, Y).is_rot = True, "=r", vbNullString) & Space(6), 6))
         End With
      Next X
      Print #1, LevLine
    Next Y
    
    Print #1, Trim(Str(NumMovable))
    
    If NumMovable > 0 Then
    
       ReDim mObjects(NumMovable - 1) As String
       Dim mCurrent                   As Long
    
       mCurrent = 0
       For Y = REAL_Y1 To REAL_Y2
         For X = REAL_X1 To CL_C - 2
            With eMatrix(X, Y)
               If Len(gMatrix(X, Y).sprite) And .is_mov Then
                  GetObject gMatrix(X, Y).sprite, oName, oDir, oCol
                  Md = GetMaxDir(oName)
                  mObjects(mCurrent) = oName & "_" & IIf(Md > 0, Md, oDir) & IIf(oCol > 0, ":" & oCol, vbNullString) & IIf(.is_rot, "=r", vbNullString)
                  mCurrent = mCurrent + 1
               End If
            End With
         Next X
       Next Y
  
       SortArray mObjects()
    
       For i = 0 To NumMovable - 1
          Print #1, mObjects(i)
       Next i
  
    End If
    
    Seek #2, 1
    Put #2, , BinHdr
  
  Close #2
  Close #1
  
  With frmScreen
     If Packname = .levPacks.List(.levPacks.ListIndex) Then
        LoadLevelNames .levPacks.List(.levPacks.ListIndex)
     End If
  End With
  
  Exit Sub
  
SaveError:
ShowError "Error occured in Game_Editor::SaveLevel" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description
  
End Sub

Sub LoadBinLevel(Packname As String, Password As String)

  Dim Entry   As WIN32_FIND_DATA
  Dim fHandle As Long
  Dim fName   As String
  Dim Object  As String * 6
   
  Dim pwmd5() As Byte
  Dim pwvalid As Boolean
        
  On Error GoTo LoadError
        
  With frmScreen
   
      .Cls
      .lstBin.Clear
   
      fHandle = FindFirstFile(App.Path & "\data\levels\" & Packname & "\*.bin", Entry)
      If fHandle <> INVALID_HANDLE_VALUE Then
         Do
            fName = Left(Entry.cFileName, InStr(1, Entry.cFileName, Chr(0)) - 1)
            If Entry.dwFileAttributes And Not vbDirectory Then
               Open App.Path & "\data\levels\" & Packname & "\" & fName For Binary As #1
                  Get #1, , BinHdr
               Close #1
               .lstBin.AddItem Val(Mid(fName, InStr(1, fName, "_") + 1)) & " - " & Trim(BinHdr.Name)
            End If
         Loop While FindNextFile(fHandle, Entry)
         FindClose fHandle
    End If

    .lstBin.ListIndex = -1
    .eLoad.Visible = True

    LockMenuMouse = True
    LockGameMouse = True

    Do
     DoEvents
    Loop Until .eLoad.Visible = False
  
    ' we are here when player clicked "OK" or "CANCEL" in the "load level to edit" list
    ' when user clicked "NEW LEVEL" we will not come here so we have to set the next 2 variables in the called procedure but_newlevel_Click()
  
    LockMenuMouse = False
    LockGameMouse = False
    
    If .lstBin.ListIndex = -1 Then Exit Sub
  
    Open App.Path & "\data\levels\" & Packname & "\level_" & Right("00" & Trim(Str(Val(.lstBin.List(.lstBin.ListIndex)))), 3) & ".bin" For Binary As #1
 
       Get #1, , BinHdr
     
       pwmd5() = MD5(Password)
       pwvalid = True
       
       For i = 1 To UBound(pwmd5)
      '    If pwmd5(i) <> BinHdr.passw(i - 1) Then pwvalid = False
       Next i
       
       If pwvalid = False Then
          ShowError "Password is incorrect, level is locked from editing."
          Close #1
          Exit Sub
       End If
       
       MapName = Trim(BinHdr.Name)
     
       ClearMatrices
     
       For Y = REAL_Y1 To REAL_Y2
          For X = REAL_X1 To CL_C - 2
             With gMatrix(X, Y)
                Get #1, , Object: .sprite = Trim(Object)
                Get #1, , eMatrix(X, Y).is_mov
                Get #1, , eMatrix(X, Y).is_rot
                Get #1, , .is_rot
                If Len(.sprite) Then .is_mov = True
             End With
          Next X
       Next Y
     
    Close #1

    DrawBorder

  End With

  With frmScreen
    CurrentLevel = Val(.lstBin.List(.lstBin.ListIndex))
  End With

  CalcBeams

  Exit Sub

LoadError:
ShowError "Error occured in Game_Editor::LoadBinLevel" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description

End Sub

Function GetNextLevelIndex(Packname As String) As String

    Dim Entry   As WIN32_FIND_DATA
    Dim fHandle As Long
    Dim fExt    As Long
    Dim cI      As Long
    
    cI = 0
   
    fHandle = FindFirstFile(App.Path & "\data\levels\" & Packname & "\*.*", Entry)
    If fHandle <> INVALID_HANDLE_VALUE Then
       Do
          fExt = Val(Mid(Entry.cFileName, InStr(1, Entry.cFileName, "_") + 1))
          If Entry.dwFileAttributes And Not vbDirectory Then
             If fExt > cI Then
                cI = fExt
             End If
          End If
       Loop While FindNextFile(fHandle, Entry)
       FindClose fHandle
    End If
   
    GetNextLevelIndex = Right("00" & Trim(Str(cI + 1)), 3)

End Function

Sub ShiftMatrix(Direction As Integer)
    
    Dim tgM()   As gMatrixInfo
    Dim teM()   As EditMatrix
    Dim Altered As Boolean
    
    On Error GoTo ShiftError
    
    tgM() = gMatrix()
    teM() = eMatrix()
    
    Select Case LCase(Direction)
    Case 38   'UP-KEY
        
        For X = REAL_X1 To REAL_X2
           tgM(X, REAL_Y2) = gMatrix(X, REAL_Y1)
           teM(X, REAL_Y2) = eMatrix(X, REAL_Y1)
        Next X
        For Y = REAL_Y1 To REAL_Y2 - 1
          For X = REAL_X1 To REAL_X2
            tgM(X, Y) = gMatrix(X, Y + 1)
            teM(X, Y) = eMatrix(X, Y + 1)
          Next X
        Next Y
        Altered = True
        
    Case 40   'DOWN-KEY
        
        For X = REAL_X1 To REAL_X2
           tgM(X, REAL_Y1) = gMatrix(X, REAL_Y2)
           teM(X, REAL_Y1) = eMatrix(X, REAL_Y2)
        Next X
        For Y = REAL_Y1 + 1 To REAL_Y2
          For X = REAL_X1 To REAL_X2
            tgM(X, Y) = gMatrix(X, Y - 1)
            teM(X, Y) = eMatrix(X, Y - 1)
          Next X
        Next Y
        Altered = True
        
    Case 37   'LEFT-KEY
        
        For Y = REAL_Y1 To REAL_Y2
           tgM(REAL_X2, Y) = gMatrix(REAL_X1, Y)
           teM(REAL_X2, Y) = eMatrix(REAL_X1, Y)
        Next Y
        For Y = REAL_Y1 To REAL_Y2
          For X = REAL_X1 To REAL_X2 - 1
            tgM(X, Y) = gMatrix(X + 1, Y)
            teM(X, Y) = eMatrix(X + 1, Y)
          Next X
        Next Y
        Altered = True
        
    Case 39   'RIGHT-KEY
        
        For Y = REAL_Y1 To REAL_Y2
           tgM(REAL_X1, Y) = gMatrix(REAL_X2, Y)
           teM(REAL_X1, Y) = eMatrix(REAL_X2, Y)
        Next Y
        For Y = REAL_Y1 To REAL_Y2
          For X = REAL_X1 + 1 To REAL_X2
            tgM(X, Y) = gMatrix(X - 1, Y)
            teM(X, Y) = eMatrix(X - 1, Y)
          Next X
        Next Y
        Altered = True
        
    End Select
  
    If Altered Then
       gMatrix() = tgM()
       eMatrix() = teM()
       CalcBeams
    End If
    
    Exit Sub
    
ShiftError:
ShowError "Error occurred in Game_Editor::ShiftMatrix" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description

End Sub

Function DoEditPrivilegesCheck() As Boolean

  Dim pwmd5() As Byte
  Dim pwvalid As Boolean
  Dim FSO As FileSystemObject
  Set FSO = New FileSystemObject
    
  DoEditPrivilegesCheck = False
    
  On Error GoTo PrivError
    
  With frmScreen
  
     If .o_edit_pack(0).Value = True Then
        
        If .lstEditPack.ListIndex = -1 Then
           ShowError "You must select a levelpack to edit."
           Exit Function
        End If
           
        Open App.Path & "\data\levels\" & .lstEditPack.List(.lstEditPack.ListIndex) & "\user.prv" For Binary As #1
          
           Get #1, , UsrPrv
       
           If .lstEditPack.List(.lstEditPack.ListIndex) <> Trim(UsrPrv.user) Then
              ShowError "User privilege file damaged or unknown packname ..."
              Close #1
              Exit Function
           End If

           pwmd5() = MD5(.ed_ver_passw)
           pwvalid = True
       
           For i = 1 To UBound(pwmd5)
              If pwmd5(i) <> UsrPrv.passw(i - 1) Then pwvalid = False
           Next i
                  
           If pwvalid = False Then
              ShowError "Incorrect password for current level pack."
              Close #1
              Exit Function
           End If
           
        Close #1
     
     Else
     
        If Len(.ed_packname) = 0 Then
           ShowError "You must enter a packname."
           Exit Function
        End If
        
        If Len(.ed_password) < 3 Then
           ShowError "Password must be at least 3 characters."
           Exit Function
        End If
        
        If .ed_password <> .ed_retype Then
           ShowError "Passwords do not match."
           Exit Function
        End If
                
        If FSO.FolderExists(App.Path & "\data\levels\" & .ed_packname) = False Then
           FSO.CreateFolder App.Path & "\data\levels\" & .ed_packname
        End If
        
        UsrPrv.user = UCase(Trim(.ed_packname))
        
        pwmd5() = MD5(.ed_password)
        
        For i = 0 To UBound(UsrPrv.passw)
           UsrPrv.passw(i) = pwmd5(i + 1)
        Next i
        
        Open App.Path & "\data\levels\" & .ed_packname & "\user.prv" For Binary As #1
           Put #1, , UsrPrv
        Close #1
     
        .ed_ver_passw = .ed_password
     
        'EngineSelect = True
        LoadLevelPackNames .ed_packname
        'EngineSelect = False
     
     End If
     
   End With

   DoEditPrivilegesCheck = True
  
   Exit Function

PrivError:
ShowError "Error occured in Game_Editor::DoEditPrivilegesCheck" & vbCrLf & "#Error: " & Err & vbCrLf & "Description: " & Err.Description
  
End Function

Function SelectedObjectIndex() As Long
  
  If YSelected * 10 + XSelected > UBound(EditObjects) Then
     YSelected = oYSelected
     XSelected = oXSelected
  End If
  
  SelectedObjectIndex = IIf(YSelected = -1, 0, YSelected * 10 + XSelected)

End Function

Sub SortArray(aArray() As String)

   Dim t    As String
   Dim S    As Boolean
      
   Do
     S = False
     For i = LBound(aArray) To UBound(aArray) - 1
       If aArray(i + 1) < aArray(i) Then
          t = aArray(i)
          aArray(i) = aArray(i + 1)
          aArray(i + 1) = t
          S = True
       End If
     Next i
   Loop Until S = False

End Sub

