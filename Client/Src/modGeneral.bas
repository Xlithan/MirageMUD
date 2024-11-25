Attribute VB_Name = "General"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds (32-bit)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'For Clear functions
Public Declare Sub ZeroMemory _
               Lib "kernel32.dll" _
               Alias "RtlZeroMemory" (Destination As Any, _
                                      ByVal length As Long)

Public Sub Main()
    
    frmSendGetData.Visible = True
    Call SetStatus("Loading...")
    
    Call SetStatus("Loading Game Data...")
    Call LoadDataFile
    
    GettingRoom = True
    'vbQuote = ChrW$(34) ' "
    
    Load frmMainGame

    ' Update the form with the game's name
    frmMainGame.Caption = "MirageMUD"
    
    ' randomize rnd's seed
    Randomize
    
    'Call InitFont
    
    Call SetStatus("Initializing TCP settings...")
    GAME_IP = Trim$(GameData.IP)
    GAME_PORT = GameData.Port
    Call TcpInit
    
    Call SetStatus("Initializing DirectX...")
    ' DX7 Master Object is already created, early binding
    Call CheckAvatars
    Call CheckNPCAvatars
    Call CheckSpells
    Call CheckItems
    
    frmSendGetData.Visible = False
    
    Load frmMainMenu ' this line also initalizes directX
    
    Init_Music
    If GameData.Music = 1 Then
        Play_Music "Menu.mp3"
    End If
    
    frmMainMenu.Visible = True
End Sub

Public Sub GameInit()
    Unload frmMainMenu
    
    frmSendGetData.Visible = False
    
    'InitRoomTables
    frmMainGame.Show
    ' Set the focus
    Call SetFocusOnChat
    Call UpdateSpells
    
    frmMainGame.Caption = "MirageMUD - " & Trim$(Player(MyIndex).name)
    
End Sub

Public Sub DestroyGame()
    ' break out of GameLoop
    InGame = False
    
    Destroy_Music
    
    Call DestroyTCP
    
    Call UnloadAllForms

    End

End Sub

Public Sub UnloadAllForms()

    Dim frm As Form
    
    For Each frm In VB.Forms

        Unload frm
    Next

End Sub

Public Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
    'DoEvents
End Sub

Public Sub AddText(ByVal msg As String)
    Dim s As String
    
    s = vbNewLine & msg
    
    With frmMainGame.txtChat
        .SelStart = Len(.Text)
        .SelRTF = BbCode2Rtf(s, .Font)
        
        .SelStart = Len(.Text) - 1
        
        ' Prevent players from name spoofing
        '.SelHangingIndent = 15
    End With
    
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, msg As String, NewLine As Boolean)

    If NewLine Then
        Txt.Text = Txt.Text + msg + vbCrLf
    Else
        Txt.Text = Txt.Text + msg
    End If
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Public Sub SetFocusOnChat()

    On Error Resume Next 'prevent RTE5, no way to handle error

    frmMainGame.txtMyChat.SetFocus
End Sub

Public Function RemoveBBCode(ByVal Text As String)
    ' Available BB Codes
    ' [b]|[/b]|[i]|[/i]|[u]|[/u]|[size=|[/size]|[color=|[/color]|[url=|[/url]|[font=|[/font]|[right]|[/right]|[center]|[/center]|[table=|[/table]|[row]|[row=|[/row]|[col]
    
    ' Can either just remove the [] brackets which breaks the BB Code, or can uncomment and remove the entire tag from the text.
    ' It's unsure if using so many Replace$ functions causes performance issues.
    
    Text = Replace$(Text, "[", vbNullString)
    Text = Replace$(Text, "]", vbNullString)
'    Text = Replace$(Text, "[b]", vbNullString)
'    Text = Replace$(Text, "[i]", vbNullString)
'    Text = Replace$(Text, "[u]", vbNullString)
'    Text = Replace$(Text, "[size=", vbNullString)
'    Text = Replace$(Text, "[color=", vbNullString)
'    Text = Replace$(Text, "[url=", vbNullString)
'    Text = Replace$(Text, "[font=", vbNullString)
'    Text = Replace$(Text, "[table=", vbNullString)
'    Text = Replace$(Text, "[row=", vbNullString)
'    Text = Replace$(Text, "[row]", vbNullString)
'    Text = Replace$(Text, "[col]", vbNullString)
'    Text = Replace$(Text, "[right]", vbNullString)
'    Text = Replace$(Text, "[center]", vbNullString)
'
'    Text = Replace$(Text, "[/b]", vbNullString)
'    Text = Replace$(Text, "[/i]", vbNullString)
'    Text = Replace$(Text, "[/u]", vbNullString)
'    Text = Replace$(Text, "[/size]", vbNullString)
'    Text = Replace$(Text, "[/color]", vbNullString)
'    Text = Replace$(Text, "[/url]", vbNullString)
'    Text = Replace$(Text, "[/font]", vbNullString)
'    Text = Replace$(Text, "[/table]", vbNullString)
'    Text = Replace$(Text, "[/row]", vbNullString)
'    Text = Replace$(Text, "[/right]", vbNullString)
'    Text = Replace$(Text, "[/center]", vbNullString)
    
    MyText = Text
End Function

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Sub RefreshEntityList(ByVal Index As Long, ByVal List As Integer)

    Dim NPCNum  As Long

    Dim ItemNum As Long

    Dim Room     As Long

    Dim i       As Long, aNPC As Long, aItem As Long

    Room = GetPlayerRoom(Index)
    aNPC = 1
    aItem = 1
    
    frmMainGame.lblRoomNum.Caption = GetPlayerRoom(MyIndex)
    
    If Room > 0 Then

        Select Case List

            Case 1
                ' Refresh NPC list
                frmMainGame.lstNPCs.Clear
        
                For i = 1 To MAX_ROOM_NPCS
                    NPCNum = RoomNpc(i, Room).num

                    If NPCNum > 0 Then
                        NPCLst(aNPC) = i
                        frmMainGame.lstNPCs.AddItem Trim$(Npc(NPCNum).name)
                        aNPC = aNPC + 1
                    End If

                Next

            Case 2
                ' Refresh item list
                frmMainGame.lstItems.Clear
                
                For i = 1 To MAX_ROOM_ITEMS
                    ItemNum = RoomItem(i, Room).num

                    If ItemNum > 0 Then
                        ItemLst(aItem) = i

                        If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                            frmMainGame.lstItems.AddItem Trim$(Item(ItemNum).name) & " (" & RoomItem(i, Room).Value & ")"
                        Else
                            frmMainGame.lstItems.AddItem Trim$(Item(ItemNum).name)
                        End If

                        aItem = aItem + 1
                    End If

                Next

        End Select

    End If
    
End Sub

Public Sub MovePicture(PB As PictureBox, _
                       Button As Integer, _
                       Shift As Integer, _
                       x As Single, _
                       y As Single)

    Dim GlobalX As Integer

    Dim GlobalY As Integer
    
    GlobalX = PB.Left
    GlobalY = PB.Top
    
    If Button = 1 Then
        PB.Left = GlobalX + x - SOffsetX
        PB.Top = GlobalY + y - SOffsetY
    End If

End Sub

Public Function isLoginLegal(ByVal Username As String, _
                             ByVal Password As String) As Boolean

    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean

    Dim i As Long
    
    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, GAME_NAME)

            Exit Function

        End If

    Next
    
    isStringLegal = True
    
End Function

