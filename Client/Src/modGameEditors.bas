Attribute VB_Name = "GameEditors"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Public Editor      As Byte ' which game editor is used

Public EditorIndex As Long ' index number

' Room attribute data
Public EditorData1 As Long

Public EditorData2 As Long

Public EditorData3 As Long

' ////////////////
' // Room Editor //
' ////////////////

Public Sub RoomEditorInit()

    On Error GoTo ErrorHandle
    
    Editor = EDITOR_ROOM
    
    EditorData1 = 0
    EditorData2 = 0
    EditorData3 = 0
    
    frmRoomEditor.Show vbModal
    
    Exit Sub
    
ErrorHandle:
    
    Select Case Err
        
        Case 380
            RoomEditorInit
    End Select
    
End Sub

Public Sub RoomEditorCancel()
    Editor = EDITOR_NONE
    
    Unload frmRoomEditor
    
    Call LoadRooms(GetPlayerRoom(MyIndex))
    Call InitRoomData
    
End Sub

Public Sub RoomEditorSend()
    Call SendRoom
    Call RoomEditorCancel
End Sub

' /////////////////
' // Item Editor //
' /////////////////

Public Sub ItemEditorInit()
    
    With frmItemEditor
        .txtName.Text = Trim$(Item(EditorIndex).name)
        .scrlPic.Value = Item(EditorIndex).Pic
        .cmbType.ListIndex = Item(EditorIndex).Type
        
        If (.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            .fraEquipment.Visible = True
            .scrlDurability.Value = Item(EditorIndex).Data1
            .scrlStrength.Value = Item(EditorIndex).Data2
        Else
            .fraEquipment.Visible = False
        End If
        
        If (.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
            .fraVitals.Visible = True
            .scrlVitalMod.Value = Item(EditorIndex).Data1
        Else
            .fraVitals.Visible = False
        End If
        
        If (.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            .fraSpell.Visible = True
            If Item(EditorIndex).Data1 > 0 Then
                .scrlSpell.Value = Item(EditorIndex).Data1
            End If
        Else
            .fraSpell.Visible = False
        End If
        
        Call ItemEditorBltItem
        
        .Show vbModal
    End With

End Sub

Public Sub ItemEditorOk()
    
    With Item(EditorIndex)
        .Data1 = 0
        .Data2 = 0
        .Data3 = 0
        
        .name = frmItemEditor.txtName.Text
        .Pic = frmItemEditor.scrlPic.Value
        .Type = frmItemEditor.cmbType.ListIndex
        
        If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) Then
            If (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
                .Data1 = frmItemEditor.scrlDurability.Value
                .Data2 = frmItemEditor.scrlStrength.Value
            End If
        End If
        
        If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) Then
            If (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
                .Data1 = frmItemEditor.scrlVitalMod.Value
            End If
        End If
        
        If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            .Data1 = frmItemEditor.scrlSpell.Value
        End If

    End With
    
    Call SendSaveItem(EditorIndex)
    
    Editor = EDITOR_NONE
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    Editor = EDITOR_NONE
    Unload frmItemEditor
End Sub

' ////////////////
' // Npc Editor //
' ////////////////

Public Sub NpcEditorInit()

    With frmNpcEditor
        .txtName.Text = Trim$(Npc(EditorIndex).name)
        .txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
        .scrlAvatar.Value = Npc(EditorIndex).Avatar
        .txtSpawnSecs.Text = CStr(Npc(EditorIndex).SpawnSecs)
        .cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
        .scrlRange.Value = Npc(EditorIndex).Range
        .txtChance.Text = CStr(Npc(EditorIndex).DropChance)
        .scrlNum.Value = Npc(EditorIndex).DropItem
        .scrlValue.Value = Npc(EditorIndex).DropItemValue
        .scrlStrength.Value = Npc(EditorIndex).Stat(Stats.Strength)
        .scrlDefense.Value = Npc(EditorIndex).Stat(Stats.Defense)
        .scrlSpeed.Value = Npc(EditorIndex).Stat(Stats.speed)
        .scrlMagic.Value = Npc(EditorIndex).Stat(Stats.Magic)
        
        .Show vbModal
    End With
    
    Call NpcEditorBltAvatar
    
End Sub

Public Sub NpcEditorOk()

    With Npc(EditorIndex)
        .name = frmNpcEditor.txtName.Text
        .AttackSay = frmNpcEditor.txtAttackSay.Text
        .Avatar = frmNpcEditor.scrlAvatar.Value
        .SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
        .Behavior = frmNpcEditor.cmbBehavior.ListIndex
        .Range = frmNpcEditor.scrlRange.Value
        .DropChance = Val(frmNpcEditor.txtChance.Text)
        .DropItem = frmNpcEditor.scrlNum.Value
        .DropItemValue = frmNpcEditor.scrlValue.Value
        .Stat(Stats.Strength) = frmNpcEditor.scrlStrength.Value
        .Stat(Stats.Defense) = frmNpcEditor.scrlDefense.Value
        .Stat(Stats.speed) = frmNpcEditor.scrlSpeed.Value
        .Stat(Stats.Magic) = frmNpcEditor.scrlMagic.Value
    End With
    
    Call SendSaveNpc(EditorIndex)
    Editor = EDITOR_NONE
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    Editor = EDITOR_NONE
    Unload frmNpcEditor
End Sub

' /////////////////
' // Shop Editor //
' /////////////////

Public Sub ShopEditorInit()

    Dim i As Long

    With frmShopEditor
        .txtName.Text = Trim$(Shop(EditorIndex).name)
        .txtJoinSay.Text = Trim$(Shop(EditorIndex).JoinSay)
        .txtLeaveSay.Text = Trim$(Shop(EditorIndex).LeaveSay)
        .chkFixesItems.Value = Shop(EditorIndex).FixesItems
        
        .cmbItemGive.Clear
        .cmbItemGive.AddItem "None"
        .cmbItemGet.Clear
        .cmbItemGet.AddItem "None"
        
        For i = 1 To MAX_ITEMS
            .cmbItemGive.AddItem i & ": " & Trim$(Item(i).name)
            .cmbItemGet.AddItem i & ": " & Trim$(Item(i).name)
        Next
        
        .cmbItemGive.ListIndex = 0
        .cmbItemGet.ListIndex = 0
        
        .Show vbModal
    End With
    
    Call UpdateShopTrade
    
End Sub

Public Sub UpdateShopTrade()

    Dim i         As Long
    
    Dim GetItem   As Long

    Dim GetValue  As Long

    Dim GiveItem  As Long

    Dim GiveValue As Long
    
    frmShopEditor.lstTradeItem.Clear
    
    For i = 1 To MAX_TRADES

        With Shop(EditorIndex).TradeItem(i)
            GetItem = .GetItem
            GetValue = .GetValue
            GiveItem = .GiveItem
            GiveValue = .GiveValue
        End With
        
        If GetItem > 0 And GiveItem > 0 Then
            frmShopEditor.lstTradeItem.AddItem i & ": " & GiveValue & " " & Trim$(Item(GiveItem).name) & " for " & GetValue & " " & Trim$(Item(GetItem).name)
        Else
            frmShopEditor.lstTradeItem.AddItem "Empty Trade Slot"
        End If

    Next
    
    frmShopEditor.lstTradeItem.ListIndex = 0
End Sub

Public Sub ShopEditorOk()

    With Shop(EditorIndex)
        .name = frmShopEditor.txtName.Text
        .JoinSay = frmShopEditor.txtJoinSay.Text
        .LeaveSay = frmShopEditor.txtLeaveSay.Text
        .FixesItems = frmShopEditor.chkFixesItems.Value
    End With
    
    Call SendSaveShop(EditorIndex)
    Editor = EDITOR_NONE
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    Editor = EDITOR_NONE
    Unload frmShopEditor
End Sub

' //////////////////
' // Spell Editor //
' //////////////////

Public Sub SpellEditorInit()

    Dim i As Long
    
    With frmSpellEditor
        .cmbClassReq.AddItem "All Classes"

        For i = 1 To Max_Classes
            .cmbClassReq.AddItem Trim$(Class(i).name)
        Next
        
        .txtName.Text = Trim$(Spell(EditorIndex).name)
        .scrlPic = Spell(EditorIndex).Pic
        .scrlMPReq.Value = Spell(EditorIndex).MPReq
        .cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
        .scrlLevelReq.Value = Spell(EditorIndex).LevelReq
        
        .cmbType.ListIndex = Spell(EditorIndex).Type

        If Spell(EditorIndex).Type <> SPELL_TYPE_GIVEITEM Then
            .fraVitals.Visible = True
            .fraGiveItem.Visible = False
            .scrlVitalMod.Value = Spell(EditorIndex).Data1
        Else
            .fraVitals.Visible = False
            .fraGiveItem.Visible = True
            .scrlItemNum.Value = Spell(EditorIndex).Data1
            .scrlItemValue.Value = Spell(EditorIndex).Data2
        End If
        
        .Show vbModal
    End With

End Sub

Public Sub SpellEditorOk()

    With Spell(EditorIndex)
        .name = frmSpellEditor.txtName.Text
        .Pic = frmSpellEditor.scrlPic.Value
        .MPReq = frmSpellEditor.scrlMPReq
        .ClassReq = frmSpellEditor.cmbClassReq.ListIndex
        .LevelReq = frmSpellEditor.scrlLevelReq.Value
        .Type = frmSpellEditor.cmbType.ListIndex
        
        If .Type <> SPELL_TYPE_GIVEITEM Then
            .Data1 = frmSpellEditor.scrlVitalMod.Value
        Else
            .Data1 = frmSpellEditor.scrlItemNum.Value
            .Data2 = frmSpellEditor.scrlItemValue.Value
        End If
        
        .Data3 = 0
    End With
    
    Call SendSaveSpell(EditorIndex)
    Editor = EDITOR_NONE
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    Editor = EDITOR_NONE
    Unload frmSpellEditor
End Sub

