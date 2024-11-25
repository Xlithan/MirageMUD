Attribute VB_Name = "Graphics"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Public Sub DrawNewChar()


    Dim ListIndexAvatar As Long
    
    ListIndexAvatar = frmMainMenu.cmbClass.ListIndex

    If ListIndexAvatar < 0 Then Exit Sub
    If Class(ListIndexAvatar + 1).Avatar = 0 Then Exit Sub
    
    frmMainMenu.picPic.Picture = LoadPicture(App.Path & AVATARS_PATH & Class(ListIndexAvatar + 1).Avatar & GFX_EXT)
    
End Sub

Public Sub DrawSelChar(ByVal i As Long)
    
    If CharAvatars(i) = 0 Or CharAvatars(i) > NumAvatars Then
        frmMainMenu.picSelChar.Picture = LoadPicture(App.Path & AVATARS_PATH & 0 & GFX_EXT)
        Exit Sub

    End If
    
    frmMainMenu.picSelChar.Picture = LoadPicture(App.Path & AVATARS_PATH & CharAvatars(i) & GFX_EXT)
    
    With frmMainMenu
        .picSelChar.ScaleMode = 3
        .picSelChar.AutoRedraw = True
        .picSelChar.PaintPicture .picSelChar.Picture, _
        0, 0, .picSelChar.ScaleWidth, .picSelChar.ScaleHeight, _
        0, 0, .picSelChar.Picture.Width / 26.46, _
        .picSelChar.Picture.Height / 26.46
    
        .picSelChar.Picture = .picSelChar.Image
    End With
End Sub

Public Sub DrawTargetChar()

    Dim pAvatar As Long
    
    pAvatar = GetPlayerAvatar(PlayerSel)

    If pAvatar < 0 Then Exit Sub
    
    frmMainGame.picTarget.Picture = LoadPicture(App.Path & AVATARS_PATH & pAvatar & GFX_EXT)
    
    With frmMainGame
        .picTarget.ScaleMode = 3
        .picTarget.AutoRedraw = True
        .picTarget.PaintPicture .picTarget.Picture, _
        0, 0, .picTarget.ScaleWidth, .picTarget.ScaleHeight, _
        0, 0, .picTarget.Picture.Width / 26.46, _
        .picTarget.Picture.Height / 26.46
    
        .picTarget.Picture = .picTarget.Image
    End With
End Sub

Public Sub NpcEditorBltAvatar()

    Dim AvatarPic As Long
    
    AvatarPic = frmNpcEditor.scrlAvatar.Value
    
    If AvatarPic < 1 Or AvatarPic > NumNPCAvatars Then
        frmNpcEditor.picAvatar.Cls

        Exit Sub

    End If
    
    frmNpcEditor.picAvatar.Picture = LoadPicture(App.Path & NPCS_PATH & AvatarPic & GFX_EXT)
    
    With frmNpcEditor
        .picAvatar.ScaleMode = 3
        .picAvatar.AutoRedraw = True
        .picAvatar.PaintPicture .picAvatar.Picture, _
        0, 0, .picAvatar.ScaleWidth, .picAvatar.ScaleHeight, _
        0, 0, .picAvatar.Picture.Width / 26.46, _
        .picAvatar.Picture.Height / 26.46
    
        .picAvatar.Picture = .picAvatar.Image
    End With
End Sub

Public Sub SpellEditorBltSpell()

    Dim SpellPic As Long
    
    SpellPic = frmSpellEditor.scrlPic.Value
    
    If SpellPic < 1 Or SpellPic > NumSpells Then
        frmSpellEditor.picPic.Cls

        Exit Sub

    End If
    
    frmSpellEditor.picPic.Picture = LoadPicture(App.Path & SPELLS_PATH & SpellPic & GFX_EXT)
    
End Sub

Public Sub BltInventory(ItemNum As Integer)

    Dim PicNum As Integer
    
    PicNum = Item(ItemNum).Pic
    
    If PicNum < 1 Or PicNum > numitems Then
        frmMainGame.picInvSelected.Cls

        Exit Sub

    End If
    
    frmMainGame.picInvSelected.Picture = LoadPicture(App.Path & ITEMS_PATH & PicNum & GFX_EXT)
    
    With frmMainGame
        .picInvSelected.ScaleMode = 3
        .picInvSelected.AutoRedraw = True
        .picInvSelected.PaintPicture .picInvSelected.Picture, _
        0, 0, .picInvSelected.ScaleWidth, .picInvSelected.ScaleHeight, _
        0, 0, .picInvSelected.Picture.Width / 26.46, _
        .picInvSelected.Picture.Height / 26.46
    
        .picInvSelected.Picture = .picInvSelected.Image
    End With
    
End Sub

Public Sub BltTarget(ByVal Entity As Byte)

    Dim AvatarPic As Integer
    
    'frmMainGame.picTarget.Cls
    Select Case Entity

        Case 0 ' Players

            'AvatarPic = Item(RoomItem(ItemSel, GetPlayerRoom(MyIndex))).Pic
        Case 1 ' NPCs
            AvatarPic = Npc(RoomNpc(NPCSel, GetPlayerRoom(MyIndex)).num).Avatar

            If AvatarPic < 1 Or AvatarPic > NumNPCAvatars Then
                frmMainGame.picTarget.Cls

                Exit Sub

            End If
            
            frmMainGame.picTarget.Picture = LoadPicture(App.Path & NPCS_PATH & AvatarPic & GFX_EXT)
            
            With frmMainGame
                .picTarget.ScaleMode = 3
                .picTarget.AutoRedraw = True
                .picTarget.PaintPicture .picTarget.Picture, _
                0, 0, .picTarget.ScaleWidth, .picTarget.ScaleHeight, _
                0, 0, .picTarget.Picture.Width / 26.46, _
                .picTarget.Picture.Height / 26.46
            
                .picTarget.Picture = .picTarget.Image
            End With

        Case 2 ' Items
            AvatarPic = Item(RoomItem(ItemSel, GetPlayerRoom(MyIndex)).num).Pic

            If AvatarPic < 1 Or AvatarPic > numitems Then
                frmMainGame.picTarget.Cls

                Exit Sub

            End If
            
            frmMainGame.picTarget.Picture = LoadPicture(App.Path & ITEMS_PATH & AvatarPic & GFX_EXT)
            
            With frmMainGame
                .picTarget.ScaleMode = 3
                .picTarget.AutoRedraw = True
                .picTarget.PaintPicture .picTarget.Picture, _
                0, 0, .picTarget.ScaleWidth, .picTarget.ScaleHeight, _
                0, 0, .picTarget.Picture.Width / 26.46, _
                .picTarget.Picture.Height / 26.46
            
                .picTarget.Picture = .picTarget.Image
            End With
    End Select
    
End Sub

Public Sub ItemEditorBltItem()

    Dim ItemPic As Integer

    Dim sRECT   As D3DRECT
    
    ItemPic = frmItemEditor.scrlPic.Value
    
    If ItemPic < 1 Or ItemPic > numitems Then
        frmItemEditor.picPic.Cls

        Exit Sub

    End If
    
    frmItemEditor.picPic.Picture = LoadPicture(App.Path & ITEMS_PATH & ItemPic & GFX_EXT)
    
    With frmItemEditor
        .picPic.ScaleMode = 3
        .picPic.AutoRedraw = True
        .picPic.PaintPicture .picPic.Picture, _
        0, 0, .picPic.ScaleWidth, .picPic.ScaleHeight, _
        0, 0, .picPic.Picture.Width / 26.46, _
        .picPic.Picture.Height / 26.46
    
        .picPic.Picture = .picPic.Image
    End With
    
End Sub
