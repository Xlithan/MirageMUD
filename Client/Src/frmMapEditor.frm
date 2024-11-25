VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRoomEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Room Editor"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   781
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgMusic 
      Left            =   11160
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtColor 
      Height          =   285
      Left            =   5280
      TabIndex        =   42
      Top             =   7080
      Width           =   2055
   End
   Begin VB.ComboBox cmbColor 
      Height          =   315
      ItemData        =   "frmMapEditor.frx":0000
      Left            =   3360
      List            =   "frmMapEditor.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Frame fraItems 
      Caption         =   "Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2775
      Left            =   7440
      TabIndex        =   31
      Top             =   3000
      Width           =   4215
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1800
         Width           =   2895
      End
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMapEditor.frx":0004
         Left            =   120
         List            =   "frmMapEditor.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1320
         Width           =   2895
      End
      Begin VB.ComboBox cmbItem 
         Height          =   315
         HelpContextID   =   4
         Index           =   5
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtAmount 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   43
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   44
         Text            =   "0"
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   46
         Text            =   "0"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   3240
         TabIndex        =   47
         Text            =   "0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtAmount 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   3240
         TabIndex        =   45
         Text            =   "0"
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Room Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4455
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtExitDesc 
         Height          =   615
         Left            =   240
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   3720
         Width           =   6495
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1320
         TabIndex        =   29
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtLongDesc 
         Height          =   1095
         Left            =   240
         MaxLength       =   512
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   2280
         Width           =   6495
      End
      Begin VB.TextBox txtShortDesc 
         Height          =   735
         Left            =   240
         MaxLength       =   256
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Label lblExitDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   3480
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Room Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Long Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Short Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2775
      Left            =   7440
      TabIndex        =   17
      Top             =   120
      Width           =   4215
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   5
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2280
         Width           =   3975
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1320
         Width           =   3975
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   840
         Width           =   3975
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   1
         ItemData        =   "frmMapEditor.frx":0008
         Left            =   120
         List            =   "frmMapEditor.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   360
         Width           =   3975
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Index           =   4
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1800
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   2895
      Begin VB.TextBox txtBootRoom 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   15
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Boot Room"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Room Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2295
      Left            =   3360
      TabIndex        =   8
      Top             =   4680
      Width           =   3975
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Height          =   255
         Left            =   2160
         TabIndex        =   52
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   255
         Left            =   720
         TabIndex        =   50
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   3000
         TabIndex        =   49
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   2160
         TabIndex        =   48
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cmbShop 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmMapEditor.frx":000C
         Left            =   720
         List            =   "frmMapEditor.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblMusic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music.mp3"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Shop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Moral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Room Links"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   2895
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         TabIndex        =   38
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   330
         Left            =   2160
         TabIndex        =   37
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtWest 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   240
         TabIndex        =   6
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtEast 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1440
         TabIndex        =   5
         Text            =   "0"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtSouth 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   840
         TabIndex        =   4
         Text            =   "0"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtNorth 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   840
         TabIndex        =   3
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblRoom 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   7080
      Width           =   3975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Room"
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   6480
      Width           =   3975
   End
End
Attribute VB_Name = "frmRoomEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Private Sub cmbColor_Click()
    Select Case cmbColor.ListIndex
    
        Case 0
            txtColor.Text = "[color=#000000]"
        Case 1
            txtColor.Text = "[color=#0000FF]"
        Case 2
            txtColor.Text = "[color=#00FF00]"
        Case 3
            txtColor.Text = "[color=#00FFFF]"
        Case 4
            txtColor.Text = "[color=#FF0000]"
        Case 5
            txtColor.Text = "[color=#FF00FF]"
        Case 6
            txtColor.Text = "[color=#964B00]"
        Case 7
            txtColor.Text = "[color=#808080]"
        Case 8
            txtColor.Text = "[color=#A9A9A9]"
        Case 9
            txtColor.Text = "[color=#8888ff]"
        Case 10
            txtColor.Text = "[color=#66ff00]"
        Case 11
            txtColor.Text = "[color=#8AFFFF]"
        Case 12
            txtColor.Text = "[color=#EE4B2B]"
        Case 13
            txtColor.Text = "[color=#FFC0CB]"
        Case 14
            txtColor.Text = "[color=#FFFF00]"
        Case 15
            txtColor.Text = "[color=#FFFFFF]"
            
    End Select
End Sub

Private Sub cmbItem_Click(Index As Integer)

    If cmbItem(Index).ListIndex < 1 Then
        cmbItem(Index).Width = 3975
        txtAmount(Index).Enabled = False
        txtAmount(Index).Text = 0
        Exit Sub
    End If
    If Item(cmbItem(Index).ListIndex).Type = ITEM_TYPE_CURRENCY Then
        cmbItem(Index).Width = 2895
        txtAmount(Index).Enabled = True
        txtAmount(Index).Text = 0
    Else
        cmbItem(Index).Width = 3975
        txtAmount(Index).Enabled = False
        txtAmount(Index).Text = 0
    End If
    
End Sub

Private Sub cmdBrowse_Click()
    
    ' ".mid", ".s3m", ".mod", ".wav", ".mp3", ".ogg", ".wma"
    dlgMusic.InitDir = App.Path & MUSIC_PATH
    dlgMusic.Filter = "Music (*.mid), (*.s3m), (*.mod), (*.wav), (*.mp3), (*.ogg), (*.wma)"
    dlgMusic.DialogTitle = "Select File"
    dlgMusic.ShowOpen
    
    lblMusic.Caption = dlgMusic.FileTitle
    
End Sub

Private Sub cmdPlay_Click()
    Play_Music lblMusic.Caption
End Sub

Private Sub cmdRemove_Click()
    lblMusic.Caption = vbNullString
End Sub

Private Sub cmdStop_Click()
    Stop_Music
End Sub

Private Sub Form_Load()

    Dim X As Long

    Dim y As Long

    Dim i As Long
    
    txtName.Text = Trim$(Room.name)
    txtShortDesc.Text = Trim$(Room.sDesc)
    txtLongDesc.Text = Trim$(Room.lDesc)
    txtExitDesc.Text = Trim$(Room.eDesc)
    txtNorth.Text = CStr(Room.North)
    txtSouth.Text = CStr(Room.South)
    txtWest.Text = CStr(Room.West)
    txtEast.Text = CStr(Room.East)
    cmbMoral.ListIndex = Room.Moral
    lblMusic.Caption = Room.Music
    txtBootRoom.Text = CStr(Room.BootRoom)
    
    With cmbColor
        .AddItem "Black" ' 0
        .AddItem "Blue" ' 1
        .AddItem "Green" ' 2
        .AddItem "Cyan" ' 3
        .AddItem "Red" ' 4
        .AddItem "Magenta" ' 5
        .AddItem "Brown" ' 6
        .AddItem "Grey" ' 7
        .AddItem "Dark Grey" ' 8
        .AddItem "Bright Blue" ' 9
        .AddItem "Bright Green" ' 10
        .AddItem "Bright Cyan" ' 11
        .AddItem "Bright Red" ' 12
        .AddItem "Pink" ' 13
        .AddItem "Yellow" ' 14
        .AddItem "White" ' 15
    End With
    
    cmbShop.AddItem "No Shop"

    For X = 1 To MAX_SHOPS
        cmbShop.AddItem X & ": " & Trim$(Shop(X).name)
    Next

    cmbShop.ListIndex = Room.Shop
    
    ' NPC list
    For X = 1 To MAX_ROOM_NPCS
        cmbNpc(X).AddItem "No NPC"
    Next
    
    For y = 1 To MAX_NPCS
        For X = 1 To MAX_ROOM_NPCS
            cmbNpc(X).AddItem y & ": " & Trim$(Npc(y).name)
        Next
    Next
    
    For i = 1 To MAX_ROOM_NPCS
        cmbNpc(i).ListIndex = Room.Npc(i)
    Next
    
    ' Room Item list
    For X = 1 To MAX_ROOM_ITEMS
        cmbItem(X).AddItem "No Item"
    Next
    
    For y = 1 To MAX_ITEMS
        For X = 1 To MAX_ROOM_ITEMS
            cmbItem(X).AddItem y & ": " & Trim$(Item(y).name)
        Next
    Next
    
    For i = 1 To MAX_ROOM_ITEMS
        cmbItem(i).ListIndex = Room.Item(i)
        If Room.Item(i) > 0 Then
            If Item(Room.Item(i)).Type = ITEM_TYPE_CURRENCY Then
                cmbItem(i).Width = 2895
                txtAmount(i).Enabled = True
                txtAmount(i).Text = Room.ItemVal(i)
            Else
                cmbItem(i).Width = 3975
                txtAmount(i).Enabled = False
                txtAmount(i).Text = 0
            End If
        Else
            cmbItem(i).Width = 3975
        End If
    Next
    
    lblRoom.Caption = "Current Room: " & GetPlayerRoom(MyIndex)
    
End Sub

Private Sub cmdCancel_Click()
    Call RoomEditorCancel
End Sub

Private Sub cmdSave_Click()

    Dim i     As Long

    Dim sTemp As Long
    
    With Room
        .name = Trim$(txtName.Text)
        .sDesc = Trim$(txtShortDesc.Text)
        .lDesc = Trim$(txtLongDesc.Text)
        .eDesc = Trim$(txtExitDesc.Text)
        .North = Val(txtNorth.Text)
        .South = Val(txtSouth.Text)
        .West = Val(txtWest.Text)
        .East = Val(txtEast.Text)
        .Moral = cmbMoral.ListIndex
        .Music = lblMusic.Caption
        .BootRoom = Val(txtBootRoom.Text)
        .Shop = cmbShop.ListIndex
        
        For i = 1 To MAX_ROOM_NPCS

            If cmbNpc(i).ListIndex > 0 Then
                
                sTemp = InStr(1, Trim$(cmbNpc(i).Text), ":", vbTextCompare)
                
                If Len(Trim$(cmbNpc(i).Text)) = sTemp Then
                    cmbNpc(i).ListIndex = 0
                End If
            End If

        Next
        
        For i = 1 To MAX_ROOM_NPCS
            .Npc(i) = cmbNpc(i).ListIndex
        Next
        
        For i = 1 To MAX_ROOM_ITEMS

            If cmbItem(i).ListIndex > 0 Then

                sTemp = InStr(1, Trim$(cmbItem(i).Text), ":", vbTextCompare)

                If Len(Trim$(cmbItem(i).Text)) = sTemp Then
                    cmbItem(i).ListIndex = 0
                End If
            End If

        Next

        For i = 1 To MAX_ROOM_ITEMS
            .Item(i) = cmbItem(i).ListIndex
            If IsNumeric(txtAmount(i).Text) Then
                If .Item(i) > 0 Then
                    If Item(.Item(i)).Type = ITEM_TYPE_CURRENCY Then
                        If CInt(txtAmount(i).Text) > 0 Then
                            .ItemVal(i) = CInt(txtAmount(i).Text)
                        End If
                    End If
                End If
            End If
        Next

    End With
    
    Call RoomEditorSend
End Sub
