VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmMainGame 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MirageMUD"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12285
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMirage.frx":08CA
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   819
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   5550
      Left            =   2370
      TabIndex        =   34
      Top             =   795
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   9790
      _Version        =   393217
      BackColor       =   2235678
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMirage.frx":168EAE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstSpells 
      Appearance      =   0  'Flat
      BackColor       =   &H00221D1E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2130
      ItemData        =   "frmMirage.frx":168F2E
      Left            =   9930
      List            =   "frmMirage.frx":168F30
      TabIndex        =   18
      Top             =   4440
      Width           =   2280
   End
   Begin VB.PictureBox picInvSelected 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   11670
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   17
      Top             =   3570
      Width           =   480
   End
   Begin VB.ListBox lstInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00221D1E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2340
      ItemData        =   "frmMirage.frx":168F32
      Left            =   9945
      List            =   "frmMirage.frx":168F34
      TabIndex        =   14
      Top             =   1140
      Width           =   2280
   End
   Begin VB.PictureBox picXP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   6780
      Picture         =   "frmMirage.frx":168F36
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   12
      Top             =   225
      Width           =   4095
   End
   Begin VB.PictureBox picMP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   4680
      Picture         =   "frmMirage.frx":16AC4E
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   11
      Top             =   225
      Width           =   1395
   End
   Begin VB.PictureBox picStamina 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2595
      Picture         =   "frmMirage.frx":16B66A
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   10
      Top             =   225
      Width           =   1395
   End
   Begin VB.PictureBox picHP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   495
      Picture         =   "frmMirage.frx":16C086
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   9
      Top             =   225
      Width           =   1395
   End
   Begin VB.ListBox lstItems 
      Appearance      =   0  'Flat
      BackColor       =   &H00221D1E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1710
      ItemData        =   "frmMirage.frx":16CAA2
      Left            =   120
      List            =   "frmMirage.frx":16CAA9
      TabIndex        =   8
      Top             =   5040
      Width           =   2055
   End
   Begin VB.ListBox lstNPCs 
      Appearance      =   0  'Flat
      BackColor       =   &H00221D1E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      ItemData        =   "frmMirage.frx":16CAB7
      Left            =   120
      List            =   "frmMirage.frx":16CABE
      TabIndex        =   7
      Top             =   3060
      Width           =   2055
   End
   Begin VB.ListBox lstPlayers 
      Appearance      =   0  'Flat
      BackColor       =   &H00221D1E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      ItemData        =   "frmMirage.frx":16CACB
      Left            =   120
      List            =   "frmMirage.frx":16CAD2
      TabIndex        =   6
      Top             =   1125
      Width           =   2055
   End
   Begin VB.PictureBox picTarget 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   660
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   4
      Top             =   7320
      Width           =   990
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   15600
      Top             =   7500
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMyChat 
      Appearance      =   0  'Flat
      BackColor       =   &H00221D1E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2400
      TabIndex        =   0
      Top             =   6690
      Width           =   7350
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   7530
      TabIndex        =   36
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label lblRoomNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8235
      TabIndex        =   35
      Top             =   8640
      Width           =   240
   End
   Begin VB.Label lblBlockChance 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006E5551&
      Height          =   240
      Left            =   6720
      TabIndex        =   33
      Top             =   8160
      Width           =   240
   End
   Begin VB.Label lblCritHit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006E5551&
      Height          =   240
      Left            =   6720
      TabIndex        =   32
      Top             =   7680
      Width           =   240
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006E5551&
      Height          =   240
      Left            =   3945
      TabIndex        =   31
      Top             =   8520
      Width           =   240
   End
   Begin VB.Label lblSpeed 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006E5551&
      Height          =   240
      Left            =   3945
      TabIndex        =   30
      Top             =   8040
      Width           =   240
   End
   Begin VB.Label lblMagi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006E5551&
      Height          =   240
      Left            =   3945
      TabIndex        =   29
      Top             =   7800
      Width           =   240
   End
   Begin VB.Label lblDef 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006E5551&
      Height          =   240
      Left            =   3945
      TabIndex        =   28
      Top             =   7560
      Width           =   240
   End
   Begin VB.Label lblStr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006E5551&
      Height          =   240
      Left            =   3945
      TabIndex        =   27
      Top             =   7320
      Width           =   240
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Block Chance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2321&
      Height          =   240
      Left            =   5640
      TabIndex        =   26
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Critical Hit Chance:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2321&
      Height          =   240
      Left            =   5160
      TabIndex        =   25
      Top             =   7440
      Width           =   1830
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2321&
      Height          =   240
      Left            =   3270
      TabIndex        =   24
      Top             =   8520
      Width           =   585
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2321&
      Height          =   240
      Left            =   3180
      TabIndex        =   23
      Top             =   8040
      Width           =   675
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magi:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2321&
      Height          =   240
      Left            =   3330
      TabIndex        =   22
      Top             =   7800
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defense:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2321&
      Height          =   240
      Left            =   2985
      TabIndex        =   21
      Top             =   7560
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Strength:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002F2321&
      Height          =   240
      Left            =   2910
      TabIndex        =   20
      Top             =   7320
      Width           =   945
   End
   Begin VB.Label lblCast 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cast"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   10635
      TabIndex        =   19
      Top             =   6705
      Width           =   855
   End
   Begin VB.Label lblUseItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   10230
      TabIndex        =   16
      Top             =   3555
      Width           =   975
   End
   Begin VB.Label lblDropItem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Drop Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   10230
      TabIndex        =   15
      Top             =   3795
      Width           =   975
   End
   Begin VB.Label lblXP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   8310
      TabIndex        =   13
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image picMinimiseButton 
      Height          =   450
      Left            =   11325
      Top             =   45
      Width           =   435
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   645
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image picInventoryButton 
      Height          =   750
      Left            =   6915
      Top             =   9810
      Width           =   825
   End
   Begin VB.Image picSpellsButton 
      Height          =   750
      Left            =   7995
      Top             =   9810
      Width           =   825
   End
   Begin VB.Image picStatsButton 
      Height          =   750
      Left            =   8985
      Top             =   9765
      Width           =   825
   End
   Begin VB.Image picTrainButton 
      Height          =   750
      Left            =   10155
      Top             =   7905
      Width           =   825
   End
   Begin VB.Image picTradeButton 
      Height          =   750
      Left            =   11205
      Top             =   7905
      Width           =   825
   End
   Begin VB.Image picQuitButton 
      Height          =   450
      Left            =   11790
      Top             =   45
      Width           =   435
   End
   Begin VB.Label lblTarget 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   285
      TabIndex        =   5
      Top             =   8580
      Width           =   1725
   End
   Begin VB.Label lblStamina 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2745
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblMP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   4845
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmMainGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' ************
' ** Events **
' ************

Private Sub Form_Load()
    'frmMainGame.Width = 10080
    txtChat.SelStart = &H7FFFFFFF
    txtChat.SelText = vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
    InGame = False
End Sub

Private Sub lstInv_DblClick()
    Call UseItem
End Sub

Private Sub lstItems_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               y As Single)
    ' Prevents the code from running when another box sets the ListIndex.
    'If lstItems.ListIndex = -1 Then Exit Sub

    Select Case Button

        Case 1

            ' If you click in an empty space it'll produce an error. This line fixes that.
            If lstItems.ListIndex = -1 Then Exit Sub
            
            ' Get the correct RoomItem number and request info from server
            ItemSel = ItemLst(lstItems.ListIndex + 1)
            Call ItemSearch(ItemSel)
            lblTarget.Caption = lstItems.Text
            
            BltTarget (2)
            
            ' Deselect the items from the other list boxes
            lstPlayers.ListIndex = -1
            lstNPCs.ListIndex = -1

        Case 2
            ' Right click event
    End Select
    
    ' Clear the other targets
    PlayerSel = -1
    NPCSel = -1
    
    SetFocusOnChat

End Sub

Private Sub lstItems_DblClick()
    Call CheckRoomGetItem
    
    ' Clear the other targets
    PlayerSel = -1
    NPCSel = -1
    
    SetFocusOnChat
End Sub

Private Sub lstNPCs_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              y As Single)
    ' Prevents the code from running when another box sets the ListIndex.
    'If lstNPCs.ListIndex = -1 Then Exit Sub

    Select Case Button

        Case 1

            ' If you click in an empty space it'll produce an error. This line fixes that.
            If lstNPCs.ListIndex = -1 Then Exit Sub
            
            ' Get the correct NPCItem number and request info from server
            NPCSel = NPCLst(lstNPCs.ListIndex + 1)
            Call NPCSearch(NPCSel)
            lblTarget.Caption = lstNPCs.Text
            
            BltTarget (1)
            
            ' Deselect the items from the other list boxes
            lstPlayers.ListIndex = -1
            lstItems.ListIndex = -1

        Case 2
            ' Right click casts spell
            Call CastSpell
    End Select
    
    ' Clear other targets
    PlayerSel = -1
    ItemSel = -1
    
    
    SetFocusOnChat

End Sub

Private Sub lstPlayers_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 y As Single)
    ' Prevents the code from running when another box sets the ListIndex.
    'If lstPlayers.ListIndex = -1 Then Exit Sub

    Select Case Button

        Case 1

            ' If you click in an empty space it'll produce an error. This line fixes that.
            If lstPlayers.ListIndex = -1 Then Exit Sub
            
            ' Set the PlayerSel variable to the ID of the player and search for their info via the name.
            PlayerSel = PlayerLst(lstPlayers.ListIndex + 1)
            Call PlayerSearch(lstPlayers.Text)
            lblTarget.Caption = lstPlayers.Text
            
            Call DrawTargetChar
            
            ' Deselect the items from the other list boxes
            lstNPCs.ListIndex = -1
            lstItems.ListIndex = -1
        
        Case 2
            ' Right click casts spell
            Call CastSpell
    End Select
    
    ' Clear the other targets
    NPCSel = -1
    ItemSel = -1
    
    SetFocusOnChat

End Sub

Private Sub picMinimiseButton_Click()
    Me.WindowState = vbMinimized
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Call HandleKeyPresses(KeyAscii)
    
    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        Case vbKeyF1
            'Call PlayerSearch(GetPlayerX(MyIndex), GetPlayerY(MyIndex))
        
        Case vbKeyF3
            Call CastSpell
        
        Case vbKeyF4
            Call UseItem
        
    End Select
    
End Sub

Private Sub txtMyChat_Change()
    MyText = txtMyChat
End Sub

Private Sub txtChat_GotFocus()
'    SetFocusOnChat
End Sub

' ***************
' ** Inventory **
' ***************

Private Sub lblUseItem_Click()
    Call UseItem
End Sub

Private Sub lblDropItem_Click()

    Dim InvNum As Long
    
    InvNum = frmMainGame.lstInv.ListIndex + 1
    
    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 Then
        If GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Show them the drop dialog
                frmDrop.Show vbModal
            Else
                Call SendDropItem(frmMainGame.lstInv.ListIndex + 1, 0)
                
                ' clear inventory graphic
                frmMainGame.picInvSelected.Cls
            End If
        End If
    End If

End Sub

Private Sub lstInv_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             y As Single)
    InventoryItemSelected = frmMainGame.lstInv.ListIndex + 1
    'lblInvSelected.Caption = "<slot " & InventoryItemSelected & ">"
    
    If GetPlayerInvItemNum(MyIndex, InventoryItemSelected) > 0 Then
        Call BltInventory(GetPlayerInvItemNum(MyIndex, InventoryItemSelected))
    Else
        frmMainGame.picInvSelected.Cls
    End If

End Sub

Private Sub lstInv_GotFocus()

    On Error Resume Next

    SetFocusOnChat
End Sub

' ************
' ** Spells **
' ************

Private Sub lblCast_Click()
    Call CastSpell
End Sub

Private Sub lstSpells_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                y As Single)
    SpellSelected = frmMainGame.lstSpells.ListIndex + 1
    'lblSpellSelected.Caption = "<slot " & SpellSelected & ">"
End Sub

Private Sub lstSpells_GotFocus()

    On Error Resume Next

    SetFocusOnChat
End Sub

' *****************
' ** GUI Buttons **
' *****************

Private Sub picSpellsButton_Click()
    Call UpdateSpells
End Sub

Private Sub picStatsButton_Click()

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CGetStats
    Call SendData(Buffer.ToArray())
End Sub

Private Sub picTrainButton_Click()
    frmTraining.Show vbModal
End Sub

Private Sub picTradeButton_Click()

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.PreAllocate 2
    Buffer.WriteInteger CTrade
    
    Call SendData(Buffer.ToArray())
End Sub

Private Sub picQuitButton_Click()

    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate 2
    Buffer.WriteInteger CQuit
    
    isLogging = True
    InGame = False
    Stop_Music
    Call SendData(Buffer.ToArray())
    Call DestroyTCP
    Call DestroyGame
End Sub

Private Sub txtMyChat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And txtMyChat = vbNullString Then
        Call AddText(COLOR_YELLOW & "<< " & COLOR_BRIGHTCYAN & Trim$(Room.name) & COLOR_YELLOW & " >>")
        Call AddText(COLOR_BRIGHTBLUE & Trim$(Room.lDesc))
        Call AddText(COLOR_BRIGHTBLUE & Trim$(Room.eDesc))
    End If
End Sub
