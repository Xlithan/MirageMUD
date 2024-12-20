VERSION 5.00
Begin VB.Form frmMainMenu 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "MirageSource 5"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMainMenu.frx":08CA
   ScaleHeight     =   299.003
   ScaleMode       =   0  'User
   ScaleWidth      =   500
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox mnuNewCharacter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   0
      Picture         =   "frmMainMenu.frx":6E6DE
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.OptionButton optFemale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   2040
         Width           =   855
      End
      Begin VB.PictureBox picPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   6360
         Picture         =   "frmMainMenu.frx":DC4F2
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   46
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1080
         Width           =   720
      End
      Begin VB.OptionButton optMale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2040
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.ComboBox cmbClass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         ItemData        =   "frmMainMenu.frx":11E3E6
         Left            =   4080
         List            =   "frmMainMenu.frx":11E3E8
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtNewCharName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4080
         MaxLength       =   20
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblAvatarPrev 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prev"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6240
         TabIndex        =   80
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lblAvatarNext 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6720
         TabIndex        =   79
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lblNewChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   6360
         TabIndex        =   78
         Top             =   3720
         Width           =   915
      End
      Begin VB.Label lblNewChar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   6360
         TabIndex        =   77
         Top             =   4080
         Width           =   915
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SPEED :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   4710
         TabIndex        =   76
         Top             =   3960
         UseMnemonic     =   0   'False
         Width           =   705
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MAGI :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   4800
         TabIndex        =   75
         Top             =   3600
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   5040
         TabIndex        =   74
         Top             =   3240
         UseMnemonic     =   0   'False
         Width           =   375
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "STR :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   4935
         TabIndex        =   73
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   480
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SP :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   3615
         TabIndex        =   72
         Top             =   3600
         UseMnemonic     =   0   'False
         Width           =   360
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MP :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   3570
         TabIndex        =   71
         Top             =   3240
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HP :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   210
         Left            =   3600
         TabIndex        =   70
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   375
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3240
         TabIndex        =   69
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   750
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3405
         TabIndex        =   68
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3345
         TabIndex        =   67
         Top             =   1080
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Character"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3960
         TabIndex        =   65
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblMAGI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   19
         Top             =   3600
         Width           =   375
      End
      Begin VB.Label lblDEF 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   18
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lblSP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   17
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label lblSPEED 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   16
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label lblMP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label lblSTR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5520
         TabIndex        =   14
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label lblHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   2880
         Width           =   375
      End
   End
   Begin VB.PictureBox mnuCredits 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   0
      Picture         =   "frmMainMenu.frx":11E3EA
      ScaleHeight     =   4500
      ScaleWidth      =   7500
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "William"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   59
         Top             =   3480
         Width           =   4095
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Solace"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   58
         Top             =   3240
         Width           =   4095
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Acer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   57
         Top             =   3000
         Width           =   4095
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Xlithan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   56
         Top             =   2760
         Width           =   4095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Special Mentions"
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
         Height          =   375
         Left            =   4080
         TabIndex        =   55
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Mirage Online originally developed by Consty. MirageSource originally released by Shannara. MirageMUD developed by Xlithan."
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   3240
         TabIndex        =   54
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "MirageSource has and always will be a collective effort by the MirageSource community since it's conception."
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   3240
         TabIndex        =   53
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Credits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4080
         TabIndex        =   52
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   51
         Top             =   4080
         Width           =   915
      End
   End
   Begin VB.PictureBox mnuNewAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   0
      Picture         =   "frmMainMenu.frx":18C1FE
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox txtNewAcctPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   4440
         MaxLength       =   20
         PasswordChar    =   "�"
         TabIndex        =   2
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox txtNewAcctName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4440
         MaxLength       =   20
         TabIndex        =   1
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   50
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   49
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a account name and password.  You can name yourself whatever you want, we have no restrictions on names."
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   3240
         TabIndex        =   34
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblCancel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6360
         TabIndex        =   33
         Top             =   4080
         Width           =   915
      End
      Begin VB.Label lblConnect 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4800
         TabIndex        =   32
         Top             =   3120
         UseMnemonic     =   0   'False
         Width           =   915
      End
      Begin VB.Label lblMainMenuT 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Account"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   31
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.PictureBox mnuChars 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   0
      Picture         =   "frmMainMenu.frx":1FA012
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.ListBox lstChars 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1455
         ItemData        =   "frmMainMenu.frx":267E26
         Left            =   3240
         List            =   "frmMainMenu.frx":267E28
         TabIndex        =   35
         Top             =   600
         Width           =   3855
      End
      Begin VB.PictureBox picSelChar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   4920
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   46
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label lblChars 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Use Character"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4320
         TabIndex        =   40
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblChars 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "New Character"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   4320
         TabIndex        =   39
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label lblChars 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Character"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   4320
         TabIndex        =   38
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblChars 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   6480
         TabIndex        =   37
         Top             =   4080
         UseMnemonic     =   0   'False
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Characters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4200
         TabIndex        =   36
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   1935
      End
   End
   Begin VB.PictureBox mnuLogin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   0
      Picture         =   "frmMainMenu.frx":267E2A
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox txtLoginName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4320
         MaxLength       =   20
         TabIndex        =   41
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtLoginPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   4320
         MaxLength       =   20
         PasswordChar    =   "�"
         TabIndex        =   5
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CheckBox chkLogin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   4320
         TabIndex        =   4
         Top             =   2400
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remember Me"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4590
         TabIndex        =   48
         Top             =   2400
         UseMnemonic     =   0   'False
         Width           =   1020
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Password : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   47
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a account name and password.  "
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3360
         TabIndex        =   46
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Name : "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   45
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3960
         TabIndex        =   44
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   2535
      End
      Begin VB.Label lblLogin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   6360
         TabIndex        =   43
         Top             =   4080
         Width           =   915
      End
      Begin VB.Label lblLogin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4800
         TabIndex        =   42
         Top             =   3000
         UseMnemonic     =   0   'False
         Width           =   915
      End
   End
   Begin VB.PictureBox mnuIPConfig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   0
      Picture         =   "frmMainMenu.frx":2D5C3E
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox txtIP 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   22
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         IMEMode         =   3  'DISABLE
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   21
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblIPConfig 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   6360
         TabIndex        =   64
         Top             =   4080
         Width           =   915
      End
      Begin VB.Label lblIPConfig 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4800
         TabIndex        =   63
         Top             =   2760
         UseMnemonic     =   0   'False
         Width           =   915
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IP Config"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3960
         TabIndex        =   62
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IP :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   61
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Port :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   60
         Top             =   1680
         Width           =   615
      End
   End
   Begin VB.PictureBox mnuGameOptions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   0
      Picture         =   "frmMainMenu.frx":343A52
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
      Begin VB.TextBox txtFonts 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   91
         Top             =   2040
         Width           =   2535
      End
      Begin VB.ComboBox cmbFonts 
         Height          =   360
         Left            =   4200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   2520
         Width           =   2535
      End
      Begin VB.CheckBox chkSound 
         BackColor       =   &H001F1F1F&
         Height          =   240
         Left            =   4200
         TabIndex        =   89
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkMusic 
         BackColor       =   &H001F1F1F&
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4200
         TabIndex        =   88
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3480
         TabIndex        =   87
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sound :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3300
         TabIndex        =   86
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Music :"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3360
         TabIndex        =   85
         Top             =   1320
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4350
         TabIndex        =   84
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblGameOpt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4800
         TabIndex        =   83
         Top             =   3360
         UseMnemonic     =   0   'False
         Width           =   915
      End
      Begin VB.Label lblGameOpt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   6360
         TabIndex        =   82
         Top             =   4080
         Width           =   915
      End
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IP Config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   4320
      TabIndex        =   30
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblMainMenuT 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   29
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credits"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   6480
      TabIndex        =   28
      Top             =   4080
      Width           =   795
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   4320
      TabIndex        =   27
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   4320
      TabIndex        =   26
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   4320
      TabIndex        =   25
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblMainMenu 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   4320
      TabIndex        =   24
      Top             =   2280
      UseMnemonic     =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

Private Sub cmbClass_Click()
    Dim n As Long
    
    n = cmbClass.ListIndex + 1
        
    lblHP.Caption = CStr(Class(n).Vital(Vitals.HP))
    lblMP.Caption = CStr(Class(n).Vital(Vitals.MP))
    lblSP.Caption = CStr(Class(n).Vital(Vitals.SP))
    
    lblSTR.Caption = CStr(Class(n).Stat(Stats.Strength))
    lblDEF.Caption = CStr(Class(n).Stat(Stats.Defense))
    lblSpeed.Caption = CStr(Class(n).Stat(Stats.speed))
    lblMAGI.Caption = CStr(Class(n).Stat(Stats.Magic))
End Sub

Private Sub cmbFonts_Click()
    txtFonts.Text = cmbFonts.Text
    txtFonts.Font = cmbFonts.Text
End Sub

Private Sub Form_Load()
    Dim filename As String
    Dim i As Long
    
    For i = 0 To Screen.FontCount - 1
        cmbFonts.AddItem Screen.Fonts(i)
    Next
    
    'Me.Caption = GAME_NAME
    
    Me.Show
    DoEvents
    
    'If App.PrevInstance = True Then
    '    MsgBox "Another MirageMUD Client is already running! Please run only one client at a time!", Error
    'End If
    
    filename = App.Path & DATA_PATH & "config.dat"
    txtIP.Text = Trim$(GameData.IP)       ' GetVar(FileName, "IPCONFIG", "IP")
    txtPort.Text = Trim$(GameData.Port)   ' GetVar(FileName, "IPCONFIG", "PORT")
    
    ' Used for Credits
    Dim result As Long
    
End Sub

'**********************************
'* Handles Character Menu Buttons *
'**********************************
Private Sub lblChars_Click(Index As Integer)

    Dim Value As Long
    
    Call Play_Sound("menu.wav")
    
    Select Case Index
        
        Case 0
            Call MenuState(MENU_STATE_USECHAR)
        
            Exit Sub

        Case 1
            Call MenuState(MENU_STATE_NEWCHAR)
        
            Exit Sub

        Case 2
            Value = MsgBox("Are you sure you wish to delete this character?", vbYesNo, GAME_NAME)

            If Value = vbYes Then
                Call MenuState(MENU_STATE_DELCHAR)
            End If
        
            Exit Sub

        Case 3
            Call DestroyTCP
            mnuLogin.Visible = True
            mnuChars.Visible = False
        
            Exit Sub

    End Select

End Sub

'**************************
'* Handles Credits Button *
'**************************
Private Sub lblCredits_Click()
    Call Play_Sound("menu.wav")
    mnuCredits.Visible = False
End Sub

'************************************
'* Handles IP Configuration Buttons *
'************************************
Private Sub lblIPConfig_Click(Index As Integer)

    Dim IP, Port As String

    Dim filename As String

    Dim fErr     As Integer

    Dim Texto    As String
    
    Call Play_Sound("menu.wav")
    
    Select Case Index
        
        Case 0
            IP = Trim$(txtIP.Text)
            Port = Val(Trim$(txtPort.Text))
            filename = App.Path & DATA_PATH & "config.dat"
        
            fErr = 0

            If fErr = 0 And Len(Trim$(IP)) = 0 Then
                fErr = 1
                Call MsgBox("Inform a correct IP.", vbCritical, GAME_NAME)

                Exit Sub

            End If

            If fErr = 0 And Port <= 0 Then
                fErr = 1
                Call MsgBox("Inform a correct Port.", vbCritical, GAME_NAME)

                Exit Sub

            End If

            If fErr = 0 Then
                GameData.IP = txtIP.Text
                GameData.Port = txtPort.Text

                Dim F As Long

                F = FreeFile
                Open filename For Binary As #F
                Put #F, , GameData
                Close #F
            End If

            mnuIPConfig.Visible = False
            Call DestroyTCP
            Call TcpInit
        
            Exit Sub

        Case 1
            mnuIPConfig.Visible = False
        
            Exit Sub

    End Select
    
End Sub

'**************************************
'* Handles Game Configuration Buttons *
'**************************************
Private Sub lblGameOpt_Click(Index As Integer)
    
    Dim filename As String
    
    Call Play_Sound("menu.wav")
    
    Select Case Index
        
        Case 0
            filename = App.Path & DATA_PATH & "config.dat"
            GameData.Music = chkMusic.Value
            GameData.Sound = chkSound.Value
            GameData.Font = Trim$(txtFonts.Text)
            
            ' Unfortunately a lot of the fancy fonts don't seem to work.
            frmMainGame.txtChat.Font = Trim$(GameData.Font)
            
            If GameData.Music = 0 Then
                Stop_Music
            ElseIf GameData.Music = 1 Then
                Play_Music "Menu.mp3"
            End If

            Dim F As Long

            F = FreeFile
            Open filename For Binary As #F
            Put #F, , GameData
            Close #F

            mnuGameOptions.Visible = False
            Exit Sub

        Case 1
            mnuGameOptions.Visible = False
        
            Exit Sub

    End Select
    
End Sub

'*************************
'* Handles Login Buttons *
'*************************
Private Sub lblLogin_Click(Index As Integer)

    Dim filename As String
    
    Call Play_Sound("menu.wav")
    
    Select Case Index
        
        Case 0
            filename = App.Path & DATA_PATH & "config.dat"
        
            If chkLogin.Value Then
                GameData.SaveLogin = 1
                GameData.Username = txtLoginName.Text
                GameData.Password = Trim$(txtLoginPassword.Text)
            Else
                GameData.SaveLogin = 0
                GameData.Username = vbNullString
                GameData.Password = vbNullString
            End If
        
            Dim F As Long

            F = FreeFile
            Open filename For Binary As #F
            Put #F, , GameData
            Close #F
        
            Call LoginConnect
        
            Exit Sub

        Case 1
            mnuLogin.Visible = False
        
            Exit Sub

    End Select
    
End Sub

'*****************************
'* Handles Main Menu Buttons *
'*****************************
Private Sub lblMainMenu_Click(Index As Integer)

    Call Play_Sound("menu.wav")
    Select Case Index
        
        Case 0
            mnuNewAccount.Visible = True
            txtNewAcctName.SetFocus
        
            Exit Sub

        Case 1

            If GameData.SaveLogin = 1 Then
                chkLogin.Value = 1
                txtLoginName.Text = Trim$(GameData.Username)
                txtLoginPassword.Text = Trim$(GameData.Password)
            End If
        
            mnuLogin.Visible = True
            txtLoginName.SetFocus
        
            Exit Sub

        Case 2
            txtIP.Text = Trim$(GameData.IP)
            txtPort.Text = Trim$(GameData.Port)
            mnuIPConfig.Visible = True
        
            txtIP.SetFocus
        
            Exit Sub

        Case 3
            chkMusic.Value = GameData.Music
            chkSound.Value = GameData.Sound
            txtFonts.Text = Trim$(GameData.Font)
            mnuGameOptions.Visible = True
        
            chkMusic.SetFocus
        
            Exit Sub
        
            Exit Sub

        Case 4
            Call DestroyGame
        
            Exit Sub

        Case 5
            mnuCredits.Visible = True
        
            Exit Sub

    End Select
    
End Sub

'************************************
'* Handles New Account Menu Buttons *
'************************************
Private Sub lblCancel_Click()
    Call Play_Sound("menu.wav")
    
    mnuNewAccount.Visible = False

    Exit Sub

End Sub

Private Sub lblConnect_Click()
    Call Play_Sound("menu.wav")
    
    Call NewAccountConnect

    Exit Sub

End Sub

'**************************************
'* Handles New Character Menu Buttons *
'**************************************
Private Sub lblNewChar_Click(Index As Integer)
        
    Call Play_Sound("menu.wav")
    Select Case Index
        
        Case 0
            Call AddCharClick
        
            Exit Sub

        Case 1
            mnuChars.Visible = True
            mnuNewCharacter.Visible = False
        
            Exit Sub

    End Select
    
End Sub

'**********************************
'* Handles Moving the Form Around *
'**********************************
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub lblAvatarNext_Click()
    Call Play_Sound("menu.wav")
    If NewCharAvatar < NumAvatars Then
        NewCharAvatar = NewCharAvatar + 1
        picPic.Picture = LoadPicture(App.Path & AVATARS_PATH & NewCharAvatar & ".bmp")
    End If
    
    picPic.ScaleMode = 3
    picPic.AutoRedraw = True
    picPic.PaintPicture picPic.Picture, _
    0, 0, picPic.ScaleWidth, picPic.ScaleHeight, _
    0, 0, picPic.Picture.Width / 26.46, _
    picPic.Picture.Height / 26.46

    picPic.Picture = picPic.Image
End Sub

Private Sub lblAvatarPrev_Click()
    Call Play_Sound("menu.wav")
    If NewCharAvatar > 1 Then
        NewCharAvatar = NewCharAvatar - 1
        picPic.Picture = LoadPicture(App.Path & AVATARS_PATH & NewCharAvatar & ".bmp")
    End If
    
    picPic.ScaleMode = 3
    picPic.AutoRedraw = True
    picPic.PaintPicture picPic.Picture, _
    0, 0, picPic.ScaleWidth, picPic.ScaleHeight, _
    0, 0, picPic.Picture.Width / 26.46, _
    picPic.Picture.Height / 26.46

    picPic.Picture = picPic.Image
End Sub

Private Sub lstChars_Click()
    DrawSelChar lstChars.ListIndex + 1
End Sub

Private Sub mnuChars_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuCredits_MouseDown(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuIPConfig_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuLogin_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuNewAccount_MouseDown(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub mnuNewCharacter_MouseDown(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      y As Single)
    Call MoveForm(frmMainMenu)
End Sub

Private Sub optFemale_Click()
    Call Play_Sound("menu.wav")
End Sub

Private Sub optMale_Click()
    Call Play_Sound("menu.wav")
End Sub
