Attribute VB_Name = "Types"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' Public data structures
Public Room As RoomRec

Public Class() As ClassRec

Public Player()  As PlayerRec

Public Item()    As ItemRec

Public Npc()     As NpcRec

Public Shop()    As ShopRec

Public Spell()   As SpellRec

Public RoomItem() As RoomItemRec

Public RoomNpc()  As RoomNpcRec

Public GameData  As DataRec

Type DataRec

    IP As String * NAME_LENGTH
    Port As Integer
    SaveLogin As Byte
    Username As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Music As Byte
    sound As Byte
    Font As String * NAME_LENGTH

End Type

Public Type PlayerInvRec

    num As Byte
    Value As Long
    Dur As Integer

End Type

Public Type PlayerRec

    ' General
    name As String * NAME_LENGTH
    Class As Byte
    Avatar As Integer
    Level As Byte
    Exp As Long
    Access As Byte
    PK As Byte
    
    Guild As String
    GuildAccess As Long
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Byte
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Byte
    
    ' Position
    Room As Integer
    Dir As Byte
    
    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    MaxStamina As Long
    Attacking As Byte
    AttackTimer As Long
    RoomGetTimer As Long
    CastedSpell As Byte

End Type

Public Type RoomRec

    name As String * NAME_LENGTH
    sDesc As String * 256
    lDesc As String * 512
    eDesc As String * 128
    Revision As Long
    Moral As Byte
    Up As Integer
    Down As Integer
    North As Integer
    East As Integer
    South As Integer
    West As Integer
    Music As String * 32
    BootRoom As Integer
    Shop As Byte
    Npc(1 To MAX_ROOM_NPCS) As Byte
    Item(1 To MAX_ROOM_ITEMS) As Byte
    ItemVal(1 To MAX_ROOM_ITEMS) As Long

End Type

Public Type ClassRec

    name As String * NAME_LENGTH
    Avatar As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
    
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long

End Type
    
Public Type ItemRec

    name As String * NAME_LENGTH
    
    Pic As Integer

    Type As Byte

    Data1 As Integer
    Data2 As Integer
    Data3 As Integer

End Type

Public Type RoomItemRec

    num As Byte
    Value As Long
    Dur As Integer

End Type

Public Type NpcRec

    name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Avatar As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte

End Type

Public Type RoomNpcRec

    num As Byte
    
    Target As Byte
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    Room As Integer
    
    ' Client use only
    Attacking As Byte
    AttackTimer As Long

End Type

Public Type TradeItemRec

    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long

End Type

Public Type ShopRec

    name As String * NAME_LENGTH
    JoinSay As String * 255
    LeaveSay As String * 255
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec

End Type

Public Type SpellRec

    name As String * NAME_LENGTH
    Pic As Integer
    MPReq As Integer
    ClassReq As Byte
    LevelReq As Byte

    Type As Byte

    Data1 As Integer
    Data2 As Integer
    Data3 As Integer

End Type
            
