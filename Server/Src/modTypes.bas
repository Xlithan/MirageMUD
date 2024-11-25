Attribute VB_Name = "modTypes"
'************************************
'**    MADE WITH MIRAGESOURCE 5    **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' Public data structures
Public Room() As RoomRec
Public RoomCache() As Cache
Public PlayersInRoom() As Long
Public Player() As AccountRec
Public TempPlayer() As TempPlayerRec
Public Class() As ClassRec
Public Item() As ItemRec
Public Npc() As NpcRec
Public RoomItem() As RoomItemRec
Public RoomNpc() As RoomNpcRec
Public Shop() As ShopRec
Public Spell() As SpellRec

Public Type PlayerInvRec
    Num As Byte
    Value As Long
    Dur As Integer
End Type

Public Type PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
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
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    
    ' Position
    Room As Integer
    Dir As Byte
End Type
Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    
    ' Characters
    ' 0 is used to prevent an RTE9 when accessing a cleared account
    Char(0 To MAX_CHARS) As PlayerRec
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    Buffer As clsBuffer
    CharNum As Byte
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    PartyPlayer As Long
    InParty As Byte
    TargetType As Byte
    Target As Byte
    CastedSpell As Byte
    PartyStarter As Byte
    GettingRoom As Byte
End Type

Public Type RoomRec
    Name As String * NAME_LENGTH
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
    Name As String * NAME_LENGTH
    
    Avatar As Integer
    
    Stat(1 To Stats.Stat_Count - 1) As Byte
End Type
    
Public Type ItemRec
    Name As String * NAME_LENGTH
    
    Pic As Integer
    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
End Type

Public Type RoomItemRec
    Num As Byte
    Value As Long
    Dur As Integer
    
    X As Byte
    y As Byte
End Type

Public Type NpcRec
    Name As String * NAME_LENGTH
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
    Num As Integer
    
    Target As Integer
    
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
End Type

Public Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Public Type ShopRec
    Name As String * NAME_LENGTH
    JoinSay As String * 255
    LeaveSay As String * 255
    FixesItems As Byte
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Public Type SpellRec
    Name As String * NAME_LENGTH
    Pic As Integer
    MPReq As Integer
    ClassReq As Byte
    LevelReq As Byte
    Type As Byte
        Data1 As Integer
        Data2 As Integer
        Data3 As Integer
End Type

Public Type Cache
    Cache() As Byte
End Type
