Attribute VB_Name = "Globals"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' Player variables
Public MyIndex                              As Long ' Index of actual player

Public PlayerInv(1 To MAX_INV)              As PlayerInvRec   ' Inventory

Public PlayerSpells(1 To MAX_PLAYER_SPELLS) As Byte

Public InventoryItemSelected                As Integer

Public SpellSelected                        As Integer

Public CharAvatars()                        As Long

Public NewCharAvatar                        As Integer

' Stops movement when updating a Room
Public CanMoveNow                           As Boolean

'Some DX8 Helpers

Public NumAvatars                           As Long

Public NumNPCAvatars                        As Long

Public numitems                             As Long

Public NumSpells                            As Long

Public AvatarCount                          As Long


Public ItemCount                            As Long


Public SpellCount                           As Long

' Entity Selection
Public PlayerSel                            As Long

Public NPCSel                               As Long

Public ItemSel                              As Long

Public PlayerLst()                          As Long

Public NPCLst(1 To MAX_ROOM_NPCS)            As Long

Public ItemLst(1 To MAX_ROOM_ITEMS)          As Long

' Debug mode
Public DebugMode                            As Boolean

Public SyncX                                As Long 'used to check sync

Public SyncY                                As Long

Public SyncRoom                              As Long

Public SentSync                             As Boolean

' Controls main gameloop
Public InGame                               As Boolean

Public isLogging                            As Boolean

' Used for improved looping
Public High_Index                           As Integer

Public High_Npc_Index                       As Integer

Public PlayersInRoomHighIndex                As Long

Public PlayersInRoom()                       As Long

' Used for dragging Picture Boxes
Public SOffsetX                             As Integer

Public SOffsetY                             As Integer

Public vbQuote                              As String

' Used to freeze controls when getting a new Room
Public GettingRoom                           As Boolean

Public GameFPS                              As Long ' frames per second rendered

' Maximum classes
Public Max_Classes                          As Byte

' Used to check if text needs to be drawn
Public BFPS                                 As Boolean ' FPS

Public BLoc                                 As Boolean ' Room, player, and mouse location
