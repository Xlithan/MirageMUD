Attribute VB_Name = "modConstants"
'************************************
'**    MADE WITH MIRAGESOURCE 5    **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' Main constants
Public Const GAME_NAME = "MirageMUD" ' Name of game
Public Const WEB_SITE = "http://ms.draignet.uk" ' Website

Public Const GAME_PORT = 7777 ' Run off What Port?
Public Const MAX_PLAYERS = 50 ' Max Players
Public Const MAX_ROOMS = 50 ' Max Rooms
Public Const MAX_ITEMS = 255 ' Max Items
Public Const MAX_SHOPS = 255 ' Max Shops
Public Const MAX_SPELLS = 255 ' Max Spells
Public Const MAX_NPCS = 255 ' Max NPCs

' path constants
Public Const ADMIN_LOG As String = "admin.log"
Public Const PLAYER_LOG As String = "player.log"

' Version constants
Public Const CLIENT_MAJOR As Byte = 1
Public Const CLIENT_MINOR As Byte = 0
Public Const CLIENT_REVISION As Byte = 0

' **********************************************************
' * The values below must match with the client's values *
' **********************************************************

' General Game-Useage constants
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_INV As Byte = 50
Public Const MAX_ROOM_ITEMS As Byte = 5
Public Const MAX_ROOM_NPCS As Byte = 5
Public Const MAX_TRADES As Byte = 8
Public Const MAX_LEVELS As Byte = 30

' Website
Public Const GAME_WEBSITE As String = "http://miragegaming.uk"

' text color constants
Public Const COLOR_BLACK                           As String = "[color=#000000]"
Public Const COLOR_BLUE                            As String = "[color=#0000FF]"
Public Const COLOR_GREEN                           As String = "[color=#00FF00]"
Public Const COLOR_CYAN                            As String = "[color=#00FFFF]"
Public Const COLOR_RED                             As String = "[color=#FF0000]"
Public Const COLOR_MAGENTA                         As String = "[color=#FF00FF]"
Public Const COLOR_BROWN                           As String = "[color=#964B00]"
Public Const COLOR_GREY                            As String = "[color=#808080]"
Public Const COLOR_DARKGREY                        As String = "[color=#A9A9A9]"
Public Const COLOR_BRIGHTBLUE                      As String = "[color=#8888ff]"
Public Const COLOR_BRIGHTGREEN                     As String = "[color=#66ff00]"
Public Const COLOR_BRIGHTCYAN                      As String = "[color=#8AFFFF]"
Public Const COLOR_BRIGHTRED                       As String = "[color=#EE4B2B]"
Public Const COLOR_PINK                            As String = "[color=#FFC0CB]"
Public Const COLOR_YELLOW                          As String = "[color=#FFFF00]"
Public Const COLOR_WHITE                           As String = "[color=#FFFFFF]"

Public Const COLOR_SAY                             As String = "[color=#FFFFFF]"
Public Const COLOR_GLOBAL                          As String = "[color=#66ff00]"
Public Const COLOR_BROADCAST                       As String = "[color=#FFFF00]"
Public Const COLOR_TELL                            As String = "[color=#8888ff]"
Public Const COLOR_EMOTE                           As String = "[color=#ff8000]"
Public Const COLOR_ADMIN                           As String = "[color=#8888ff]"
Public Const COLOR_HELP                            As String = "[color=#00c000]"
Public Const COLOR_WHO                             As String = "[color=#ff0000]"
Public Const COLOR_JOINLEFT                        As String = "[color=#a4a4a4]"
Public Const COLOR_NPC                             As String = "[color=#00c000]"
Public Const COLOR_ALERT                           As String = "[color=#c40000]"
Public Const COLOR_NEWROOM                         As String = "[color=#88c6ff]"

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Account constants
Public Const NAME_LENGTH As Byte = 20
Public Const MAX_CHARS As Byte = 3

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Room constants
Public Const MAX_ROOMX As Byte = 15
Public Const MAX_ROOMY As Byte = 11
Public Const ROOM_MORAL_NONE As Byte = 0
Public Const ROOM_MORAL_SAFE As Byte = 1
Public Const ROOM_MORAL_INN As Byte = 2
Public Const ROOM_MORAL_ARENA As Byte = 3

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_POTIONADDHP As Byte = 5
Public Const ITEM_TYPE_POTIONADDMP As Byte = 6
Public Const ITEM_TYPE_POTIONADDSP As Byte = 7
Public Const ITEM_TYPE_POTIONSUBHP As Byte = 8
Public Const ITEM_TYPE_POTIONSUBMP As Byte = 9
Public Const ITEM_TYPE_POTIONSUBSP As Byte = 10
Public Const ITEM_TYPE_CURRENCY As Byte = 11
Public Const ITEM_TYPE_SPELL As Byte = 12

' Direction constants
Public Const DIR_NORTH As Byte = 0
Public Const DIR_SOUTH As Byte = 1
Public Const DIR_WEST As Byte = 2
Public Const DIR_EAST As Byte = 3

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' Npc constants
Public Const Npc_BEHAVIOR_ATTACKONSIGHT As Byte = 0
Public Const Npc_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1
Public Const Npc_BEHAVIOR_FRIENDLY As Byte = 2
Public Const Npc_BEHAVIOR_SHOPKEEPER As Byte = 3
Public Const Npc_BEHAVIOR_GUARD As Byte = 4

' Spell constants
Public Const SPELL_TYPE_ADDHP As Byte = 0
Public Const SPELL_TYPE_ADDMP As Byte = 1
Public Const SPELL_TYPE_ADDSP As Byte = 2
Public Const SPELL_TYPE_SUBHP As Byte = 3
Public Const SPELL_TYPE_SUBMP As Byte = 4
Public Const SPELL_TYPE_SUBSP As Byte = 5
Public Const SPELL_TYPE_GIVEITEM As Byte = 6

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_Npc As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

' **********************************************
' Default starting location [Server Only]
Public Const START_ROOM = 1

