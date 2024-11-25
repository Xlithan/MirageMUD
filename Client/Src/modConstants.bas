Attribute VB_Name = "Constants"
'************************************
'**       MADE WITH MIRAGEMUD      **
'#       Maintained by Xlithan     #'
'************************************
Option Explicit

' API Declares
Public Declare Sub CopyMemory _
               Lib "kernel32.dll" _
               Alias "RtlMoveMemory" (Destination As Any, _
                                      Source As Any, _
                                      ByVal length As Long)

Public Declare Function CallWindowProc _
               Lib "user32" _
               Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                        ByVal hwnd As Long, _
                                        ByRef msg() As Byte, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long

' RichText Transparency Declares
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE       As Long = (-20)

Public Const WS_EX_TRANSPARENT As Long = &H20&

' Move Form Declares
Public Type TextSize

    Width As Long
    Height As Long

End Type

Public Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long
'Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As TextSize) As Long

' Winsock globals
Public GAME_IP                               As String

' Varriables for Moving Forms
Public Const WM_NCLBUTTONDOWN                As Long = &HA1

' path constants
Public Const DATA_PATH                       As String = "\data\"

Public Const SOUND_PATH                      As String = "\data\sfx\"

Public Const MUSIC_PATH                      As String = "\data\music\"
Public Const SFX_PATH                        As String = "\data\sfx\"

' Log Path and variables
Public Const LOG_DEBUG                       As String = "debug.txt"

Public Const LOG_PATH                        As String = "\data\Logs\"

' Room Path and variables
Public Const ROOM_PATH                        As String = "\data\rooms\"

Public Const ROOM_EXT                         As String = ".dat"

' Gfx Path and variables
Public Const GFX_PATH                        As String = "\data\gfx\"
Public Const AVATARS_PATH                    As String = "\data\gfx\avatars\Players\"
Public Const NPCS_PATH                       As String = "\data\gfx\avatars\NPCs\"
Public Const ITEMS_PATH                      As String = "\data\gfx\items\"
Public Const SPELLS_PATH                     As String = "\data\gfx\spells\"

Public Const GFX_EXT                         As String = ".bmp"

' Menu states
Public Const MENU_STATE_NEWACCOUNT           As Byte = 0

Public Const MENU_STATE_DELACCOUNT           As Byte = 1

Public Const MENU_STATE_LOGIN                As Byte = 2

Public Const MENU_STATE_GETCHARS             As Byte = 3

Public Const MENU_STATE_NEWCHAR              As Byte = 4

Public Const MENU_STATE_ADDCHAR              As Byte = 5

Public Const MENU_STATE_DELCHAR              As Byte = 6

Public Const MENU_STATE_USECHAR              As Byte = 7

Public Const MENU_STATE_INIT                 As Byte = 8

' Avatar, item, spell size constants
Public Const SIZE_X                          As Integer = 32

Public Const SIZE_Y                          As Integer = 32

' ********************************************************
' * The values below must match with the server's values *
' ********************************************************

' General constants
Public GAME_NAME                             As String

Public GAME_PORT                             As Integer

Public MAX_PLAYERS                           As Integer

Public MAX_ITEMS                             As Integer

Public MAX_NPCS                              As Integer

Public MAX_SHOPS                             As Integer

Public MAX_SPELLS                            As Integer

Public Const MAX_PLAYER_SPELLS               As Byte = 20

Public Const MAX_INV                         As Byte = 50

Public Const MAX_ROOM_ITEMS                   As Byte = 5

Public Const MAX_ROOM_NPCS                    As Byte = 5

Public Const MAX_TRADES                      As Byte = 8

Public Const MAX_LEVELS                      As Byte = 30

' Website
Public GAME_WEBSITE                          As String

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
Public Const COLOR_GLOBAL                          As String = "[color=#000000]"
Public Const COLOR_BROADCAST                       As String = "[color=#FFFF00]"
Public Const COLOR_TELL                            As String = "[color=#8888ff]"
Public Const COLOR_EMOTE                           As String = "[color=#ff8000]"
Public Const COLOR_ADMIN                           As String = "[color=#8888ff]"
Public Const COLOR_HELP                            As String = "[color=#00c000]"
Public Const COLOR_WHO                             As String = "[color=#ff0000]"
Public Const COLOR_JOINLEFT                        As String = "[color=#a4a4a4]"
Public Const COLOR_NPC                             As String = "[color=#00c000]"
Public Const COLOR_ALERT                           As String = "[color=#c40000]"
Public Const COLOR_NEWROOM                          As String = "[color=#88c6ff]"

' Boolean constants
Public Const NO                              As Byte = 0

Public Const YES                             As Byte = 1

' Account constants
Public Const NAME_LENGTH                     As Byte = 20

Public Const MAX_CHARS                       As Byte = 3

' Sex constants
Public Const SEX_MALE                        As Byte = 0

Public Const SEX_FEMALE                      As Byte = 1

' Room constants
Public MAX_ROOMS                              As Long

Public Const ROOM_MORAL_NONE                  As Byte = 0

Public Const ROOM_MORAL_SAFE                  As Byte = 1

Public Const ROOM_MORAL_INN                   As Byte = 2

Public Const ROOM_MORAL_ARENA                 As Byte = 3

' Item constants
Public Const ITEM_TYPE_NONE                  As Byte = 0

Public Const ITEM_TYPE_WEAPON                As Byte = 1

Public Const ITEM_TYPE_ARMOR                 As Byte = 2

Public Const ITEM_TYPE_HELMET                As Byte = 3

Public Const ITEM_TYPE_SHIELD                As Byte = 4

Public Const ITEM_TYPE_POTIONADDHP           As Byte = 5

Public Const ITEM_TYPE_POTIONADDMP           As Byte = 6

Public Const ITEM_TYPE_POTIONADDSP           As Byte = 7

Public Const ITEM_TYPE_POTIONSUBHP           As Byte = 8

Public Const ITEM_TYPE_POTIONSUBMP           As Byte = 9

Public Const ITEM_TYPE_POTIONSUBSP           As Byte = 10

Public Const ITEM_TYPE_CURRENCY              As Byte = 11

Public Const ITEM_TYPE_SPELL                 As Byte = 12

' Direction constants
Public Const DIR_NORTH                       As Byte = 0

Public Const DIR_SOUTH                       As Byte = 1

Public Const DIR_WEST                        As Byte = 2

Public Const DIR_EAST                        As Byte = 3

' Admin constants
Public Const ADMIN_MONITOR                   As Byte = 1

Public Const ADMIN_MAPPER                    As Byte = 2

Public Const ADMIN_DEVELOPER                 As Byte = 3

Public Const ADMIN_CREATOR                   As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOR_ATTACKONSIGHT      As Byte = 0

Public Const NPC_BEHAVIOR_ATTACKWHENATTACKED As Byte = 1

Public Const NPC_BEHAVIOR_FRIENDLY           As Byte = 2

Public Const NPC_BEHAVIOR_SHOPKEEPER         As Byte = 3

Public Const NPC_BEHAVIOR_GUARD              As Byte = 4

' Spell constants
Public Const SPELL_TYPE_ADDHP                As Byte = 0

Public Const SPELL_TYPE_ADDMP                As Byte = 1

Public Const SPELL_TYPE_ADDSP                As Byte = 2

Public Const SPELL_TYPE_SUBHP                As Byte = 3

Public Const SPELL_TYPE_SUBMP                As Byte = 4

Public Const SPELL_TYPE_SUBSP                As Byte = 5

Public Const SPELL_TYPE_GIVEITEM             As Byte = 6

' Game editor constants
Public Const EDITOR_NONE                     As Byte = 0

Public Const EDITOR_ITEM                     As Byte = 1

Public Const EDITOR_NPC                      As Byte = 2

Public Const EDITOR_SPELL                    As Byte = 3

Public Const EDITOR_SHOP                     As Byte = 4

Public Const EDITOR_ROOM                      As Byte = 5

' Target type constants
Public Const TARGET_TYPE_NONE                As Byte = 0

Public Const TARGET_TYPE_PLAYER              As Byte = 1

Public Const TARGET_TYPE_NPC                 As Byte = 2

