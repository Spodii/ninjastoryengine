Attribute VB_Name = "EngineGlobals"
'*********************************************************************************
'Required global variables, constants and types by all of the parts of the engine
'*********************************************************************************

Option Explicit

'Directions
Public Const NORTH As Byte = 1
Public Const NORTHEAST As Byte = 2
Public Const EAST As Byte = 3
Public Const SOUTHEAST As Byte = 4
Public Const SOUTH As Byte = 5
Public Const SOUTHWEST As Byte = 6
Public Const WEST As Byte = 7
Public Const NORTHWEST As Byte = 8

'Grid size
Public Const GRIDSIZE As Long = 32

'If game windows have to stay in the screen
Public Const STAYINSCREEN As Boolean = True

'Client screen resolution
Public Const ScreenWidth As Long = 1024         'Screen resolution width
Public Const ScreenHeight As Long = 768         'Screen resolution height

'Grh data (the stuff you write in the GrhRaw.txt)
Public Type tGrhData
    X As Integer        'Source X
    Y As Integer        'Source Y
    Width As Integer    'Source width
    Height As Integer   'Source height
    TextureID As Long   'Non-animated: The texture file number (ie 9 for 9.png)
    Speed As Single     'Animated: How fast the animation moves
    NumFrames As Long   'Animated: How many frames there are (UBound of Frames())
    Frames() As Long    'Animated: List of frames (by GrhIndex) in the animation
End Type

'Each individual instance of a graphic
Public Const ANIMTYPE_STATIONARY As Byte = 0
Public Const ANIMTYPE_LOOPONCE As Byte = 1
Public Const ANIMTYPE_LOOP As Byte = 2
Public Type tGrh
    AnimType As Byte        'Which animation type to use (see consts above)
    GrhIndex As Long        'Points to the GrhData index to be used
    Frame As Single         'Current frame (for animated only)
    LastUpdated As Long     'Tick count the Grh was last updated at
End Type

'Map information
Public Const MAP_MAXSIZE As Integer = 1000
Public Const TILETYPE_NOTHING As Byte = 0   'No attributes
Public Const TILETYPE_BLOCKED As Byte = 1   'Blocked in all directions (can not occupy same space as tile)
Public Const TILETYPE_PLATFORM As Byte = 2  'Can walk/jump through tile, but can also land on it
Public Const TILETYPE_LADDER As Byte = 3    'Tile is a ladder
Public Const TILETYPE_SPAWN As Byte = 4     'NPCs can spawn on the tile
Public Type tMapInfo
    Name As String                  'Name of the map (max length 30)
    TileWidth As Integer            'Width of the map in tiles (max size MAP_MAXSIZE)
    TileHeight As Integer           'Height of the map in tiles (max size MAP_MAXSIZE)
    HasFloatingBlocks As Boolean    'If there are any blocks on the map floating (if false, performance is improved)
    MusicID As Integer              'ID of the music for the map
    TileInfo() As Byte              'Information of each individual tile
End Type

'Body size information
Public Type tBodyInfo
    Width As Integer
    Height As Integer
    PunchWidth As Integer
    PunchTime As Long
End Type

'2d location with bytes, used for tile location
Public Type tWorldTilePos
    Map As Integer
    TileX As Integer
    TileY As Integer
End Type
Public Type tTilePos
    TileX As Integer
    TileY As Integer
End Type

'2d location with integers, used for pixel location
Public Type tWorldPixelPos
    Map As Integer
    X As Integer
    Y As Integer
End Type
Public Type tPixelPos
    X As Integer
    Y As Integer
End Type

'NPC map spawning information
Public Type tMapSpawn
    NPCID As Integer    'ID of the NPC to spawn
    Amount As Byte      'Amount of the NPC to spawn
End Type

'Map graphics
Public Type tMapGrh
    Grh As tGrh             'Graphic information for the tile
    X As Integer            'X pixel co-ordinate of the top-left corner of the graphic
    Y As Integer            'Y pixel co-ordinate of the top-left corner of the graphic
    Behind As Byte          'If the graphic is behind the user, True for in front of them
End Type

'Map backgrounds
Public Const NumBGLayers As Long = 10
Public Const BGGridSize As Long = 512
Public Type tBackground
    Segment() As tGrh       'Each segment of the background
End Type

'Map items on the server
Public Type tServerMapItem
    ItemIndex As Integer
    Amount As Integer
    X As Integer
    Y As Integer
    Life As Long
End Type

'How long an item last on the ground of a map
Public Const MAPITEMLIFE As Long = 45000

'Map items on the client
Public Type tClientMapItem
    Grh As tGrh                 'Graphic information
    ItemIndex As Integer        'Item's index
    Amount As Integer           'Amount of the item
    X As Integer                'Current X pixel co-ordinate
    Y As Integer                'Current Y pixel co-ordinate
    DestY As Integer            'The item must land on this Y co-ordinate (prevents excessive drops)
    DestX As Integer            'The item may not pass 45 pixels in either direction from this point
    Xv As Single                'X velocity
    Yv As Single                'Y velocity
    LastPickupAttempt As Long   'When the user last tried to pick up the item - used to prevent server spam
    Moving As Boolean           'If the map item is moving (Xv <> 0, Yv <> 0, Yv < DestY)
End Type

'Specifically for map items that are fading out (so we can reuse the slot without waiting for the fade)
'Items that are fading out only do that - fade out. You can not interact with them, pick them up, they do
'not move, etc. This is only for graphical effect, thats it.
Public Type tClientFadeItem
    Grh As tGrh         'Graphic information
    X As Integer        'X pixel co-ordinate
    Y As Integer        'Y pixel co-ordinate
    Alpha As Single     'Alpha value of the item (goes from 255 to 0)
End Type

'How long the user must wait between attempts to pickup a single item again (not the same as picking
'up different items - this is only for LastPickupAttempt)
Public Const PICKUPSAMEITEMTIME As Long = 1000

'How long the user must wait between pickup attempts in general
Public Const PICKUPITEMTIME As Long = 200

'Twice of how far away from an item the user can be to grab it (simple take the max pixel distance in any direction, multiply by two)
Public Const PICKUPITEMDISTCLIENT As Long = 40 * 2

'Twice of how far the server will allow the user's distance from the client
Public Const PICKUPITEMDISTSERVER As Long = 160 * 2

'Character movement constants
Public Const MOVESPEED As Single = 0.2
Public Const GRAVITY As Single = 0.7
Public Const JUMPDECAY As Single = 0.004
Public Const JUMPHEIGHT As Single = 2.5
Public Const HITTIME As Integer = 500 - 50  '- 50 for lag compensation

'Character action information (jumping/walking not included as an action)
Public Enum eCharAction
    eNone = 0
    ePunch = 1
    eHit = 2
    eDeath = 3
End Enum

'Characters on the client
Public Type tClientChar
    Name As String          'Displayed name of the character
    HPP As Byte              'Character's displayed HP percent
    MPP As Byte              'Character's displayed MP percent
    Jump As Single          'The characters's jumping strength (0 = not jumping)
    Body As tGrh            'Body grh (if a NPC, Body is the same as Sprite)
    BodyIndex As Byte       'Body paper-doll index (if a NPC, BodyIndex is the same as SpriteIndex)
    X As Single             'X pixel co-ordinate of the top-left corner of the character
    Y As Single             'Y pixel co-ordinate of the top-left corner of the character
    DrawX As Single         'Draw pixel co-ordinates of the top-left corner of the character, used to
    DrawY As Single         'allow us to smooth out the character position corrections
    LastX As Single         'Last pixel co-ordinate of the top-left corner of the character - used to
    LastY As Single         'determine what direction they have moved to update their animation accordingly
    LastTileX As Integer    'Last tile position of the top-left corner of the character - used to
    LastTileY As Integer    'check for collisions inside of blocked areas
    Width As Integer        'The width of the character's collision (in pixels)
    Height As Integer       'The height of the character's collision (in pixels)
    OnGround As Byte        'If the character is currently on the ground
    MoveDir As Byte         'The direction the character is moving
    IdleFrames As Byte      'How many frames the user has been idle for
    Heading As Byte         'If the character is facing left or right
    Used As Boolean         'If the character index is used
    IsNPC As Boolean        'If the character is a NPC
    Action As eCharAction   'The action the character is currently performing (jumping/walking not included as an action)
End Type

'Equipped items
Public Const EQUIPSLOT_CAP As Byte = 1
Public Const EQUIPSLOT_FOREHEAD As Byte = 2
Public Const EQUIPSLOT_RING1 As Byte = 3
Public Const EQUIPSLOT_RING2 As Byte = 4
Public Const EQUIPSLOT_RING3 As Byte = 5
Public Const EQUIPSLOT_RING4 As Byte = 6
Public Const EQUIPSLOT_EYEACC As Byte = 7
Public Const EQUIPSLOT_EARACC As Byte = 8
Public Const EQUIPSLOT_MANTLE As Byte = 9
Public Const EQUIPSLOT_CLOTHES As Byte = 10
Public Const EQUIPSLOT_PENDANT As Byte = 11
Public Const EQUIPSLOT_WEAPON As Byte = 12
Public Const EQUIPSLOT_SHIELD As Byte = 13
Public Const EQUIPSLOT_GLOVES As Byte = 14
Public Const EQUIPSLOT_PANTS As Byte = 15
Public Const EQUIPSLOT_SHOES As Byte = 16

'Item types
Public Const ITEMTYPE_USEONCE As Byte = 0
Public Const ITEMTYPE_WEAPON As Byte = 1
Public Const ITEMTYPE_CLOTHES As Byte = 2
Public Const ITEMTYPE_CAP As Byte = 3
Public Const ITEMTYPE_RING As Byte = 4
Public Const ITEMTYPE_FOREHEAD As Byte = 5
Public Const ITEMTYPE_EYEACC As Byte = 6
Public Const ITEMTYPE_EARACC As Byte = 7
Public Const ITEMTYPE_MANTLE As Byte = 8
Public Const ITEMTYPE_PENDANT As Byte = 9
Public Const ITEMTYPE_SHIELD As Byte = 10
Public Const ITEMTYPE_GLOVES As Byte = 11
Public Const ITEMTYPE_PANTS As Byte = 12
Public Const ITEMTYPE_SHOES As Byte = 13

'User stats (used by the client and server)
Public Type tUserStats
    HP As Integer
    MaxHP As Integer
    MP As Integer
    MaxMP As Integer
    Level As Integer
    EXP As Long
    Str As Integer
    Dex As Integer
    Intl As Integer
    Luk As Integer
    ModStr As Integer
    ModDex As Integer
    ModIntl As Integer
    ModLuk As Integer
    MinHit As Integer
    MaxHit As Integer
    Def As Integer
    Ryu As Long
End Type

'Item information for the client
Public Type tClientItem
    Name As String
    Desc As String
    GrhIndex As Long
    ItemType As Byte
    HP As Integer
    MP As Integer
    MaxHP As Integer
    MaxMP As Integer
    Str As Integer
    Dex As Integer
    Intl As Integer
    Luk As Integer
    MinHit As Integer
    MaxHit As Integer
    Def As Integer
    Stacking As Integer
End Type

'Specifically for items that have been picked up by a user, and is playing the
'animation of the item heading towards the user
'We use a separate UDT and array to free up the slot in case a new items is made
'in that slot, allowing us to have the new item displayed and the item still
'flying to the person who grabbed it
Public Type tClientPickupItem
    X As Single             'Current X pixel co-ordinate
    Y As Single             'Current Y pixel co-ordinate
    Width As Integer        'Width of the item
    Height As Integer       'Height of the item
    ToCharIndex As Integer  'The character index the item is going to
    Grh As tGrh             'Graphic information for the item
    Used As Boolean         'If the slot in use
End Type

'Inventory slots
Public Const USERINVSIZE As Byte = 23
Public Type tServerInvSlot
    ItemIndex As Integer
    Amount As Integer
End Type
Public Type tClientInvSlot
    ItemIndex As Integer
    Amount As Integer
    Grh As tGrh
End Type

'Paper-dolling information
Public Type tPDBody
    Width As Integer        'Collision width used when using this body
    Height As Integer       'Collision height used when using this body
    Stand As Long           'GrhIndex for standing
    Walk As Long            'GrhIndex for walking
    JumpUp As Long          'GrhIndex for jumping (going up)
    JumpDown As Long        'GrhIndex for jumping (going down)
    Punch As Long           'GrhIndex for punching
    PunchTime As Long       'How long the punch animation lasts
    PunchWidth As Long      'The size of the punch collision area
    Hit As Integer          'GrhIndex for being hit
    Death As Integer        'GrhIndex for dieing
End Type
Public Type tPDSprite
    Width As Integer        'Collision width used when using this body
    Height As Integer       'Collision height used when using this body
    Stand As Long           'GrhIndex for standing
    Walk As Long            'GrhIndex for walking
    JumpUp As Long          'GrhIndex for jumping (going up)
    JumpDown As Long        'GrhIndex for jumping (going down)
    Punch As Long           'GrhIndex for punching
    PunchTime As Long       'How long the punch animation lasts
    PunchWidth As Long      'The size of the punch collision area
    Hit As Integer          'GrhIndex for being hit
    Death As Integer        'GrhIndex for dieing
End Type

'Client NPC information
Public Type tClientNPC
    Name As String
    Sprite As Byte
    Width As Integer
    Height As Integer
    Level As Integer
    Heading As Byte
End Type

'Server NPC information
Public Type tServerNPCDrop
    ItemIndex As Integer
    Amount As Integer
    Chance As Single
End Type
Public Type tServerNPCFlags
    Status As Long
End Type
Public Type tServerNPCStats
    HP As Integer
    MaxHP As Integer
    MP As Integer
    MaxMP As Integer
    Level As Integer
    Str As Integer
    Dex As Integer
    Ryu As Integer
    EXP As Integer
    Intl As Integer
    Luk As Integer
    ModStr As Integer
    ModDex As Integer
    ModIntl As Integer
    ModLuk As Integer
    BaseMaxHP As Integer
    BaseMaxMP As Integer
    Def As Integer
    MinHit As Integer
    MaxHit As Integer
End Type
Public Type tServerNPC
    Name As String
    TemplateID As Integer
    CharIndex As Integer
    Sprite As Byte
    Map As Integer
    X As Single
    Y As Single
    Spawn As Integer
    Heading As Byte
    Flags As tServerNPCFlags
    Stats As tServerNPCStats
    NumDrops As Byte
    Drops() As tServerNPCDrop
End Type

'Converting between degrees and radians
Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

'VB doesn't support these following keys for some reason
Public Const vbKeyAlt As Integer = 18

'APIs
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
