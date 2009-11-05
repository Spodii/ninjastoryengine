Attribute VB_Name = "Declares"
Option Explicit

'If the client files are created
Public MakeClientInfo As Boolean

'If the priority is currently set to high
Public IsHighPriority As Boolean

'If the server loop is running
Public ServerRunning As Boolean

'Holds visual data for a user or NPC character
Public Type tCharData
    Index As Integer        'Points to the index of the NPC or User
    CharType As Byte        'States whether it is a PC or NPC
End Type
Public LastChar As Integer          'Highest CharList index used
Public CharListUBound As Integer    'Size of the CharList array
Public NumCharsFree As Integer      'Number of free indicies in the CharList array
Public CharList() As tCharData
Public Const CHARTYPE_NONE As Byte = 0
Public Const CHARTYPE_PC As Byte = 1
Public Const CHARTYPE_NPC As Byte = 2

'User connection status flags
Public Const PCSTATUSFLAG_DISCONNECTING As Long = 2 ^ 0

'NPC status flags
Public Const NPCSTATUSFLAG_SPAWNED As Long = 2 ^ 0

'Array of the user classes
Public UserList() As User
Public UserListUBound As Integer    'UBound of the UserList() array
Public LastUser As Integer          'Highest used index in the UserList() array

'Array of the NPC classes
Public NPCList() As NPC
Public NPCListUBound As Integer     'UBound of the NPCList() array
Public LastNPC As Integer           'Highest used index in the NPCList() array
Public NumNPCsFree As Integer       'Number of free indicies in the NPCList() array

'Array of the map classes
Public Maps() As Map
Public MapsUBound As Integer        'UBound of the Maps array

'Array for the body information
Public BodyInfo() As tBodyInfo
Public BodyInfoUBound As Byte

'Array for the sprite information
Public SpriteInfo() As tBodyInfo
Public SpriteInfoUBound As Byte

'Array for the items
Public Items() As Item
Public ItemsUBound As Integer

'Cached packets for server messages with no parameters
Public Const NumcMessages As Long = 100
Public Type cMessageData
    Data(0 To 1) As Byte
End Type
Public cMessage(1 To NumcMessages) As cMessageData

'Internal socket ID
Public LocalSocketID As Long
