VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ID of the item
Private ItemID As Integer

'Type of the item
Private cItemType As Byte

'How deep the item can stack
Private cStacking As Integer

'Stats that the item can raise
Private cHP As Integer
Private cMP As Integer
Private cMaxHP As Integer
Private cMaxMP As Integer
Private cStre As Integer
Private cDex As Integer
Private cIntl As Integer
Private cLuk As Integer
Private cMinHit As Integer
Private cMaxHit As Integer
Private cDef As Integer

Public Sub Init(ByVal pItemID As Integer)
'*********************************************************************************
'Load up the values of an item
'*********************************************************************************

    'Make the database query
    DB_RS.Open "SELECT * FROM items WHERE `id`='" & pItemID & "'", DB_Conn, adOpenStatic, adLockOptimistic

    'Check that the user exists
    If DB_RS.EOF Then
        DB_RS.Close
        Exit Sub
    End If

    'Fill the values
    ItemID = pItemID
    HP = DB_RS!HP
    MP = DB_RS!MP
    MaxHP = DB_RS!MaxHP
    MaxMP = DB_RS!MaxMP
    Str = DB_RS!Str
    Dex = DB_RS!Dex
    Intl = DB_RS!Intl
    Luk = DB_RS!Luk
    MinHit = DB_RS!MinHit
    MaxHit = DB_RS!MaxHit
    Def = DB_RS!Def
    Stacking = DB_RS!Stacking

    'Get the item type
    Select Case Trim$(DB_RS!Type)
    Case "USEONCE":     ItemType = ITEMTYPE_USEONCE
    Case "WEAPON":      ItemType = ITEMTYPE_WEAPON
    Case "CLOTHES":     ItemType = ITEMTYPE_CLOTHES
    Case "CAP":         ItemType = ITEMTYPE_CAP
    Case "RING":        ItemType = ITEMTYPE_RING
    Case "FOREHEAD":    ItemType = ITEMTYPE_FOREHEAD
    Case "EYEACC":      ItemType = ITEMTYPE_EYEACC
    Case "EARACC":      ItemType = ITEMTYPE_EARACC
    Case "MANTLE":      ItemType = ITEMTYPE_MANTLE
    Case "PENDANT":     ItemType = ITEMTYPE_PENDANT
    Case "SHIELD":      ItemType = ITEMTYPE_SHIELD
    Case "GLOVES":      ItemType = ITEMTYPE_GLOVES
    Case "PANTS":       ItemType = ITEMTYPE_PANTS
    Case "SHOES":       ItemType = ITEMTYPE_SHOES
    End Select
    
    'Close the recordset
    DB_RS.Close

End Sub

Public Property Get Stacking() As Integer
    Stacking = cStacking
End Property

Public Property Let Stacking(ByVal Value As Integer)
    cStacking = Value
End Property

Public Property Get ItemType() As Byte
    ItemType = cItemType
End Property

Private Property Let ItemType(ByVal Value As Byte)
    cItemType = Value
End Property

Public Property Get MinHit() As Integer
    MinHit = cMinHit
End Property

Private Property Let MinHit(ByVal Value As Integer)
    cMinHit = Value
End Property

Public Property Get MaxHit() As Integer
    MaxHit = cMaxHit
End Property

Private Property Let MaxHit(ByVal Value As Integer)
    cMaxHit = Value
End Property

Public Property Get Def() As Integer
    Def = cDef
End Property

Private Property Let Def(ByVal Value As Integer)
    cDef = Value
End Property

Public Property Get HP() As Integer
    HP = cHP
End Property

Private Property Let HP(ByVal Value As Integer)
    cHP = Value
End Property

Public Property Get MP() As Integer
    MP = cMP
End Property

Private Property Let MP(ByVal Value As Integer)
    cMP = Value
End Property

Public Property Get MaxHP() As Integer
    MaxHP = cMaxHP
End Property

Private Property Let MaxHP(ByVal Value As Integer)
    cMaxHP = Value
End Property

Public Property Get MaxMP() As Integer
    MaxMP = cMaxMP
End Property

Private Property Let MaxMP(ByVal Value As Integer)
    cMaxMP = Value
End Property

Public Property Get Str() As Integer
    Str = cStre
End Property

Private Property Let Str(ByVal Value As Integer)
    cStre = Value
End Property

Public Property Get Dex() As Integer
    Dex = cDex
End Property

Private Property Let Dex(ByVal Value As Integer)
    cDex = Value
End Property

Public Property Get Intl() As Integer
    Intl = cIntl
End Property

Private Property Let Intl(ByVal Value As Integer)
    cIntl = Value
End Property

Public Property Get Luk() As Integer
    Luk = cLuk
End Property

Private Property Let Luk(ByVal Value As Integer)
    cLuk = Value
End Property
