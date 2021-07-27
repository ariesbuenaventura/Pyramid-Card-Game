VERSION 5.00
Begin VB.UserControl Card 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1065
   ScaleHeight     =   1440
   ScaleWidth      =   1065
   ToolboxBitmap   =   "ucCard.ctx":0000
End
Attribute VB_Name = "Card"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CPS = 13                   ' Cards Per Suit

Private Const OffsetCardRankResId = 101
Private Const OffsetCardDeckResId = 201

Public Enum DeckConstants
    crdDCDefault
    crdDCDeck001
    crdDCDeck002
End Enum

Public Enum FaceConstants
    crdFCFaceUp
    crdFCFaceDn
End Enum

Public Enum RankConstants
    crdRCAce
    crdRCTwo
    crdRCThree
    crdRCFour
    crdRCFive
    crdRCSix
    crdRCSeven
    crdRCEight
    crdRCNine
    crdRCTen
    crdRCJack
    crdRCQueen
    crdRCKing
End Enum

Public Enum SuitConstants
    crdSCClubs
    crdSCSpades
    crdSCHearts
    crdSCDiamond
End Enum

Private Type CardProperties
    Data     As String
    Deck     As DeckConstants
    Face     As FaceConstants
    Rank     As RankConstants
    Suit     As SuitConstants
    Selected As Boolean
End Type

Event Click()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLECompleteDrag(Effect As Long)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)

Dim MyProp As CardProperties

Public Property Get Data() As String
    Data = MyProp.Data
End Property

Public Property Let Data(New_Data As String)
    MyProp.Data = New_Data
End Property

Public Property Get Deck() As DeckConstants
    Deck = MyProp.Deck
End Property

Public Property Let Deck(New_Deck As DeckConstants)
    MyProp.Deck = New_Deck
    
    If MyProp.Face = crdFCFaceDn Then Call RedrawCard
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Face() As FaceConstants
    Face = MyProp.Face
End Property

Public Property Let Face(New_Face As FaceConstants)
    MyProp.Face = New_Face
    
    Call RedrawCard
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Rank() As RankConstants
    Rank = MyProp.Rank
End Property

Public Property Let Rank(New_Rank As RankConstants)
    MyProp.Rank = New_Rank
    
    If MyProp.Face = crdFCFaceUp Then Call RedrawCard
End Property

Public Property Get Suit() As SuitConstants
    Suit = MyProp.Suit
End Property

Public Property Let Suit(New_Suit As SuitConstants)
    MyProp.Suit = New_Suit
    
    If MyProp.Face = crdFCFaceUp Then Call RedrawCard
End Property

Public Property Get Selected() As Boolean
    Selected = MyProp.Selected
End Property

Public Property Let Selected(New_Selected As Boolean)
    MyProp.Selected = New_Selected
    
    Call RedrawCard
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    MyProp.Data = vbNullString
    MyProp.Deck = crdDCDefault
    MyProp.Face = crdFCFaceUp
    MyProp.Rank = crdRCAce
    MyProp.Suit = crdSCClubs
    MyProp.Selected = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Public Property Get OLEDropMode() As Integer
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    MyProp.Data = PropBag.ReadProperty("Data", vbNullString)
    MyProp.Deck = PropBag.ReadProperty("Deck", crdDCDefault)
    MyProp.Face = PropBag.ReadProperty("Face", crdFCFaceUp)
    MyProp.Rank = PropBag.ReadProperty("Rank", crdRCAce)
    MyProp.Suit = PropBag.ReadProperty("Suit", crdSCClubs)
    MyProp.Selected = PropBag.ReadProperty("Selected", False)
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
End Sub

Private Sub UserControl_Resize()
    Call RedrawCard
End Sub

Private Sub UserControl_Show()
    Call RedrawCard
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Data", MyProp.Data, vbNullString
    PropBag.WriteProperty "Deck", MyProp.Deck, crdDCDefault
    PropBag.WriteProperty "Face", MyProp.Face, crdFCFaceUp
    PropBag.WriteProperty "Rank", MyProp.Rank, crdRCAce
    PropBag.WriteProperty "Suit", MyProp.Suit, crdSCClubs
    PropBag.WriteProperty "Selected", MyProp.Selected, False
    
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, vbDefault)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
End Sub

Private Sub RedrawCard()
    Dim resId As Integer
    
    If MyProp.Face = crdFCFaceUp Then
        resId = OffsetCardRankResId + MyProp.Rank + MyProp.Suit * CPS
    Else
        resId = OffsetCardDeckResId + MyProp.Deck
    End If
    
    UserControl.Cls
    UserControl.PaintPicture LoadResPicture(resId, vbResBitmap), 0, 0, _
                             UserControl.ScaleWidth, UserControl.ScaleHeight, , , , , _
                             IIf(MyProp.Selected, vbSrcInvert, vbSrcCopy)
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - Screen.TwipsPerPixelX, _
                             UserControl.ScaleHeight - Screen.TwipsPerPixelY), _
                             IIf(MyProp.Selected, vbWhite, vbBlack), B
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Function hDC() As Long
    hDC = UserControl.hDC
End Function

Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub

