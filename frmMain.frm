VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Pyramid 1.0 by Aris Buenaventura"
   ClientHeight    =   7800
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   689
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTable 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   7455
      Left            =   0
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      Begin prjPyramid.Card crdPlayingCard 
         Height          =   1440
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   2540
      End
   End
   Begin VB.PictureBox picScore 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Height          =   315
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   689
      TabIndex        =   0
      Tag             =   "0"
      Top             =   7485
      Width           =   10335
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameDeal 
         Caption         =   "&Deal"
      End
      Begin VB.Menu mnuGameBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuGameDeck 
         Caption         =   "Deck..."
      End
      Begin VB.Menu mnuGameBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameCardSize 
         Caption         =   "&Card Size"
         Begin VB.Menu mnuGameCardSizeSel 
            Caption         =   "&Small"
            Index           =   0
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuGameCardSizeSel 
            Caption         =   "&Normal"
            Index           =   1
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuGameCardSizeSel 
            Caption         =   "&Large"
            Index           =   2
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu mnuGameLevel 
         Caption         =   "&Level"
         Begin VB.Menu mnuGameLevelNormal 
            Caption         =   "&Normal"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuGameLevelDifficult 
            Caption         =   "&Difficult"
         End
      End
      Begin VB.Menu mnuGameBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameShowHint 
         Caption         =   "Show &Hint"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuGameDemo 
         Caption         =   "Show &Demo"
      End
      Begin VB.Menu mnuGameBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHint 
      Caption         =   "Hint!"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuStopDemo 
      Caption         =   "Stop Demo"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CSTOCK = 0
Private Const CWASTE = 1
Private Const CPYRAMID = 2

Private Const CARD_WIDTH = 71
Private Const CARD_HEIGHT = 96

Private Const TOTAL_CARDS = 52

Private Type PosInfo
    X As Long
    Y As Long
End Type

Private Type UndoInfo
    Data          As String
    Enabled       As Boolean
    Face          As FaceConstants
    MousePointer  As MousePointerConstants
    Selected      As Boolean
    Tag           As String
    Visible       As Boolean
    arrCardIndex  As Integer
End Type

Dim UndoScore    As Integer
Dim UndoHint     As Boolean
Dim IsFormLoaded As Boolean
Dim StopDemo     As Boolean
Dim IsDragDrop                     As Boolean
Dim IsDragDropOn                   As Boolean
Dim MousePos                       As PosInfo
Dim OriginalLocation               As PosInfo
Dim Undo(1 To TOTAL_CARDS)         As UndoInfo
Dim arrCardIndex(1 To TOTAL_CARDS) As Integer  ' stores all index of unwanted cards

Private Sub crdPlayingCard_Click(Index As Integer)
    If IsDragDropOn Then
        IsDragDropOn = False
        Exit Sub
    End If
    
    If GetCountCard(CPYRAMID) = 0 Then Exit Sub
    
    If mnuGameLevelNormal.Checked Then
        If (crdPlayingCard(Index).Data = "@") And _
           (crdPlayingCard(Index).Tag <> "S") Then
            Exit Sub
        End If
    Else
        If (crdPlayingCard(Index).Face = crdFCFaceDn) And _
           (crdPlayingCard(Index).Tag <> "S") Then
            Exit Sub
        End If
    End If
    
    Select Case crdPlayingCard(Index).Tag
    Case Is = "P", "W" ' P - cards that form pyramid; W - waste cards
        Dim i    As Integer ' iteration
        Dim sum  As Integer ' sum of selected ranks
        Dim bval As Boolean
        
        If crdPlayingCard(Index).Face = crdFCFaceUp Then
            crdPlayingCard(Index).Selected = Not crdPlayingCard(Index).Selected
            crdPlayingCard(Index).Refresh
            
            For i = crdPlayingCard.LBound To crdPlayingCard.UBound
                If crdPlayingCard(i).Selected Then
                    sum = sum + crdPlayingCard(i).Rank + 1
                End If
            Next i
        End If
        
        If (sum = 13) Or (GetTotalSelected = 2) Then bval = True
        
        If bval Then
            picTable.Refresh
            Sleep 600 ' pause for at least 0.6 second
            If sum = 13 Then Call SaveMove
            
            For i = crdPlayingCard.LBound To crdPlayingCard.UBound
                If crdPlayingCard(i).Selected Then
                    crdPlayingCard(i).Selected = False
                    If sum = 13 Then Process i
                End If
            Next i
        End If
    Case Is = "S" ' stock cards
        Dim cntr As Integer ' counter
        
        Call SaveMove
        
        If mnuGameLevelNormal.Checked Then
            If crdPlayingCard(Index).Rank = crdRCKing Then
                ' Kings are discarded singly
                crdPlayingCard(Index).Visible = False
                crdPlayingCard(Index).Tag = "U"
                
                For i = crdPlayingCard.LBound To crdPlayingCard.UBound
                    If arrCardIndex(i) = 0 Then
                        arrCardIndex(i) = Index
                        Exit For
                    End If
                Next i
                
                Call picTable_Resize
                Exit Sub
            End If
        End If
        
        cntr = 1
        For i = crdPlayingCard.UBound To crdPlayingCard.LBound Step -1
            If crdPlayingCard(i).Visible Then
                If crdPlayingCard(i).Tag = "S" Then
                    crdPlayingCard(i).Tag = "W"
                    crdPlayingCard(i).Face = crdFCFaceUp
                    crdPlayingCard(i).Move CInt(picTable.ScaleWidth / 2) + 10, _
                                           picTable.ScaleHeight - GetCardHeight - 5
                    crdPlayingCard(i).ZOrder vbBringToFront
        
                    If (cntr = 3) Or mnuGameLevelNormal.Checked Then
                        Exit For
                    Else
                        cntr = cntr + 1
                    End If
                ElseIf crdPlayingCard(i).Tag = "W" Then
                    If crdPlayingCard(i).Selected Then
                        crdPlayingCard(i).Selected = False
                    End If
                End If
            End If
        Next i
        
        Call SearchHint
    End Select
    
    Call StatGame
End Sub

Private Sub crdPlayingCard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (crdPlayingCard(Index).Rank = crdRCKing) And _
       (crdPlayingCard(Index).Data <> "@") And _
       (crdPlayingCard(Index).Face = crdFCFaceUp) And _
       (GetTotalSelected = 0) Then
        Call SaveMove
        Process Index
        Call StatGame
        Exit Sub
    End If
        
    If Button And vbLeftButton Then
        IsDragDrop = IsCardClickable(Index)
    
        If IsDragDrop Then
            MousePos.X = ScaleX(X, vbTwips, vbPixels)
            MousePos.Y = ScaleY(Y, vbTwips, vbPixels)
                    
            OriginalLocation.X = crdPlayingCard(Index).Left
            OriginalLocation.Y = crdPlayingCard(Index).Top
        End If
    End If
        
    If IsCardClickable(Index) Then
        crdPlayingCard(Index).MousePointer = vbCustom
        Set crdPlayingCard(Index).MouseIcon = LoadResPicture(102, vbResCursor)
        crdPlayingCard(Index).ZOrder vbBringToFront
    End If
End Sub

Private Sub crdPlayingCard_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsDragDrop Then Exit Sub
    
    Dim idx    As Integer
    Dim NewPos As PosInfo
    
    Static OldGetSelIndex As Integer
    
    If Button And vbLeftButton Then
        NewPos.X = crdPlayingCard(Index).Left - (MousePos.X - ScaleX(X, vbTwips, vbPixels))
        NewPos.Y = crdPlayingCard(Index).Top - (MousePos.Y - ScaleY(Y, vbTwips, vbPixels))

        crdPlayingCard(Index).Move NewPos.X, NewPos.Y
        picTable.Refresh
        
        idx = GetSelIndex(Index)
        If idx <> 0 Then
            If (idx <> OldGetSelIndex) And (OldGetSelIndex <> 0) Then
                If crdPlayingCard(OldGetSelIndex).Selected Then
                    crdPlayingCard(OldGetSelIndex).Selected = False
                End If
            End If
              
            If Not crdPlayingCard(idx).Selected Then
                crdPlayingCard(idx).Selected = True
            End If
            
            OldGetSelIndex = idx
        Else
            If OldGetSelIndex <> 0 Then
                If crdPlayingCard(OldGetSelIndex).Selected Then
                    crdPlayingCard(OldGetSelIndex).Selected = False
                End If
            End If
        End If
        
        If ((OriginalLocation.X < crdPlayingCard(Index).Left - 3) Or _
            (OriginalLocation.X > crdPlayingCard(Index).Left + 3)) Or _
           ((OriginalLocation.Y < crdPlayingCard(Index).Top - 3) Or _
            (OriginalLocation.Y > crdPlayingCard(Index).Top + 3)) Then
            If Not IsDragDropOn Then IsDragDropOn = True
        Else
            If IsDragDropOn Then IsDragDropOn = False
        End If
    End If
End Sub

Private Sub crdPlayingCard_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If crdPlayingCard(Index).MousePointer = vbCustom Then
        Set crdPlayingCard(Index).MouseIcon = LoadResPicture(101, vbResCursor)
    End If
    
    If IsCardClickable(Index) Then
        crdPlayingCard(Index).ZOrder vbBringToFront
    End If
    
    If Not IsDragDrop Then Exit Sub
    
    Dim i   As Integer
    Dim idx As Integer
    
    idx = GetSelIndex(Index)
    If idx <> 0 Then
        Call SaveMove
        For i = crdPlayingCard.LBound To crdPlayingCard.UBound
            If crdPlayingCard(i).Visible Then
                If (i = Index) Or (i = idx) Then
                    If crdPlayingCard(i).Selected Then
                        crdPlayingCard(i).Selected = False
                    End If
                    
                    Process i
                End If
            End If
        Next i
                
        Call StatGame
    Else
        LockWindowUpdate Me.hwnd
        crdPlayingCard(Index).Left = OriginalLocation.X
        crdPlayingCard(Index).Top = OriginalLocation.Y
        LockWindowUpdate 0
    End If
        
    IsDragDrop = False
End Sub

Private Sub Form_Load()
    Dim i   As Integer ' iteration
    Dim lvl As Integer ' level
    Dim sz  As Integer ' card size
    
    IsFormLoaded = False
    
    ' restore setting
    lvl = GetSetting("Pyramid", "Setting", "Level", 1)
    sz = GetSetting("Pyramid", "Setting", "Size", 1)
    
    If lvl = 0 Then
        mnuGameLevelNormal.Checked = True
        mnuGameLevelDifficult.Checked = False
    Else
        mnuGameLevelNormal.Checked = False
        mnuGameLevelDifficult.Checked = True
    End If
    
    mnuGameCardSizeSel_Click sz
    mnuGameShowHint.Checked = _
        GetSetting("Pyramid", "Setting", "Hint", vbChecked)
    crdPlayingCard(crdPlayingCard.LBound).Deck = _
        GetSetting("Pyramid", "Setting", "Deck", crdDCDefault)
    mnuHint.Visible = CBool(mnuGameShowHint.Checked)
    
    ' set playing card
    For i = crdPlayingCard.LBound + 1 To TOTAL_CARDS
        Load crdPlayingCard(i) ' create new card/control
        crdPlayingCard(i).Move -GetCardWidth, -GetCardHeight, _
                                GetCardWidth, GetCardHeight
    
        Set crdPlayingCard(i).MouseIcon = LoadResPicture(101, vbResCursor)
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HtmlHelp Me.hwnd, "", HH_CLOSE_ALL, 0&
End Sub

Private Sub Form_Resize()
    '   Do not update the window yet until all cards/controls
    ' are placed in their designated position.
    LockWindowUpdate Me.hwnd ' for fast display
    
    picTable.Height = Me.ScaleHeight - picScore.Height
    
    ' Unlock window update
    LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim SelSZ As Integer ' selected size
    
    If mnuGameCardSizeSel(0).Checked Then
        SelSZ = 0
    ElseIf mnuGameCardSizeSel(1).Checked Then
        SelSZ = 1
    Else
        SelSZ = 2
    End If
    
    ' save setting
    SaveSetting "Pyramid", "Setting", "Level", _
                IIf(mnuGameLevelNormal.Checked, 0, 1)
    SaveSetting "Pyramid", "Setting", "Size", SelSZ
    SaveSetting "Pyramid", "Setting", "Deck", _
                crdPlayingCard(crdPlayingCard.LBound).Deck
    SaveSetting "Pyramid", "Setting", "Hint", mnuGameShowHint.Checked
    
    
    SaveSetting "14613165", "15431", "SD", GetSetting("14613165", "15431", "SD", 15) - 1
    
    End
End Sub

Private Sub mnuGameCardSizeSel_Click(Index As Integer)
    Dim i As Integer ' iteration
    
    For i = mnuGameCardSizeSel.LBound To mnuGameCardSizeSel.UBound
        If i <> Index Then
            mnuGameCardSizeSel(i).Checked = False
        Else
            mnuGameCardSizeSel(i).Checked = True
        End If
    Next i
    
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        crdPlayingCard(i).Width = CARD_WIDTH * GetSizePercentage
        crdPlayingCard(i).Height = CARD_HEIGHT * GetSizePercentage
    Next i
    
    Call picTable_Resize
End Sub

Private Sub mnuGameDeal_Click()
    Dim i                         As Integer ' iteration
    Dim ll                        As Integer ' lower limit
    Dim ul                        As Integer ' upper limit
    Dim cntr                      As Integer ' counter
    Dim rndval                    As Integer ' random value
    Dim isExist                   As Boolean ' test if value exist
    Dim arrTemp(1 To TOTAL_CARDS) As Integer ' temporary array
    
    mnuGameDemo.Enabled = True
    mnuGameUndo.Enabled = False
    Erase arrCardIndex
    
    cntr = 1
    With crdPlayingCard
        ll = .LBound
        ul = .UBound
        
        Call Randomize
        
        Do While cntr < ul + 1
            isExist = False
            rndval = Int((ul - ll + 1) * Rnd + ll)
            
            For i = 1 To cntr
                If arrTemp(i) = rndval Then
                    isExist = True
                    Exit For
                End If
            Next i
            
            If Not isExist Then
                arrTemp(cntr) = rndval
                cntr = cntr + 1
            End If
        Loop
    End With
        
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        If crdPlayingCard(i).Selected Then crdPlayingCard(i).Selected = False
        
        crdPlayingCard(i).Data = ""
        
        If mnuGameLevelNormal.Checked Then
            If i < 22 Then
                crdPlayingCard(i).Data = "@"
            End If
            
            crdPlayingCard(i).Face = crdFCFaceUp
        Else
            If (i > 21) And (i < 29) Then
                crdPlayingCard(i).Face = crdFCFaceUp
            Else
                crdPlayingCard(i).Face = crdFCFaceDn
            End If
        End If
        
        If (i < 22) Then
            crdPlayingCard(i).MousePointer = vbDefault
        Else
            crdPlayingCard(i).MousePointer = vbCustom
        End If
        
        If i < 29 Then
            crdPlayingCard(i).Tag = "P" ' P stands for pyramid (cards that form pyramid)
        Else
            crdPlayingCard(i).Tag = "S" ' S stands for stock
        End If
        
        ' each suit has 13 cards
        crdPlayingCard(i).Rank = (arrTemp(i) - 1) Mod 13
        ' their are 4 suits (clubs, spades, hearts, diamond)
        crdPlayingCard(i).Suit = (arrTemp(i) - 1) Mod 4
        crdPlayingCard(i).Visible = True
        crdPlayingCard(i).Enabled = True
    Next i
    
    If Not IsFormLoaded Then IsFormLoaded = True
    
    picScore.Tag = 0
    
    Call AlignPlayingCards
    Call ShowScore
    Call SearchHint
End Sub

Private Sub mnuGameDeck_Click()
    frmDeck.Show vbModal, Me
End Sub

Private Sub mnuGameDemo_Click()
    Dim Sel1               As Integer
    Dim Sel2               As Integer
    Dim IsHint             As Boolean
    Dim TotalActiveCard    As Integer
    Dim OldTotalActiveCard As Integer
            
    mnuStopDemo.Visible = True
    mnuGameDeal.Enabled = False
    mnuGameDemo.Enabled = False
    mnuGameLevelNormal.Enabled = False
    mnuGameLevelDifficult.Enabled = False
    mnuHint.Enabled = False
    Me.Refresh
    
    StopDemo = False
    Do While mnuStopDemo.Visible
        If GetHint(Sel1, Sel2) Then
            If Sel1 <> 0 Then crdPlayingCard(Sel1).Selected = True
            If Sel2 <> 0 Then crdPlayingCard(Sel2).Selected = True
            picTable.Refresh
            Sleep 800
                    
            If Sel1 <> 0 Then crdPlayingCard(Sel1).Selected = False
            If Sel2 <> 0 Then crdPlayingCard(Sel2).Selected = False
                    
            If Sel1 <> 0 Then Process Sel1
            If Sel2 <> 0 Then Process Sel2
            Sleep 200
        Else
            If GetCountCard(CSTOCK) <> 0 Then
                crdPlayingCard_Click GetTopStock
            Else
                TotalActiveCard = GetCountCard(CSTOCK) + _
                                  GetCountCard(CWASTE) + _
                                  GetCountCard(CPYRAMID)
                        
                If OldTotalActiveCard <> TotalActiveCard Then
                    Dim xpos As Integer
                    Dim ypos As Integer
                
                    xpos = CInt(picTable.ScaleWidth / 2) - GetCardWidth - 10
                    ypos = picTable.ScaleHeight - GetCardHeight - 5
            
                    picTable_MouseDown vbLeftButton, 0, xpos + GetCardHeight / 2, _
                                                        ypos + GetCardWidth / 2
                    OldTotalActiveCard = TotalActiveCard
                Else
                    Exit Do
                End If
            End If
        End If
                        
        If GetCountCard(CPYRAMID) = 0 Then Exit Do
        DoEvents
    Loop
            
    mnuStopDemo.Visible = False
    mnuGameDeal.Enabled = True
    mnuGameDemo.Enabled = True
    mnuGameLevelNormal.Enabled = True
    mnuGameLevelDifficult.Enabled = True
    
    If (GetCountCard(CPYRAMID) <> 0) Then
        If Not StopDemo Then
            MsgBox "No more moves available", vbInformation, "Demo"
        End If
    Else
        Dim i As Integer
        
        For i = crdPlayingCard.LBound To crdPlayingCard.UBound
            crdPlayingCard(i).Enabled = False
        Next i
        Exit Sub
    End If
        
    Call SearchHint
End Sub

Private Sub mnuGameExit_Click()
    Unload Me
End Sub

Private Sub AlignPlayingCards()
    Dim X      As Integer ' x-position
    Dim Y      As Integer ' y-position
    Dim SW     As Integer ' scale width
    Dim SH     As Integer ' scale height
    Dim row    As Integer ' row
    Dim col    As Integer ' column
    Dim curIdx As Integer ' current index
    Dim i      As Integer ' iteration
    
    SW = picTable.ScaleWidth
    SH = picTable.ScaleHeight
    
    curIdx = 1
    For row = 1 To 7
        For col = row To 1 Step -1
            X = (SW - (GetCardWidth + GetCardWidth * 0.05) * row) / 2
            X = SW - (X + col * (GetCardWidth + GetCardWidth * 0.05)) + 2
            Y = (row - 1) * GetCardHeight * 0.5 + 5
            
            crdPlayingCard(curIdx).Move X, Y
            crdPlayingCard(curIdx).ZOrder vbBringToFront
            curIdx = curIdx + 1 ' next card
        Next col
    Next row
    
    X = SW / 2
    Y = SH - GetCardHeight - 5
        
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        If crdPlayingCard(i).Tag = "S" Then
            crdPlayingCard(i).Move X - GetCardWidth - 10, Y
            crdPlayingCard(i).ZOrder vbBringToFront
        ElseIf crdPlayingCard(i).Tag = "W" Then
            crdPlayingCard(i).Move X + 10, Y
            crdPlayingCard(i).ZOrder vbSendToBack
        End If
    Next i
    
    With picTable
        .Cls
        .DrawWidth = 4
        .ForeColor = vbBlack
        .FillStyle = vbDiagonalCross
        
        picTable.Line (X - GetCardWidth - 10, Y)-(X - 10, Y + GetCardHeight), , B
        picTable.Line (X + 10, Y)-(X + GetCardWidth + 10, Y + GetCardHeight), , B
        
        .DrawWidth = 15
        .ForeColor = vbGreen
        .FillStyle = vbFSTransparent
            
        If mnuGameLevelNormal.Checked Then
            picTable.Line (X - GetCardWidth - 10 + GetCardWidth / 2 - 20, _
                           Y + GetCardHeight / 2 - 20)- _
                          (X - GetCardWidth - 10 + GetCardWidth / 2 + 20, _
                           Y + GetCardHeight / 2 + 20)
            picTable.Line (X - GetCardWidth - 10 + GetCardWidth / 2 + 20, _
                           Y + GetCardHeight / 2 - 20)- _
                          (X - GetCardWidth - 10 + GetCardWidth / 2 - 20, _
                           Y + GetCardHeight / 2 + 20)
        Else
            picTable.Circle (X - GetCardWidth - 10 + GetCardWidth / 2, _
                             Y + GetCardHeight / 2), GetCardWidth / 2 - .DrawWidth * 0.8
        End If
    End With
End Sub

Private Sub ShowScore()
    Dim BH As Integer ' border height
    Dim SW As Integer ' scale width
    Dim SH As Integer ' scale height
    
    With picScore
        SW = .ScaleWidth
        SH = .ScaleHeight
        
        picScore.Line (-1, 0)-(SW + 1, SH), vbWhite, BF
        picScore.Line (-1, 0)-(SW + 1, SH), vbBlack, B
        
        .CurrentX = SW - .TextWidth("Score : " & .Tag) - .TextWidth("W") / 2
        .CurrentY = (SH - .TextHeight(.Tag)) / 2
        picScore.Print "Score : " & .Tag ' display score
    End With
End Sub

Private Sub mnuGameLevelDifficult_Click()
    If Not mnuGameLevelDifficult.Checked Then
        mnuGameLevelNormal.Checked = False
        mnuGameLevelDifficult.Checked = True
    End If
    
    Call mnuGameDeal_Click
End Sub

Private Sub mnuGameLevelNormal_Click()
    If Not mnuGameLevelNormal.Checked Then
        mnuGameLevelNormal.Checked = True
        mnuGameLevelDifficult.Checked = False
    End If
    
    Call mnuGameDeal_Click
End Sub

Private Sub mnuGameShowHint_Click()
    mnuGameShowHint.Checked = Not mnuGameShowHint.Checked
    
    mnuHint.Visible = CBool(mnuGameShowHint.Checked)
    If mnuHint.Visible Then Call SearchHint
End Sub

Private Sub mnuGameUndo_Click()
    Dim i As Integer
    
    mnuGameUndo.Enabled = False
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        crdPlayingCard(i).Data = Undo(i).Data
        crdPlayingCard(i).Enabled = Undo(i).Enabled
        crdPlayingCard(i).Face = Undo(i).Face
        crdPlayingCard(i).MousePointer = Undo(i).MousePointer
        crdPlayingCard(i).Selected = Undo(i).Selected
        crdPlayingCard(i).Tag = Undo(i).Tag
        crdPlayingCard(i).Visible = Undo(i).Visible
        arrCardIndex(i) = Undo(i).arrCardIndex
    Next i
        
    picScore.Tag = UndoScore
    mnuHint.Enabled = UndoHint
    
    Call picTable_Resize
    Call picScore_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    HtmlHelp Me.hwnd, App.Path & "\pyramid.chm", HH_DISPLAY_TOPIC, ByVal "Pyramid.htm"
End Sub

Private Sub mnuHint_Click()
    Dim Sel1 As Integer
    Dim Sel2 As Integer
    
    If GetHint(Sel1, Sel2) Then
        If Sel1 <> 0 Then crdPlayingCard(Sel1).Selected = True
        If Sel2 <> 0 Then crdPlayingCard(Sel2).Selected = True
        
        picTable.Refresh ' refresh the card first so that we can see the effect
        Sleep 300        ' pause for at least 0.3 second
        
        If Sel1 <> 0 Then crdPlayingCard(Sel1).Selected = False
        If Sel2 <> 0 Then crdPlayingCard(Sel2).Selected = False
    End If
End Sub

Private Sub mnuStopDemo_Click()
    mnuStopDemo.Visible = False
    mnuGameDeal.Enabled = True
    mnuGameDemo.Enabled = True
    mnuGameLevelNormal.Enabled = True
    mnuGameLevelDifficult.Enabled = True
    StopDemo = True
    Call SearchHint
End Sub

Private Sub picScore_Resize()
    Call ShowScore
End Sub

Private Sub picTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not crdPlayingCard(crdPlayingCard.LBound).Enabled Then Exit Sub
    If (GetCountCard(CSTOCK) <> 0) Or mnuGameLevelNormal.Checked Then Exit Sub
    
    Dim i    As Integer
    Dim xpos As Integer
    Dim ypos As Integer
    
    xpos = CInt(picTable.ScaleWidth / 2) - GetCardWidth - 10
    ypos = picTable.ScaleHeight - GetCardHeight - 5
    
    If (X >= xpos) And (X <= (xpos + GetCardWidth)) Then
        If (Y >= ypos) And (Y <= (ypos + GetCardHeight)) Then
            If Button And vbLeftButton Then
                Call SaveMove
                
                If GetCountCard(CWASTE) <> 0 Then
                    picTable.MousePointer = vbCustom
                    Set picTable.MouseIcon = LoadResPicture(102, vbResCursor)
                Else
                    picTable.MousePointer = vbDefault
                End If
    
                For i = crdPlayingCard.LBound To crdPlayingCard.UBound
                    If crdPlayingCard(i).Tag = "W" Then
                        crdPlayingCard(i).Tag = "S"
                        crdPlayingCard(i).Face = IIf(mnuGameLevelNormal.Checked, _
                                                     crdFCFaceUp, crdFCFaceDn)
                        crdPlayingCard(i).Move xpos, ypos
                        crdPlayingCard(i).ZOrder vbBringToFront
                    End If
                    
                    If crdPlayingCard(i).Selected Then
                        crdPlayingCard(i).Selected = False
                    End If
                Next i
            End If
        End If
    End If
End Sub

Private Sub picTable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xpos     As Integer
    Dim ypos     As Integer
    Dim IsRegion As Integer
    
    xpos = picTable.ScaleWidth / 2 - GetCardWidth - 10
    ypos = picTable.ScaleHeight - GetCardHeight - 5
    
    If (X >= xpos) And (X <= (xpos + GetCardWidth)) Then
        If (Y >= ypos) And (Y <= (ypos + GetCardHeight)) Then
            IsRegion = True
        End If
    End If
    
    If IsRegion Then
        If Not mnuGameLevelNormal.Checked Then
            If crdPlayingCard(crdPlayingCard.LBound).Enabled Then
                If GetCountCard(CWASTE) <> 0 Then
                    picTable.MousePointer = vbCustom
                    picTable.MouseIcon = LoadResPicture(101, vbResCursor)
                Else
                    picTable.MousePointer = vbDefault
                End If
            Else
                picTable.MousePointer = vbDefault
            End If
        Else
            picTable.MousePointer = vbDefault
        End If
    Else
        picTable.MousePointer = vbDefault
    End If
End Sub

Private Sub picTable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If picTable.MousePointer = vbCustom Then
        picTable.MouseIcon = LoadResPicture(101, vbResCursor)
    End If
End Sub

Private Sub picTable_Resize()
    If IsFormLoaded Then
        Call AlignPlayingCards
        
        Dim i As Integer
        
        For i = LBound(arrCardIndex()) To UBound(arrCardIndex())
            If arrCardIndex(i) <> 0 Then
                If i <= TOTAL_CARDS / 2 Then
                    BitBlt picTable.hDC, 3, (i - 1) * 17 + 5, _
                                         GetCardWidth, GetCardHeight, _
                           crdPlayingCard(arrCardIndex(i)).hDC, 0, 0, vbSrcCopy
                Else
                    BitBlt picTable.hDC, picTable.ScaleWidth - GetCardWidth - 3, _
                                        (i Mod ((TOTAL_CARDS / 2) + 1)) * 17 + 5, _
                                         GetCardWidth, GetCardHeight, _
                           crdPlayingCard(arrCardIndex(i)).hDC, 0, 0, vbSrcCopy
                End If
            End If
        Next i
    End If
End Sub

Private Sub Process(i As Integer)
    crdPlayingCard(i).Visible = False
    
    If crdPlayingCard(i).Tag = "P" Then
        Dim idx As Integer ' index of the card
                                                            
        idx = GetCalcIndex(i) ' calculate the index of the card in the pyramid
                                    
        If idx = GetCalcIndex(i - 1) Then
            If Not crdPlayingCard(i - 1).Visible Then
                ' opens the card above left
                If mnuGameLevelNormal.Checked Then
                    crdPlayingCard(i - idx).Data = ""
                    picScore.Tag = CInt(picScore.Tag) + 5
                Else
                    crdPlayingCard(i - idx).Face = crdFCFaceUp
                    picScore.Tag = CInt(picScore.Tag) + 10
                End If
                                            
                crdPlayingCard(i - idx).MousePointer = vbCustom
            End If
        End If
                                    
        If idx = GetCalcIndex(i + 1) Then
            If Not crdPlayingCard(i + 1).Visible Then
                ' opens the card above right
                If mnuGameLevelNormal.Checked Then
                    crdPlayingCard(i - idx + 1).Data = ""
                    picScore.Tag = CInt(picScore.Tag) + 5
                Else
                    crdPlayingCard(i - idx + 1).Face = crdFCFaceUp
                    picScore.Tag = CInt(picScore.Tag) + 10
                End If
                                            
                crdPlayingCard(i - idx + 1).MousePointer = vbCustom
            End If
        End If
    End If
       
    Dim j As Integer
                        
    For j = LBound(arrCardIndex()) To UBound(arrCardIndex())
        If arrCardIndex(j) = 0 Then
            arrCardIndex(j) = crdPlayingCard(i).Index
                                
            If j <= TOTAL_CARDS / 2 Then
                BitBlt picTable.hDC, 3, (j - 1) * 17 + 5, _
                                     GetCardWidth, GetCardHeight, _
                       crdPlayingCard(i).hDC, 0, 0, vbSrcCopy
            Else
                BitBlt picTable.hDC, picTable.ScaleWidth - GetCardWidth - 3, _
                                    (j Mod ((TOTAL_CARDS / 2) + 1)) * 17 + 5, _
                                     GetCardWidth, GetCardHeight, _
                       crdPlayingCard(i).hDC, 0, 0, vbSrcCopy
            End If

            picTable.Refresh
            Exit For
        End If
    Next j
    
    If crdPlayingCard(i).Rank = crdRCKing Then
        picScore.Tag = CInt(picScore.Tag) + IIf(mnuGameLevelNormal.Checked, 2, 4)
    Else
        picScore.Tag = CInt(picScore.Tag) + IIf(mnuGameLevelNormal.Checked, 4, 6)
    End If
    
    crdPlayingCard(i).Tag = "U" ' U - unwanted card
    Call ShowScore
    Call SearchHint
End Sub

Private Function GetCalcIndex(ByVal n As Integer)
    Dim col As Integer ' column
    Dim row As Integer ' row
    Dim i   As Integer ' iteration
    
    '        card index (7x7)
    '            |
    '           \ /
    '            v
    '
    '            1
    '           2 3
    '          4 5 6
    '         7 8 9 10
    '      11 12 13 14 15
    '    16 17 18 19 20 21
    '   22 23 24 25 26 27 28
    
     GetCalcIndex = 0

     For row = 1 To 7
        For col = 1 To row
            i = i + 1
            If i = n Then
                GetCalcIndex = row
                Exit Function
            End If
        Next
     Next
End Function

Private Function GetCountCard(ByVal Op As Integer) As Integer
    Dim i As Integer
    
    GetCountCard = 0
    
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        If Op = CSTOCK Then
            ' count all stock cards
            If crdPlayingCard(i).Tag = "S" Then
                GetCountCard = GetCountCard + 1
            End If
        ElseIf Op = CWASTE Then
            ' count all waste cards
            If crdPlayingCard(i).Tag = "W" Then
                GetCountCard = GetCountCard + 1
            End If
        ElseIf Op = CPYRAMID Then
            ' count all pyramid cards
            If crdPlayingCard(i).Tag = "P" Then
                GetCountCard = GetCountCard + 1
            End If
        End If
    Next i
End Function

Private Sub SaveMove()
    Dim i As Integer
    
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        Undo(i).Data = crdPlayingCard(i).Data
        Undo(i).Enabled = crdPlayingCard(i).Enabled
        Undo(i).Face = crdPlayingCard(i).Face
        Undo(i).MousePointer = crdPlayingCard(i).MousePointer
        Undo(i).Tag = crdPlayingCard(i).Tag
        Undo(i).Visible = crdPlayingCard(i).Visible
        Undo(i).arrCardIndex = arrCardIndex(i)
    Next i
    
    UndoScore = picScore.Tag
    UndoHint = mnuHint.Enabled
    mnuGameUndo.Enabled = True
End Sub

Private Sub StatGame()
    If GetCountCard(CPYRAMID) = 0 Then
        Dim i As Integer
        
        For i = crdPlayingCard.LBound To crdPlayingCard.UBound
            crdPlayingCard(i).Enabled = False
        Next i
        
        mnuHint.Enabled = False
        mnuGameDemo.Enabled = False
        picScore.Tag = CInt(picScore.Tag) + _
                       IIf(mnuGameLevelNormal.Checked, 50, 100)
        Call ShowScore
        
        MsgBox "Congratulations!", vbInformation Or vbOKOnly, "Pyramid 1.0"
    End If
End Sub

Private Function IsCardClickable(Index As Integer) As Boolean
    If mnuGameLevelNormal.Checked Then
        If (crdPlayingCard(Index).Data <> "@") Then
            IsCardClickable = True
        End If
    Else
        If (crdPlayingCard(Index).Face <> crdFCFaceDn) Then
            IsCardClickable = True
        End If
    End If
End Function

Private Function GetTopStock() As Integer
    Dim i As Integer
    
    For i = crdPlayingCard.UBound To crdPlayingCard.LBound Step -1
        If (crdPlayingCard(i).Tag = "S") And (crdPlayingCard(i).Visible) Then
            GetTopStock = i
            Exit Function
        End If
    Next i
End Function

Private Function GetTopWaste() As Integer
    Dim i As Integer
    
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        If (crdPlayingCard(i).Tag = "W") And (crdPlayingCard(i).Visible) Then
            GetTopWaste = i
            Exit Function
        End If
    Next i
End Function

Private Function GetTotalSelected() As Integer
    Dim i As Integer
    
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        If crdPlayingCard(i).Visible Then
            If crdPlayingCard(i).Selected Then
                GetTotalSelected = GetTotalSelected + 1
            End If
        End If
    Next i
End Function

Private Function GetCardWidth() As Long
    GetCardWidth = CARD_WIDTH * GetSizePercentage
End Function

Private Function GetCardHeight() As Long
    GetCardHeight = CARD_HEIGHT * GetSizePercentage
End Function

Private Function GetSizePercentage() As Single
    If mnuGameCardSizeSel(0).Checked Then
        GetSizePercentage = 0.95 ' 95%
    ElseIf mnuGameCardSizeSel(1).Checked Then
        GetSizePercentage = 1    ' 100%
    Else
        GetSizePercentage = 1.05 ' 105%
    End If
End Function

Private Function GetSelIndex(Index As Integer) As Integer
    Dim i        As Integer
    Dim rcSrc1   As RECT
    Dim rcSrc2   As RECT
    Dim rcDest   As RECT
    Dim rcTemp   As RECT
    Dim DestArea As Integer
    Dim TempArea As Integer
    Dim bval     As Boolean
    
    SetRect rcTemp, 0, 0, 0, 0
    SetRect rcSrc1, crdPlayingCard(Index).Left, crdPlayingCard(Index).Top, _
                    crdPlayingCard(Index).Width + crdPlayingCard(Index).Left, _
                    crdPlayingCard(Index).Top + crdPlayingCard(Index).Height
    
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        If crdPlayingCard(i).Visible Then
            If mnuGameLevelNormal.Checked Then
                If (crdPlayingCard(i).Data <> "@") Then bval = True
            Else
                If (crdPlayingCard(i).Face <> crdFCFaceDn) Then bval = True
            End If
            
            If bval Then
                If crdPlayingCard(i).Tag = "S" Then
                    If GetTopStock = i Then
                        bval = True
                    Else
                        bval = False
                    End If
                ElseIf crdPlayingCard(i).Tag = "W" Then
                    If GetTopWaste = i Then
                        bval = True
                    Else
                        bval = False
                    End If
                End If
            End If
            
            If bval Then
                If i <> Index Then
                    SetRect rcSrc2, crdPlayingCard(i).Left, crdPlayingCard(i).Top, _
                                    crdPlayingCard(i).Width + crdPlayingCard(i).Left, _
                                    crdPlayingCard(i).Top + crdPlayingCard(i).Height
                                    
                    If IntersectRect(rcDest, rcSrc1, rcSrc2) Then
                        TempArea = (rcTemp.Right - rcTemp.Left) * _
                                   (rcTemp.Bottom - rcTemp.Top)
                        DestArea = (rcDest.Right - rcDest.Left) * _
                                   (rcDest.Bottom - rcDest.Top)
                                  
                        If TempArea < DestArea Then
                            If ((crdPlayingCard(Index).Rank + 1) + (crdPlayingCard(i).Rank + 1)) = 13 Then
                                CopyRect rcTemp, rcDest
                                GetSelIndex = i
                            End If
                        End If
                    End If
                End If
            End If
            
            bval = False
        End If
    Next i
End Function

Private Function GetHint(ByRef Sel1 As Integer, ByRef Sel2 As Integer) As Boolean
    Dim i               As Integer
    Dim j               As Integer
    Dim IsCardClickable As Boolean
    
    Sel1 = 0
    Sel2 = 0
    
    For i = crdPlayingCard.LBound To crdPlayingCard.UBound
        If crdPlayingCard(i).Visible Then
            If (crdPlayingCard(i).MousePointer = vbCustom) And _
               (crdPlayingCard(i).Face <> crdFCFaceDn) Then
                IsCardClickable = True
                    
                If crdPlayingCard(i).Tag = "S" Then
                    If GetTopStock = i Then
                        IsCardClickable = True
                    Else
                        IsCardClickable = False
                    End If
                ElseIf crdPlayingCard(i).Tag = "W" Then
                    If GetTopWaste = i Then
                        IsCardClickable = True
                    Else
                        IsCardClickable = False
                    End If
                End If
                    
                If IsCardClickable Then
                    If crdPlayingCard(i).Rank <> crdRCKing Then
                        For j = crdPlayingCard.LBound To crdPlayingCard.UBound
                            If crdPlayingCard(j).Visible Then
                                If (crdPlayingCard(j).MousePointer = vbCustom) And _
                                   (crdPlayingCard(j).Face <> crdFCFaceDn) Then
                                    If j <> i Then
                                        IsCardClickable = True
                                            
                                        If crdPlayingCard(j).Tag = "S" Then
                                            If GetTopStock = j Then
                                                IsCardClickable = True
                                            Else
                                                IsCardClickable = False
                                            End If
                                        ElseIf crdPlayingCard(j).Tag = "W" Then
                                            If GetTopWaste = j Then
                                                IsCardClickable = True
                                            Else
                                                IsCardClickable = False
                                            End If
                                        End If
                                            
                                        If IsCardClickable Then
                                            If crdPlayingCard(i).Rank + crdPlayingCard(j).Rank + 2 = 13 Then
                                                Sel1 = i
                                                Sel2 = j
                                                GoTo HighlightCard
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next j
                    Else
                        Sel1 = i
                        Sel2 = 0
                        GoTo HighlightCard
                    End If
                End If
            End If
        End If
    Next i
    
HighlightCard:
    If (Sel1 <> 0) Or (Sel2 <> 0) Then GetHint = True
End Function

Private Sub SearchHint()
    If Not mnuStopDemo.Visible Then
        If mnuHint.Visible Then
            Dim Sel1 As Integer
            Dim Sel2 As Integer
            
            mnuHint.Enabled = GetHint(Sel1, Sel2)
        End If
    End If
End Sub


