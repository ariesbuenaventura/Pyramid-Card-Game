VERSION 5.00
Begin VB.Form frmDeck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Card Back"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   Icon            =   "frmDeck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3495
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1260
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin prjPyramid.Card crdDeck 
      Height          =   1455
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2566
      Face            =   1
   End
   Begin prjPyramid.Card crdDeck 
      Height          =   1455
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2566
      Deck            =   1
      Face            =   1
   End
   Begin prjPyramid.Card crdDeck 
      Height          =   1455
      Index           =   2
      Left            =   2340
      TabIndex        =   2
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2566
      Deck            =   2
      Face            =   1
   End
End
Attribute VB_Name = "frmDeck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub crdDeck_Click(Index As Integer)
    Dim i As Integer
        
    For i = crdDeck.LBound To crdDeck.UBound
        If i <> Index Then
            crdDeck(i).Selected = False
        Else
            crdDeck(i).Selected = True
        End If
    Next i
    
    With frmMain
        For i = .crdPlayingCard.LBound To .crdPlayingCard.UBound
            .crdPlayingCard(i).Deck = Index
        Next i
    End With
End Sub

Private Sub Form_Load()
    With frmMain
        crdDeck(.crdPlayingCard(.crdPlayingCard.LBound).Deck).Selected = True
    End With
End Sub
