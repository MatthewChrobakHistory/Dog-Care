VERSION 5.00
Begin VB.Form frmShop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shop"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4350
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblCost 
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblCost 
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblCost 
      Height          =   255
      Index           =   6
      Left            =   3720
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCost 
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCost 
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   6
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCost 
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCost 
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCost 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblCost 
      Height          =   255
      Index           =   9
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblPoints 
      Alignment       =   2  'Center
      Caption         =   "Points: 0"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Click on the item you wish to purchase"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Image imgItem 
      Height          =   495
      Index           =   8
      Left            =   1920
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image imgItem 
      Height          =   495
      Index           =   7
      Left            =   1320
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image imgItem 
      Height          =   495
      Index           =   6
      Left            =   3720
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgItem 
      Height          =   495
      Index           =   5
      Left            =   3120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgItem 
      Height          =   495
      Index           =   4
      Left            =   2520
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgItem 
      Height          =   495
      Index           =   3
      Left            =   1320
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgItem 
      Height          =   495
      Index           =   2
      Left            =   720
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgItem 
      Height          =   495
      Index           =   9
      Left            =   2520
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image imgItem 
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ShowShop()
Dim i As Byte

For i = 1 To 9
    imgItem(i).Picture = LoadPicture(App.Path & "\graphics\items\" & Item(i).Picture & ".bmp")
    lblCost(i).Caption = ShopItem(i).Cost
Next

End Sub

Private Sub imgItem_Click(Index As Integer)

Call BuyItem(Index, ShopItem(Index).Cost)

End Sub
