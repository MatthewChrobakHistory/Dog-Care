VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dog Care"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton cmdShop 
         Caption         =   "$"
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   2520
         Width           =   495
      End
      Begin VB.CheckBox chkSellMode 
         Caption         =   "Sell"
         Height          =   195
         Left            =   1200
         TabIndex        =   10
         Top             =   2760
         Width           =   1215
      End
      Begin VB.PictureBox picFace 
         Height          =   1335
         Left            =   600
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   85
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar PBHunger 
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar PBThirst 
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar PBEnergy 
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblPoints 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Energy"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Thirst"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Hunger"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Image imgInventoryItem 
         Height          =   495
         Index           =   7
         Left            =   1320
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image imgInventoryItem 
         Height          =   495
         Index           =   6
         Left            =   720
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image imgInventoryItem 
         Height          =   495
         Index           =   5
         Left            =   120
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image imgInventoryItem 
         Height          =   495
         Index           =   4
         Left            =   1920
         Top             =   3120
         Width           =   495
      End
      Begin VB.Image imgInventoryItem 
         Height          =   495
         Index           =   3
         Left            =   1320
         Top             =   3120
         Width           =   495
      End
      Begin VB.Image imgInventoryItem 
         Height          =   495
         Index           =   2
         Left            =   720
         Top             =   3120
         Width           =   495
      End
      Begin VB.Image imgInventoryItem 
         Height          =   495
         Index           =   8
         Left            =   1920
         Top             =   3720
         Width           =   495
      End
      Begin VB.Image imgInventoryItem 
         Height          =   495
         Index           =   1
         Left            =   120
         Top             =   3120
         Width           =   495
      End
   End
   Begin VB.PictureBox picMap 
      Height          =   4095
      Left            =   120
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   285
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblInfo 
      Caption         =   "Game Loaded"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4320
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShop_Click()

frmShop.Show
Call frmShop.ShowShop

End Sub

Private Sub Form_Load()

frmMain.Visible = True
Call LoadGame
Call GameLoop

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub

Private Sub imgInventoryItem_Click(Index As Integer)

If chkSellMode.Value = 1 Then
    Call SellItem(Index, Inventory(Index).ActionValue, Inventory(Index).ItemType)
    Exit Sub
End If

If Dog.InAction = True Then Exit Sub

Call UseItem(Index)

End Sub

Private Sub tmerLoop_Timer()

Call GameLoop

End Sub
