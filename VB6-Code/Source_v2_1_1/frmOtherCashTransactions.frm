VERSION 5.00
Begin VB.Form frmOtherCashTransactions 
   Caption         =   "Other Cash Transactions"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Select type of Transaction:"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.OptionButton Option4 
         Caption         =   "Other Expenses"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Stock Purchase"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Normal  Stock Sold"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pawn Stock Sold"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmOtherCashTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
