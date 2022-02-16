VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cold Fusion LUA Decompiler"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblDesc 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label lblAuthors 
      Alignment       =   1  'Right Justify
      Caption         =   "-By 4E534B (4E534B@gmail.com) "
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   570
      Width           =   4695
   End
   Begin VB.Label lblVers 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Caption = "Cold Fusion LUA Decompiler"
    
    lblVers.Caption = "CFLuaDC" & _
        " v" & _
        App.Major & _
        "." & _
        App.Minor & _
        "." & _
        App.Revision & _
        IIf(App.Comments = "-", "", " " & App.Comments)
        
    lblDesc.Caption = "To decompile a LUA, call " & vbCrLf & _
        "      CFLuaDC <file name>" & vbCrLf & vbCrLf & _
        "For example: CFLuaDC Vgr_BattleCruiser.ship will decompile the file into Vgr_BattleCruiser_DC.ship" _
        & vbCrLf & vbCrLf & "If you encounter any bugs, please let me know about it."
End Sub

