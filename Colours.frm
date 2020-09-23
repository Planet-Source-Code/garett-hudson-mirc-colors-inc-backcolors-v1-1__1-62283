VERSION 5.00
Begin VB.Form frmColors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colour Index"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   570
   ScaleWidth      =   2505
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   15
      Left            =   2205
      TabIndex        =   15
      Top             =   300
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   14
      Left            =   1890
      TabIndex        =   14
      Top             =   300
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   13
      Left            =   1575
      TabIndex        =   13
      Top             =   300
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   12
      Left            =   1260
      TabIndex        =   12
      Top             =   300
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   945
      TabIndex        =   11
      Top             =   285
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   630
      TabIndex        =   10
      Top             =   300
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   315
      TabIndex        =   9
      Top             =   300
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   2205
      TabIndex        =   7
      Top             =   0
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   1890
      TabIndex        =   6
      Top             =   0
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   1575
      TabIndex        =   5
      Top             =   0
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   1260
      TabIndex        =   4
      Top             =   0
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   945
      TabIndex        =   3
      Top             =   0
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   630
      TabIndex        =   2
      Top             =   0
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   315
      TabIndex        =   1
      Top             =   0
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   285
   End
   Begin VB.Label PicColor 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   300
      Width           =   285
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    Form1.Command.SelText = Chr(KeyAscii)
    Me.Hide
End Sub

Private Sub Form_Load()
Dim X As Integer

    Me.Left = Form1.Screen.Left + 200
    Me.Top = Form1.Screen.Top + Form1.Screen.Height - 1600

    PicColor(0).BackColor = RGB(255, 255, 255)
    PicColor(1).BackColor = RGB(0, 0, 0)
    PicColor(2).BackColor = RGB(0, 0, 123)
    PicColor(3).BackColor = RGB(0, 146, 0)
    PicColor(4).BackColor = RGB(255, 0, 0)
    PicColor(5).BackColor = RGB(123, 0, 0)
    PicColor(6).BackColor = RGB(156, 0, 156)
    PicColor(7).BackColor = RGB(255, 125, 0)
    PicColor(8).BackColor = RGB(255, 255, 0)
    PicColor(9).BackColor = RGB(0, 255, 0)
    PicColor(10).BackColor = RGB(0, 146, 148)
    PicColor(11).BackColor = RGB(0, 255, 255)
    PicColor(12).BackColor = RGB(0, 0, 255)
    PicColor(13).BackColor = RGB(255, 0, 255)
    PicColor(14).BackColor = RGB(123, 125, 123)
    PicColor(15).BackColor = RGB(214, 211, 214)
    For X = 0 To 15
'        If X = 1 Then ' BLACK Background, I want a White Foreground so it shows
'            PicColor(X).ForeColor = RGB(255, 255, 255)
'        End If
        PicColor(X).Caption = X
    Next X
    
End Sub

Private Sub PicColor_Click(Index As Integer)
       Form1.Command.SelText = Index
       Me.Hide
End Sub
