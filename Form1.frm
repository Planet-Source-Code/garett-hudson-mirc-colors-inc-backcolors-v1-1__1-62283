VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "mIRC Colors for your IRC App v1.01 (v0rtexx@yahoo.com)"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   7920
   Begin VB.TextBox Command 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   7815
   End
   Begin RichTextLib.RichTextBox Screen 
      Height          =   4935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8705
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' IRC Colors in your IRC App. V1.0 (C) V0RTEXX (v0rtexx@yahoo.com)
'
' Quick Background
' ~~~~~~~~~~~~~~~~
' Okay. I have searched long and hard for a source code that will
' highlight text in a Richtextbox, just like mIRC does. My search
' revealed nothing that I could use. This source (built from scratch)
' allows you to do so. I have included a "Color picker" so you can
' see which colors you are going to choose (Like mIRC), before you
' choose them.
'
' Misc Notes
' ~~~~~~~~~~~~~~~~
' I am sure there are going to be some minor bugs. If so, please email
' me about them and I will try to get them fixed as soon as possible,
' and release the new version on www.planetsourcecode.com/vb/
'
' Also email me for questions or comments and I will be happy to respond.
'
' I just want this released so you all can have background colours
' in your IRC apps! Peace!

' Disclaimer (hah)
' ~~~~~~~~~~~~~~
' You may use any of this code or all of it, I don't care. Give me a
' shout-out, or not.. I don't care :) Happy programming all!

'
' Versions
' ~~~~~~~~

'1.01 [Aug23,05]
'     - Got of the dependency of having to use Fixedsys Font that I
'       included in v1.00 (for unknown reasons). We now use the standard
'       fixedsys that is already build into VB's Font listing
'     - Created New caption of my program for display. This shows the
'       options I can do, and the color picker that the other pic did not.
'     - Fixed major bug in ManageRTF. \highlight# and \cf# were showing
'       up twice in the RTFCodes confusing the colors at times. This fix
'       clears up conflicting backgrounds and foregrounds and the spaces
'       in between.
'     - Easy to understand RTF-Formatting tutorial written in the module
'     - Fixed spelling mistakes, typos. Duhh
'     - Fixed major bug in ManageRTF. "(B)(U)Text(U)Text2" will now display
'       properly. Before "Text2" would not even show up. Had to add a space
'       at the end of RTFCodes due to the fact I trimmed it before editing.
'     - Due to popular demand, (And easier beta testing) I made it easier
'       to add chars(Bold,Underline,Color) marks in mid-sentance, instead of
'       only at the end. Still not perfect though :)
'
'1.00 [Aug23,05]
'     - Initial Release - (Pre-Mature? You decide) :)
'
' NOTES: (B)=Bold Char Code (U)=Underline Char Code (K)=Color Char Code


Private Sub Command_KeyPress(KeyAscii As Integer)
' To all who've complained... Here's a new fix for ya... cheers :)
'
' Again I must stress, its the Color Function that needs your input, not
' silly Keypress Subs!
Dim OldSelStart As Integer
    If KeyAscii = 11 Then ' Control-K
        OldSelStart = Command.SelStart
        Command = Mid(Command, 1, Command.SelStart) & Chr(3) & Mid(Command, Command.SelStart + 1)
        Command.SelStart = OldSelStart + 1
        frmColors.Show
    ElseIf KeyAscii = 21 Then
        OldSelStart = Command.SelStart
        Command = Mid(Command, 1, Command.SelStart) & Chr(31) & Mid(Command, Command.SelStart + 1)
        Command.SelStart = OldSelStart + 1
    ElseIf KeyAscii = 2 Then
        OldSelStart = Command.SelStart
        Command = Mid(Command, 1, Command.SelStart) & Chr(2) & Mid(Command, Command.SelStart + 1)
        Command.SelStart = OldSelStart + 1
    End If
    
    If KeyAscii = 13 Then ' If the <ENTER> key is pressed
        Call DoColors(Me, Command)
        Command = ""
    End If
End Sub

Private Sub Form_Load()
Dim X As Integer

' Start the main screen. You can play with this if you like
    
    Call DoColors(Me, Chr(2) & "Syntax: " & Chr(2) & Chr(31) & "(CTRL-K)##,##" & Chr(31) & " " & Chr(31) & "(CTRL-B)BOLD" & Chr(31) & " " & Chr(31) & "(CTRL-U)UNDERLINE")
    Screen.SelText = vbCrLf & vbCrLf
    Call DoColors(Me, Chr(3) & "11,12Just Like mIRC")
    
    For X = 0 To 15
        Call DoColors(Me, Chr(3) & X & "Color" & X & vbTab & vbTab & Chr(3) & "01," & X & "BackColor Too!")
    Next X
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Screen.Width = Me.ScaleWidth
    Screen.Height = Me.ScaleHeight - Command.Height
    Command.Width = Me.ScaleWidth
    Command.Top = Me.ScaleHeight - Command.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Screen_Click()
    Screen.SelStart = Len(Screen)
    Command.SetFocus
End Sub
