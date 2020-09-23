Attribute VB_Name = "Module1"
Global RTFCodes As String  ' This will hold any rtf codes we need for correct formatting
                           ' eg: {\b\ul\cf#\highlight#} ... etc

' This should be the exact color table that mIRC uses!
Public Const TblColor = "{\rtf1{\colortbl\red255\green255\blue255;\red0\green0\blue0;\red0\green0\blue123;\red0\green146\blue0;\red255\green0\blue0;\red123\green0\blue0;\red156\green0\blue156;\red255\green125\blue0;\red255\green255\blue0;\red0\green255\blue0;\red0\green146\blue148;\red0\green255\blue255;\red0\green0\blue255;\red255\green0\blue255;\red123\green125\blue123;\red214\green211\blue214;}"

' ----------------------------------
' Short Tutorial on RTF Formatting.
' ----------------------------------
'
' Okay, to properly understand this program, you need to know a little about RTF
' formatting. I must admit that until 2 days ago I didn't know about it either so
' bare with me if you are more advanced. I use only the basic in RTF; mainly Bold,
' Underline, Foreground Color and Highlighting. These are represented as:
'   "\b"            Bold
'   "\ul"           Underline
'   "\highlight#"   Highlight (#=Color Index)
'   "\cf#"          Forground Color (#=Color Index)
'
' The # for the color is all setup for you in the Color Table (Represented here as
' TblColor; Look above) It may look like a lot of technical junk but really, it only
' sets up the colors that we use. {\rtf1} indicates we are using rtf formating to the
' Richtextbox. {\colortbl} Sets up the color table and the red\green\blue values
' calculate the colour index (Similar to the RGB Function in VB). Each Index is separated
' by a ";". So my colour Index above states 0 as "\red255\green255\blue255;" 1 as
' "\red0\green0\blue0\". Simple as that. Just call up the Index via \highlight# or \cf#
' and you get the correct color. Now I set this color table up to match mIRC exactly. You
' may change this table to any color you wish. Just keep the format EXACTLY the same or
' you will have some problems.
'
' Very Simple Example: {\rtf1{\colortbl\red255\green255\blue255;}{\cf0 Red Text}{\par}}
' Type this into a Richtextbox without a vbCrLf at the end and "Red Text" will show in Red.
' Note that we have only one color in our color table. We can only use \cf0 or \highlight0
' since that is the only color we have declared.

' {\par} is used instead of vbCrLf or Chr(13) or whatever. This will call a new line to
' the RTF Format and display your colors that you have called upon. You may still use
' vbCrLf but if there is any RTF Formatting involved you will not see it and the text
' will be jumbled up

' ---------- End of Tutorial (Hope you found it useful)


' hah. No more vbCrlF :) NOTE: You must use this if you are using formatting of any kind
Public Const CrLf = "{\par}}"

Function ManageRTF(sAddDel As String) '  Add or DEL setting from the RTFCodes variable
'Dim TempString As String
Dim X As Integer
    RTFCodes = Trim(RTFCodes)
    
    ' Clear any previous foreground codes, if we need to replace it
    If (sAddDel Like "\cf#") Or (sAddDel Like "\cf##") Then
        If InStr(RTFCodes, "\cf") > 0 Then ' We have it set. Lets delete it now
            For X = 2 To 15 ' We cant include "1" because it will conflict
                            ' with 10, 11, 12, 13, 14, 15 as "\cf1...0...1..etc
                If InStr(RTFCodes, "\cf" & X) Then
                    RTFCodes = Replace(RTFCodes, "\cf" & X, "")
                    Exit For
                End If
            Next X
        
            ' If 2-15 doesn't get rid of it, 0 or 1 will!
            If InStr(RTFCodes, "\cf0") > 0 Then RTFCodes = Replace(RTFCodes, "\cf0", "")
            If InStr(RTFCodes, "\cf1") > 0 Then RTFCodes = Replace(RTFCodes, "\cf1", "")
        End If
    ElseIf (sAddDel Like "\highlight#") Or (sAddDel Like "\highlight##") Then
        If InStr(RTFCodes, "\highlight") > 0 Then ' We have it set. Lets delete it now
            For X = 2 To 15 ' We cant include "1" because it will conflict
                            ' with 10, 11, 12, 13, 14, 15 as "\cf1...0...1..etc
                If InStr(RTFCodes, "\highlight" & X) Then
                    RTFCodes = Replace(RTFCodes, "\highlight" & X, "")
                    Exit For
                End If
            Next X
        
            ' If 2-15 doesn't get rid of it, 0 or 1 will!
            If InStr(RTFCodes, "\highlight0") > 0 Then RTFCodes = Replace(RTFCodes, "\highlight0", "")
            If InStr(RTFCodes, "\highlight1") > 0 Then RTFCodes = Replace(RTFCodes, "\highlight1", "")
        End If
    End If
    
    
    If InStr(RTFCodes, sAddDel) > 0 Then ' It Exists. Remove code
        RTFCodes = Replace(RTFCodes, sAddDel, "")
        
        ' Need to add the space back due to the fact I trimmed it above. Otherwise text
        ' will not display
        
'        If RTFCodes = "\b" Then RTFCodes = "\b "
'        If RTFCodes = "\ul" Then RTFCodes = "\ul "
        RTFCodes = RTFCodes & " "
    Else ' Doesn't exist. Add code
        RTFCodes = RTFCodes & sAddDel & " "
    End If
End Function


' Okay, here's what you have been waiting for. To Implement in your own program
' you need to display it using this function. I added sFormName due to the fact that
' many like to use MDI's for their IRC Client layout. This way you can specify the form
' you want to display on (keeping in mind that you have all display control's(Richtextboxes)
' with the same name). For this example I used "Screen". Anyways. Have fun!

Function DoColors(sFormName As Form, sText As String)
Dim X As Integer
Dim OutPutLine As String
Dim ForeColor As String
Dim BackColor As String
    
    ' Check if incomming text has codes (Bold, Underline or Color)
    ' If NOT, then output text via normal methods
    If (InStr(sText, Chr(2)) = 0) And (InStr(sText, Chr(3)) = 0) And (InStr(sText, Chr(31)) = 0) Then
        sFormName.Screen.SelText = sText & vbCrLf
        Exit Function
    End If
       
    ' Reset RTFCodes and OutPutLine since this is a new line.
    RTFCodes = ""
    OutPutLine = TblColor & "{"
    
    For X = 1 To Len(sText)
        If Mid(sText, X, 1) = Chr(2) Then ' Bold Character Code
            ManageRTF ("\b")
            OutPutLine = OutPutLine & "}{" & RTFCodes
            ' Check if there is a space after Bold Symbol, if so.. We need to add an extra space to the OutPutLine
            'If Mid(sText, X + 1, 1) = " " Then OutPutLine = OutPutLine & " "
        ElseIf Mid(sText, X, 1) = Chr(31) Then ' Underline Character Code
            ManageRTF ("\ul")
            OutPutLine = OutPutLine & "}{" & RTFCodes
            ' Check if there is a space after Underline Symbol, if so.. We need to add an extra space to the OutPutLine
            'If Mid(sText, X + 1, 1) = " " Then OutPutLine = OutPutLine & " "
        ElseIf Mid(sText, X, 1) = Chr(3) Then ' Color Character Code
            If Mid(sText, X + 1, 2) Like "##" Then ' Double Forecolor selection
                
                If Mid(sText, X + 1, 1) = "0" Then   ' Eg forecolor selection 04, or 05 etc...
                    ForeColor = Mid(sText, X + 2, 1) ' We want get rid of the "0"
                Else
                    ForeColor = Mid(sText, X + 1, 2)
                End If
                
                ManageRTF ("\cf" & ForeColor)
                X = X + 2
                
                ' We now have the forecolor selection from the ##. Lets now see
                ' if any background wants to be set
                
                If Mid(sText, X + 1, 1) = "," Then ' We could have background if numbers
                                                   ' follow this point
                    ' Check for number
                    If Mid(sText, X + 2, 2) Like "##" Then ' We have double digit background selection
                        ' Set background (TODO GET RID OF LEFT(X,1) = "0"
                        BackColor = Mid(sText, X + 2, 2)
                        ManageRTF ("\highlight" & BackColor)
                        OutPutLine = OutPutLine & "}{" & RTFCodes
                        X = X + 3
                    ElseIf Mid(sText, X + 2, 1) Like "#" Then ' We have single digit background selection
                        BackColor = Mid(sText, X + 2, 1)
                        ManageRTF ("\highlight" & BackColor)
                        OutPutLine = OutPutLine & "}{" & RTFCodes
                        X = X + 2
                    Else
                        ' No number following "," Continue with just the forecolor
                        OutPutLine = OutPutLine & "}{" & RTFCodes
                    End If
                Else
                    ' No Backgrounds, so continue with just the forecolor
                    OutPutLine = OutPutLine & "}{" & RTFCodes
                End If
            ElseIf Mid(sText, X + 1, 1) Like "#" Then ' Single Forecolor Selection
                ForeColor = Mid(sText, X + 1, 1)
                ManageRTF ("\cf" & ForeColor)
                X = X + 1
                If Mid(sText, X + 1, 1) = "," Then ' We may have background. Lets check
                    If Mid(sText, X + 2, 2) Like "##" Then ' We have double digit background selection
                        BackColor = Mid(sText, X + 2, 2)
                        ManageRTF ("\highlight" & BackColor)
                        OutPutLine = OutPutLine & "}{" & RTFCodes
                        X = X + 3
                    ElseIf Mid(sText, X + 2, 1) Like "#" Then ' We have single digit background selection
                        BackColor = Mid(sText, X + 2, 1)
                        ManageRTF ("\highlight" & BackColor)
                        OutPutLine = OutPutLine & "}{" & RTFCodes
                        X = X + 2
                    Else
                        ' We didn't find any numbers after the "," so just use the forecolor
                        OutPutLine = OutPutLine & "}{" & RTFCodes
                    End If
                Else
                    ' No Background, so continue with forecolor
                    OutPutLine = OutPutLine & "}{" & RTFCodes
                End If
            End If
        Else
            ' No Color/Bold/Underline codes to process, get the single character.
            OutPutLine = OutPutLine & Mid(sText, X, 1)
        End If
            
    Next X
    
    ' Finnish off the OutPutLine so it's ready for display
    OutPutLine = OutPutLine & "}" & CrLf
    
    ' Display it
    sFormName.Screen.SelText = OutPutLine
End Function

' We're Done
