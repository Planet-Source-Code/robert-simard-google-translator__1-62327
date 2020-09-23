Attribute VB_Name = "modTranslate"
Option Explicit

Global DefLang As String

'URL de traduction
Const WebURL As String = "http://translate.google.com/translate_t"

'Chaine situé juste avant le début de la traduction
Const SearchString As String = "result_box dir="  'Changer par Google le 2008-01-12

'Chaine situé à la fin de la traduction
Const EndString As String = "</"


Public Function Traduction(InputText As String, LangueTrad As String) As String
'But        = Traduire un texte en une autre langue
'
'InputText  = Texte à traduire
'
'LangueTrad = Traduire dans la langue..
'
'Modifié le = 18 septembre 2008

Dim TMPString  As String
Dim StartPos   As Long
Dim DebString  As String
Dim InitString As String

If IsConnected = False Then
   MsgBox "Vous devez être connecté à Internet pour traduire un texte !", vbInformation, "Traduction"
   Traduction = ""
   Exit Function
End If

'InputText
TMPString = GetHTMLFromURL(WebURL & "?langpair=" & LangueTrad & "&text=" & InputText)

InitString = SearchString & Chr(34) & "ltr" & Chr(34) & ">"

StartPos = InStr(1, TMPString, InitString, vbTextCompare)
If StartPos = 0 Then
   Traduction = ""
   Exit Function
End If

DebString = Right(TMPString, Len(TMPString) - (StartPos + Len(InitString) - 1))

StartPos = InStr(1, DebString, EndString, vbTextCompare)

Traduction = ReplaceHTMLString(Left(DebString, StartPos - 1))

End Function



Public Function ReplaceHTMLString(InputString) As String
Dim retValue As String

retValue = Replace(InputString, "&#39;", "'")

ReplaceHTMLString = retValue
End Function
