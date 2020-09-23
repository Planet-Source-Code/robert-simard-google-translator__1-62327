VERSION 5.00
Begin VB.Form frmTranslate 
   Caption         =   "Text translator from Robert Simard (Logipro) logipro2000@hotmail.com"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   300
      Left            =   7200
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "Translate"
      Height          =   300
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox cboLang 
      Height          =   315
      ItemData        =   "frmTranslate.frx":0000
      Left            =   600
      List            =   "frmTranslate.frx":0002
      TabIndex        =   2
      Top             =   120
      Width           =   3495
   End
   Begin VB.TextBox txtOutputText 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3000
      Width           =   8175
   End
   Begin VB.TextBox txtInputText 
      Appearance      =   0  'Flat
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   8175
   End
   Begin VB.Label Label3 
      Caption         =   "Text for translatation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Translate text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   495
   End
End
Attribute VB_Name = "frmTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TradLang As String


Private Sub cmdCopy_Click()
  Clipboard.SetText Me.txtOutputText
End Sub

Private Sub cmdTranslate_Click()
Dim MyLang As Variant

'// Array of Language
MyLang = Array("en|es", "en|fr", "en|it", _
               "en|pt", "en|ja", "en|ko", _
               "en|zh-CN", "de|en", "de|fr", _
               "es|en", "fr|en", "fr|de", _
               "it|en", "pt|en", "ja|en", "zh-CN|en")

 '// Selected Language
 TradLang = MyLang(Me.cboLang.ListIndex)
 Me!txtOutputText = Traduction(Me.txtInputText, TradLang)
End Sub

Private Sub Form_Load()

'// Fill combo
With Me.cboLang
 .AddItem "English to Spanish", 0
 .AddItem "English to French", 1
 .AddItem "English to Italian", 2
 .AddItem "English to Portuguese", 3
 .AddItem "English to Japanese BETA", 4
 .AddItem "English to Korean BETA", 5
 .AddItem "English to Chinese(Simplified) BETA", 6
 .AddItem "German to English", 7
 .AddItem "German to French", 8
 .AddItem "Spanish to English", 9
 .AddItem "French to English", 10
 .AddItem "French to German", 11
 .AddItem "Italian to English", 12
 .AddItem "Portuguese to English", 13
 .AddItem "Japanese to English BETA", 14
 .AddItem "Chinese(Simplified) to English BETA", 15
 .ListIndex = 1
End With
  
End Sub
