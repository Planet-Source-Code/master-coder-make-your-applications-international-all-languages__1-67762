VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Language Demo"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Select a language"
      Height          =   3795
      Left            =   105
      TabIndex        =   1
      Top             =   180
      Width           =   5625
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   345
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   975
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   6
         Top             =   1395
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   2040
         TabIndex        =   5
         Top             =   1830
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   2040
         TabIndex        =   4
         Top             =   2250
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   2040
         TabIndex        =   3
         Top             =   2685
         Width           =   3510
      End
      Begin VB.CommandButton cmdChangeLanguage 
         Caption         =   "Change Language"
         Height          =   555
         Left            =   4020
         TabIndex        =   2
         Top             =   3180
         Width           =   1545
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   2730
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   2295
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1875
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   1020
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   555
      Left            =   4230
      TabIndex        =   0
      Top             =   4200
      Width           =   1545
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------------------
' MODULE     : Form2
' FILENAME   :
' PROJECT    : Project1
' CREATED BY : Bryan Utley
'         ON : 1/30/2007 at 19:06
' COPYRIGHT  : Copyright 2007 - All Rights Reserved
'              The World Wide Web Programmer's Consortium
'
' DESCRIPTION:
'
' COMMENTS   :
'
' WEB SITE   : http://www.thewwwpc.com
' E-MAIL     : bryan@thewwwpc.com
'
' MODIFICATION HISTORY:
'
' 1.0.0   MODIFIED ON   : 1/30/2007 at 19:06
'         MODIFIED BY   : Bryan Utley
'         MODIFICATIONS : Initial Version
' ---------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Dim TagsSetToCaption As Boolean
Dim Lang             As New Translator

Private Sub cmdChangeLanguage_Click()

    ChangeLanguage Combo1.Text

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    With Lang
        .LanguageFilePath = App.Path & "\languages"   '// Set path to language files (*.lng)
        .LookUpMethod = Phrase                        '// Set Lookup Method (Index, Key or Phrase)
        .UseFormName = "Form1"                        '// Set the form name for the lookup, allows duplicate Index/Key for each form
        .BaseLanguage = "English"                     '// Set the base Language Name
        .UsingBase = True                             '// If 'True' then use base language, 'False' use translation
    End With

    With Combo1
        .AddItem "English"
        .AddItem "Spanish"
        .AddItem "German"
        .AddItem "French"
        .ListIndex = 0
    End With
    
    MoveCaptionsToTag
    
    ChangeLanguage "English"                        '// Set Language to 'English'

End Sub

Private Sub ChangeLanguage(ByVal Language As String)
Dim ctl As Control
On Error Resume Next

    Lang.LanguageFile = Language
    
    For Each ctl In Me.Controls
    
        ctl.Caption = Lang.Say(ctl.Tag)
    
    Next ctl

End Sub

Private Sub MoveCaptionsToTag()
Dim ctl As Control
On Error Resume Next

    For Each ctl In Me.Controls
    
        ctl.Tag = ctl.Caption
    
    Next ctl

    TagsSetToCaption = True

End Sub
