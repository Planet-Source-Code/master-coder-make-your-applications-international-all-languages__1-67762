VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDemoCaptionTrans 
      Caption         =   "Demonstrate using caption text to lookup translation."
      Height          =   645
      Left            =   135
      TabIndex        =   14
      Top             =   4095
      Width           =   2220
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Command1"
      Height          =   555
      Left            =   4230
      TabIndex        =   13
      Top             =   4200
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3795
      Left            =   105
      TabIndex        =   0
      Top             =   180
      Width           =   5625
      Begin VB.CommandButton cmdChangeLanguage 
         Caption         =   "Command1"
         Height          =   555
         Left            =   4020
         TabIndex        =   12
         Top             =   3150
         Width           =   1545
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   4
         Left            =   2040
         TabIndex        =   11
         Top             =   2685
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   3
         Left            =   2040
         TabIndex        =   10
         Top             =   2250
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   2
         Left            =   2040
         TabIndex        =   9
         Top             =   1830
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Top             =   1395
         Width           =   3510
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   975
         Width           =   3510
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Text            =   "Combo1"
         Top             =   345
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         TabIndex        =   5
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         TabIndex        =   4
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         TabIndex        =   3
         Top             =   1875
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         TabIndex        =   2
         Top             =   2295
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         TabIndex        =   1
         Top             =   2730
         Width           =   675
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------------------
' MODULE     : Form1
' FILENAME   : Form1.frm
' PROJECT    : Project1
' CREATED BY : Bryan Utley
'         ON : 1/30/2007 at 18:49
' COPYRIGHT  : Copyright 2007 - All Rights Reserved
'              The World Wide Web Programmer's Consortium
'
' DESCRIPTION: Test applicatin for the Translator class
'
' COMMENTS   :
'
' WEB SITE   : http://www.thewwwpc.com
' E-MAIL     : bryan@thewwwpc.com
'
' MODIFICATION HISTORY:
'
' 1.0.0   MODIFIED ON   : 1/30/2007 at 18:49
'         MODIFIED BY   : Bryan Utley
'         MODIFICATIONS : Initial Version
' ---------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Dim Lang As New Translator

Private Sub cmdChangeLanguage_Click()

    ChangeLanguage Combo1.Text

End Sub

Private Sub cmdDemoCaptionTrans_Click()
    Form2.Show vbModal
End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    With Lang
        .LanguageFilePath = App.Path & "\languages"   '// Set path to language files (*.lng)
        .LookUpMethod = Index                         '// Set Lookup Method (Index, Key or Phrase)
        .UseFormName = Me.Name                        '// Set the form name for the lookup, allows duplicate Index/Key for each form
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

    ChangeLanguage "English"                        '// Set Language to 'English'

End Sub

Private Sub ChangeLanguage(ByVal Language As String)
On Error Resume Next

    With Lang
        
        .LanguageFile = Language

        .LookUpMethod = Key
        
            Me.Caption = .Say("Form Title")

        .LookUpMethod = Index
        
            Frame1.Caption = .Say(10)

        .LookUpMethod = Key
        
            Label1.Caption = .Say("label1")
            Label2.Caption = .Say("label2")
            Label3.Caption = .Say("label3")
            Label4.Caption = .Say("label4")
            Label5.Caption = .Say("label5")

        .LookUpMethod = Phrase

            cmdChangeLanguage.Caption = .Say("Change Language")
        
        .LookUpMethod = Index
        
            cmdExit.Caption = .Say(1000)
    
    End With

End Sub
