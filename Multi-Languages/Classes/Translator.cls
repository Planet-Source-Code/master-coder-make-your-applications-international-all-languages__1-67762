VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Translator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------------------
' MODULE     : Translator
' FILENAME   : Translator.cls
' PROJECT    :
' CREATED BY : Bryan Utley
'         ON : 1/30/2007 at 18:37
' COPYRIGHT  : Copyright 2007 - All Rights Reserved
'              The World Wide Web Programmer's Consortium
'
' DESCRIPTION: A class to allow support of multiple language files within an application
'
' COMMENTS   : Dependencies--> NONE
'
' WEB SITE   : http://www.thewwwpc.com
' E-MAIL     : bryan@thewwwpc.com
'
' MODIFICATION HISTORY:
'
' 1.0.0   MODIFIED ON   : 1/30/2007 at 18:37
'         MODIFIED BY   : Bryan Utley
'         MODIFICATIONS : Initial Version
' ---------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Private Type type_LanguageArrayStructure
    BasePhrase  As String
    Form        As String
    Index       As Long
    Key         As String
    TransPhrase As String
End Type

Enum enum_LookupMethod
    Index = 1
    Key = 2
    Phrase = 3
End Enum
#If False Then
Private Index, Key, Phrase
#End If

Dim m_LookupMethod     As Integer
Dim m_LanguageFile     As String
Dim m_LanguageFilePath As String
Dim m_UseFormName      As String
Dim m_BaseLanguage     As String
Dim m_UsingBase        As Boolean

Dim LangFiles          As New Collection
Dim Lang()             As type_LanguageArrayStructure

'---------------------------------------------------------------------------------------------------------------------------------
' PROCEDURE   : Class Module ReadLanguageFile
' PURPOSE     : Read language file and parse into 'Lang' array
' DESCRIPTION :
' RETURN VALUE:
' CREATED BY  : Bryan Utley
'         ON  : 1/30/2007 at 18:37
' MODULE      : Translator
' PROJECT     :
' COMMENTS    : Update to use Raise.Error in Error Handling
' ---------------------------------------------------------------------------------------------------------------------------------
Private Sub ReadLanguageFile()

Dim strQuoteMark As String * 1
Dim strFormName  As String
Dim lngFileNum   As Long
Dim lngCounter   As Long
Dim strInput  As String

    On Error GoTo Err_Handler

    lngCounter = 0
    strQuoteMark = Chr$(34)
    lngFileNum = FreeFile()

    Open LanguageFilePath & LanguageFile For Input As #lngFileNum

    While Not EOF(lngFileNum)

        Line Input #lngFileNum, strInput

        strInput = Trim(strInput)

        If (Len(strInput) > 0) And (Left(strInput, 1) <> "'") Then

            If Left(strInput, 1) = "[" Then
                strFormName = StringBetween(strInput, "[]")
            End If

            lngCounter = lngCounter + 1
            ReDim Preserve Lang(1 To lngCounter)

            '// Store FormName to array
            Lang(lngCounter).Form = strFormName

            '// Store Index to array
            Lang(lngCounter).Index = CLng(Val(strInput))

            '// See if we have an Index or Key being used
            If InStr(1, strInput, ":", vbTextCompare) > 0 Then
                Lang(lngCounter).Key = StringBetween(strInput, "{}")
            End If

            '// See if both Base and Translated phrases exist
            If InStr(1, strInput, "|", vbTextCompare) > 0 Then
                Lang(lngCounter).BasePhrase = StringBetween(Left(strInput, InStr(1, strInput, "|") - 1), strQuoteMark)
                Lang(lngCounter).TransPhrase = StringBetween(Mid(strInput, InStr(1, strInput, "|") + 1), strQuoteMark)
            End If

        End If
    Wend

Err_Resume:
    Close #lngFileNum
    On Error GoTo 0

Exit Sub

Err_Handler:
    MsgBox "Language File '" & LanguageFile & "' not found!", vbCritical, "File not found"
    GoTo Err_Resume

End Sub

'---------------------------------------------------------------------------------------------------------------------------------
' PROCEDURE   : Class Module Say
' PURPOSE     : Return translated phrase
' DESCRIPTION :
' RETURN VALUE:
' CREATED BY  : Bryan Utley
'         ON  : 1/30/2007 at 18:38
' MODULE      : Translator
' PROJECT     :
' COMMENTS    : Update to use a single For/Next
' ---------------------------------------------------------------------------------------------------------------------------------
Public Function Say(ByVal SearchMethod As Variant) As String

Dim i As Long

    Say = ""

    If LookUpMethod = Phrase Then

        For i = 1 To UBound(Lang)
            If UCase(SearchMethod) = UCase(Lang(i).BasePhrase) Then
                If (UseFormName > "") And (LCase(UseFormName) = LCase(Lang(i).Form)) Then
                    If UsingBase Then
                        Say = Lang(i).BasePhrase
                    Else
                        Say = Lang(i).TransPhrase
                    End If
                Else
                    If UsingBase Then
                        Say = Lang(i).BasePhrase
                    Else
                        Say = Lang(i).TransPhrase
                    End If
                End If
                Exit Function
            End If
        Next

    ElseIf LookUpMethod = Index Then

        For i = 1 To UBound(Lang)
            If UCase(SearchMethod) = UCase(Lang(i).Index) Then
                If (UseFormName > "") And (LCase(UseFormName) = LCase(Lang(i).Form)) Then
                    If UsingBase Then
                        Say = Lang(i).BasePhrase
                    Else
                        Say = Lang(i).TransPhrase
                    End If
                Else
                    If UsingBase Then
                        Say = Lang(i).BasePhrase
                    Else
                        Say = Lang(i).TransPhrase
                    End If
                End If
                Exit Function
            End If
        Next

    ElseIf LookUpMethod = Key Then

        For i = 1 To UBound(Lang)
            If UCase(SearchMethod) = UCase(Lang(i).Key) Then
                If (UseFormName > "") And (LCase(UseFormName) = LCase(Lang(i).Form)) Then
                    If UsingBase Then
                        Say = Lang(i).BasePhrase
                    Else
                        Say = Lang(i).TransPhrase
                    End If
                Else
                    If UsingBase Then
                        Say = Lang(i).BasePhrase
                    Else
                        Say = Lang(i).TransPhrase
                    End If
                End If
                Exit Function
            End If
        Next

    End If

End Function

'---------------------------------------------------------------------------------------------------------------------------------
' PROCEDURE   : Class Module StringBetween
' PURPOSE     : Return the contents of a string between specified seperator
' DESCRIPTION :
' RETURN VALUE:
' CREATED BY  : Bryan Utley
'         ON  : 1/30/2007 at 18:38
' MODULE      : Translator
' PROJECT     :
' COMMENTS    :
' ---------------------------------------------------------------------------------------------------------------------------------
Private Function StringBetween(ByVal SourceString As String, ByVal Seperators As String) As String

Dim strLeftSeperator  As String * 1
Dim strRightSeperator As String * 1
Dim lngLeftPosition   As Long
Dim lngRightPosition  As Long

    strLeftSeperator = Left(Seperators, 1)
    strRightSeperator = Right(Seperators, 1)

    lngLeftPosition = InStr(1, SourceString, strLeftSeperator, vbTextCompare)

    If lngLeftPosition > 0 Then
        lngRightPosition = InStrRev(SourceString, strRightSeperator, -1, vbTextCompare)
    End If

    If lngRightPosition > lngLeftPosition Then
        StringBetween = Mid(SourceString, lngLeftPosition + 1, lngRightPosition - lngLeftPosition - 1)
    End If

End Function

'---------------------------------------------------------------------------------------------------------------------------------
' Properties
'---------------------------------------------------------------------------------------------------------------------------------

Public Property Let UsingBase(ByVal value As Boolean)

    m_UsingBase = value

End Property

Public Property Get UsingBase() As Boolean

    UsingBase = m_UsingBase

End Property

Public Property Let BaseLanguage(ByVal Language As String)

    m_BaseLanguage = Language

End Property

Public Property Get BaseLanguage() As String

    BaseLanguage = m_BaseLanguage

End Property

Public Property Let LookUpMethod(ByVal Method As enum_LookupMethod)

    m_LookupMethod = Method

End Property

Public Property Get LookUpMethod() As enum_LookupMethod

    LookUpMethod = m_LookupMethod

End Property

Public Property Let UseFormName(ByVal strFormName As String)

    m_UseFormName = strFormName

End Property

Public Property Get UseFormName() As String

    UseFormName = m_UseFormName

End Property

Public Property Let LanguageFilePath(ByVal FilePath As String)

Dim DirName As String

    On Error Resume Next

        m_LanguageFilePath = IIf(Right(FilePath, 1) = "\", FilePath, FilePath & "\")

        LangFiles.Clear
        DirName = Dir(m_LanguageFilePath & "*.lng")

        Do While DirName <> ""
            If DirName <> "." And DirName <> ".." Then
                LangFiles.Add DirName, DirName
            End If
            DirName = Dir
        Loop

End Property

Public Property Get LanguageFilePath() As String

    LanguageFilePath = m_LanguageFilePath

End Property

Public Property Let LanguageFile(ByVal Filename As String)

    If LCase(Filename) <> LCase(BaseLanguage) Then
        m_LanguageFile = Filename & ".lng"
        UsingBase = False
        ReadLanguageFile
    Else
        UsingBase = True
        m_LanguageFile = LangFiles.Item(1)
        ReadLanguageFile
    End If

End Property

Public Property Get LanguageFile() As String

    LanguageFile = m_LanguageFile

End Property
