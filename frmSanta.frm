VERSION 5.00
Begin VB.Form frmSanta 
   Caption         =   "Secret Santa Generator"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   Icon            =   "frmSanta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   11
      Left            =   3960
      TabIndex        =   56
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   10
      Left            =   3960
      TabIndex        =   55
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   9
      Left            =   3960
      TabIndex        =   54
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   8
      Left            =   3960
      TabIndex        =   53
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   7
      Left            =   3960
      TabIndex        =   52
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   51
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   50
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   49
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   3
      Left            =   3960
      TabIndex        =   48
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   47
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   46
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtExclusion 
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   45
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   11
      Left            =   120
      TabIndex        =   43
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   11
      Left            =   2040
      TabIndex        =   42
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   11
      Left            =   6240
      TabIndex        =   41
      Top             =   5640
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   10
      Left            =   120
      TabIndex        =   40
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   10
      Left            =   2040
      TabIndex        =   39
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   10
      Left            =   6240
      TabIndex        =   38
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton cmdSaveNames 
      Caption         =   "&Save Names"
      Height          =   495
      Left            =   4560
      TabIndex        =   37
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdXML 
      Caption         =   "Import &Name File"
      Height          =   495
      Left            =   3000
      TabIndex        =   36
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import &Text"
      Height          =   495
      Left            =   1440
      TabIndex        =   35
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add More Names"
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   9
      Left            =   6240
      TabIndex        =   33
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   9
      Left            =   2040
      TabIndex        =   32
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   31
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   8
      Left            =   6240
      TabIndex        =   30
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   8
      Left            =   2040
      TabIndex        =   29
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   7
      Left            =   6240
      TabIndex        =   27
      Top             =   3720
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   7
      Left            =   2040
      TabIndex        =   26
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   6
      Left            =   6240
      TabIndex        =   24
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   6
      Left            =   2040
      TabIndex        =   23
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   5
      Left            =   6240
      TabIndex        =   21
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   5
      Left            =   2040
      TabIndex        =   20
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   18
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   4
      Left            =   2040
      TabIndex        =   17
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   3
      Left            =   6240
      TabIndex        =   15
      Top             =   1800
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   3
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   2
      Left            =   6240
      TabIndex        =   12
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   2
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   9
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   3
      Top             =   360
      Width           =   3255
   End
   Begin VB.TextBox txtPartner 
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Go"
      Height          =   495
      Left            =   6120
      TabIndex        =   0
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Exclusion (optional - prev year):"
      Height          =   255
      Left            =   3960
      TabIndex        =   44
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Email Address:"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Partner (optional):"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmSanta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSaveNames_Click()

Dim xmlString As String
xmlString = ""
xmlString = xmlString & "<PARTICIPANTS>"

Dim i As Long
For i = 0 To txtName.UBound
    If Trim(txtName(i).Text) <> "" Then
        xmlString = xmlString & "<PARTICIPANT>"
        xmlString = xmlString & "<NAME>" & txtName(i).Text & "</NAME>"
        xmlString = xmlString & "<EXCLUSION>" & txtPartner(i).Text & "</EXCLUSION>"
        xmlString = xmlString & "<EXCLUSIONPREV>" & txtExclusion(i).Text & "</EXCLUSIONPREV>"
        xmlString = xmlString & "<EMAIL>" & txtEmail(i).Text & "</EMAIL>"
        xmlString = xmlString & "</PARTICIPANT>"
    End If
Next
xmlString = xmlString & "</PARTICIPANTS>"

Dim xmlDom As DOMDocument
Set xmlDom = New DOMDocument

xmlDom.loadXML xmlString
xmlDom.Save App.Path & "\ParticipantList.xml"

'Dim fso As FileSystemObject
'Dim ts As TextStream
'
'Set fso = New FileSystemObject
'Set ts = fso.OpenTextFile(App.Path & "\ParticipantList.xml", ForWriting, True)
'ts.Write xmlString
'ts.Close
'Set ts = Nothing
'Set fso = Nothing


End Sub

Private Sub cmdXML_Click()
Dim xmlDom As DOMDocument
Set xmlDom = New DOMDocument

xmlDom.Load App.Path & "\ParticipantList.xml"
Dim i As Long
If xmlDom.parseError.errorCode = 0 Then
    For i = 0 To xmlDom.documentElement.childNodes.length - 1
        txtName(i).Text = xmlDom.documentElement.childNodes(i).selectSingleNode("NAME").nodeTypedValue
        txtPartner(i).Text = xmlDom.documentElement.childNodes(i).selectSingleNode("EXCLUSION").nodeTypedValue
        txtExclusion(i).Text = xmlDom.documentElement.childNodes(i).selectSingleNode("EXCLUSIONPREV").nodeTypedValue
        txtEmail(i).Text = xmlDom.documentElement.childNodes(i).selectSingleNode("EMAIL").nodeTypedValue
    Next i
Else
    MsgBox "There has been an error with your import. Please try to resave your information.", vbCritical, "There has been a problem."
End If
End Sub

Private Sub Command1_Click()
Dim people() As String 'use a random number
Dim people2() As String
Dim toMail() As String

Dim i As Long
Dim x As Long: x = -1
For i = 0 To txtName.UBound - 1
  If Trim(txtName(i).Text) <> "" Then
    x = x + 1
  End If
Next i

StartOver:
ReDim people(3, x)

For i = 0 To UBound(people, 2)
    people(0, i) = txtName(i).Text
    people(1, i) = txtPartner(i).Text
    people(2, i) = txtEmail(i).Text
    people(3, i) = txtExclusion(i).Text
Next i

'people(0, 0) = "Scott"  'indiv
'people(1, 0) = "Jen"    'spouse
'people(2, 0) = "scott.mcvay@westgroup.com"
'
'people(0, 1) = "Jen"
'people(1, 1) = "Scott"
'people(2, 1) = "jlmcvay@mninter.net"
'
'people(0, 2) = "Kyle"
'people(1, 2) = "Shannon"
'people(2, 2) = "kzemlicka1@earthlink.net"
'
'people(0, 3) = "Shannon"
'people(1, 3) = "Kyle"
'people(2, 3) = "szemlicka@earthlink.net"
'
'people(0, 4) = "Jay"
'people(1, 4) = "Lori"
'people(2, 4) = "jayj@insigniapops.com"
'
'people(0, 5) = "Lori"
'people(1, 5) = "Jay"
'people(2, 5) = "Lori.Jorgenson@TENNANTCO.com"
'
'people(0, 6) = "Dan"
'people(1, 6) = "Katie"
'people(2, 6) = "holleyk@earthlink.net"
'
'people(0, 7) = "Katie"
'people(1, 7) = "Dan"
'people(2, 7) = "holleyk@earthlink.net"
'
'people(0, 8) = "Cynthia"
'people(1, 8) = "Matthew"
'people(2, 8) = "johnmz@uslink.net"
'
'people(0, 9) = "Matthew"
'people(1, 9) = "Cynthia"
'people(2, 9) = "johnmz@uslink.net"

ReDim toMail(UBound(people, 2))
ReDim people2(3, x)
people2() = people()

Dim indivFirst As String
Dim indivSecond As String
Dim intFirst As Integer
Dim intSecond As Integer
'Dim i As Integer
Dim iTrys As Integer

For i = 0 To UBound(people, 2)

Randomize
indivFirst = people(0, i)
indivSecond = indivFirst

iTrys = 0
Do Until indivSecond <> indivFirst
    If iTrys > UBound(people2(), 2) + 50 Then 'big number is fine
        Debug.Print indivFirst & ": " & indivSecond
        Debug.Print "Try Again"
        Debug.Print "------------------------------"
        Debug.Print vbCrLf
        GoTo StartOver
    End If
    Randomize
    intSecond = Int((UBound(people2(), 2) + 1) * Rnd)
    indivSecond = people2(0, intSecond)
    
    'Exceptions
    If (indivSecond = people(1, i)) Or (indivSecond = people(3, i)) Then 'the exclusion
        indivSecond = indivFirst
    End If
    
    iTrys = iTrys + 1
Loop

'so, generate email :)
toMail(i) = indivFirst & ", your secret Santa recipient is: " & indivSecond
txtExclusion(i).Text = indivSecond
Debug.Print indivFirst & ": " & indivSecond

'remove person from people2
For y = intSecond To UBound(people2(), 2) - 1
    people2(0, y) = people2(0, y + 1)
    people2(1, y) = people2(1, y + 1)
    people2(2, y) = people2(2, y + 1)
    people2(3, y) = people2(3, y + 1)
Next
If UBound(people2(), 2) - 1 <> -1 Then
    ReDim Preserve people2(3, UBound(people2(), 2) - 1)
End If
'check, if you loop more than x times, start over

Next i

Dim ol As New Outlook.Application
Dim myitem As Outlook.MailItem

For i = 0 To UBound(people(), 2)

    Set myitem = ol.CreateItem(olMailItem)
    myitem.To = people(2, i)
    myitem.Subject = "Secret Santa Target - if you get your own spouse or yourself, please let me know"
    myitem.Body = toMail(i)
    
    myitem.Send
    Set myitem = Nothing
Next i

cmdSaveNames.Value = True


Set ol = Nothing
Debug.Print "--------------------------------"
End Sub

