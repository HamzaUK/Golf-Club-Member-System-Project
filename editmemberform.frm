VERSION 5.00
Begin VB.Form editmemberform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "editmemberform"
   ClientHeight    =   6390
   ClientLeft      =   3300
   ClientTop       =   3030
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11910
   Begin VB.ComboBox cboteam 
      Height          =   315
      Left            =   5880
      TabIndex        =   31
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox confirmpnumbertext 
      Height          =   615
      Left            =   9120
      TabIndex        =   16
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox pnumbertext 
      Height          =   615
      Left            =   9120
      TabIndex        =   15
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox infotext 
      Height          =   1815
      Left            =   9120
      TabIndex        =   14
      Top             =   4440
      Width           =   2655
   End
   Begin VB.TextBox ad1text 
      Height          =   615
      Left            =   3000
      TabIndex        =   13
      Top             =   4440
      Width           =   2655
   End
   Begin VB.ComboBox cbogender 
      Height          =   315
      Left            =   2880
      TabIndex        =   12
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Searchcmd 
      Caption         =   "Search"
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton saveeditcmd 
      Caption         =   "Save Edit"
      Height          =   615
      Left            =   5640
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox searchtext 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox fnametext 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox snametext 
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox memidtext 
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox pcodetext 
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox towntext 
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox agetext 
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox htext 
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox wtext 
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox emailtext 
      Height          =   615
      Left            =   9120
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "+44"
      Height          =   375
      Left            =   8760
      TabIndex        =   34
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "+44"
      Height          =   375
      Left            =   8760
      TabIndex        =   33
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "Team"
      Height          =   375
      Left            =   5880
      TabIndex        =   32
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label agelabel 
      Caption         =   "Age"
      Height          =   375
      Left            =   5880
      TabIndex        =   30
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label adlabel 
      Caption         =   "Address"
      Height          =   375
      Left            =   3000
      TabIndex        =   29
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label townlabel 
      Caption         =   "Town"
      Height          =   255
      Left            =   3000
      TabIndex        =   28
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label pnumlabel 
      Caption         =   "Phone Number"
      Height          =   375
      Left            =   9120
      TabIndex        =   27
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label emaillabel 
      Caption         =   "Email"
      Height          =   375
      Left            =   9120
      TabIndex        =   26
      Top             =   960
      Width           =   855
   End
   Begin VB.Label genderlabel 
      Caption         =   "Gender"
      Height          =   375
      Left            =   3000
      TabIndex        =   25
      Top             =   960
      Width           =   855
   End
   Begin VB.Label plabel 
      Caption         =   "Postcode"
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "PIN"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label snamelabel 
      Caption         =   "Surname"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Forename(s)"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   975
   End
   Begin VB.Label medlabel 
      Caption         =   "Medical Information"
      Height          =   375
      Left            =   9120
      TabIndex        =   20
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label heightlabel 
      Caption         =   "Height (M)"
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label weightlabel 
      Caption         =   "Weight (Kg)"
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label conpnumberlabel 
      Caption         =   "Confirm Phone Number"
      Height          =   375
      Left            =   9120
      TabIndex        =   17
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Menu home 
      Caption         =   "Home"
   End
   Begin VB.Menu training 
      Caption         =   "&Training"
      Begin VB.Menu gap1 
         Caption         =   "-"
      End
      Begin VB.Menu newtraining 
         Caption         =   "New Training"
      End
      Begin VB.Menu compare 
         Caption         =   "Compare Training"
      End
   End
   Begin VB.Menu team 
      Caption         =   "Team"
   End
   Begin VB.Menu member 
      Caption         =   "&Member"
      Begin VB.Menu gap2 
         Caption         =   "-"
      End
      Begin VB.Menu newmem 
         Caption         =   "New Member"
      End
      Begin VB.Menu editmem 
         Caption         =   "Edit Member"
      End
      Begin VB.Menu searchmem 
         Caption         =   "Search Member"
      End
   End
End
Attribute VB_Name = "editmemberform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
cboteam.AddItem "Yes"
cboteam.AddItem "No"
cbogender.AddItem "Male"
cbogender.AddItem "Female"
End Sub

Private Sub editmem_Click()
MsgBox "You are already on the Edit Members page", vbOKOnly
End Sub

Private Sub home_Click()
editmemberform.Hide
homeform.Show 1
End Sub

Private Sub newmem_Click()
editmemberform.Hide
newmemberform.Show 1
End Sub

Private Sub newtraining_Click()
editmemberform.Hide
Trainingform.Show 1
End Sub

Private Sub saveeditcmd_Click()
    Dim tempteam As Boolean
    Dim tempgender As Boolean
    Dim valemail As Boolean
        valemail = False
    Dim valmember As Boolean
        valmember = True
    Dim pchannel As Integer
    Dim p As Person
        pchannel = FreeFile
    If snametext.Text = "" Then
        valmember = False
    ElseIf fnametext.Text = "" Then
        valmember = False
    ElseIf agetext.Text = "" Then
        valmember = False
    ElseIf cboteam.Text = "" Then
        valmember = False
    ElseIf ad1text.Text = "" Then
        valmember = False
    ElseIf cbogender.Text = "" Then
        valmember = False
    ElseIf towntext.Text = "" Then
        valmember = False
    ElseIf wtext.Text = "" Then
        valmember = False
    ElseIf htext.Text = "" Then
        valmember = False
    ElseIf pcodetext.Text = "" Then
        valmember = False
    ElseIf emailtext.Text = "" Then
        valmember = False
    ElseIf pnumbertext.Text = "" Then
        valmember = False
    ElseIf Len(pnumbertext.Text) <> 11 Then
        valmember = False
    End If
    If pnumbertext.Text = confirmpnumbertext.Text Then
        valmember = True
    Else
        valmember = False
    End If
    If valmember = True Then
    Dim pcount As Integer
    pcount = 1
    Open pfile For Random As pchannel Len = plength
    Get pchannel, pcount, p
    Do While Not EOF(pchannel)
        If p.MemID = memidtext.Text Then
            p.sname = snametext.Text
            p.fname = fnametext.Text
            p.team = tempteam
            p.MemID = memidtext.Text
            p.email = emailtext.Text
            p.age = agetext.Text
            p.ad1 = ad1text.Text
            p.pnumber = pnumbertext.Text
            p.pnumber = confirmpnumbertext
            p.town = towntext.Text
            p.weightp = wtext.Text
            p.heightp = htext.Text
            p.postcode = pcodetext.Text
            p.gender = tempgender
            p.info = infotext.Text
            Put pchannel, pcount, p
        End If
        pcount = pcount + 1
        Get pchannel, pcount, p
    Loop
    Close pchannel
    snametext.Text = ""
    fnametext.Text = ""
    memidtext.Text = ""
    emailtext.Text = ""
    agetext.Text = ""
    ad1text.Text = ""
    pnumbertext.Text = ""
    confirmpnumbertext = ""
    towntext.Text = ""
    wtext.Text = ""
    htext.Text = ""
    pcodetext.Text = ""
    infotext.Text = ""
    confirmpnumbertext.Text = ""
        MsgBox "Member record was successfully altered!", vbOKOnly
    Else
        MsgBox "You have left some fields blank, please go back and fill them in correctly", vbOKOnly
    End If
End Sub

Private Sub Searchcmd_Click()
Dim p As Person
Dim pchannel As Integer
Dim x As Integer
Dim foundatleastonerecord As Boolean
foundatleastonerecord = False
x = 1
pchannel = FreeFile
Open pfile For Random As pchannel Len = plength
Get pchannel, x, p
Do While Not EOF(pchannel)
    If Trim(p.fname) = Trim(searchtext.Text) Then
        If p.gender = "False" Then
            cbogender = "Male"
                    cbogender = "Female"
        End If
        If p.team = "True" Then
            cboteam = "Yes"
                Else
                    cboteam = "No"
        End If
        snametext.Text = p.sname
        fnametext.Text = p.fname
        memidtext.Text = p.MemID
        emailtext.Text = p.email
        agetext.Text = p.age
        ad1text.Text = p.ad1
        pnumbertext.Text = p.pnumber
        confirmpnumbertext.Text = p.pnumber
        towntext.Text = p.town
        wtext.Text = p.weightp
        htext.Text = p.heightp
        pcodetext.Text = p.postcode
        foundatleastonerecord = True
    End If
    x = x + 1
    Get pchannel, x, p
Loop
Close pchannel
If foundatleastonerecord = True Then
    MsgBox "The member has been found.", vbOKOnly
ElseIf foundatleastonerecord = False Then
    MsgBox "There is no such person in this file, please try again or add this member.", vbOKOnly
End If
End Sub

Private Sub searchmem_Click()
editmemberform.Hide
searchform.Show 1
End Sub

Private Sub team_Click()
editmemberform.Hide
Teamform.Show 1
End Sub

Private Sub compare_click()
editmemberform.Hide
compareform.Show 1
End Sub



