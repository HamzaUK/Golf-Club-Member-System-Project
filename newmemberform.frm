VERSION 5.00
Begin VB.Form newmemberform 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5655
   ClientLeft      =   3300
   ClientTop       =   3180
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11895
   Begin VB.ComboBox cboteam 
      Height          =   315
      Left            =   5880
      TabIndex        =   29
      Top             =   3720
      Width           =   2775
   End
   Begin VB.TextBox memidtext 
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox confirmpnumbertext 
      Height          =   615
      Left            =   9120
      TabIndex        =   13
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox infotext 
      Height          =   1815
      Left            =   9120
      TabIndex        =   12
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox fnametext 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox snametext 
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox pcodetext 
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox emailtext 
      Height          =   615
      Left            =   9120
      TabIndex        =   8
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox pnumbertext 
      Height          =   615
      Left            =   9120
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox htext 
      Height          =   615
      Left            =   5880
      TabIndex        =   6
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox wtext 
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox towntext 
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox ad1text 
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox agetext 
      Height          =   615
      Left            =   5880
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.ComboBox cbogender 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton createcmd 
      Caption         =   "Create"
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "+44"
      Height          =   375
      Left            =   8760
      TabIndex        =   32
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "+44"
      Height          =   375
      Left            =   8760
      TabIndex        =   31
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "Team"
      Height          =   375
      Left            =   5880
      TabIndex        =   30
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label agelabel 
      Caption         =   "Age"
      Height          =   375
      Left            =   5880
      TabIndex        =   28
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label adlabel 
      Caption         =   "Address"
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label townlabel 
      Caption         =   "Town"
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label pnumlabel 
      Caption         =   "Phone Number"
      Height          =   375
      Left            =   9120
      TabIndex        =   25
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label emaillabel 
      Caption         =   "Email"
      Height          =   375
      Left            =   9120
      TabIndex        =   24
      Top             =   240
      Width           =   855
   End
   Begin VB.Label genderlabel 
      Caption         =   "Gender"
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   240
      Width           =   855
   End
   Begin VB.Label plabel 
      Caption         =   "Postcode"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label PINlabel 
      Caption         =   "PIN"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label snamelabel 
      Caption         =   "Surname"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label fnamelabel 
      Caption         =   "Forename(s)"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin VB.Label medlabel 
      Caption         =   "Medical Information"
      Height          =   255
      Left            =   9120
      TabIndex        =   18
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label heightlabel 
      Caption         =   "Height (M)"
      Height          =   375
      Left            =   5880
      TabIndex        =   17
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label weightlabel 
      Caption         =   "Weight (Kg)"
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label conpnumberlabel 
      Caption         =   "Confirm Phone Number"
      Height          =   255
      Left            =   9120
      TabIndex        =   14
      Top             =   2400
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
Attribute VB_Name = "newmemberform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub createcmd_Click()
    Dim tempteam As Boolean
    Dim tempgender As Boolean
    Dim valemail As Boolean
        valemail = False
    Dim valmember As Boolean
        valmember = True
    Dim pchannel As Integer
    Dim p As Person
        pchannel = FreeFile
    If cbogender.Text = "Male" Then
        tempgender = True
    Else
        tempgender = False
    End If
    If cboteam.Text = "Yes" Then
        tempteam = True
    Else
        tempteam = False
    End If
    If cbogender.Text = "Male" Then
        tempgender = True
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
    ElseIf Len(pnumbertext.Text) <> 10 Then
        valmember = False
    End If
    If pnumbertext.Text = confirmpnumbertext.Text Then
        valmember = True
    Else
        valmember = False
    End If
    If valmember = True Then
        Open pfile For Random As pchannel Len = plength
            p.sname = snametext.Text
            p.fname = fnametext.Text
            p.team = tempteam
            p.MemID = memidtext.Text
            p.email = emailtext.Text
            p.age = agetext.Text
            p.ad1 = ad1text.Text
            p.pnumber = pnumbertext.Text
            p.town = towntext.Text
            p.weightp = wtext.Text
            p.heightp = htext.Text
            p.postcode = pcodetext.Text
            p.gender = tempgender
            p.info = infotext.Text
        pactive = pactive + 1
    MsgBox "Member sucessfully added! " & Trim(p.MemID) & " is your pin, vbOKOnly"
        Put pchannel, pactive, p
        Close pchannel
    snametext.Text = ""
    fnametext.Text = ""
    memidtext.Text = ""
    emailtext.Text = ""
    agetext.Text = ""
    ad1text.Text = ""
    pnumbertext.Text = ""
    towntext.Text = ""
    wtext.Text = ""
    htext.Text = ""
    pcodetext.Text = ""
    infotext.Text = ""
    confirmpnumbertext.Text = ""
        Dim PIN As Integer
        Randomize
        PIN = Int((9999 - 1000 + 1) * Rnd + 1000)
        memidtext = PIN
    Else
        MsgBox "You have either left blank filed(s) or incorrectly filled out some fields, please go back and correct it.", vbOKOnly
    End If
    End If
End Sub

Private Sub Form_Load()
Dim PIN As Integer
Randomize
    PIN = Int((9999 - 1000 + 1) * Rnd + 1000)
    memidtext = PIN
    cbogender.AddItem "Male"
    cbogender.AddItem "Female"
    cboteam.AddItem "Yes"
    cboteam.AddItem "No"
End Sub

Private Sub editmem_Click()
    newmemberform.Hide
    editmemberform.Show 1
End Sub

Private Sub home_Click()
    newmemberform.Hide
    homeform.Show 1
End Sub

Private Sub newmem_Click()
    MsgBox "You are already on the New Members page", vbOKOnly
End Sub

Private Sub newtraining_Click()
    newmemberform.Hide
    Trainingform.Show 1
End Sub

Private Sub searchmem_Click()
    newmemberform.Hide
    searchform.Show 1
End Sub

Private Sub team_Click()
    newmemberform.Hide
    Teamform.Show 1
End Sub

Private Sub compare_click()
    newmemberform.Hide
    compareform.Show 1
End Sub
