VERSION 5.00
Begin VB.Form searchform 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5925
   ClientLeft      =   3300
   ClientTop       =   3030
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11895
   Begin VB.ComboBox cboteam 
      Height          =   315
      Left            =   5880
      TabIndex        =   30
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox infotext 
      Height          =   1815
      Left            =   9120
      TabIndex        =   26
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox fnametext 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox snametext 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox memidtext 
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox pcodetext 
      Height          =   615
      Left            =   3000
      TabIndex        =   12
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox emailtext 
      Height          =   615
      Left            =   9120
      TabIndex        =   11
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox pnumbertext 
      Height          =   615
      Left            =   9120
      TabIndex        =   10
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox htext 
      Height          =   615
      Left            =   5880
      TabIndex        =   9
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox wtext 
      Height          =   615
      Left            =   5880
      TabIndex        =   8
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox towntext 
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox ad1text 
      Height          =   615
      Left            =   3000
      TabIndex        =   6
      Top             =   4440
      Width           =   2655
   End
   Begin VB.TextBox agetext 
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.ComboBox cbogender 
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Searchcmd 
      Caption         =   "Search"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox searchtext 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton editmemcmd 
      Caption         =   "Edit Member"
      Height          =   495
      Left            =   10440
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton newmemcmd 
      Caption         =   "New Member"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "+44"
      Height          =   375
      Left            =   8760
      TabIndex        =   32
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "Team"
      Height          =   375
      Left            =   5880
      TabIndex        =   31
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label weightlabel 
      Caption         =   "Weight (Kg)"
      Height          =   255
      Left            =   5880
      TabIndex        =   29
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label heightlabel 
      Caption         =   "Height (M)"
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label medlabel 
      Caption         =   "Medical Information"
      Height          =   255
      Left            =   9120
      TabIndex        =   27
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label fnamelabel 
      Caption         =   "Forename(s)"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   960
      Width           =   975
   End
   Begin VB.Label snamelabel 
      Caption         =   "Surname"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label PINlabel 
      Caption         =   "PIN"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label plabel 
      Caption         =   "Postcode"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label genderlabel 
      Caption         =   "Gender"
      Height          =   375
      Left            =   3000
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.Label emaillabel 
      Caption         =   "Email"
      Height          =   255
      Left            =   9120
      TabIndex        =   20
      Top             =   960
      Width           =   495
   End
   Begin VB.Label pnumlabel 
      Caption         =   "Phone Number"
      Height          =   255
      Left            =   9120
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label townlabel 
      Caption         =   "Town"
      Height          =   255
      Left            =   3000
      TabIndex        =   18
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label adlabel 
      Caption         =   "Address"
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label agelabel 
      Caption         =   "Age"
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   960
      Width           =   375
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
Attribute VB_Name = "searchform"
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

Private Sub Newmemcmd_Click()
searchform.Hide
newmemberform.Show 1
End Sub

Private Sub Editmemcmd_Click()
searchform.Hide
editmemberform.Show 1
End Sub

Private Sub Searchcmd_Click()
Dim p As Person
Dim pchannel As Integer
Dim x As Integer
Dim foundatleastonerecord As Boolean
    foundatleastonerecord = True
x = 1
pchannel = FreeFile
Open pfile For Random As pchannel Len = plength
Get pchannel, x, p
Do While Not EOF(pchannel)
    If Trim(p.fname) = Trim(searchtext.Text) Then
        snametext.Text = p.sname
        fnametext.Text = p.fname
        memidtext.Text = p.MemID
        emailtext.Text = p.email
        cboteam.Text = p.team
        agetext.Text = p.age
        ad1text.Text = p.ad1
        cbogender.Text = p.gender
        pnumbertext.Text = p.pnumber
        towntext.Text = p.town
        wtext.Text = p.weightp
        htext.Text = p.heightp
        pcodetext.Text = p.postcode
        foundatleastonerecord = True
    MsgBox "Member found!", vbOKOnly
    End If
    x = x + 1
    Get pchannel, x, p
Loop
Close pchannel
If foundatleastonerecord = False Then
    MsgBox "There is no such person in this file, please try again or add this member.", vbOKOnly
End If
End Sub

Private Sub editmem_Click()
searchform.Hide
editmemberform.Show 1
End Sub

Private Sub home_Click()
searchform.Hide
homeform.Show 1
End Sub

Private Sub newmem_Click()
searchform.Hide
newmemberform.Show 1
End Sub

Private Sub newtraining_Click()
searchform.Hide
Trainingform.Show 1
End Sub

Private Sub searchmem_Click()
MsgBox "You are already on the Search page", vbOKOnly
End Sub

Private Sub team_Click()
searchform.Hide
Teamform.Show 1
End Sub

Private Sub compare_click()
searchform.Hide
compareform.Show 1
End Sub
