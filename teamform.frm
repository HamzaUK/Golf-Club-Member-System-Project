VERSION 5.00
Begin VB.Form teamform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Team"
   ClientHeight    =   3615
   ClientLeft      =   7020
   ClientTop       =   3030
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4335
   Begin VB.ListBox teamlist 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Total Minutes of Training"
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Forename"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "PIN"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
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
Attribute VB_Name = "Teamform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub editmem_Click()
Teamform.Hide
editmemberform.Show 1
End Sub

Private Sub Form_Load()
Dim p As Person
Dim t As training
Dim pchannel As Integer
Dim tChannel As Integer
Dim x As Integer
Dim y As Integer
Dim totalmins As Integer
Dim sort As Boolean
Dim temp As String
Dim listcount As Integer
Dim i As Integer
    pchannel = FreeFile
        Open pfile For Random As pchannel Len = plength
        x = 1
        Get pchannel, x, p
            Do While Not EOF(pchannel)
            totalmins = 0
            y = 1
            tChannel = FreeFile
                Open tFile For Random As tChannel Len = tLength
                Get tChannel, y, t
                    Do While Not EOF(tChannel)
                    If p.MemID = t.PIN And p.team = True Then
                        totalmins = totalmins + t.time
                            teamlist.AddItem Trim(p.MemID) & "                    " & totalmins & "                                      " & Trim(p.fname)
                    End If
                    If p.team = False Then
                            End If
                y = y + 1
                Get tChannel, y, t
                Loop
                Close tChannel
        x = x + 1
        Get pchannel, x, p
    Loop
    Close pchannel
    Do
        sort = False
        For i = 0 To teamlist.listcount - 2
            If teamlist.List(i) < teamlist.List(i + 1) Then
            sort = True
            temp = teamlist.List(i)
            teamlist.List(i) = teamlist.List(i + 1)
            teamlist.List(i + 1) = temp
        End If
    Next i
Loop Until sort = False
listcount = teamlist.listcount
Do While listcount > 8
    teamlist.RemoveItem 8
    listcount = teamlist.listcount
Loop
End Sub

Private Sub home_Click()
    Teamform.Hide
    homeform.Show 1
End Sub

Private Sub newmem_Click()
    Teamform.Hide
    newmemberform.Show 1
End Sub


Private Sub newtraining_Click()
    Teamform.Hide
    Trainingform.Show 1
End Sub

Private Sub searchmem_Click()
    Teamform.Hide
    searchform.Show 1
End Sub

Private Sub team_Click()
    MsgBox "You are already on the Team page", vbOKOnly
End Sub

Private Sub compare_click()
    Teamform.Hide
    compareform.Show 1
End Sub

