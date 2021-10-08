VERSION 5.00
Begin VB.Form homeform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home"
   ClientHeight    =   5550
   ClientLeft      =   6705
   ClientTop       =   3030
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4935
   Begin VB.ListBox bestlist 
      Height          =   3375
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "PIN"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Forename"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Total Calories Burnt"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label toplabel 
      Caption         =   "Top Performers"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label homelabel 
      Caption         =   "Please use the menu bar as navigation of The Parkwood Vale Harriers Running Club program."
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   2535
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
         Caption         =   "&New Member"
      End
      Begin VB.Menu editmem 
         Caption         =   "&Edit Member"
      End
      Begin VB.Menu searchmem 
         Caption         =   "&Search Member"
      End
   End
End
Attribute VB_Name = "homeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub editmem_Click()
    editmemberform.Show 1
    Unload Me
End Sub

Private Sub Form_Load()
Dim p As Person
Dim t As training
Dim pchannel As Integer
Dim tChannel As Integer
Dim x As Integer
Dim y As Integer
Dim totalcal As Integer
Dim sort As Boolean
Dim temp As String
Dim listcount As Integer
Dim i As Integer
    pchannel = FreeFile
        Open pfile For Random As pchannel Len = plength
        x = 1
        Get pchannel, x, p
            Do While Not EOF(pchannel)
            totalcal = 0
            y = 1
            tChannel = FreeFile
                Open tFile For Random As tChannel Len = tLength
                Get tChannel, y, t
                    Do While Not EOF(tChannel)
                    If p.MemID = t.PIN Then
                        totalcal = totalcal + t.calburn
                    End If
                y = y + 1
                Get tChannel, y, t
                Loop
                Close tChannel
            bestlist.AddItem Trim(p.MemID) & "                    " & totalcal & "                                      " & Trim(p.fname)
        x = x + 1
        Get pchannel, x, p
    Loop
    Close pchannel
    Do
        sort = False
        For i = 0 To bestlist.listcount - 2
            If bestlist.List(i) < bestlist.List(i + 1) Then
            sort = True
            temp = bestlist.List(i)
            bestlist.List(i) = bestlist.List(i + 1)
            bestlist.List(i + 1) = temp
        End If
    Next i
Loop Until sort = False
listcount = bestlist.listcount
Do While listcount > 8
    bestlist.RemoveItem 8
    listcount = bestlist.listcount
Loop
End Sub

Private Sub home_Click()
    MsgBox "You are already on the Homepage", vbOKOnly
End Sub

Private Sub newmem_Click()
    homeform.Hide
    newmemberform.Show 1
End Sub

Private Sub newtraining_Click()
    homeform.Hide
    Trainingform.Show 1
End Sub

Private Sub searchmem_Click()
    homeform.Hide
    searchform.Show 1
End Sub

Private Sub team_Click()
    homeform.Hide
    Teamform.Show 1
End Sub

Private Sub compare_click()
    homeform.Hide
    compareform.Show 1
End Sub




