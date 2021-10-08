VERSION 5.00
Begin VB.Form compareform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compare"
   ClientHeight    =   6030
   ClientLeft      =   4950
   ClientTop       =   3090
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8790
   Begin VB.ComboBox teamcbo1 
      Height          =   315
      Left            =   120
      TabIndex        =   37
      Top             =   5640
      Width           =   2775
   End
   Begin VB.ComboBox teamcbo2 
      Height          =   315
      Left            =   4680
      TabIndex        =   36
      Top             =   5640
      Width           =   2775
   End
   Begin VB.TextBox timetext2 
      Height          =   495
      Left            =   6000
      TabIndex        =   35
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox timetext1 
      Height          =   495
      Left            =   1440
      TabIndex        =   34
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox distancetext2 
      Height          =   495
      Left            =   6000
      TabIndex        =   29
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox distancetext1 
      Height          =   495
      Left            =   1440
      TabIndex        =   28
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox averagespeedtext2 
      Height          =   495
      Left            =   6000
      TabIndex        =   27
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox averagespeedtext1 
      Height          =   495
      Left            =   1440
      TabIndex        =   26
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox calburnttext2 
      Height          =   495
      Left            =   6000
      TabIndex        =   23
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox calburnttext1 
      Height          =   495
      Left            =   1440
      TabIndex        =   20
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox trainingnumbertext2 
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   1440
      Width           =   2655
   End
   Begin VB.ComboBox activitycbo2 
      Height          =   315
      Left            =   6000
      TabIndex        =   14
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox IDtext2 
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox datetext2 
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox trainingnumbertext1 
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   1440
      Width           =   2655
   End
   Begin VB.ComboBox activitycbo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox IDtext1 
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox datetext1 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton searchcmd2 
      Caption         =   "Search"
      Height          =   615
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton searchcmd1 
      Caption         =   "Search"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox searchmemtext2 
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox searchmemtext1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label teamlabel 
      Caption         =   "Team"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label teamlabel2 
      Caption         =   "Team"
      Height          =   495
      Left            =   4680
      TabIndex        =   38
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label timelabel2 
      Caption         =   "Time (Mins)"
      Height          =   495
      Left            =   4680
      TabIndex        =   33
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label timelabel 
      Caption         =   "Time (Mins)"
      Height          =   495
      Left            =   120
      TabIndex        =   32
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label dislabel2 
      Caption         =   "Distance (M)"
      Height          =   495
      Left            =   4680
      TabIndex        =   31
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label dislabel 
      Caption         =   "Distance (M)"
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label speedlabel 
      Caption         =   "Average Speed (Mph)"
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label speedlabel2 
      Caption         =   "Average Speed (Mph)"
      Height          =   495
      Left            =   4680
      TabIndex        =   24
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label calburnlabel2 
      Caption         =   " Calories Burnt"
      Height          =   495
      Left            =   4680
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label calburn 
      Caption         =   " Calories Burnt"
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label numlabel2 
      Caption         =   "Training Number"
      Height          =   495
      Left            =   4680
      TabIndex        =   18
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label IDlabel2 
      Caption         =   "Member ID"
      Height          =   495
      Left            =   4680
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label actlabel2 
      Caption         =   "Activity"
      Height          =   375
      Left            =   4680
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label datelabel2 
      Caption         =   "Date"
      Height          =   495
      Left            =   4680
      TabIndex        =   15
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label numlabel 
      Caption         =   "Training Number"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label IDlabel 
      Caption         =   "Member ID"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label actlabel 
      Caption         =   "Activity"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label datelabel 
      Caption         =   "Date"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1215
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
Attribute VB_Name = "compareform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub editmem_Click()
editmemberform.Show 1
compareform.Hide
End Sub

Private Sub Form_Load()
    activitycbo1.AddItem "Running"
    activitycbo1.AddItem "Cycling"
    activitycbo1.AddItem "Freestyle, Slow"
    activitycbo1.AddItem "Freestyle, Fast"
    activitycbo1.AddItem "Backstroke"
    activitycbo1.AddItem "Breaststroke"
    activitycbo1.AddItem "Butterfly"
    activitycbo2.AddItem "Running"
    activitycbo2.AddItem "Cycling"
    activitycbo2.AddItem "Freestyle, Slow"
    activitycbo2.AddItem "Freestyle, Fast"
    activitycbo2.AddItem "Backstroke"
    activitycbo2.AddItem "Breaststroke"
    activitycbo2.AddItem "Butterfly"
    teamcbo1.AddItem "Male"
    teamcbo1.AddItem "female"
    teamcbo2.AddItem "Male"
    teamcbo2.AddItem "female"
End Sub

Private Sub home_Click()
    compareform.Hide
    homeform.Show 1
End Sub

Private Sub newmem_Click()
    compareform.Hide
    newmemberform.Show 1
End Sub

Private Sub newtraining_Click()
    compareform.Hide
    Trainingform.Show 1
End Sub

Private Sub searchcmd1_Click()
    Dim p As Person
    Dim t As training
    Dim pchannel As Integer
    Dim tChannel As Integer
    Dim x As Integer
    Dim y As Integer
    Dim foundatleastonerecord As Boolean
        foundatleastonerecord = False
    x = 1
    y = 1
    pchannel = FreeFile
    Open pfile For Random As pchannel Len = plength
    Get pchannel, y, p
    Do While Not EOF(pchannel)
        If Trim(p.fname) = Trim(searchmemtext1.Text) Then
            IDtext1.Text = p.MemID
            foundatleastonerecord = True
        End If
        y = y + 1
        Get pchannel, y, p
    Loop
    Close pchannel
    If foundatleastonerecord = False Then
        MsgBox "There is no such training in this file, please try again or add this training.", vbOKOnly
    End If
    tChannel = FreeFile
    Open tFile For Random As tChannel Len = tLength
    If foundatleastonerecord = True Then
        Do While Not EOF(tChannel)
            If Trim(t.PIN) = Trim(IDtext1.Text) Then
                activitycbo1.Text = t.Acttype
                averagespeedtext1.Text = t.speed
                distancetext1.Text = t.distance
                timetext1.Text = t.time
                calburnttext1.Text = t.calburn
                datetext1.Text = t.DTime
                trainingnumbertext1.Text = t.trainingnum
            End If
            x = x + 1
            Get tChannel, x, t
        Loop
        Close tChannel
    End If
End Sub
Private Sub searchcmd2_Click()
    Dim p As Person
    Dim t As training
    Dim pchannel As Integer
    Dim tChannel As Integer
    Dim x As Integer
    Dim y As Integer
    Dim foundatleastonerecord As Boolean
        foundatleastonerecord = False
    x = 1
    y = 1
    pchannel = FreeFile
    Open pfile For Random As pchannel Len = plength
    Get pchannel, y, p
    Do While Not EOF(pchannel)
        If Trim(p.fname) = Trim(searchmemtext2.Text) Then
            IDtext2.Text = p.MemID
            foundatleastonerecord = True
        End If
        y = y + 1
        Get pchannel, y, p
    Loop
    Close pchannel
    If foundatleastonerecord = False Then
        MsgBox "There is no such training in this file, please try again or add this training.", vbOKOnly
    End If
    tChannel = FreeFile
    Open tFile For Random As tChannel Len = tLength
    If foundatleastonerecord = True Then
        Do While Not EOF(tChannel)
            If Trim(t.PIN) = Trim(IDtext2.Text) Then
                activitycbo2.Text = t.Acttype
                Averagespeedtext.Text = t.speed
                datetext2.Text = t.DTime
                calburnttext2.Text = t.calburn
                trainingnumbertext2.Text = t.trainingnum
            End If
            x = x + 1
            Get tChannel, x, t
        Loop
        Close tChannel
    End If
End Sub

Private Sub searchmem_Click()
    compareform.Hide
    searchform.Show 1
    
End Sub

Private Sub team_Click()
    compareform.Hide
    Teamform.Show 1
End Sub

Private Sub compare_click()
    MsgBox "You are already on the Compare page", vbOKOnly
End Sub
