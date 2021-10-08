VERSION 5.00
Begin VB.Form startupform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parkwood Vale Harriers Running Club"
   ClientHeight    =   2925
   ClientLeft      =   6390
   ClientTop       =   2730
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5340
   Begin VB.CommandButton startcmd 
      Caption         =   "Start"
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.CommandButton exitcmd 
      Caption         =   "Exit"
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label welcomelabel 
      Caption         =   "Welcome to Parkwood Vale Harriers Running Club. Please press start to open the program, or press exit to close this program."
      Height          =   855
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "I liek eggs"
      Height          =   375
      Left            =   13680
      TabIndex        =   0
      Top             =   9360
      Width           =   855
   End
End
Attribute VB_Name = "startupform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub exitcmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim pchannel As Integer
    Dim p As Person
    pfile = App.Path + "\people_details.dat"
    plength = Len(p)
    pchannel = FreeFile
    Open pfile For Random As pchannel Len = plength
    pactive = FileLen(pfile) / plength
    Close pchannel
    
    Dim tChannel As Integer
    Dim t As training
    tFile = App.Path + "\training_details.dat"
    tLength = Len(t)
    tChannel = FreeFile
    Open tFile For Random As tChannel Len = tLength
    tActive = FileLen(tFile) / tLength
    Close tChannel
End Sub

Private Sub startcmd_Click()
    startupform.Hide
    homeform.Show 1
End Sub



