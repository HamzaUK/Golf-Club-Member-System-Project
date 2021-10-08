Attribute VB_Name = "Module1"
Option Explicit

Type Person
    sname As String * 15
    fname As String * 30
    gender As String * 6
    MemID As Integer
    age As Integer
    ad1 As String * 30
    town As String * 30
    postcode As String * 10
    email As String * 20
    weightp As Single
    heightp As Single
    info As String * 150
    pnumber As String * 10
    team As Boolean
End Type

Global pfile As String
Global plength As Long
Global pactive As Integer

Type training
    trainingnum As Integer
    PIN As Integer
    Acttype As String * 10
    DTime As Date
    distance As Single
    time As Single
    calburn As Single
    speed As String * 50
    running As Integer
    cycling As Integer
    swimming As Integer
End Type

Global tFile As String
Global tLength As Long
Global tActive As Integer

