VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--> MAKE SURE YOU REGISTER THE TLB FILE <--
'
'If you down know how..google for "tlb reggie register" reggie is a free
'tlb registration program you can find on the net..adds right click option
'for it.
'
'This is a VB only implementation of IActiveScript which lets you integrate
'scripting support in your apps without the need for the MSScript control.
'
'I started working on this because I wanted to work torwards adding full
'IActiveScript support including debugging because the MS Script control
'sometimes isnt enough.
'
'Anyway, this is the first step. It may not be 100% perfectly implemented
'but it runs and works with objects you pass to it.
'
'You are free to use it in any commercial or non commercial applications
'as you so desire. A brief line "This product contains software written by
'David Zimmer" would be nice, but is not required.
'
'if you want to compile your own tlb file you need the mktyplib.exe app
'that comes with VC in the \bin folder. It may also come with the sdk or
'possibly on vb cd.
'
'
'Enjoy
'
'http://sandsprite.com


Private WithEvents sc As CActiveScript
Attribute sc.VB_VarHelpID = -1

Private Sub Form_Load()
    Set sc = New CActiveScript
    sc.AddObject "frm", Me
End Sub

Private Sub Command1_Click()
    
    Const script As String = "frm.caption = ""test""" & vbCrLf & _
                             "frm.command1.caption = ""fart""" & vbCrLf & _
                             "msgbox ""Width="" & frm.width" & vbCrLf & _
                             "frm.throwErrorTest"
                             
    
    On Error Resume Next
    
    sc.RunCode script
    
    
    
End Sub

Private Sub Command2_Click()

    Const script As String = "function test()" & vbCrLf & _
                             "   msgbox ""function test ran""" & vbCrLf & _
                             "end function" & vbCrLf & _
                             "function test2(a,b)" & vbCrLf & _
                             "   test2=a+b" & vbCrLf & _
                             "end function"
    
    
    sc.RunCode script
    sc.CallFunction "test"
    MsgBox "Return Value from test2(1,1) =" & sc.CallFunction("test2", 1, 1)
    
    
                            
End Sub

Private Sub sc_Error(description As String, ScriptSource As String, lineNumber As Long, charposition As Long)
    
    MsgBox "Script Error: " & vbCrLf & vbCrLf & _
            "Description: " & description & vbCrLf & _
            "Source: " & ScriptSource & vbCrLf & _
            "Line: " & lineNumber
    
End Sub
