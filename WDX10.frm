VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form WDX10 
   Caption         =   "WDX10"
   ClientHeight    =   870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   ScaleHeight     =   870
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   2280
      Top             =   480
   End
   Begin MSCommLib.MSComm comX10 
      Left            =   1200
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "WDX10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Main()
   Dim i As Integer
   Dim s As String
   Dim Args, Arg1, Arg2 As String
   
   Args = Command
   If Len(Args) = 7 Then
      Arg1 = Args(0) & Args(1) & Args(2)
      Arg2 = Args(4) & Args(5) & Args(6)
   
      comX10.CommPort = 2
      comX10.InputLen = 4
      comX10.Settings = "2400,N,8,2"
      comX10.PortOpen = True
      s = comX10.Input
      comX10.Output = "G02" & Chr(13)
      s = comX10.Input
      comX10.Output = "GON" & Chr(13)
      s = comX10.Input
   End If
   
End Sub


Private Sub Form_Load()

End Sub
