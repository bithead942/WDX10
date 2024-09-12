Attribute VB_Name = "Module1"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Main()
   'Dim s As String
   Dim i As Integer
   Dim Args, Arg() As String
   Dim comX10 As MSComm
   
   On Error GoTo HandleError
   Args = Command
   If Len(Args) = 7 Then
      Arg() = Split(Args, " ", 2, vbTextCompare)
   
      Set comX10 = CreateObject("MSCommLib.MSComm")
      comX10.CommPort = 1
      comX10.InputLen = 4
      comX10.Settings = "2400,N,8,2"
      comX10.PortOpen = True
      Sleep (50)
      comX10.Output = Arg(0) & Chr(13)
      For i = 1 To 2
         Sleep (575)
         comX10.Output = Arg(1) & Chr(13)
      Next i
      comX10.PortOpen = False
   End If

HandleError:
   
End Sub
