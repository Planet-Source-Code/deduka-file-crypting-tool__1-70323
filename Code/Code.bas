Attribute VB_Name = "Code"
Public AppPath As String
Public DATA1() As Byte
Public DATA2() As Byte
Public ToByte As Long


Public Sub Main()
  AppPath = App.Path
  AppPath = Trim(AppPath)
  If Right(AppPath, 1) = "\" Then AppPath = left(AppPath, Len(AppPath) - 1)
  frm_main.Show
End Sub
