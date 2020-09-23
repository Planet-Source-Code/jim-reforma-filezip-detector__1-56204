Attribute VB_Name = "Module"
'==================================================================
'Comments!! Imagination is more important than knowledge,
'           for knowledge is limited while imagination embraces
'           the entire world.""
'==================================================================Public Comment As String, ListErrores As String
' Copyright Version (c), July 2000
' Jim Reforma [virushacker23@yahoo.com]

Public NuMar As Long
Public Accion As Boolean
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndINsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Sub AgregarError(ErrNum As String)
  Dim Mes As String

  Select Case ErrNum
    Case 0
      Mes = "Error in file no. " + Format$(NuMar) + ": Filename lenght invalid"
    Case 1
      Mes = "The quantity of files in the archive does not match the value saved in the archive"
    Case 2
      Mes = "Invalid value found"
    Case 3
      Mes = "Warning in file no. " + Format$(NuMar) + ": Filename lenght is longer than 255 characters"
  End Select
  If ListaErrores = "" Then
    ListaErrores = Mes
  Else
    ListaErrores = ListaErrores + vbCrLf + Mes
  End If
End Sub

'==================================================================
'Comments!! Imagination is more important than knowledge,
'           for knowledge is limited while imagination embraces
'           the entire world.""
'==================================================================

