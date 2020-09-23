Attribute VB_Name = "modOS"
' Written by José Luis Farías.
' Chile 1446 - Salto - Uruguay - CP 50.000
' JoseloFarias[at]adinet.com.uy
' ¡¡¡Vamo' arriba Uruguay, carajo!!!
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
' ¡PLEASE!, if you use this Code sendme your Name and Country
' And if you like, emailme a program copy (source code if better)
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*

Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public ID As String * 40, Version As String * 5, Build As String * 10
Public Sub GetVersion()
     Dim OSinfo As OSVERSIONINFO
     Dim RetValue As Integer
     OSinfo.dwOSVersionInfoSize = 148
     OSinfo.szCSDVersion = Space$(128)
     RetValue = GetVersionExA(OSinfo)
     With OSinfo
     Select Case .dwPlatformId
      Case 1
          Select Case .dwMinorVersion
              Case 0
                  ID = "Microsoft Windows 95"
              Case 10
                  If .dwBuildNumber >= 2183 Then
                      ID = "Microsoft Windows 98 SE"
                  Else
                      ID = "Microsoft Windows 98"
                  End If
              Case 90
                  ID = "Microsoft Windows Millennium Edition"
          End Select
      Case 2
          Select Case .dwMajorVersion
              Case 3
                  ID = "Microsoft Windows NT 3.51"
              Case 4
                  ID = "Microsoft Windows NT 4.0"
              Case 5
                  If .dwMinorVersion = 0 Then
                      ID = "Microsoft Windows 2000"
                  ElseIf .dwMinorVersion = 1 Then
                      ID = "Microsoft Windows XP"
                  Else
                      ID = "Microsoft Windows Sever 2003"
                  End If
          End Select
      Case Else
         ID = "Failed!!"
    End Select
     Version = .dwMajorVersion & "." & .dwMinorVersion
     Build = .dwBuildNumber
     End With
End Sub
