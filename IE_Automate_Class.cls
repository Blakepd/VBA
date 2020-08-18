'This is a Class Module

' Make sure the following WIN API for Sleep is in one of the modules you use this class with:
'Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IE_Automate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'IE_Automate Class

  ' OBJECTS
  Private classIE
  Private classWSS
  Private classSHELL
 
  ' EVENTS
  Private Sub Class_Initialize()
    Set classIE = Nothing
    Set classWSS = CreateObject("WScript.Shell")
    Set classSHELL = CreateObject("Shell.Application")
  End Sub
  
  Private Sub Class_Terminate()
    Set classIE = Nothing
    Set classWSS = Nothing
    Set classSHELL = Nothing
  End Sub
     
  ' PROPERTIES:
  
  'Lists all ie processes active in IE
  Public Property Get Avail()
    Dim process, list: list = Array()
    For Each process In classSHELL.Windows
      If IsIE(process) Then
        ReDim Preserve list(UBound(list) + 1)
        Set list(UBound(list)) = process
      End If
    Next
    Avail = list
  End Property
  
  ' The ie base - in module should look like the following:
  ' ieobj.Base.document.getElemetById("UserName").Value = "User"
  Public Property Get Base()
    If classIE Is Nothing Then
      Call noInitMsg
    End If
    Set Base = classIE
  End Property
  
  ' URL property used with naviate subs below
  Public Property Get url()
    If classIE Is Nothing Then
      Call noInitMsg
      url = ""
    Else
      url = classIE.LocationURL
    End If
  End Property
  
  ' Allows for the changing of ie title is desired
  Public Property Get Title()
    If classIE Is Nothing Then
      Call noInitMsg
      Title = ""
    Else
      Title = classIE.LocationName
    End If
  End Property
  
  ' METHODS:
  ' Closes current ie window
  Public Sub Closer()
    If classIE Is Nothing Then
      Call noInitMsg
    Else
      classIE.Quit
    End If
  End Sub
  
  ' Closes all ie windows
  Public Sub CloseAll()
    Dim window
    For Each window In classSHELL.Windows
      If IsIE(window) Then: window.Quit
    Next
  End Sub
  
  ' Close an ie tab by passing ie title ex:
  ' ie.Close_Per_Title("Google")
  Public Sub Close_Per_Title(win_Title As String)
    Dim window
    For Each window In classSHELL.Windows
      If IsIE(window) Then
        'Debug.Print window.Document.Title
        If win_Title = "" Then
            If window.document.Title = win_Title Then
                window.Quit
            End If
        ElseIf window.document.Title Like "*" & win_Title & "*" Then
          window.Quit
        End If
      End If
    Next
  End Sub
  
  ' Show ie window the same as ieobj.visisble = True
  Public Sub Show()
    Call autoInit
    classIE.Visible = True
  End Sub
  
  ' Hide ie window the same as ieobj.visible = False
  Public Sub Hide()
    Call autoInit
    classIE.Visible = False
  End Sub
  
  ' Center ie window on screen (current monitor screen)
  Public Sub Center()
    Call WaitForLoad
    On Error Resume Next
    With classIE.document.ParentWindow.screen
      classIE.Left = (.Width - classIE.Width) / 2
      classIE.Top = (.Height - classIE.Height) / 2
    End With
    If err.Number = 505 Then
       Navigate "about:blank"
    End If
  End Sub
  
  Public Sub WaitOnElement(error)
    Sleep 500
  End Sub
  
  ' Initiate new ie window (regular)
  Public Sub Navigate(url)
    Call autoInit
    classIE.Navigate2 url
  End Sub
  
  ' Initiate new ie tab in current ie window (if no ie window then new will open)
  Public Sub NavigateTab(url)
    Call autoInit
    classIE.Navigate2 url, 2048
    Sleep (750)
  End Sub
  
  
  Public Sub NavigateBgTab(url)
    Call autoInit
    classIE.Navigate2 url, 4096
  End Sub
  
  ' Wait for webpage to load
  Public Sub WaitForLoad()
    If classIE Is Nothing Then
      Call noInitMsg
    Else
      While (classIE.Busy) And Not (classIE.ReadyState = 4): Sleep (400): Wend
    End If
  End Sub
  
  ' Wait for webpage to load - interactive pages load on 3
  Public Sub WaitForLoad2()
    If classIE Is Nothing Then
      Call noInitMsg
    Else
      While (classIE.Busy) And Not (classIE.ReadyState = 3): Sleep (400): Wend
    End If
  End Sub

  ' Wait for element load
  Public Sub DeepWaitForLoad(elem)
    While Not (elem.ReadyState = "complete"): Sleep (400): Wend
  End Sub
    
  Public Sub ReBase(ie)
    Init (ie)
  End Sub
  
  ' Set ie to a different (still current) ie window using url
  Public Sub RePoint(url)
    Dim window
    For Each window In classSHELL.Windows
      If IsIE(window) And (LCase(window.LocationURL) = LCase(url)) Then
        Set classIE = window
        Exit Sub
      End If
    Next
    Call ErrorOut(strURL, "Internet Explorer")
  End Sub
  
  ' Seting classIE to most recent tab opened in ie use this after navigate ex:
  ' ie.navigateTab url
  ' ie.Latest_Tab
  Public Sub Latest_tab()
    Dim window
    For Each window In classSHELL.Windows
      If IsIE(window) Then
        Set classIE = window
      End If
    Next
    Call autoInit
  End Sub
  
  ' Sets classie to first ie window
  Public Sub Latest()
    Dim window
    For Each window In classSHELL.Windows
      If IsIE(window) Then
        Set classIE = window
        Exit Sub
      End If
    Next
    Call autoInit
  End Sub
    
  Private Function IsIE(obj)
    IsIE = CBool(Right(LCase(obj.FullName), 12) = "iexplore.exe")
  End Function
  
  ' Returns the first element that matches the given CSS selector
  Public Function Query(squery)
    Call WaitForLoad
    On Error Resume Next
    Dim element
    Set element = classIE.document.querySelector(squery)
    If err.Number = 0 Then
      Set Query = element
    Else
      Call ErrorOut(squery, classIE.LocationURL)
    End If
  End Function
  
  ' Returns all elements that match the given CSS selector
  Public Function QueryAll(squery)
    Call WaitForLoad
    On Error Resume Next
    Dim elements
    Set elements = classIE.document.querySelectorAll(squery)
    If err.Number = 0 Then
      Set QueryAll = elements
    Else
      Call ErrorOut(squery, classIE.LocationURL)
    End If
  End Function
  
  ' Similar to Query Function - Added Error Handling
  Public Function Deeper(squery)
    Call WaitForLoad
    On Error Resume Next
    Dim element
    Set element = classIE.document.querySelector(squery)
    If err.Number = 0 Then
      Call DeepWaitForLoad(element)
      Set Deeper = element.contentDocument
      If err.Number = -2147024891 Then
        MsgBox "ERROR: Deeper(""" & squery & """)" & vbLf & vbLf & "Same Origin Policy Violated.", vbCritical, "IE_Automate: " & err.Description
        WScript.Quit
      End If
    Else
      Call ErrorOut(squery, classIE.LocationURL)
    End If
  End Function
  
  ' ERROR HANDLER
  Private Sub ErrorOut(item, at)
    MsgBox _
      "CANNOT FIND [ " & item & " ]" & vbLf & _
      "AT [ " & at & " ]", vbCritical
    WScript.Quit
  End Sub
  
  Private Sub autoInit()
    If classIE Is Nothing Then
      Set classIE = CreateObject("InternetExplorer.Application")
    End If
  End Sub
  
  Private Sub noInitMsg()
    classWSS.PopUp "Not yet initialized.", 1, "EasyIEAutomate"
  End Sub



