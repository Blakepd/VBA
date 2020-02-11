Attribute VB_Name = "IE_Mod"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub IE_Automate()
Dim ie As IE_Automate

Set ie = New IE_Automate

Const URL = "https://www.google.com/"

ie.Latest
ie.NavigateTab URL
ie.Latest_tab
ie.Show

ie.Base.document.getElementsByClassName("gLFyf gsfi")(0).Value = "Current Weather"

ie.Base.document.getElementsByClassName("gNO89b")(0).Click

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'A Few Other Common Options

ie.Base.document.getElementbyName("q").Value = "Current Weather" ' Name
ie.Base.document.getElementbyName("q").OnChange

ie.Base.document.getElementbyName("btnK").Click ' Name Click

ie.Base.document.getElementbyId("Test").Value = "Test" ' Classic ID
ie.Base.document.getElementbyId("Test").OnChange ' Keep field on change (if you change a fields value and it disappears use this)

Dim obj As Object
Set obj = ie.Base.document.getElementsByTagName("iframe")(0).contentDocument ' Set obj = ie frame

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub

