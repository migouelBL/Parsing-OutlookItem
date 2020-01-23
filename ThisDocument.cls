VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Dim OutApp As Object
Dim OutMail As Object
Dim EmailBody As String

'---------------Read Email Body----------------
Set OutApp = CreateObject("Outlook.Application")
Select Case TypeName(OutApp.ActiveWindow)
        Case "Explorer"
            Set OutMail = OutApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set OutMail = OutApp.ActiveInspector.CurrentItem
    End Select

EmailBody = OutMail.Body



'---------Declare Variable with value from EmailBody------------
RequestNumber = RegExing("Request Number:", EmailBody)
RequestType = RegExing("Service Request Type:", EmailBody)
StartDate = RegExing("Start Date:", EmailBody)
EndDate = RegExing("End Date:", EmailBody)


'-------------Plugging Variable into Word FormField-------------
With FormFields
    .Item("RequestNumber").result = RequestNumber
    .Item("RequestType").result = RequestType
    .Item("StartDate").result = StartDate
    .Item("EndDate").result = EndDate
End With


Set OutApp = Nothing
Set OutMail = Nothing
End Sub

Function RegExing(strPattern As String, strInput As String) As String

Dim RegEx As Object
Set RegEx = CreateObject("VBScript.RegExp")

With RegEx
    .pattern = "[\t]*" & strPattern & "[\t]*([^\n\r]*)"
    .Global = True
End With

Set Matches = RegEx.Execute(strInput)
For Each Match In Matches
    RegExing = Match.Value
    RegExing = Right(RegExing, Len(RegExing) - InStr(RegExing, ":") - 1)
    RegExing = Replace$(Trim$(RegExing), vbTab, "")
    Debug.Print RegExing
Next Match

Set Match = Nothing
Set Matches = Nothing
Set RegEx = Nothing

End Function

