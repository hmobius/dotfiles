' http://www.slipstick.com/outlook/email/automatically-bcc-all-message/

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_ItemSend(ByVal Item As Object, _
                                 Cancel As Boolean)
    Dim objRecip As Recipient
    Dim strMsg As String
    Dim res As Integer
    Dim strBcc As String
    On Error Resume Next
 
    strBcc = "dama@manning.com"
 
' Use the account name as it appears in Account Settings
If Item.SendUsingAccount = "Manning" Then
 
    Set objRecip = Item.Recipients.Add(strBcc)
    objRecip.Type = olBCC
    If Not objRecip.Resolve Then
        strMsg = "Could not resolve the Bcc recipient. " & _
                 "Do you want to send the message?"
        res = MsgBox(strMsg, vbYesNo + vbDefaultButton1, _
                "Could Not Resolve Bcc")
        If res = vbNo Then
            Cancel = True
        End If
    End If
     
End If
 
    Set objRecip = Nothing
End Sub


