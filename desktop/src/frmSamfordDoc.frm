VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSamfordDoc 
   Caption         =   "Create Document"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmSamfordDoc.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmSamfordDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnCreate_Click()
    If CheckSendDoc.Value = True Then
        If CheckZappedDoc.Value = False Then
            ' Only create send doc
            Me.Hide
            SamfordTools.AssembleSpeech
        Else
            ' Create send and zapped docs
            Me.Hide
            SamfordTools.AssembleSpeechAndZap
        End If
    Else
        If CheckZappedDoc.Value = False Then
            ' Prompt user for selection
            MsgBox "Please select an option"
        Else
            ' Only create zapped doc
            Me.Hide
            SamfordTools.Zapper
        End If
    End If
            
End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub CheckZappedDoc_Click()

End Sub
