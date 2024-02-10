Attribute VB_Name = "SamfordTools"
'*************************************************************************************
'* CUSTOM SAMFORD FUNCTIONS                                                                    *
'*************************************************************************************

Public Sub MakeNewDocOptions()
    UI.ShowForm "SamfordDoc"
End Sub

Public Sub SendDoc()
 If GetSetting("Verbatim", "Paperless", "MakeZappedDoc") = True Then
    Call AssembleSpeechAndZap
 Else
    Call AssembleSpeech
 End If
End Sub

Public Sub AssembleSpeechAndZap()
'
' AssembleSpeechAndZap Macro
'
'
Dim newFile As Document
Dim oldFile As Document

Set oldFile = ActiveDocument
Set newFile = Documents.Add("Debate.dotm")

oldFile.Range.Select
oldFile.Range.Copy
newFile.Activate
newFile.Range.Paste

Call AggressiveDeleteAnalytics

Call SaveToDesktop(newFile, "Send ", oldFile.Name)
If GetSetting("Verbatim", "Paperless", "CloseSendDocAuto", True) = True Then newFile.Close

oldFile.Activate

Call Zapper
End Sub

Public Sub AssembleSpeech(Optional ByVal prefix As String = "Send ")
'
' AssembleSpeech Macro
'
'
Dim newFile As Document
Dim oldFile As Document

Set oldFile = ActiveDocument
Set newFile = Documents.Add("Debate.dotm")

oldFile.Range.Select
oldFile.Range.Copy
newFile.Activate
newFile.Range.Paste

Call AggressiveDeleteAnalytics

Call SaveToDesktop(newFile, prefix, oldFile.Name)
If prefix = "Marked " Then Exit Sub
If GetSetting("Verbatim", "Paperless", "CloseSendDocAuto", True) = True Then newFile.Close

oldFile.Activate

End Sub

Public Sub AggressiveDeleteAnalytics()
'
' AggressiveDeleteAnalytics Macro
'
'
Dim p1 As Paragraph
Dim p2 As Paragraph
Dim p1Count As Long
Dim p2Count As Long
Dim totalDel As Long
Dim AnCount As Long
Dim AnalyticRange As Range
Dim i As Long

Application.ScreenUpdating = False

For i = 1 To 2
AnCount = 0
p1Count = 1
p2Count = 1
totalDel = 0

For Each p1 In ActiveDocument.Paragraphs
  If p1.Range.Text <> vbCr And p1.OutlineLevel = 4 Then 'Ignore blank paragraphs
    If ActiveDocument.Paragraphs.Count < (p1Count + 3 - totalDel) Then
        Set AnalyticRange = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(p1Count - totalDel).Range.Start, End:=ActiveDocument.Paragraphs.Last.Range.End)
        Else:
        Set AnalyticRange = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(p1Count - totalDel).Range.Start, End:=ActiveDocument.Paragraphs(p1Count + 3 - totalDel).Range.End)
    End If
    AnalyticRange.Select
    For Each p2 In Selection.Paragraphs
      If (p2Count >= 1 And p2Count < 3) And p1.OutlineLevel = 4 And (p2.OutlineLevel = 4 Or p2.OutlineLevel = 3 Or p2.OutlineLevel = 2 Or p2.OutlineLevel = 1) Then
        AnCount = AnCount + 1
        If p1.OutlineLevel = 4 And (p2.OutlineLevel = 4 Or p2.OutlineLevel = 3 Or p2.OutlineLevel = 2 Or p2.OutlineLevel = 1) And AnCount > 1 Then
            p1.Range.Delete
            totalDel = totalDel + 1
        End If
      End If
      p2Count = p2Count + 1
    Next p2
    
  End If
  
  'Reset Duplicate Counter
    AnCount = 0
    p2Count = 1
    p1Count = p1Count + 1

Next p1

Dim oPara As Paragraph
For Each oPara In ActiveDocument.Paragraphs
   If Len(oPara.Range.Text) = 1 Then
      oPara.Range.Delete
   End If
Next

With ActiveDocument.Paragraphs
    Set p1 = ActiveDocument.Paragraphs.Last
    If p1.OutlineLevel = 4 Then p1.Range.Delete
End With
Next

Application.ScreenUpdating = True

End Sub

Public Sub Zapper()
'
' Zapper Macro
'
'
Dim newFile As Document
Dim oldFile As Document

Set oldFile = ActiveDocument
Set newFile = Documents.Add("Debate.dotm")

Selection.Collapse Direction:=wdCollapseStart
Selection.FormattedText = oldFile.Range
Set myRange = Selection.FormattedText
newFile.Activate

'Adds paragraph breaks before all tags and fix incorrect formatting
For Each p1 In newFile.Paragraphs
    If p1.OutlineLevel = 4 Or p1.OutlineLevel = 3 Or p1.OutlineLevel = 2 Then
        p1.Range.InsertParagraphBefore
        p1.Previous.Style = "Tag"
    End If
    
    Select Case p1.OutlineLevel
    Case 1 'Pocket
        p1.Style = "Pocket"
    Case 2 'Hat
        p1.Style = "Hat"
    Case 3 'Block
        p1.Style = "Block"
    Case 4 'Tag
        p1.Style = "Tag"
        'Marks cite with white highlight
        p1.Next.Range.HighlightColorIndex = wdWhite
        p1.Next.Range.Font.Underline = False
    End Select
Next
    

'Finds cite and unhighlights any non-bolded portion
Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting
With Selection.Find
    .Text = ""
    .Font.Bold = False
    .Font.Underline = False
    .Replacement.Text = ""
    .Replacement.Highlight = False
    .Forward = True
    .Wrap = wdFindContinue
    .Format = True
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
End With
Selection.Find.Execute Replace:=wdReplaceAll

'Removes card nonhighlighting
With newFile.Range.Find
    .ClearFormatting
    .Format = True
    .Text = ""
    .Replacement.Text = " "
    .Wrap = wdFindContinue
    .Style = "Normal"
    .Highlight = False
    .Execute Replace:=wdReplaceAll
End With




Call SaveToDesktop(newFile, "Zapped ", oldFile.Name)

Application.ScreenUpdating = True
End Sub

Public Sub SaveToDesktop(saveFile, prefixName, fileName)
'
' SaveToDesktop Macro
'
'
Dim strPath As String
    If GetSetting("Verbatim", "Paperless", "SaveSendToDesktop") = True Then
       #If Mac Then 'On Mac

           strPath = "/Users/" & Environ("USER") & "/Desktop/"

           saveFile.SaveAs2 (strPath & prefixName & fileName)

      #Else 'On Windows
          strPath = Environ("USERPROFILE") & "\Desktop\"
          saveFile.SaveAs2 (strPath & prefixName & fileName)
      #End If
    Else
      #If Mac Then 'On Mac
           strPath = GetSetting("Verbatim", "Paperless", "SendDocDir")
           saveFile.SaveAs2 (strPath & "/" & prefixName & fileName)
      #Else 'On Windows
          strPath = GetSetting("Verbatim", "Paperless", "SendDocDir")
          saveFile.SaveAs2 (strPath & "\" & prefixName & fileName)
      #End If
    End If
End Sub


Public Sub MakeMarkedDoc()
'
'MakeMarkedDoc Macro
'
'
Dim zappedFile As Document
Set zappedFile = ActiveDocument

Dim originalDoc As Document
Dim markedDoc As Document

Dim nameLength As Integer
nameLength = Len(ActiveDocument.Name)

Dim isZappedDoc As Boolean
isZappedDoc = False

Dim foundDoc As Boolean
foundDoc = False

If Left(ActiveDocument.Name, 6) = "Zapped" Then isZappedDoc = True
If isZappedDoc = False Then
    If MsgBox("Would you like to make a marked doc out of the current document?", vbYesNo, "Make Marked Doc?") = vbNo Then
        Exit Sub
    Else
        Set originalDoc = ActiveDocument
        Call AssembleSpeech("Marked ")
    End If
Else

    ' Look for a document with the same title as current minus Zapped
    foundDoc = False
    For Each d In Application.Documents
        If d.Name = Right(zappedFile.Name, nameLength - 7) Then
            Set originalDoc = d
            foundDoc = True
        End If
    Next d
    
    If foundDoc = False Then
        noSpeechError = MsgBox("Please open the original speech document", vbOKOnly, "Main speech doc not detected")
        Exit Sub
    End If
        
            
    ' Make marked doc
    originalDoc.Activate
    Call AssembleSpeech("Marked ")
    
    ' Get marked doc
    For Each d In Application.Documents
        If d.Name = "Marked " & Right(zappedFile.Name, nameLength - 7) Then
            Set markedDoc = d
        End If
    Next d
    
    If MsgBox("Would you like to remove unread cards?", vbYesNo, "Remove Unread Cards?") = vbYes Then
        Call SetMarks(zappedFile, markedDoc, True)
    ' Else
        ' Call SetMarks(zappedFile, markedDoc, False)
    End If
        
    
End If
End Sub

Public Sub SetMarks(ByVal zapDoc As Document, ByVal markDoc As Document, ByVal removeUnreadCards As Boolean)
'
'Helper Mark function
'
'

'Detect 2 asterisks at the beginning of tag in zapped doc and remove corresponding card in marked doc
If removeUnreadCards Then
    zapDoc.Activate
    For Each p1 In zapDoc.Paragraphs
        If p1.OutlineLevel = 4 Or p1.OutlineLevel = 3 Or p1.OutlineLevel = 2 Then
            detectSkipped = Left(p1.Range, 2)
            If detectSkipped = "**" Then
                tagText = Right(p1.Range, Len(p1.Range) - 2)
                Debug.Print ("Card to be deleted found. Tag = " & tagText)
                Call FindCard(markDoc, tagText, True)
            End If
        End If
    Next
    ' Call FindHighlighting(zapDoc)
' Else
    ' Call FindHighlighting(zapDoc)
End If

End Sub

Public Sub FindCard(ByVal searchDoc As Document, ByVal searchTag As String, ByVal deleteCard As Boolean)
'
'Helper function to find and delete a card by being given its tag
'
'
searchDoc.Activate

Dim oRange As Range
Set oRange = ActiveDocument.Content
oRange.Collapse Direction:=wdCollapseStart

With Selection.Find
    .Forward = True
    .Text = searchTag
    .MatchCase = True
    .Wrap = wdFindContinue
    .Execute
End With

Selection.StartOf
Paperless.SelectHeadingAndContent
If deleteCard Then Selection.Delete

End Sub

