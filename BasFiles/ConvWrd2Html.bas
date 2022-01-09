Attribute VB_Name = "ConvWrd2Html"
Option Explicit
Public Entities, Quotes, CharFonts As Variant
Sub transform_doc2MultiPChoice()
' This macro converts the active document into a document with HTML tags for layout.
' - thanks to Version 2.8 - Toxaris
' - added hard spaces in routine replace_empty_paragraphs
' - added space enter in replace_pilcrow

' - Geveling 220102

Dim i As Integer
Dim DocumentID As String
Dim ReplyEmailAddress As String

    ActiveDocument.Save
    Application.ScreenUpdating = False

With Options
   .AutoFormatAsYouTypeReplaceQuotes = False
End With


    'Execute functions...
    DocumentID = InputBox("DocumentID ? give a short name no spaces no special characters")
    ReplyEmailAddress = InputBox("What is your reply email address ?")
    LoadArrays
    HideAcco
    replace_pilcrow 'replace pilcrows by real line break.
    remove_formating 'remove certain layouts which are more rare in ePUB
    replace_headers 'change headers
    replace_notes 'change footnotes in endnotes and convert endnotes
    replace_bookmarks 'change bookmarks
    replace_hyper 'change hyperlinks
    For i = 0 To UBound(CharFonts) 'change italic, bold and underline
                replace_formating CStr(CharFonts(i))
    Next i
    replace_lists 'change simple lists (1 level only)
    replace_tables 'change tables (no merged cells!)
    'replace_customparagraphs 'change own custom styles
    replace_smallcaps 'change smallcaps
    replace_paragraphs 'change remaining paragraphs
    replace_empty_paragraphs 'change planned empty lines (section changes)
    replace_new_line 'change soft enters
    For i = 0 To UBound(Entities, 1) 'change special characters in HTML (more can be added)
          replace_specials CStr(Entities(i, 0)), CStr(Entities(i, 1)) 'omzetten special characters in HTML codes
    Next i
    For i = 0 To UBound(Quotes, 1)
          replace_specials CStr(Quotes(i, 0)), CStr(Quotes(i, 1)) 'omzetten quotes in HTML codes
    Next i
    replace_pics 'export and change pictures
    place_headerfooter 'insert HTML header
    addRadioButtons
    Call addSubmitButton(CStr(ReplyEmailAddress), CStr(DocumentID))
    Call AddStyle
    
    
    saveashtml (DocumentID) 'save HTML file
    
    Application.ScreenUpdating = True
With Options
   .AutoFormatAsYouTypeReplaceQuotes = True
End With
    
End Sub



Sub addSubmitButton(ReplyEmailAddress As String, DocumentID As String)

     Dim cScript As String
     
     cScript = "<script>" & vbCrLf & _
               " function clickFunction() { " & vbCrLf & _
               " document.getElementById(" & Chr(34) & "Sbmt" & Chr(34) & ").style.color = " & Chr(34) & "red" & Chr(34) & "; " & vbCrLf & _
               " var radios = document.getElementsByTagName('input'); " & vbCrLf & _
               " var value; " & vbCrLf & _
               " var rs;  " & vbCrLf & _
               " rs = " & Chr(34) & "" & Chr(34) & ";   " & vbCrLf & _
               " for (var i = 0; i < radios.length; i++) { " & vbCrLf & _
               " if (radios[i].type === 'radio' && radios[i].checked) {  " & vbCrLf & _
               "   value = radios[i].name.concat(radios[i].value); " & vbCrLf & _
               "   rs = rs.concat(value); " & vbCrLf & _
               "  } " & vbCrLf & _
               "} " & vbCrLf & _
               " rs = rs.concat(" & Chr(34) & " ; " & Chr(34) & ", document.getElementById('cname').value)" & vbCrLf & _
               " var link = " & Chr(34) & "mailto:" & ReplyEmailAddress & "" & Chr(34) & " " & vbCrLf & _
               " + " & Chr(34) & "?cc=                         " & Chr(34) & " " & vbCrLf & _
               " + " & Chr(34) & "&subject=" & Chr(34) & " + encodeURIComponent(" & Chr(34) & DocumentID & Chr(34) & ") " & vbCrLf & _
               " + " & Chr(34) & "&body=" & Chr(34) & " + encodeURIComponent(rs) " & vbCrLf & _
               " ; " & vbCrLf & _
               " window.location.href = link; " & vbCrLf & _
               " } " & vbCrLf & _
               " </script>" & vbCrLf & _
               ""
     
 
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
    Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=2
    Selection.TypeText Text:=cScript
    
    cScript = "<p><label for=" & Chr(34) & "fname" & Chr(34) & ">Naam:</label>" & vbCrLf & _
              "  <input type=" & Chr(34) & "text" & Chr(34) & " id=" & Chr(34) & "cname" & Chr(34) & " name=" & Chr(34) & "cname" & Chr(34) & "><br><br></p>" & vbCrLf & _
              "<p><button class=" & Chr(34) & "button button1" & Chr(34) & " id=" & Chr(34) & "Sbmt" & Chr(34) & " onclick=" & Chr(34) & "clickFunction()" & Chr(34) & ">Submit</button></p></body>"
     
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "</body>"
        .Replacement.Text = cScript
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
     
End Sub

Sub AddStyle()
   Dim cStyle
   
   cStyle = "<head><style type=" & Chr(34) & "text/css" & Chr(34) & "> p {font-family: sans-serif}</style>"
 
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<head>"
        .Replacement.Text = cStyle
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
  

End Sub

Function HideAcco()

    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToAbsolute, Count:=1
     Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "\{" & "*" & "\}"
        .Replacement.Text = ""
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    



End Function

Function replace_pilcrow()

'Repair wrong usage of pilcrows instead of real line breaks.

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "^13"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

            Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = " ^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Function

Function remove_formating()

'Remove layout which usually is not used in ePUB.
Dim oRg As Range
Dim answer As String

Set oRg = ActiveDocument.Range
answer = vbYes

StatusBar = "Cleanup undesired layout..."

'Do While answer <> vbNo
'    answer = MsgBox("Remove bold layout?", vbQuestion + vbYesNo, "Layout")
'    If answer = vbNo Then Exit Do
     oRg.Font.Bold = False
'    answer = vbNo
'Loop

answer = vbYes

'Do While answer <> vbNo
'    answer = MsgBox("Remove underlines?", vbQuestion + vbYesNo, "Layout")
'    If answer = vbNo Then Exit Do
'    oRg.Font.Underline = False
'    answer = vbNo
'Loop

End Function

Function replace_headers()
' change headers to HTML code
Dim i, headnum As Integer
Dim oRg As Range

Set oRg = ActiveDocument.Range
StatusBar = "Change headers to HTML..."

For i = -2 To -7 Step -1
headnum = Abs(i) - 1
    oRg.Find.ClearFormatting
    oRg.Find.Style = ActiveDocument.Styles(i) 'search only for headers
    oRg.Find.Text = "" ' Search for anything, as long as it is the header
    oRg.Find.Wrap = wdFindContinue
        Do While oRg.Find.Execute = True ' Execute replacements
            If oRg.Characters.Count > 1 Then ' Catch empty headers
                oRg.Style = -1 'change to normal style to prevent further selection
                oRg.Find.Replacement.Font.Reset
                While oRg.Characters.Last = " " Or oRg.Characters.Last = vbCr
                    oRg.MoveEnd Unit:=wdCharacter, Count:=-1 ' remove line break from selection
                Wend
                oRg.InsertBefore "<h" & headnum & ">"
                oRg.InsertAfter "</h" & headnum & ">"
                oRg.Find.Replacement.ClearFormatting
            Else
                oRg.Style = -1
            End If
            oRg.Move
        Loop
Next i

End Function

Function addRadioButtons()
Dim i
Dim oRng As Range
Dim oPar As Paragraph

For Each oPar In ActiveDocument.Paragraphs
 
 Set oRng = oPar.Range
 
 If InStr(oRng.Text, "#") > 0 Then
    Dim s
    Dim e
    Dim cTag
    Dim nQ
    Dim cA
    Dim cNew
    Dim cLineText
    
    s = InStr(oRng.Text, "#")
    e = InStr(oRng.Text, " ")
    cTag = Mid(oRng.Text, s, e - s)
    nQ = 0
    
    For i = 2 To e - s
        If IsNumeric(Mid(cTag, i, 1)) Then
            nQ = nQ * 10 + CInt(Mid(cTag, i, 1))
        Else
            cA = Mid(cTag, i, 1)
        End If
    Next i
    
    cNew = "<input type=" & Chr(34) & "radio" & Chr(34) & " id=" & Chr(34) & CStr(nQ) & cA & Chr(34) & " name=" & Chr(34) & "q" & CStr(nQ) & Chr(34) & " value=" & Chr(34) & cA & Chr(34) & "><label for=" & Chr(34) & cA & Chr(34) & ">"
    
    'Replace the cTag with cNew
    
    oRng.Find.ClearFormatting
    oRng.Find.Replacement.ClearFormatting
    With oRng.Find
        .Text = cTag
        .Replacement.Text = cNew
    End With
    oRng.Find.Execute Replace:=wdReplaceAll
    
    'insert </label> just before th </p>
    
    cLineText = oRng.Text
    oRng.Text = Mid(cLineText, 1, Len(cLineText) - 5) & "</label></p>" & vbCrLf
    
    
    
        
       
 End If
    
Next

End Function


Function replace_formating(tg As String)
' Replace any text in italic, bold or underline and change this to HTML
    Dim oRg As Range
    Dim para As Paragraph
    Dim ParaText As String
        Dim newFnt As Font

        Set oRg = ActiveDocument.Range
        Set newFnt = New Font

StatusBar = "Change layout like bold/italic/underline/etc..."

With oRg.Find
.ClearFormatting
.Text = " ^13"
.Replacement.Text = vbCr
.Execute Replace:=wdReplaceAll, Wrap:=wdFindContinue
End With

If Not (oRg Is Nothing) Then Set oRg = Nothing

    Selection.HomeKey wdStory
    With Selection.Find
        .ClearFormatting
        .Text = ""

        'you can add different formats/fonts to look for by creating a seperate Case
        Select Case tg
        Case "i"
            .Font.Italic = True
            newFnt.Italic = False
        Case "b"
            .Font.Bold = True
            newFnt.Bold = False
        Case "u"
            .Font.Underline = True
            newFnt.Underline = False
        Case "s"
            .Font.StrikeThrough = True
            newFnt.StrikeThrough = False
        Case "sup"
            .Font.Superscript = True
            newFnt.Superscript = False
        Case "sub"
            .Font.Subscript = True
            newFnt.Subscript = False
        Case Else       'if the tag is not listed above then exit the function
            If Not (oRg Is Nothing) Then Set oRg = Nothing
            If Not (newFnt Is Nothing) Then Set newFnt = Nothing
            Exit Function
        End Select
    End With
        
Do While Selection.Find.Execute = True
   Set oRg = Selection.Range
   oRg.Font = newFnt
   If oRg.Characters.Count > 0 Then
      oRg.MoveStartWhile Cset:=vbCr & Chr(11) & " ‘“", Count:=wdForward
      If oRg.Characters(1) = "<" Then
        oRg.MoveStartUntil Cset:=">", Count:=wdForward
        oRg.MoveStart Count:=1
      End If
      oRg.MoveEndWhile Cset:=vbCr & ChrW(11) & " .,’”–", Count:=wdBackward
      If oRg.Characters(oRg.Characters.Count) = ">" Then
        oRg.MoveEndUntil Cset:="<", Count:=wdBackward
        oRg.MoveEnd Count:=-1
      End If
      If oRg.Characters.Count > 0 Then
        oRg.InsertBefore "<" & tg & ">"
        oRg.InsertAfter "</" & tg & ">"
      End If
      Selection.Collapse Direction:=wdCollapseEnd
   Else
      Selection.Collapse Direction:=wdCollapseEnd
   End If
Loop

   'centreer omzetten
    For Each para In ActiveDocument.Paragraphs
        If para.Alignment = wdAlignParagraphCenter Then
        para.Alignment = wdAlignParagraphLeft
        Set oRg = para.Range
         oRg.MoveEndWhile Cset:=vbCr, Count:=wdBackward    'enter niet meenemen
            oRg.InsertBefore "<center>"
            oRg.InsertAfter "</center>"
        End If
        Next para
If Not (oRg Is Nothing) Then Set oRg = Nothing


End Function
Function replace_smallcaps()
' Search for smallcaps (style) and change this to a HTML style. This style must exist in the stylesheet
Dim oRg As Range
Dim answer, invoer As String

answer = vbYes

'Do While answer <> vbNo
'    answer = MsgBox("Change smallcaps?", vbQuestion + vbYesNo, "Layout")
'    If answer = vbYes Then Exit Do
'    answer = vbNo
'Loop

If answer = vbYes Then
'invoer = ""
'On Error Resume Next
'    invoer = InputBox("Name of the smallcaps style?")
'    On Error GoTo 0
invoer = "smllcps"

Set oRg = ActiveDocument.Range

StatusBar = "Change smallcaps..."

With oRg.Find
.ClearFormatting
.Text = " ^13"
.Replacement.Text = vbCr
.Execute Replace:=wdReplaceAll, Wrap:=wdFindContinue
End With

If Not (oRg Is Nothing) Then Set oRg = Nothing

Selection.HomeKey wdStory
Selection.Find.ClearFormatting
Selection.Find.Font.SmallCaps = True
Selection.Find.Text = ""
Do While Selection.Find.Execute = True
   Set oRg = Selection.Range
   oRg.Font.SmallCaps = False
   If oRg.Characters.Count > 0 Then
      oRg.MoveStartWhile Cset:=vbCr & " ‘“", Count:=wdForward
      If oRg.Characters(1) = "<" Then
        oRg.MoveStartUntil Cset:=">", Count:=wdForward
        oRg.MoveStart Count:=1
      End If
      oRg.MoveEndWhile Cset:=vbCr & ChrW(11) & " .,’”–", Count:=wdBackward
      If oRg.Characters(oRg.Characters.Count) = ">" Then
        oRg.MoveEndUntil Cset:="<", Count:=wdBackward
        oRg.MoveEnd Count:=-1
      End If
      If oRg.Characters.Count > 0 Then
        oRg.Case = wdUpperCase
        oRg.InsertBefore "<span class=""" & invoer & """>"
        oRg.InsertAfter "</span>"
      End If
      Selection.Collapse Direction:=wdCollapseEnd
   Else
      Selection.Collapse Direction:=wdCollapseEnd
   End If
Loop
End If

If Not (oRg Is Nothing) Then Set oRg = Nothing


End Function
Function replace_specials(strFind As String, strReplace As String)
' Replace special characters (e.g. Ellips) to HTML entities …

Dim oRg As Range

Set oRg = ActiveDocument.Range
StatusBar = "Replace special characters to HTML entity..."

    With oRg.Find
        .ClearFormatting    ' Vorige opmaak wissen
        .Text = strFind    ' spaties aan het eind van een regel verwijderen
        .Replacement.Text = strReplace    ' vervangen door alleen een enter
        .Execute Replace:=wdReplaceAll, Wrap:=wdFindContinue    ' Voer de vervangingen uit
    End With


If Not (oRg Is Nothing) Then Set oRg = Nothing

End Function

Function replace_notes()
' Change footnotes to endnotes
' Change endnotes in HTML with references.
Dim num As Long
Dim myString As String

StatusBar = "Convert foot- and endnotes..."

If ActiveDocument.Footnotes.Count > 0 Then
With ActiveDocument.Sections.Last.Range
    .Collapse Direction:=wdCollapseEnd
    .InsertParagraphAfter
    .InsertAfter "<hr />" & vbCr
    Selection.EndKey Unit:=wdStory
    Selection.ClearFormatting
End With
    ActiveDocument.Footnotes.Convert
End If
        
If ActiveDocument.Endnotes.Count = 0 Then
    Exit Function
End If

With Selection
    .HomeKey wdStory
    For num = 1 To ActiveDocument.Endnotes.Count
        .GoToNext wdGoToEndnote
        .TypeText Text:="<a href=" & Chr(34) & "#end" & CStr(num) & Chr(34) & " id=" & Chr(34) & "endref" & CStr(num) & Chr(34) & "><sup>" & CStr(num) & "</sup></a>"
        .Expand wdWord
        With ActiveDocument.Endnotes(1)
            myString = myString & "<a href=" & Chr(34) & "#endref" & CStr(num) & Chr(34) & " id=" & Chr(34) & "end" & CStr(num) & Chr(34) & "><sup>" & CStr(num) & "</sup></a>" & ". " & .Range.Text & vbCrLf
            .Delete
        End With
    Next
    .EndKey wdStory
    .InsertAfter myString
    .Collapse Direction:=wdCollapseEnd
End With

End Function

Function replace_lists()
' change simple lists (1 level!)
Dim lijst As List
Dim para As Paragraph
Dim i As Long

StatusBar = "Convert lists..."

For Each para In ActiveDocument.ListParagraphs
    With para.Range
    For i = 1 To .ListFormat.ListLevelNumber
        .MoveEnd Unit:=wdCharacter, Count:=-1
        .InsertBefore "<li>"
        .InsertAfter "</li>"
    Next i
    End With
Next para

For Each lijst In ActiveDocument.Lists
    With lijst.Range
        .MoveEnd Unit:=wdCharacter, Count:=-1
        If .ListFormat.ListType = wdListBullet Then
            .InsertBefore "<ul>" & vbCr
            .InsertAfter "</ul>"
        Else
            .InsertBefore "<ol>" & vbCr
            .InsertAfter "</ol>"
        End If
        .ListFormat.RemoveNumbers
    End With
Next lijst

End Function

Function replace_tables()
' convert simple tables (no merged cells!)
Dim oRow As Row
Dim oCell As Cell
Dim sCellText As String
Dim tTable As Table
Dim noRows, noCells As Long

StatusBar = "Convert tables..."
   
For Each tTable In ActiveDocument.Tables
    For Each oRow In tTable.Rows
        For Each oCell In oRow.Cells
            sCellText = oCell.Range
            sCellText = Left$(sCellText, Len(sCellText) - 2)
            If Len(sCellText) = 0 Then sCellText = "&nbsp;"
            sCellText = "<td>" & sCellText & "</td>"
            oCell.Range = sCellText
        Next oCell
        sCellText = oRow.Cells(1).Range
        sCellText = Left$(sCellText, Len(sCellText) - 2)
        sCellText = "<tr>" & vbCr & sCellText
        oRow.Cells(1).Range = sCellText
        sCellText = oRow.Cells(oRow.Cells.Count).Range
        sCellText = Left$(sCellText, Len(sCellText) - 2)
        sCellText = sCellText & vbCr & "</tr>"
        oRow.Cells(oRow.Cells.Count).Range = sCellText
    Next oRow
    sCellText = tTable.Rows(1).Cells(1).Range
    sCellText = Left$(sCellText, Len(sCellText) - 2)
    sCellText = "<table>" & vbCr & sCellText
    tTable.Rows(1).Cells(1).Range = sCellText
    noRows = tTable.Rows.Count
    noCells = tTable.Rows(noRows).Cells.Count
    sCellText = tTable.Rows(noRows).Cells(noCells).Range
    sCellText = Left$(sCellText, Len(sCellText) - 2)
    sCellText = sCellText & vbCr & "</table>"
    tTable.Rows(noRows).Cells(noCells).Range = sCellText
    
    tTable.ConvertToText Separator:=wdSeparateByParagraphs
Next tTable

End Function

Function replace_bookmarks()
'Omzetten bookmarks
    Dim addr As String
    Dim bmark As Bookmark

        StatusBar = "Convert bookmarks..."
        
    For Each bmark In ActiveDocument.Bookmarks
        addr = bmark.Name
        bmark.Range.InsertBefore "<a id=" & Chr(34) & addr & Chr(34) & "></a>"    'create reference
    Next bmark

    Selection.Find.Execute Replace:=wdReplaceAll


End Function

Function replace_hyper()
'Convert hyperlinks
Dim hyper As Hyperlink
Dim hypercount, i As Long
Dim addr As String

StatusBar = "Convert hyperlinks..."

hypercount = ActiveDocument.Hyperlinks.Count
If hypercount > 0 Then
    For i = 1 To hypercount
        Set hyper = ActiveDocument.Hyperlinks(1)
        If hyper.SubAddress <> "" Then
            addr = "#" & hyper.SubAddress 'internal hyperlink
        Else
            addr = hyper.Address    'external hyperlink
        End If
        hyper.Delete    'Verwijder hyperlink, niet de tekst!
        hyper.Range.InsertBefore "<a href=" & Chr(34) & addr & Chr(34) & ">"    'place HTML link
        hyper.Range.InsertAfter "</a>"
    Next i
End If

End Function

Function replace_pics()
Dim sDir
Dim iDir, num As Integer
Dim oPlaatje As Word.InlineShape
Dim oShape As Word.Shape
Dim HuidigeMap, ExportMap As String
Dim imgname, oldname As String

StatusBar = "Export images and create links..."

HuidigeMap = ActiveDocument.Path & Application.PathSeparator
ExportMap = HuidigeMap & "Save_As_HTML_files" & Application.PathSeparator

On Error Resume Next
Kill HuidigeMap & "Save_As_HTML.html"

On Error Resume Next
Kill ExportMap & "*.*"

On Error Resume Next
RmDir ExportMap
    
Application.Documents.Add ActiveDocument.FullName
ActiveDocument.SaveAs HuidigeMap & "Save_As_HTML.html", FileFormat:=wdFormatHTML
ActiveDocument.Close

num = 1

For Each oShape In ActiveDocument.Shapes
   oShape.ConvertToInlineShape
Next
  
For Each oPlaatje In ActiveDocument.InlineShapes
   With oPlaatje.Range
       imgname = "image" & Format(num, "000") & ".jpg"
       oldname = ExportMap & "image" & Format(num, "000") & ".jpg"
       .InsertBefore "<img src=" & Chr(34) & imgname & Chr(34) & " />"
       oPlaatje.Delete
       imgname = HuidigeMap & Application.PathSeparator & imgname
       FileCopy oldname, imgname
       num = num + 1
   End With
Next

On Error Resume Next
Kill HuidigeMap & "Save_As_HTML.html"

On Error Resume Next
Kill ExportMap & "*.*"

On Error Resume Next
RmDir ExportMap

End Function

Function replace_customparagraphs()
'Convert own paragraph styles

Dim oRg As Range
Dim para As Paragraph
Dim answer, invoer As String
Dim s As Style
Dim verwerkt As Boolean

Set oRg = ActiveDocument.Range
answer = vbYes
verwerkt = False

StatusBar = "Convert own Word Styles..."

Do While answer <> vbNo
    answer = MsgBox("Convert other paragraph styles?", vbQuestion + vbYesNo, "Paragraphs")
    If answer = vbNo Then Exit Function
    On Error Resume Next
    invoer = InputBox("Convert which style?")

    For Each para In ActiveDocument.Paragraphs
        Set oRg = para.Range
            If oRg.Style = invoer Then
                  verwerkt = True
              oRg.Style = -1
              oRg.Font.Reset
              If oRg.Text = vbCr Then oRg.InsertBefore "&nbsp;"
              oRg.MoveEndWhile Cset:=vbCr, Count:=wdBackward
              oRg.InsertBefore "<p class=" & Chr(34) & invoer & Chr(34) & ">"
              oRg.InsertAfter "</p>"
        End If
        If Not (oRg Is Nothing) Then Set oRg = Nothing
    Next

    If Not verwerkt Then MsgBox ("Style " & invoer & " does not exist.")
    verwerkt = False
Loop

On Error GoTo 0

If Not (oRg Is Nothing) Then Set oRg = Nothing

End Function

Function replace_paragraphs()
'Convert paragraphs in HTML. Skip headers, lists, tables and paragraphs in other styles

Dim oRg As Range
Dim GeenPara(14), firstchar As String
Dim para As Paragraph


'list of codes of tags for which paragraphs must be skipped. Only first three characters of tag required
GeenPara(0) = "<h1"
GeenPara(1) = "<h2"
GeenPara(2) = "<h3"
GeenPara(3) = "<h4"
GeenPara(4) = "<h5"
GeenPara(5) = "<h6"
GeenPara(6) = "<ol"
GeenPara(7) = "<ul"
GeenPara(8) = "<li"
GeenPara(9) = "<ta"
GeenPara(10) = "<td"
GeenPara(11) = "<tr"
GeenPara(12) = "<p "
GeenPara(13) = "<ce"
GeenPara(14) = "</t"

StatusBar = "Convert paragraphs..."

For Each para In ActiveDocument.Paragraphs
    Set oRg = para.Range
    oRg.Style = -1
    oRg.Font.Reset
    If oRg.Text = vbCr Then oRg.InsertBefore "&nbsp;"
    oRg.MoveEndWhile Cset:=vbCr, Count:=wdBackward
    firstchar = Left(oRg.Text, 3)
    If InStr(Join(GeenPara), firstchar) = 0 Then
    oRg.InsertBefore "<p>"
    oRg.InsertAfter "</p>"
    End If
    If Not (oRg Is Nothing) Then Set oRg = Nothing
Next

If Not (oRg Is Nothing) Then Set oRg = Nothing
If Not (para Is Nothing) Then Set para = Nothing

End Function

Function replace_empty_paragraphs()
'Convert planned empty lines

Dim oRg As Range

Set oRg = ActiveDocument.Range

StatusBar = "Convert paragraphs..."

With oRg.Find
    .ClearFormatting
    .MatchWildcards = True
    .Text = "^13^13"
    .Replacement.Text = vbCr & "<p>&nbsp;</p>" & vbCr
    .Replacement.Style = -1
    .Execute Replace:=wdReplaceAll, Wrap:=wdFindContinue
End With

With oRg.Find
.ClearFormatting
.MatchWildcards = False
.Text = "^s"
.Replacement.Text = "&nbsp;"
.Execute Replace:=wdReplaceAll, Wrap:=wdFindContinue
End With

If Not (oRg Is Nothing) Then Set oRg = Nothing

End Function


Function replace_new_line()
' Change soft line break in HTML

Dim oRg As Range

Set oRg = ActiveDocument.Range
StatusBar = "Convert special characters to HTML code..."

With oRg.Find
    .ClearFormatting
    .Text = "^11"
    .Replacement.Text = "<br />"
    .Execute Replace:=wdReplaceAll, Wrap:=wdFindContinue
End With

If Not (oRg Is Nothing) Then Set oRg = Nothing



End Function

Function place_headerfooter()
Dim MyText, invoer As String
Dim myRange As Object

StatusBar = "Insert header and stylesheet link..."

Set myRange = ActiveDocument.Range
invoer = ""
On Error Resume Next
    invoer = "" 'InputBox("Name external stylesheet?")
    If invoer <> "" Then
        invoer = "<link rel=" & Chr(34) & "stylesheet" & Chr(34) & " type=" & Chr(34) & "text/css" & Chr(34) & " href=" & Chr(34) & "..\Styles\" & invoer & Chr(34) & ">" & vbCr
    End If
        invoer = invoer + "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & " content=" & Chr(34) & "text/html; charset=UTF-8" & Chr(34) & "/>"
    On Error GoTo 0

MyText = "<html>" & vbCr & "<head>" & vbCr & invoer & "</head>" & vbCr & "<body>" & vbCr
myRange.InsertBefore (MyText)
MyText = "</body>" & vbCr & "</html>"
myRange.InsertAfter (MyText)

End Function

Function saveashtml(DocumentID As String)
Dim Bestandsnaam, answer As String
Dim extPos As Integer
 
' answer = MsgBox("Save HTML?", vbQuestion + vbYesNo, "Save")
answer = vbYes
If answer = vbNo Then Exit Function


    
extPos = InStrRev(ActiveDocument.FullName, ".")
Bestandsnaam = Left(ActiveDocument.FullName, extPos - 1) & ".html"
Bestandsnaam = DocumentID & ".html"
ActiveDocument.SaveAs FileName:=Bestandsnaam, FileFormat:=wdFormatText, Encoding:=msoEncodingUTF8

MsgBox "File saved as " & Bestandsnaam, vbInformation + vbOKOnly, "Done!"


ActiveDocument.Close

End Function

Function LoadArrays()
'if you wish to add entity codes make sure you update the first number in the redim statement(s)

'List of characters and their corresponding entity codes
ReDim CharFonts(6) As String
    CharFonts(0) = "i"                                                'Italics
    CharFonts(1) = "b"                                                'Bold
    CharFonts(2) = "s"                                                'Strikethrough
    CharFonts(3) = "sup"                                              'Superscript
    CharFonts(4) = "sub"                                              'Subscript
    CharFonts(5) = "u"                                                'Underline

'List of characters and their corresponding entity codes
ReDim Entities(91, 1) As String
    Entities(0, 0) = "—": Entities(0, 1) = "&mdash;"                  'Em dash
    Entities(1, 0) = "&#8212;": Entities(1, 1) = "&mdash;"            'Em dash
    Entities(2, 0) = "–": Entities(2, 1) = "&ndash;"                  'En dash
    Entities(3, 0) = "&#8211;": Entities(3, 1) = "&ndash;"            'En dash
    Entities(4, 0) = "&#8230;": Entities(4, 1) = "&hellip;"           'Horizontal ellipse
    Entities(5, 0) = "…": Entities(5, 1) = "&hellip;"                 'Horizontal ellipse
    Entities(6, 0) = "¡": Entities(6, 1) = "&iexcl;"                'Inverted Exclamation
    Entities(7, 0) = "&#161;": Entities(7, 1) = "&iexcl;"           'Inverted Exclamation
    Entities(8, 0) = "©": Entities(8, 1) = "&copy;"                 'Copyright
    Entities(9, 0) = "&#169;": Entities(9, 1) = "&copy;"            'Copyright
    Entities(10, 0) = "®": Entities(10, 1) = "&reg;"                  'Registered trademark
    Entities(11, 0) = "&#174;": Entities(11, 1) = "&reg;"             'Registered trademark
    Entities(12, 0) = "°": Entities(12, 1) = "&deg;"                  'Degree sign
    Entities(13, 0) = "&#176;": Entities(13, 1) = "&deg;"             'Degree sign
    Entities(14, 0) = "±": Entities(14, 1) = "&plusmn;"               'Plus or minus
    Entities(15, 0) = "&#177;": Entities(15, 1) = "&plusmn;"          'Plus or minus
    Entities(16, 0) = "µ": Entities(16, 1) = "&micro;"                'Micro sign
    Entities(17, 0) = "&#181;": Entities(17, 1) = "&micro;"           'Micro sign
    Entities(18, 0) = "·": Entities(18, 1) = "&middot;"               'Middle dot
    Entities(19, 0) = "&#183;": Entities(19, 1) = "&middot;"          'Middle dot
    Entities(20, 0) = "¼": Entities(20, 1) = "&frac14;"               'Fraction one-fourth
    Entities(21, 0) = "&#188;": Entities(21, 1) = "&frac14;"          'Fraction one-fourth
    Entities(22, 0) = "½": Entities(22, 1) = "&frac12;"               'Fraction one-half
    Entities(23, 0) = "&#189;": Entities(23, 1) = "&frac12;"          'Fraction one-half
    Entities(24, 0) = "¾": Entities(24, 1) = "&frac34;"               'Fraction three-fourths
    Entities(25, 0) = "&#190;": Entities(25, 1) = "&frac34;"          'Fraction three-fourths
    Entities(26, 0) = "¿": Entities(26, 1) = "&iquest;"               'Inverted question mark
    Entities(27, 0) = "&#191;": Entities(27, 1) = "&iquest;"          'Inverted question mark
    Entities(28, 0) = "Ø": Entities(28, 1) = "&Oslash;"               'Capital O, slash
    Entities(29, 0) = "&#216;": Entities(29, 1) = "&Oslash;"          'Capital O, slash
    Entities(30, 0) = "÷": Entities(30, 1) = "&divide;"               'Division sign
    Entities(31, 0) = "&#247;": Entities(31, 1) = "&divide;"          'Division sign
    Entities(32, 0) = "ø": Entities(32, 1) = "&oslash;"               'Small o, slash
    Entities(33, 0) = "&#248;": Entities(33, 1) = "&oslash;"          'Small o, slash
    Entities(34, 0) = "ƒ": Entities(34, 1) = "&fnof;"                 'florin
    Entities(35, 0) = "&#402;": Entities(35, 1) = "&fnof;"            'florin
    Entities(36, 0) = "†": Entities(36, 1) = "&dagger;"               'Dagger
    Entities(37, 0) = "&#8224;": Entities(37, 1) = "&dagger;"         'Dagger
    Entities(38, 0) = "‡": Entities(38, 1) = "&Dagger;"               'Double-dagger
    Entities(39, 0) = "&#8225;": Entities(39, 1) = "&Dagger;"         'Double-dagger
    Entities(40, 0) = "•": Entities(40, 1) = "&bull;"                 'Bullseye circle
    Entities(41, 0) = "&#8226;": Entities(41, 1) = "&bull;"           'Bullseye circle
    Entities(42, 0) = "‰": Entities(42, 1) = "&permil;"               'Per mille (1,000)
    Entities(43, 0) = "&#8240;": Entities(43, 1) = "&permil;"         'Per mille (1,000)
    Entities(44, 0) = "€": Entities(44, 1) = "&euro;"                 'Euro sign
    Entities(45, 0) = "&#8364;": Entities(45, 1) = "&euro;"           'Euro sign
    Entities(46, 0) = "™": Entities(46, 1) = "&trade;"                'Trademark
    Entities(47, 0) = "&#8482;": Entities(47, 1) = "&trade;"          'Trademark

'List of quotes and their corresponding entity codes
ReDim Quotes(22, 1) As String
    Quotes(0, 0) = "‘": Quotes(0, 1) = "&lsquo;"                      'Left single quote
    Quotes(1, 0) = "&#8216;": Quotes(1, 1) = "&lsquo;"                'Left single quote
    Quotes(2, 0) = "^0145": Quotes(2, 1) = "&lsquo;"                  'Left single quote
    Quotes(3, 0) = "’": Quotes(3, 1) = "&rsquo;"                      'Right single quote
    Quotes(4, 0) = "&#8217;": Quotes(4, 1) = "&rsquo;"                'Right single quote
    Quotes(5, 0) = "^0146": Quotes(5, 1) = "&rsquo;"                  'Right single quote
    Quotes(6, 0) = "“": Quotes(6, 1) = "&ldquo;"                      'Left double quote
    Quotes(7, 0) = "&#8220;": Quotes(7, 1) = "&ldquo;"                'Left double quote
    Quotes(8, 0) = "^0147": Quotes(8, 1) = "&ldquo;"                  'Left double quote
    Quotes(9, 0) = "”": Quotes(9, 1) = "&rdquo;"                      'Right double quote
    Quotes(10, 0) = "&#8221;": Quotes(10, 1) = "&rdquo;"              'Right double quote
    Quotes(11, 0) = "^0148": Quotes(11, 1) = "&rdquo;"                'Right double quote
    Quotes(12, 0) = "„": Quotes(12, 1) = "&dbquo;"                    'Double bottom quote
    Quotes(13, 0) = "»": Quotes(13, 1) = "&raquo;"                    'Right angle quote, guillemot right
    Quotes(14, 0) = "&#187;": Quotes(14, 1) = "&raquo;"               'Right angle quote, guillemot right
    Quotes(15, 0) = "«": Quotes(15, 1) = "&laquo;"                    'Left angle quote, guillemot left
    Quotes(16, 0) = "&#171;": Quotes(16, 1) = "&laquo;"               'Left angle quote, guillemot left
    Quotes(17, 0) = "‹": Quotes(17, 1) = "&lsaquo;"                   'Left Single angle quote
    Quotes(18, 0) = "&#8249;": Quotes(18, 1) = "&lsaquo;"             'Left Single angle quote
    Quotes(19, 0) = "›": Quotes(19, 1) = "&rsaquo;"                   'Right Single angle quote
    Quotes(20, 0) = "&#8250;": Quotes(20, 1) = "&rsaquo;"             'Right Single angle quote
End Function
