const module_name  = "LineWizard"
const module_ver   = "0.1"

Dim textToInsert
Dim textToFind
Dim startPosition
Dim j

sub about
    Dim msg
    msg = "Line Wizard v0.1. Written by Shpati Koleka. MIT License."
    MsgBox msg
end sub

Private Function ChangeEachLine(action)
    ' Set the Editor
    set Editor = NewEditor()
    Editor.assignActiveEditor()
    Dim modifiedText
    modifiedText = ""
    Dim words
    Dim i
    i = 0
    ' Iterate through each line and add/remove the text to the specified position
    For each item in Editor
        
        If action = "add text at the start" Then
            modifiedText = modifiedText & textToInsert & item & vbNewLine
        End If
        If action = "add text at the end" Then
            modifiedText = modifiedText & item & textToInsert & vbNewLine
        End If
        If action = "remove first word" Then
            ' Split the line into words
            words = Split(item, " ")
            i = 0
            For each word in words
                If i > 0 and i < UBound(words) + 1 Then modifiedText = modifiedText & " " & word
                i = i + 1
            Next
            modifiedText = modifiedText & vbNewLine
        End If
        If action = "remove last word" Then
            ' Split the line into words
            words = Split(item, " ")
            i = 0
            for each word in words
                If i = 0 and UBound(words) > 1 Then modifiedText = modifiedText & word
                If i > 0 and i < UBound(words) Then modifiedText = modifiedText & " " & word
                i = i + 1
            Next
            modifiedText = modifiedText & vbNewLine
        End If
        If action = "keep first word" Then
            ' Split the line into words
            words = Split(item, " ")
            modifiedText = modifiedText & words(0) & vbNewLine
        End If
        If action = "keep last word" Then
            ' Split the line into words
            words = Split(item, " ")
            modifiedText = modifiedText & words(UBound(words)) & vbNewLine
        End If
        If action = "keep lines containing a given text" Then
            If InStr(1, item, textToInsert, vbTextCompare) > 0 Then
                modifiedText = modifiedText & item & vbNewLine
            End If
        End If
        If action = "remove lines containing a given text" Then
            If InStr(1, item, textToInsert, vbTextCompare) = 0 Then
                modifiedText = modifiedText & item & vbNewLine
            End If
        End If
        If action = "remove blank lines" Then
            If item <> "" Then modifiedText = modifiedText & item & vbNewLine
        End If
        If action = "remove redundant blank lines" Then
            If item <> "" Then i = 0
            if i = 0 then modifiedText = modifiedText & item & vbNewLine
            If item = "" Then i = 1
        End If
        If action = "remove redundant spaces" Then
            Dim item1
            Do ' Loop to remove redundant spaces
                item1 = item
                item = Replace(item, "  ", " ") ' Replace double spaces with a single space
                If item = item1 Then Exit Do
            Loop
            modifiedText = modifiedText & item & vbNewLine
        End If
        If action = "remove leading spaces" Then
            item = LTrim(item)
            modifiedText = modifiedText & item & vbNewLine
        End If
        If action = "remove trailing spaces" Then
            item = RTrim(item)
            modifiedText = modifiedText & item & vbNewLine
        End If
        If action = "add line break before text" Then
            textToInsert = vbNewLine & textToFind
            item = Replace(item, textToFind, textToInsert)
            modifiedText = modifiedText & item & vbNewLine
        End If
        If action = "add line break after text" Then
            textToInsert = textToFind & vbNewLine
            item = Replace(item, textToFind, textToInsert)
            modifiedText = modifiedText & item & vbNewLine
        End If
        If action = "replace text" Then
            item = Replace(item, textToFind, textToInsert)
            modifiedText = modifiedText & item & vbNewLine
        End If
        If action = "replace text after" Then
            startPosition = InStr(1, item, textToFind, vbTextCompare)
            If startPosition > 0 Then
                item = Left(item, startPosition + 1 - j) & textToInsert
            End If
            modifiedText = modifiedText & item & vbNewLine
        End If
        If action = "replace text before" Then
            startPosition = InStr(1, item, textToFind, vbTextCompare)
            If startPosition > 0 Then
                item = textToInsert & Mid(item, startPosition + j)
            End If
            modifiedText = modifiedText & item & vbNewLine
        End If
    Next
    
    ' Remove the last vbNewLine
    modifiedText = Left(modifiedText, Len(modifiedText) - 2)
    
    ' Set the document text to the modified text
    Editor.command("ecSelectAll")
    Editor.selText(modifiedText)
End Function

Sub AddStartingText
    textToInsert = InputBox("Enter the text to add to the start of each line:")
    If textToInsert = "" Then Exit Sub
    ChangeEachLine("add text at the start")
end Sub

Sub AddEndingText
    textToInsert = InputBox("Enter the text to add to the end of each line:")
    If textToInsert = "" Then Exit Sub
    ChangeEachLine("add text at the end")
end Sub

Sub RemoveFirstWord
    ChangeEachLine("remove first word")
    ChangeEachLine("remove leading spaces")
end Sub

Sub RemoveLastWord
    ChangeEachLine("remove last word")
end Sub

Sub KeepFirstWord
    ChangeEachLine("keep first word")
end Sub

Sub KeepLastWord
    ChangeEachLine("keep last word")
end Sub

Sub KeepLinesContainingText
    textToInsert = InputBox("Enter the text for filtering the lines to keep:")
    If textToInsert = "" Then Exit Sub
    ChangeEachLine("keep lines containing a given text")
end Sub

Sub RemoveLinesContainingText
    textToInsert = InputBox("Enter the text for filtering the lines to remove:")
    If textToInsert = "" Then Exit Sub
    ChangeEachLine("remove lines containing a given text")
end Sub

Sub RemoveBlankLines
    ChangeEachLine("remove blank lines")
end Sub

Sub RemoveRedundantBlankLines
    ChangeEachLine("remove redundant blank lines")
end Sub

Sub RemoveRedundantSpaces
    ChangeEachLine("remove redundant spaces")
end Sub

Sub RemoveLeadingSpaces
    ChangeEachLine("remove leading spaces")
end Sub

Sub RemoveTrailingSpaces
    ChangeEachLine("remove trailing spaces")
end Sub

Sub RemoveLeadingTrailingSpaces
    ChangeEachLine("remove leading spaces")
    ChangeEachLine("remove trailing spaces")
end Sub

Sub AddLineBreakBeforeText
    textToFind = InputBox("Enter the text before which the line should break:")
    If textToFind = "" Then Exit Sub
    textToInsert = vbNewLine & textToFind
    ChangeEachLine("replace text")
end Sub

Sub AddLineBreakAfterText
    textToFind = InputBox("Enter the text after which the line should break:")
    If textToFind = "" Then Exit Sub
    textToInsert = textToFind & vbNewLine
    ChangeEachLine("replace text")
end Sub

Sub ReplaceText
    textToFind = InputBox("Enter the text to find:")
    If textToFind = "" Then Exit Sub
    textToInsert = InputBox("Enter the text to replace it with:")
    If textToInsert = "" Then Exit Sub
    ChangeEachLine("replace text")
end Sub

Sub ReplaceLineTextAfterStringExc
    textToFind = InputBox("Enter the text string to find:")
    If textToFind = "" Then Exit Sub
    textToInsert = InputBox("Enter the text to write between the given string and the end of the line:")
    If textToInsert = "" Then Exit Sub
    j = 0
    ChangeEachLine("replace text after")
end Sub

Sub ReplaceLineTextBeforeStringExc
    textToFind = InputBox("Enter the text string to find:")
    If textToFind = "" Then Exit Sub
    textToInsert = InputBox("Enter the text to write between the beginning of the line and the given string:")
    If textToInsert = "" Then Exit Sub
    j = 0
    ChangeEachLine("replace text before")
end Sub

Sub ReplaceLineTextAfterStringInc
    textToFind = InputBox("Enter the text string to find:")
    If textToFind = "" Then Exit Sub
    textToInsert = InputBox("Enter the text to write between the given string and the end of the line:")
    If textToInsert = "" Then Exit Sub
    j = Len(textToFind)
    ChangeEachLine("replace text after")
end Sub

Sub ReplaceLineTextBeforeStringInc
    textToFind = InputBox("Enter the text string to find:")
    If textToFind = "" Then Exit Sub
    textToInsert = InputBox("Enter the text to write between the beginning of the line and the given string:")
    If textToInsert = "" Then Exit Sub
    j = Len(textToFind)
    ChangeEachLine("replace text before")
end Sub

'The sub "Init" is required to create the menu items during initialization
sub Init
    
    addMenuItem "1.1 Add text at the start of each line...","Line Wizard", "AddStartingText"
    addMenuItem "1.2 Add text at the end of each line...","Line Wizard", "AddEndingText"
    addMenuItem "1.3 Add a line break before a given text...","Line Wizard", "AddLineBreakBeforeText"
    addMenuItem "1.4 Add a line break after a given text...","Line Wizard", "AddLineBreakAfterText"
    addMenuItem "2.1 Remove the first word of each line","Line Wizard", "RemoveFirstWord"
    addMenuItem "2.2 Remove the last word of each line","Line Wizard", "RemoveLastWord"
    addMenuItem "3.1 Keep only the first word of each line","Line Wizard", "KeepFirstWord"
    addMenuItem "3.2 Keep only the last word of each line","Line Wizard", "KeepLastWord"
    addMenuItem "4.1 Keep only lines containing a given text...","Line Wizard", "KeepLinesContainingText"
    addMenuItem "4.2 Remove lines containing a given text...","Line Wizard", "RemoveLinesContainingText"
    addMenuItem "5.1 Remove blank lines","Line Wizard", "RemoveBlankLines"
    addMenuItem "5.2 Remove redundant blank lines","Line Wizard", "RemoveRedundantBlankLines"
    addMenuItem "5.3 Remove redundant spaces","Line Wizard", "RemoveRedundantSpaces"
    addMenuItem "5.4 Remove leading spaces","Line Wizard", "RemoveLeadingSpaces"
    addMenuItem "5.5 Remove trailing spaces","Line Wizard", "RemoveTrailingSpaces"
    addMenuItem "5.6 Remove leading and trailing spaces","Line Wizard", "RemoveLeadingTrailingSpaces"
    addMenuItem "6.1 Replace text between a given string (exclusive) and the end of the line...","Line Wizard", "ReplaceLineTextAfterStringExc"
    addMenuItem "6.2 Replace text between a given string (inclusive) and the end of the line...","Line Wizard", "ReplaceLineTextAfterStringInc"
    addMenuItem "6.3 Replace text between the beginning of a line and a given string (exclusive)...","Line Wizard", "ReplaceLineTextBeforeStringExc"
    addMenuItem "6.4 Replace text between the beginning of a line and a given string (inclusive)...","Line Wizard", "ReplaceLineTextBeforeStringInc"
    addMenuItem "6.5 Find and Replace text...","Line Wizard", "ReplaceText"
    addMenuItem "About","Line Wizard", "about"
    
end sub

