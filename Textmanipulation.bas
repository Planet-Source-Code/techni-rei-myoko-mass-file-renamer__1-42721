Attribute VB_Name = "Textmanipulation"
Option Explicit
'text functions that i needed to test a lot so i put in a bas file so i can use them in the immediate window

Public Function chkdir(folder As String, file As String) As String
    If Right(folder, 1) <> "\" Then chkdir = folder & "\" & file Else chkdir = folder & file
End Function
Public Function removetext(text As String, start As Long, finish As Long, Optional exclusive As Boolean = True) As String
    If exclusive = True Then
        removetext = Left(text, start - 1) & Right(text, Len(text) - finish)
    Else
        removetext = Mid(text, start, finish - start)
    End If
End Function
Public Function removebrackets(ByVal text As String, leftb As String, rightb As String) As String
    Do While InStr(text, leftb) > 0 And InStr(text, rightb) > InStr(text, leftb)
        text = removetext(text, InStr(text, leftb), InStr(text, rightb))
    Loop
    removebrackets = text
End Function

Public Sub removeindex(strarray, index As Long)
    Dim temp As Long
    For temp = index + 1 To UBound(strarray)
        strarray(temp - 1) = strarray(temp)
    Next
    If UBound(strarray) > 0 Then
        ReDim Preserve strarray(LBound(strarray) To UBound(strarray) - 1)
    Else
        ReDim strarray(0)
    End If
End Sub

Public Function islike(filter As String, expression As String) As Boolean
On Error Resume Next
Dim tempstr() As String, count As Long
If Replace(filter, ";", Empty) <> filter Then
tempstr = Split(filter, ";")
islike = False
For count = LBound(tempstr) To UBound(tempstr)
    If LCase(expression) Like LCase(tempstr(count)) Then islike = True
Next
Else
If expression Like filter Then islike = True Else islike = False
End If
End Function
Public Function removeallbutlast(ByVal text As String, char As String, replacewith As String) As String
    Dim count As Long
    For count = 1 To countchars(text, char)
        If InStr(text, char) < InStrRev(text, char) Then
            text = Left(text, InStr(text, char) - 1) & replacewith & Right(text, Len(text) - InStr(text, char))
        End If
    Next
    removeallbutlast = text
End Function
Public Function killchars(ByVal text As String, filter As String, Optional replacewith As String = Empty) As String
Dim count As Long
For count = 1 To Len(text)
    If Replace(filter, Mid(text, count, 1), Empty) <> filter Then
        text = Left(text, count - 1) & replacewith & Right(text, Len(text) - count)
    End If
Next
killchars = text
End Function
Public Function replacedoubles(ByVal text As String, char As String) As String
    Do While InStr(text, char & char) > 0
        text = Replace(text, char & char, char)
    Loop
    replacedoubles = text
End Function
Public Function removefirst(ByVal text As String, char As String) As String
Do Until Left(text, Len(char)) <> char
    text = Right(text, Len(text) - Len(char))
Loop
removefirst = text
End Function
Public Function animename(ByVal text As String) As String
Dim tempstr() As String, count As Long 'ActiveMovie. anime amv (45343433) 4 .mp3
    'Remove profanity (no one needs to see that)
    text = Replace(text, "shit", Empty, , , vbTextCompare)
    text = Replace(text, "fuck", Empty, , , vbTextCompare)

    'Remove anything inside brackets of any kind
    text = removebrackets(text, "<", ">")
    text = removebrackets(text, "(", ")")
    text = removebrackets(text, "[", "]")
    text = removebrackets(text, "{", "}")
    
    'Replace underscores and commas with spaces (for god sakes, why do ppl use them)
    text = Replace(text, "_", " ")
    text = Replace(text, ",", " ")
    
    'Replace all but the last period with spaces
    text = removeallbutlast(text, ".", " ")
    
tempstr = Split(text, " ")

Do Until count > UBound(tempstr)
    'These words aren't needed as they are implied or unneeded
    If islike("ep*", tempstr(count)) = True Then
        tempstr(count) = Replace(tempstr(count), "ep", Empty, , , vbTextCompare)
    End If
    If islike("episode;anime;amv;ova;oav;divx;volume;ep;tv;dvd;dvd*;vol;sd", tempstr(count)) = True Then
        removeindex tempstr, count
        count = count - 1
    Else
        'If a word is Ep9 or something
        If islike("ep*;ep?", tempstr(count)) = True Then
            tempstr(count) = Replace(tempstr(count), "ep", "- ", , , vbTextCompare)
        End If
        
            If IsNumeric(tempstr(count)) Then
                'Make numerical words 2 digits minimum
                If Len(tempstr(count)) = 1 Then tempstr(count) = "0" & tempstr(count)
                'Make sure theres a minus sign delimeting it
                tempstr(count) = " - " & tempstr(count)
            Else
                'make sure its a word that gets capitalized
                If islike("the;an;no;and;a;or;in;with;on;of;over;at", tempstr(count)) = False Then
                    'Capitalize first letter, lower case the rest
                    If Len(tempstr(count)) > 1 Then
                        tempstr(count) = UCase(Left(tempstr(count), 1)) & LCase(Right(tempstr(count), Len(tempstr(count)) - 1))
                    Else
                        tempstr(count) = UCase(tempstr(count))
                    End If
                End If
            End If

    End If
    If tempstr(count) = Empty Then
        removeindex tempstr, count
        count = count - 1
    End If
count = count + 1
Loop

text = Join(tempstr, " ")

    'Remove Double spaces
    text = replacedoubles(text, " ")
    
    'Remove the first char if its a space or minus
    text = removefirst(text, " ")
    text = removefirst(text, "-")
    
    'Makes sure each '-' is surrounded by spaces, and never 2 in a row
    text = Replace(text, " -", "-")
    text = Replace(text, "- ", "-")
    text = replacedoubles(text, "-")
    text = Replace(text, "-", " - ")
    
    'Remove spaces before extentions
    text = Replace(text, " .", ".")
    
    'Kill off all non alphanumeric characters
    text = killnonalpha(text)
    
    'Force numbers to be 2 digits or more
    text = stringformat(text, 2)
    
    animename = text
End Function
Public Function killnonalpha(ByVal text As String) As String
Dim temp As Long
Do Until temp >= Len(text)
    temp = temp + 1
    If isalphanumeric(Mid(text, temp, 1)) = False Then
        text = Replace(text, Mid(text, temp, 1), Empty)
    End If
Loop
killnonalpha = text
End Function
Public Function isalphanumeric(text As String) As Boolean
Const chars As String = "_.'- "
isalphanumeric = False
text = Left(LCase(text), 1)
If text >= "a" And text <= "z" Then isalphanumeric = True
If text >= "0" And text <= "9" Then isalphanumeric = True
If Replace(chars, text, Empty) <> chars Then isalphanumeric = True
End Function

Public Function countchars(text As String, char As String) As Long
    Dim count As Long, counter As Long
    counter = 0
    For count = 1 To Len(text)
        If Mid(text, count, Len(char)) = char Then counter = counter + 1
    Next
    countchars = counter
End Function
Public Function stringformat(ByVal text As String, Optional numofzeros As Long = 16) As String
    Dim count As Long, numstart As Long, tempnumber As Long, dontdo As Long
    Const append As String = "randomness"
    text = text & append 'function wont work if theres numbers on the end, just like another complex function i wrote, odd isnt it?
    
    If killchars(text, "0123456789") <> text Then
        
        Do Until count >= Len(text)
        count = count + 1
        If count > dontdo Then
            If IsNumeric(Mid(text, count, 1)) = True Then
                If numstart = 0 Then numstart = count
            Else
                If numstart > 0 And count - numstart < numofzeros Then
                    text = Left(text, numstart - 1) & Format(Mid(text, numstart, count - numstart), String(numofzeros, "0")) & Right(text, Len(text) - count + 1)
                    dontdo = Len(Left(text, numstart - 1) & Format(Mid(text, numstart, count - numstart), String(numofzeros, "0")))
                    numstart = 0
                End If
            End If
        End If
        Loop
    End If
    stringformat = Left(text, Len(text) - Len(append))
End Function
Public Function getfromquotes(text As String) As String
    If InStr(text, """") = 0 Then getfromquotes = text: Exit Function
    getfromquotes = Mid(text, InStr(text, """") + 1, InStrRev(text, """") - 1 - InStr(text, """"))
End Function
Public Function containsword(text As String, word As String) As Boolean
    containsword = Len(text) <> Len(Replace(text, word, Empty, , , vbTextCompare))
End Function
Public Function renameinbox(firstname As String, Optional newname As String) As String
    Dim middlename As String, lastname As String, foldername As String
    If firstname <> Empty Then
        foldername = Left(firstname, InStrRev(firstname, "\"))
        middlename = Right(firstname, Len(firstname) - InStrRev(firstname, "\"))
        lastname = InputBox(firstname, "Please select a new name for this file", IIf(newname = Empty, middlename, newname))
        If lastname <> Empty Then renameinbox = foldername & lastname
    End If
End Function
Public Function movetofolder(filename As String, Optional hwnd As Long) As String
    Dim middlename As String, newfolder As String
    newfolder = BrowseForFolder(hwnd, "Please select the destination")
    If newfolder = Empty Then movetofolder = filename: Exit Function
    middlename = Right(filename, Len(filename) - InStrRev(filename, "\"))
    movetofolder = uniquefilename(chkdir(newfolder, middlename))
End Function
Public Function uniquefilename(filename As String) As String
    Dim temp1 As String, temp2 As String, temp3 As Long
    uniquefilename = filename
    
    If FileExists(filename) Then
        Dim count As Long
        count = 1
        temp3 = InStrRev(filename, ".")
        temp1 = filename
        If temp3 > 0 Then
            temp1 = Left(filename, temp3 - 1)
            temp2 = Right(filename, Len(filename) - temp3 + 1)
        End If
        Do Until FileExists(temp1 & "(" & count & ")" & temp2) = False
            count = count + 1
        Loop
        uniquefilename = temp1 & "(" & count & ")" & temp2
    End If
End Function
Public Function FileExists(filename As String) As Boolean
On Error Resume Next 'Checks to see if a file exists
FileExists = Dir(filename) <> Empty
End Function
Public Function isadir(text As String) As Boolean
isadir = InStr(text, ":\") > 0
End Function
