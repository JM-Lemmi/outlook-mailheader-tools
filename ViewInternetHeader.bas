Sub ViewInternetHeader()
    Dim olItem As Outlook.MailItem, olMsg As Outlook.MailItem
    Dim strheader As String

    For Each olItem In Application.ActiveExplorer.Selection
        strheader = GetInetHeaders(olItem)
    
    MsgBox (strheader)
       
    Next
    Set olMsg = Nothing
End Sub

Function GetInetHeaders(olkMsg As Outlook.MailItem) As String
    ' Purpose: Returns the internet headers of a message.'
    ' Written: 4/28/2009'
    ' Author:  BlueDevilFan'
    ' //techniclee.wordpress.com/
    ' Outlook: 2007'
    Const PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    Dim olkPA As Outlook.PropertyAccessor
    Set olkPA = olkMsg.PropertyAccessor
    GetInetHeaders = olkPA.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)
    Set olkPA = Nothing
End Function

Public Function RegExpReplace(text As String, pattern As String, text_replace As String, Optional instance_num As Integer = 0, Optional match_case As Boolean = True) As String
    ' https://www.ablebits.com/office-addins-blog/excel-regex-replace/#function
    Dim text_result, text_find As String
    Dim matches_index, pos_start As Integer
 
    On Error GoTo ErrHandle
    text_result = text
    Set regEx = CreateObject("VBScript.RegExp")
 
    regEx.pattern = pattern
    regEx.Global = True
    regEx.MultiLine = True
 
    If True = match_case Then
        regEx.ignorecase = False
    Else
        regEx.ignorecase = True
    End If
 
    Set matches = regEx.Execute(text)
 
    If 0 < matches.Count Then
        If (0 = instance_num) Then
            text_result = regEx.Replace(text, text_replace)
        Else
            If instance_num <= matches.Count Then
                pos_start = 1
                For matches_index = 0 To instance_num - 2
                    pos_start = InStr(pos_start, text, matches.Item(matches_index), vbBinaryCompare) + Len(matches.Item(matches_index))
                Next matches_index
 
                text_find = matches.Item(instance_num - 1)
                text_result = Left(text, pos_start - 1) & Replace(text, text_find, text_replace, pos_start, 1, vbBinaryCompare)
            End If
        End If
    End If
 
    RegExpReplace = text_result
    Exit Function
 
ErrHandle:
    RegExpReplace = CVErr(xlErrValue)
End Function

Public Function GetStringFromPattern(search_str As String, pattern As String)
' https://stackoverflow.com/a/10904299
Dim regEx As New VBScript_RegExp_55.RegExp
Dim matches
    GetStringFromPattern = ""
    regEx.pattern = pattern
    regEx.Global = True
    If regEx.test(search_str) Then
        Set matches = regEx.Execute(search_str)
        GetStringFromPattern = matches(0).SubMatches(0)
    End If
End Function

Sub ViewMessagePath()
    Dim olItem As Outlook.MailItem, olMsg As Outlook.MailItem
    Dim strheader As String

    For Each olItem In Application.ActiveExplorer.Selection
        strheader = GetInetHeaders(olItem)
    
    lineheader = RegExpReplace(strheader, "\r\n\t", " ")
    
    Dim Table() As Variant
    Dim Sender() As Variant
    Dim Receiver() As Variant
    Dim Software()
    Dim Envelope() As Variant
    Dim Time() As Variant
    
    ' loop over lines of lineheader and extract all lines starting with "Received:" into an Array
    ' loop over that array to extract the information with Regex and put into respective arrays.
    ' from: 'from(.*)by' , by, with, for, ; timedate
    
    ' asseble Table
    ' display table
    
    Next
    Set olMsg = Nothing
End Sub
