Attribute VB_Name = "General"
'///////////////////////////////////////////////////////////////
' This little section is nice if you are going to be setting
' the path for multiple platforms.  Windows NT does it different
' to windows 98.  This cleans up the process.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Function NormalisePath(ByVal strPath As String) As String
    If Right$(strPath, 1) = "\" Then    ' if the end of the string does have a "\"
        NormalisePath = strPath         ' Then return the string as is
    Else
        NormalisePath = strPath & "\"   ' else add a "\" to the end of the string
    End If
End Function

