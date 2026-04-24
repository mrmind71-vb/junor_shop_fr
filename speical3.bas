Attribute VB_Name = "Sepcial"
Function addInsertUpdate(aInsert As Variant, pTable As String, pCondition As String, pFieldName As String)
Dim cInsert As New ChilkatStringBuilder, cUpdate As New ChilkatStringBuilder

cInsert.Append "INSERT INTO " & pTable & "("
For i = 0 To UBound(aInsert) Step 2
    cInsert.Append aInsert(i) & ","
Next
cInsert.Shorten 1
cInsert.AppendLine ")", 1

cInsert.Append "VALUES("
For i = 0 To UBound(aInsert) Step 2
    cInsert.Append aInsert(i + 1) & ","
Next
cInsert.Shorten 1
cInsert.Append ")"

cUpdate.AppendLine "UPDATE " & pTable & " SET ", 1
For i = 0 To UBound(aInsert) Step 2
    cUpdate.Append aInsert(i) & " = " & aInsert(i + 1) & ","
Next
cUpdate.Shorten 1
cUpdate.AppendLine "", 1
If pCondition <> "" Then
    cUpdate.Append "WHERE " & pCondition
End If

Dim sb As New ChilkatStringBuilder
sb.AppendLine "IF dbo.DOC_" & pTable & "(" & pFieldName & ") = 0", 1
'sb.append "BEGIN"
sb.AppendLine cInsert.GetAsString, 1
'sb.append "END"
sb.AppendLine "ELSE ", 1
'sb.append "BEGIN "
sb.Append cUpdate.GetAsString
'sb.Append "END"

addInsertUpdate = sb.GetAsString()
End Function

