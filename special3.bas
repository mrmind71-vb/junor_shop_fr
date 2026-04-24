Attribute VB_Name = "special3"
Public Function fnDateSales(con As ADODB.Connection) As String
Dim loctable As New ADODB.Recordset
Dim cString As String
Set loctable = cmd("SELECT dSales FROM DSALES", con).Execute
If Not loctable.EOF Then
    fnDateSales = myFormat(loctable!dSales)
End If
loctable.Close
End Function
Public Function IsDgt(ByVal strData As String) As Boolean
    ' Check if string is empty
    If Len(Trim(strData)) = 0 Then
        IsDgt = False
        Exit Function
    End If
    
    ' Pattern checks if the entire string consists only of 0-9
    ' The String(Len(strData), "#") creates a pattern like "####"
    IsDgt = (Trim(strData) Like String(Len(Trim(strData)), "#"))
End Function
Sub ItemsLook(oForm As Form, oSearch As Form, Optional pOrderNo As String = "", Optional pOffer As String = "")
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(9, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "SELECT  Top 1000 item," & _
                  "file1_10.mosm," & _
                  "fact.desca," & _
                  "file1_10.modelfact0," & _
                  "file1_10.supp," & _
                  "file1_10.desca," & _
                  "file1_10.scal," & _
                  "file1_10.color," & _
                  "PRICE," & _
                  "PRICE_2 " & _
                  "from file1_10 " & _
                  "inner join fact on fact.code = file1_10.fact" & _
                  " WHERE ISSTOP = 0 "
If pOrderNo <> "" Then
    Generalarray(1) = Generalarray(1) & " AND ITEM IN (SELECT ITEM FROM FILE6_50 WHERE DOC_NO = " & MyParn(pOrderNo) & ")"
End If

If pOffer <> "" Then
    Generalarray(1) = Generalarray(1) & " AND FILE1_10.MODELNO IN (SELECT MODELNO FROM FILE0_90 WHERE FILE0_90.DOC_NO = " & MyParn(pOffer) & ")"
End If


Generalarray(2) = "ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL "
Generalarray(3) = 12000
Generalarray(5) = False

listarray(0, 0) = "≈”„ «·’‰› ° «·„’‰⁄ ° —ﬁ„ „ÊœÌ· & „ﬁ«” "
listarray(0, 1) = "(%%FILE1_10.DESCA%% OR %%file1_10.MODELFACT%% OR %%FACT.DESCA%%)"

listarray(1, 0) = "„Ê”„ "
listarray(1, 1) = "(%%File1_10.MOSM%%)"

listarray(2, 0) = "„ﬁ«”"
listarray(2, 1) = "%%File1_10.SCAL%%"

listarray(3, 0) = "«··Ê‰"
listarray(3, 1) = "(%%File1_10.COLOR%%)"
 
GrdArray(0, 0) = "»«—ﬂÊœ"
GrdArray(0, 1) = 900

GrdArray(1, 0) = "„Ê”„"
GrdArray(1, 1) = 800

GrdArray(2, 0) = "«·„’‰⁄"
GrdArray(2, 1) = 1800

GrdArray(3, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "„ﬂ »"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·’‰›"
GrdArray(5, 1) = 4500

GrdArray(6, 0) = "«·„ﬁ«”"
GrdArray(6, 1) = 1300

GrdArray(7, 0) = "«··Ê‰"
GrdArray(7, 1) = 800

GrdArray(8, 0) = " «·”⁄—"
GrdArray(8, 1) = 1000

GrdArray(9, 0) = " ”⁄— «Êﬂ«“ÊÌ‰"
GrdArray(9, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Public Function fnBalance(pItem As String, con As ADODB.Connection, Optional pstore As String = "", Optional pDate As String = "", Optional pId As String = "") As Long
Dim cmBalance As New ADODB.command
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "ITEM", pItem)
If pstore <> "" Then aPrm = AddFlag(aPrm, "STORE", pstore)
If IsDate(pDate) Then aPrm = AddFlag(aPrm, "DATE", myFormat_sp(pDate))
If pId <> "" Then aPrm = AddFlag(aPrm, "ID", pId)

Set cmBalance = cmd("dbo.sp_balance", con, adStoredProc, aPrm)
cmBalance.Execute

fnBalance = Val(cmBalance.Parameters("@BALANCE") & "")
Set cmBalance = Nothing
End Function
Public Function IsFormOpen(ByVal FormName As String) As Boolean
    Dim frm As Form
    For Each frm In Forms
        If UCase(frm.Name) = UCase(FormName) Then
            IsFormOpen = True
            Exit Function
        End If
    Next
    IsFormOpen = False
End Function
Public Function UpdatePhones(pPhone As String, pName As String, pDate As String, con As ADODB.Connection) As Boolean
Dim cmdPhone As New ADODB.command
Set cmdPhone = cmd("dbo.sp_cust_phone", con, adStoredProc, AddFlag(Empty, "phone", pPhone))
cmdPhone.Execute

Dim aInsert As Variant
If Trim(pName) <> "" Then
    If Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Or Trim(pName) <> cmdPhone.Parameters("@DESCA") Then
        aInsert = AddFlag(aInsert, "DESCA", addstring(pName))
    End If
ElseIf (Not IsNull(cmdPhone.Parameters("@DESCA").Value)) And Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
    aInsert = AddFlag(aInsert, "DESCA", addstring(cmdPhone.Parameters("@DESCA").Value))
End If

If Not IsNull(cmdPhone.Parameters("@F_DATE").Value) Then
    If myFormat(cmdPhone.Parameters("@F_DATE").Value) < myFormat(pDate) Then
        aInsert = AddFlag(aInsert, "F_DATE", addDate(pDate))
    ElseIf Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
        aInsert = AddFlag(aInsert, "F_DATE", addDate(cmdPhone.Parameters("@F_DATE").Value))
    End If
ElseIf IsDate(pDate) Then
    aInsert = AddFlag(aInsert, "F_DATE", addDate(pDate))
End If

If Val(cmdPhone.Parameters("@COUNT") & "") <> 1 Then
    aInsert = AddFlag(aInsert, "PHONE", addstring(pPhone))
End If

If Not IsEmpty(aInsert) Then
    aInsert = AddFlag(aInsert, "isnew", "1")
    con.BeginTrans
    If Val(cmdPhone.Parameters("@COUNT") & "") > 1 Then
        con.Execute "DELETE FROM SUBCUST WHERE PHONE = " & MyParn(pPhone)
    End If
    
    If Val(cmdPhone.Parameters("@COUNT") & "") = 1 Then
        con.Execute addUpdate(aInsert, "SUBCUST", "PHONE = " & MyParn(pPhone))
    Else
        con.Execute addInsert(aInsert, "SUBCUST")
    End If
    con.CommitTrans
End If
UpdatePhones = True
Exit Function
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Public Function fnPhoneName(pPhone As String, con As ADODB.Connection) As String
Dim cmdPhone As New ADODB.command
Set cmdPhone = cmd("dbo.sp_cust_phone", con, adStoredProc, AddFlag(Empty, "phone", pPhone))
cmdPhone.Execute

fnPhoneName = cmdPhone.Parameters("@DESCA").Value & ""
End Function
Public Function IsValidMobile(ByVal strNumber As String) As Boolean
    Dim i As Integer
    Dim prefix As String
    
    ' 1. ?????? ?? ????? (??? ?? ???? 11 ????? ??????)
    If Len(strNumber) <> 11 Then
        IsValidMobile = False
        Exit Function
    End If
    
    ' 2. ?????? ?? ?? ???? ???????? ????? ???
    For i = 1 To 11
        If Not IsNumeric(Mid(strNumber, i, 1)) Then
            IsValidMobile = False
            Exit Function
        End If
    Next i
    
    ' 3. ?????? ?? ??????? (??? ?? ???? ?? 01 ?? 0 ?? 1 ?? 2 ?? 5)
    prefix = Left(strNumber, 3)
    Select Case prefix
        Case "010", "011", "012", "015"
            IsValidMobile = True
        Case Else
            IsValidMobile = False
    End Select
End Function
Public Function UpdateDiscount(pDoc_no As String, con As ADODB.Connection, Optional pDiscount_add As Double, Optional pDiscount_total As Double) As Boolean
Dim cmDiscount As New ADODB.command
Set cmDiscount = cmd("dbo.sp_offer_discount", con, adStoredProc, AddFlag(Empty, "DOC_NO", pDoc_no))
cmDiscount.Execute

If Not IsNull(cmDiscount.Parameters("@OFFER_NO").Value) Then
    If IsMissing(pDiscount_add) Then
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = " & cmDiscount.Parameters("@DISCOUNT").Value & "," & _
                    "FILE6_20H.DISCOUNT = FILE6_20H.DISCOUNT_ADD + " & cmDiscount.Parameters("@DISCOUNT").Value & "," & _
                    "FILE6_20H.IS_OFFER = " & IIf(cmDiscount.Parameters("@OFFER_NO").Value > 0, "1", "0") & _
                    " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
    Else
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = " & cmDiscount.Parameters("@DISCOUNT").Value & "," & _
                    "FILE6_20H.DISCOUNT_ADD = " & pDiscount_add & "," & _
                    "FILE6_20H.DISCOUNT = " & pDiscount_add + cmDiscount.Parameters("@DISCOUNT").Value & "," & _
                    "FILE6_20H.IS_OFFER = " & IIf(cmDiscount.Parameters("@OFFER_NO").Value > 0, "1", "0") & _
                    " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
    End If
Else
    If pDiscount_add = 0 Then
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = 0," & _
                    "FILE6_20H.DISCOUNT_Add = 0," & _
                    "FILE6_20H.DISCOUNT = " & pDiscount_total & "," & _
                    "FILE6_20H.IS_OFFER = 0" & _
                    " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
    Else
        con.Execute "UPDATE FILE6_20H " & _
                    "SET FILE6_20H.DISCOUNT_OFFER = 0," & _
                    "FILE6_20H.DISCOUNT_Add = 0," & _
                    "FILE6_20H.DISCOUNT = " & pDiscount_add & " ," & _
                    "FILE6_20H.IS_OFFER = 0" & _
                    " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
    End If
End If

'If Not IsNull(cmDiscount.Parameters("@OFFER_NO").value) Then
'    con.Execute "UPDATE FILE6_20H " & _
'                "SET FILE6_20H.DISCOUNT = FILE6_20H.DISCOUNT_ADD +  " & cmDiscount.Parameters("@DISCOUNT").value & "," & _
'                "FILE6_20H.DISCOUNT_OFFER = " & cmDiscount.Parameters("@DISCOUNT").value & "," & _
'                "FILE6_20H.IS_OFFER = " & IIf(cmDiscount.Parameters("@OFFER_NO").value > 0, "1", "0") & _
'                " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
'ElseIf bDelDiscount Then
'    con.Execute "UPDATE FILE6_20H " & _
'                "SET FILE6_20H.DISCOUNT = DISCOUNT_ADD," & _
'                "FILE6_20H.DISCOUNT_OFFER = 0," & _
'                "FILE6_20H.DISCOUNT_ADD = 0," & _
'                "FILE6_20H.IS_OFFER = 0" & _
'                " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
'End If
End Function
Public Function RoundToNearest5(ByVal Value As Double) As Double
    RoundToNearest5 = Int((Value / 5) + 0.5) * 5
End Function
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




