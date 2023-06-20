Attribute VB_Name = "Z"
Sub CopyAttachments(objSourceItem, objTargetItem, MyKind As Integer)
    
    Const PR_ATTACH_MIME_TAG = "http://schemas.microsoft.com/mapi/proptag/0x370E001E"
    Const PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
    Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
    Const PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003"
    Const PR_ATTACH_CONTENT_LOCATION = "http://schemas.microsoft.com/mapi/proptag/0x3713001E"
    Const PR_ATTACH_METHOD = "http://schemas.microsoft.com/mapi/proptag/0x37050003"
    
    Dim FSO
    Dim fldTemp
    Dim strPath As String, strFile As String
    Dim objAtt
    Dim ObjAttDest
   'Dim MyItem
   'Dim MyAttachments
    Dim pa As PropertyAccessor
    Dim c As Integer
    Dim cid As String
    Dim body As String
    Dim test

    body = objSourceItem.HTMLBody
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set fldTemp = FSO.GetSpecialFolder(2) ' TemporaryFolder
    strPath = fldTemp.Path & "\"
    For Each objAtt In objSourceItem.Attachments
        Set pa = objAtt.PropertyAccessor
        cid = pa.GetProperty(PR_ATTACH_CONTENT_ID)

        If Len(cid) > 0 Then
            If InStr(body, cid) Then
                If MyKind = 1 Or MyKind = 3 Then
                    If pa.GetProperty(PR_ATTACHMENT_HIDDEN) Then
                        strFile = strPath & objAtt.FileName
                        objAtt.SaveAsFile strFile
                        Set ObjAttDest = objTargetItem.Attachments.Add(strFile, olByValue, 0, objAtt.DisplayName)
                        ObjAttDest.PropertyAccessor.SetProperty PR_ATTACH_MIME_TAG, pa.GetProperty(PR_ATTACH_MIME_TAG)
                        ObjAttDest.PropertyAccessor.SetProperty PR_ATTACH_CONTENT_ID, pa.GetProperty(PR_ATTACH_CONTENT_ID)
                        ObjAttDest.PropertyAccessor.SetProperty PR_ATTACHMENT_HIDDEN, True
                        
                        objTargetItem.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/8514000B", True
                        FSO.DeleteFile strFile
                        objTargetItem.Save
                    End If
                End If
            Else
                If MyKind = 2 Or MyKind = 3 Then

                    'In case that PR_ATTACHMENT_HIDDEN does not exists,
                    'an error will occur. We simply ignore this error and
                    'treat it as false.
                    On Error Resume Next
                    If Not pa.GetProperty(PR_ATTACHMENT_HIDDEN) Then
                        strFile = strPath & objAtt.FileName
                        objAtt.SaveAsFile strFile
                        objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
                        FSO.DeleteFile strFile
                    End If
                    On Error GoTo 0
                End If
            End If
        Else
            strFile = strPath & objAtt.FileName
            objAtt.SaveAsFile strFile
            objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
            FSO.DeleteFile strFile
        End If
    Next
 
   Set fldTemp = Nothing
   Set FSO = Nothing
End Sub

Function StringToBase36(str As String) As String
    Dim strLen As Integer: strLen = Len(str)
    Dim result1 As Double, result2 As Double
    Dim i As Integer
    For i = 1 To strLen
        Dim charCode As Integer: charCode = Asc(Mid(str, i, 1))
        If charCode >= 48 And charCode <= 57 Then ' Caractere numérico
            result1 = result1 * 10 + (charCode - 48)
        ElseIf charCode >= 65 And charCode <= 90 Then ' Caractere alfabético maiúsculo
            result1 = result1 * 36 + (charCode - 55)
        ElseIf charCode >= 97 And charCode <= 122 Then ' Caractere alfabético minúsculo
            result1 = result1 * 36 + (charCode - 87)
        End If
        If i = strLen - 5 Then
            result2 = result1
            result1 = 0
        End If
    Next i
    If result2 = 0 Then
        StringToBase36 = Hex(result1)
    Else
        StringToBase36 = Hex(result2) & Hex(result1)
    End If
End Function

' ###################################################################################################
' FUNÇÃO MoveToFolder
' RESPONSÁVEL POR MOVIMENTAR UM EMAIL DA CAIXA DE ENTRADA PARA ALGUMA PASTA ESCOLHIDA
'
' objMail = EMAIL QUE SERÁ MOVIMENTADO
' nameFolder = NOME DA PASTA ONDE DEVERÁ SER MOVIDO O EMAIL
' emailToMove = ENDEREÇO DE EMAIL QUE SERÁ CONSIDERADO PARA A BUSCA DO NOME DA PASTA
'
Function MoveToFolder(objMail, nameFolder As String, emailToMove As String) As String

    Dim objFolder As Outlook.Folder
    Dim oAccount As account
    Dim objNamespace As Outlook.NameSpace
    Set objNamespace = GetNamespace("MAPI")
        
    For Each oAccount In Session.Accounts
        For Each fldr In objNamespace.Folders
            For Each foldersOutlook In fldr.Folders
    
                If fldr.Name = emailToMove Then
                    If foldersOutlook = nameFolder Then
                    
                        objMail.UnRead = False
                        objMail.Move foldersOutlook
                        
                        Exit Function
                        
                    End If
                End If
                
            Next
        Next
    Next

End Function
' ###################################################################################################
