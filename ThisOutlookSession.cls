VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private WithEvents sentItems As Outlook.items
Attribute sentItems.VB_VarHelpID = -1
        
'##############################################################################################################################################################
'######## TODO E-MAIL QUE CHEGAR NESSE ENDEREÇO NO ENDEREÇO primaryAccountEmail ###############################################################################
'######## SERÁ MOVIMENTADO PARA A PASTA QUE SE ENCONTRA NA RAIZ, NOMEADA COMO namePrimaryFolderMove ###########################################################
'######## SERÁ RESPONDIDO COM A MENSAGEM AUTOMÁTICA automaticMessage1 and automaticMessage2 ###################################################################
'##############################################################################################################################################################
'######## CASO O REMETENTE SEJA O primaryAccountEmail O EMAIL SERÁ MOVIMENTADO PARA A PASTA QUE SE ENCONTRA NA RAIZ, NOMEADA COMO nameArchiveFolderMove #######
'##############################################################################################################################################################

Const primaryAccountEmail As String = "SUBSTITUA-AQUI@outlook.com.br"
Const namePrimaryFolderMove As String = "SUBSTITUA-AQUI-COM-O-NOME-DA-PASTA"
Const nameArchiveFolderMove As String = "SUBSTITUA-AQUI-COM-O-NOME-DA-PASTA"
Const automaticMessage1 As String = "Obrigado por entrar em contato conosco. Este é um e-mail automático para confirmar que recebemos sua mensagem. Iremos responder assim que possível."
Const automaticMessage2 As String = "Atenciosamente, SUBSTITUA-AQUI-NOME-DA-EMPRESA."

'#########################################################################################################
'######## ESSA FUNÇÃO É COMUM PARA UTILIZAÇÃO EM GRUPOS QUE DIRECIONAM UM E-MAIL PARA O PRINCIPAL ########
'#########################################################################################################
'######## TODO E-MAIL QUE CHEGAR NESSE ENDEREÇO NO ENDEREÇO secondAccountEmail ###########################
'######## SERÁ MOVIMENTADO PARA A PASTA QUE SE ENCONTRA NA RAIZ, NOMEADA COMO nameSecondFolderMove #######
'######## PARA DESABILITAR, DEFINA functionMoveSecond COMO false #########################################
'#########################################################################################################

Const secondAccountEmail As String = "SUBSTITUA-AQUI-CASO-PRECISE@outlook.com.br"
Const nameSecondFolderMove As String = "SUBSTITUA-AQUI-COM-O-NOME-DA-PASTA"
Const functionMoveSecond As Boolean = True

'###################################################################################################

Function setEvent()

    Dim objFolder As Outlook.Folder
    Dim objNamespace As Outlook.NameSpace
    Set objNamespace = GetNamespace("MAPI")

    Dim oAccount As account
    For Each oAccount In Session.Accounts
        For Each fldr In objNamespace.Folders
            For Each foldersOutlook In fldr.Folders
                If foldersOutlook = "Itens Enviados" Then
                    If fldr.Name = primaryAccountEmail Then
                        Set sentItems = foldersOutlook.items
                        
                        Set objNamespace = Nothing
                        Set objFolder = Nothing
                        Set oAccount = Nothing
                        
                        Exit Function
                    End If
                End If
            Next
        Next
    Next

End Function

Private Sub sentItems_ItemAdd(ByVal Item As Object)

    If TypeOf Item Is Outlook.MailItem Then
        Dim objMail As Outlook.MailItem
        Set objMail = Item

        If Left(objMail.body, Len(automaticMessage1)) = automaticMessage1 Then
        
            Dim objFolder As Outlook.Folder
            Dim objNamespace As Outlook.NameSpace
            Set objNamespace = GetNamespace("MAPI")
        
            Dim oAccount As account
            For Each oAccount In Session.Accounts
                For Each fldr In objNamespace.Folders
                    If fldr.Name = primaryAccountEmail Then
                        For Each foldersOutlook In fldr.Folders
                            If foldersOutlook = "Caixa de Entrada" Then
                                
                                objMail.UnRead = True
                                objMail.Move foldersOutlook
                                
                                Set objNamespace = Nothing
                                Set objFolder = Nothing
                                Set oAccount = Nothing
                                
                                Exit Sub
                            End If
                        Next
                    End If
                Next
            Next
        End If
    End If
End Sub

Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)


    Dim EntryID As Variant
    For Each EntryID In Split(EntryIDCollection, ",")
    
        
        '####### SETANDO O E-MAIL QUE SERÁ TRABALHADO ########
        Set objMail = Application.Session.GetItemFromID(EntryID)
        
        '############################################################
        '######## DEFININDO QUAL ENDEREÇO QUE GEROU A MACRO #########
        '############################################################
        
        Debug.Print "----------- SESSÃO A ----------------"
        Debug.Print "DEFINIDO ENDEREÇO QUE DISPAROU MACRO"
        Debug.Print "#####################################"
        
        Dim emailAdressApplication As String
        Dim itm As Object
        
        Dim objNamespace As Outlook.NameSpace
        Set objNamespace = GetNamespace("MAPI")
        
        Dim arr() As String
        arr = Split(EntryIDCollection, ",")
        
        For i = 0 To UBound(arr)
        
            Set itm = objNamespace.GetItemFromID(arr(i))
            If itm.Class = olMail Then
                emailAdressApplication = itm.Parent.Parent.Name
            End If
        
        Next
        
        Debug.Print "ENDEREÇO:" & emailAdressApplication
        Debug.Print "-------------------------------------"
            
        '####### CASO O ENDEREÇO QUE GEROU A MACRO SEJA O ENDEREÇO DEFINIDO COMO PRIMARIO #######
        If InStr(1, emailAdressApplication, primaryAccountEmail, vbTextCompare) > 0 Then
                         
            '###########################################################
            '####### DEFININDO OS ENDEREÇOS DOS EMAILS COPIADOS ########
            '###########################################################
            
            Debug.Print "----------- SESSÃO B ----------------"
            Debug.Print "DEFINIDO LISTA DE 'PARA' E EM 'COPIA'"
            Debug.Print "#####################################"
            
            Dim recipients As Outlook.recipients
            Dim recipient As Outlook.recipient
            Set recipients = objMail.recipients
            
            Dim ccList As String
            Dim toList As String
            
            ccList = ""
            toList = ""
            
            For Each recipient In recipients
                
                Dim addressTemp As String
                addressTemp = ""
                
                If recipient.AddressEntry.Address <> "" Then
                    If Not recipient.AddressEntry Is Nothing Then
                        If recipient.AddressEntry.AddressEntryUserType = olExchangeDistributionListAddressEntry Then
                            addressTemp = recipient.AddressEntry.GetExchangeDistributionList.PrimarySmtpAddress
                        ElseIf recipient.AddressEntry.AddressEntryUserType = OlAddressEntryUserType.olExchangeUserAddressEntry Then
                            addressTemp = recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
                        Else
                            addressTemp = recipient.Address
                        End If
                    End If
                End If
                
                If recipient.Type = olCC Then ccList = ccList & addressTemp & "; "
                If recipient.Type = olTo Then toList = toList & addressTemp & "; "
                
            Next recipient
            
            Debug.Print "ENVIADO PARA " & toList
            Debug.Print "COPIADO PARA " & ccList
            Debug.Print "-------------------------------------"
              
            '#####################################################
            '############## DEFININDO O EMITENTE #################
            '#####################################################
            
            Debug.Print "--------- SESSÃO C ---------"
            Debug.Print "DEFINIDO QUEM ENVIOU O EMAIL"
            Debug.Print "############################"
            
            If objMail.Sender.AddressEntryUserType = olExchangeUserAddressEntry Then
                SenderAddress = objMail.Sender.GetExchangeUser().PrimarySmtpAddress
            Else
                SenderAddress = objMail.Sender.Address
            End If
            
            Debug.Print "ENVIADO POR: " & SenderAddress
            Debug.Print "---------------------------"
            
            '#####################################################
            '################# MOVENDO O EMAIL ###################
            '#####################################################
            
            Debug.Print "----------- SESSÃO D ----------------"
            Debug.Print "RETIRANDO EMAIL DA CAIXA DE ENTRADA"
            Debug.Print "#####################################"
            
            If InStr(1, SenderAddress, primaryAccountEmail, vbTextCompare) > 0 Then
            
                Call Z.MoveToFolder(objMail, nameArchiveFolderMove, primaryAccountEmail)
                Debug.Print "MOVIMENTANDO PARA RESOLVIDOS"
                Debug.Print "-------------------------------------"
                    
            ElseIf InStr(1, toList & ccList, primaryAccountEmail, vbTextCompare) > 0 Then
                
                '####### CASO O EMAIL RECEBIDO NÃO TENHA UM NÚMERO DE TICKET AINDA, GERA UM NOVO E MOVIMENTA#######
                If InStr(1, objMail.Subject, "[TICKET #", vbTextCompare) = 0 Then
                
                    Call Z.MoveToFolder(objMail, namePrimaryFolderMove, primaryAccountEmail)
                    Debug.Print "MOVIMENTANDO PARA HISTÓRICO"
                    Debug.Print "-------------------------------------"
                    
                    Call generateTicket(objMail, toList, ccList)
                End If
                
            ElseIf functionMoveSecond And InStr(1, toList & ccList, secondAccountEmail, vbTextCompare) > 0 Then
                
                Call Z.MoveToFolder(objMail, nameSecondFolderMove, primaryAccountEmail)
                Debug.Print "MOVIMENTANDO PARA CADASTRO"
                Debug.Print "-------------------------------------"
               
            End If
            
        End If
        
    Next EntryID
    
End Sub

Function generateTicket(objMail, toList As String, ccList As String)

    Dim objReply As Outlook.MailItem
    Set objReply = objMail.ReplyAll
    
    '#####################################################
    '############# DEFININDO QUEM ENVIOU #################
    '#####################################################
                            
    Dim SenderAddress As String
    
    If objMail.Sender.AddressEntryUserType = olExchangeUserAddressEntry Then
        SenderAddress = objMail.Sender.GetExchangeUser().PrimarySmtpAddress
    Else
        SenderAddress = objMail.Sender.Address
    End If
    
    '#####################################################
    '########## CONSTRUINDO O CORPO DO EMAIL #############
    '#####################################################
    
    Dim sBody As String
    sBody = "<p style='margin: 0;'><font face='Calibri' style='font-size: 15px;'>" & automaticMessage1 & "<br>" & automaticMessage2 & "</font></p><br></br>" & _
            "<hr style='border:1px solid #ccc'>" & _
            "<p style='margin: 0;'><font face='Calibri' style='font-size: 15px;'><b>De: </b> " & SenderAddress & "<br>" & _
            "<b>Enviada em: </b>" & Format(objMail.ReceivedTime, "dddd, d \de MMMM \de yyyy HH:mm") & "<br>"
    
    If Len(toList) > 0 Then
        sBody = sBody & "<b>Para:</b> " & toList & "<br>"
    End If
    
    If Len(ccList) > 0 Then
        sBody = sBody & "<b>Cc:</b> " & ccList & "<br>"
    End If
    
    sBody = sBody & "<b>Assunto: </b>" & objMail.Subject & "<br><br></font></p>" & _
            "<p style='margin-bottom:5px;'><font face=""Akko W1G"" size=""10pt"">" & objMail.HTMLBody
            
    '#####################################################
    '################# ENVIANDO O EMAIL ##################
    '#####################################################
    
    Call setEvent
    
    objReply.Subject = "RE: [TICKET #" & y.CompactarCodigo(CStr(Replace(Format(Now, "yymmddhhnnss"), " ", ""))) & "] " & objMail.Subject
    objReply.HTMLBody = sBody
    Z.CopyAttachments objMail, objReply, 3
    objReply.Send

End Function






