VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BAIXA 
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   OleObjectBlob   =   "BAIXA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BAIXA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indice As Variant
Dim numero_primeiro_certificado As Variant
Private Sub pesquisa_Click()

If txt_Numero_Pedido = "" Then
    MsgBox ("Campo vazio! Necessário preencher o número do pedido!"), vbCritical
    Call Limpar
    Exit Sub
ElseIf txt_Numero_Pedido = Null Then
    MsgBox ("Campo vazio! Necessário preencher o número do pedido!"), vbCritical
    Call Limpar
    Exit Sub
ElseIf txt_Numero_Pedido = vbNullString Then
    MsgBox ("Campo vazio! Necessário preencher o número do pedido!"), vbCritical
    Call Limpar
    Exit Sub
ElseIf Len(txt_Numero_Pedido) <> 13 Then
    MsgBox ("Número do pedido incorreto!"), vbCritical
    Call Limpar
    Exit Sub
Else
    Call Procurar
End If

End Sub
Private Sub Procurar()

Application.ScreenUpdating = False

Dim numeroPedido As Variant

numeroPedido = txt_Numero_Pedido.Text
    
    With Sheets("Banco de Dados").Range("F6:F40000")
        Set primeiro_certificado = Range("F6:F40000").Find(numeroPedido, _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlNext, _
                                MatchCase:=False)
                                
        Set ultimo_certificado = Range("F6:F40000").Find(numeroPedido, _
                                LookIn:=xlValues, _
                                LookAt:=xlWhole, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlPrevious, _
                                MatchCase:=False)
    
        If Not primeiro_certificado Is Nothing Then
            Application.Goto primeiro_certificado, True
        Else
            MsgBox "Não Encontrado!", vbCritical
            Call Limpar
            Exit Sub
        End If
    End With
        
primeiro = Range(Replace(primeiro_certificado.Address, "F", "B")).Value
ultimo = Range(Replace(ultimo_certificado.Address, "F", "B")).Value

numero_primeiro_certificado = Left(primeiro, Len(primeiro) - 5)
numero_ultimo_certificado = Left(ultimo, Len(ultimo) - 5)

If numero_primeiro_certificado = numero_ultimo_certificado Then
    txt_Numero_Certificado = numero_primeiro_certificado
Else
    txt_Numero_Certificado = numero_primeiro_certificado & "/" & numero_ultimo_certificado
End If

txt_Nome = Range(Replace(primeiro_certificado.Address, "F", "N")).Value
txt_Email = Range(Replace(primeiro_certificado.Address, "F", "AD")).Value
txt_Hoje = Date

ActiveWindow.ScrollColumn = 1
Application.ScreenUpdating = True

End Sub
Private Sub button_Imprimir_Click()

Application.ScreenUpdating = False
    
    numero_primeiro_certificado = numero_primeiro_certificado + 5
    
    For i = numero_primeiro_certificado To 30
      
        numeroPedido = "F" & i
        celulaBaixa = "AB" & i
        numeroCertificado = "B" & i
        
        If Range(numeroPedido).Value = txt_Numero_Pedido.Text Then
    
            Sheets("Banco de Dados").Range(celulaBaixa).Value = txt_Hoje.Text
            
            celulaNome = Sheets("Enviar E-mail").Range("A200").End(xlUp).Offset(1, 0).Address
            celulaNumero = Sheets("Enviar E-mail").Range("B200").End(xlUp).Offset(1, 0).Address
            celulaCertificados = Sheets("Enviar E-mail").Range("C200").End(xlUp).Offset(1, 0).Address
            celulaEmail = Sheets("Enviar E-mail").Range("D200").End(xlUp).Offset(1, 0).Address
            
            Sheets("Enviar E-mail").Range(celulaNome).Value = txt_Nome.Text
            Sheets("Enviar E-mail").Range(celulaNumero).Value = txt_Numero_Pedido.Text
            Sheets("Enviar E-mail").Range(celulaCertificados).Value = Range(numeroCertificado).Value
            Sheets("Enviar E-mail").Range(celulaEmail).Value = txt_Email.Text
            
        End If

    Next
    
    Unload Me
    
Application.ScreenUpdating = True

End Sub
Private Sub button_Limpar_Click()

Limpar

End Sub
Private Sub Cancelar_Click()

Unload Me

End Sub
Private Sub txt_Numero_Pedido_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

Call pesquisa_Click

End Sub
Private Sub txt_Numero_Pedido_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

 If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
    
End Sub
Private Sub UserForm_Initialize()

button_Imprimir.Enabled = False

End Sub
Private Sub Limpar()

txt_Nome.Text = ""
txt_Numero_Pedido.Text = ""
txt_Numero_Certificado.Text = ""
txt_Hoje.Text = ""
txt_Email.Text = ""

End Sub
Private Sub txt_Hoje_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
txt_Hoje.MaxLength = 10

Select Case KeyAscii

    Case 8
    Case 13: SendKeys "{TAB}"
    Case 48 To 57
    
    If txt_Hoje.SelStart = 2 Then txt_Hoje.SelText = "/"
    If txt_Hoje.SelStart = 5 Then txt_Hoje.SelText = "/"
    Case Else: KeyAscii = 0
    
End Select
    
End Sub
Private Sub txt_Numero_Certificado_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
txt_Numero_Certificado.MaxLength = 11

Select Case KeyAscii

    Case 8
    Case 13: SendKeys "{TAB}"
    Case 48 To 57
    
    If txt_Numero_Certificado.SelStart = 5 Then txt_Numero_Certificado.SelText = "/"
    Case Else: KeyAscii = 0
    
End Select
    
End Sub
Public Sub Valida()

If txt_Nome.Text = Empty Or txt_Numero_Pedido.Text = Empty Or txt_Numero_Certificado.Text = Empty Or txt_Hoje.Text = Empty Or txt_Email = Empty Then
    button_Imprimir.Enabled = False
    
ElseIf txt_Nome.Text <> Empty Or txt_Numero_Pedido.Text <> Empty Or txt_Numero_Certificado.Text <> "" Or txt_Hoje.Text <> Empty Or txt_Email.Text <> Empty Then
    button_Imprimir.Enabled = True
    
End If

End Sub
Private Sub txt_Hoje_Change()

Valida

End Sub
Private Sub txt_Numero_Pedido_Change()

Valida

End Sub
Private Sub txt_Numero_Certificado_Change()

Valida

End Sub
Private Sub txt_Nome_Change()

Valida

End Sub
Private Sub txt_Email_Change()

Valida

End Sub
