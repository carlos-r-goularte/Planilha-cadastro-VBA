VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ETIQUETA 
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   OleObjectBlob   =   "ETIQUETA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ETIQUETA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
txt_Data = Range(Replace(primeiro_certificado.Address, "F", "AA")).Value
taxa_urgencia = Range(Replace(primeiro_certificado.Address, "F", "Z")).Value


ActiveWindow.ScrollColumn = 1
Application.ScreenUpdating = True

End Sub
Private Sub button_Imprimir_Click()

Application.ScreenUpdating = False
    
    Sheets("Plan2").Range("B1").Value = txt_Nome.Text
    Sheets("Plan2").Range("B3").Value = txt_Numero_Pedido.Text
    Sheets("Plan2").Range("C4").Value = txt_Numero_Certificado.Text
    Sheets("Plan2").Range("C5").Value = txt_Data.Text
    Sheets("Plan2").Range("E4").Value = taxa_urgencia.Text
    

   Sheets("Plan2").PrintOut from:=1, to:=1, copies:=1, ActivePrinter:="ZDesigner ZD230-203dpi ZPL"
    
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

Dim taxa(1 To 2) As String

taxa(1) = "Sim"
taxa(2) = "Não"

taxa_urgencia.List = taxa

button_Imprimir.Enabled = False

End Sub
Private Sub Limpar()

txt_Nome.Text = ""
txt_Numero_Pedido.Text = ""
txt_Numero_Certificado.Text = ""
txt_Data.Text = ""
taxa_urgencia.Text = ""

End Sub
Private Sub txt_Data_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
txt_Data.MaxLength = 10

Select Case KeyAscii

    Case 8
    Case 13: SendKeys "{TAB}"
    Case 48 To 57
    
    If txt_Data.SelStart = 2 Then txt_Data.SelText = "/"
    If txt_Data.SelStart = 5 Then txt_Data.SelText = "/"
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

If txt_Nome.Text = Empty Or txt_Numero_Pedido.Text = Empty Or txt_Numero_Certificado.Text = Empty Or txt_Data.Text = Empty Or taxa_urgencia = Empty Then
    button_Imprimir.Enabled = False
    
ElseIf txt_Nome.Text <> Empty Or txt_Numero_Pedido.Text <> Empty Or txt_Numero_Certificado.Text <> "" Or txt_Data.Text <> Empty Or taxa_urgencia.Text <> Empty Then
    button_Imprimir.Enabled = True
    
End If

End Sub
Private Sub txt_Data_Change()

Valida

End Sub
Private Sub txt_Numero_Pedido_Change()

Valida

End Sub
Private Sub taxa_urgencia_Change()

Valida

End Sub
Private Sub txt_Numero_Certificado_Change()

Valida

End Sub
Private Sub txt_Nome_Change()

Valida

End Sub

