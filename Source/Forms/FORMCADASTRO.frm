VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FORMCADASTRO 
   Caption         =   "CADASTRO DE CLIENTES"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13215
   OleObjectBlob   =   "FORMCADASTRO.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FORMCADASTRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBCadastro_Click()

Application.ScreenUpdating = False
    
For X = 5 To 999999

    If Plan1.Cells(X, 1) = txtCOD.Text Then
    
    MsgBox ("Este código de cliente já foi cadastrado"), vbCritical
    txtCOD.SetFocus
    CBCadastro.Enabled = False
    
    Exit Sub
    
    ElseIf Plan1.Cells(X, 1) = "" Then
        Exit For
    End If

Next

Plan1.Cells(X, 1) = Me.txtCOD.Text
Plan1.Cells(X, 2) = Me.txtCLIENTE.Text
Plan1.Cells(X, 3) = Me.txtEND.Text

Application.ScreenUpdating = True

Ordenar2

MsgBox ("CLIENTE CADASTRADO COM SUCESSO")

Unload Me

End Sub
Private Sub CBCancelar_Click()

Ordenar2

Unload Me

End Sub
Private Sub CBLimpar_Click()

Me.txtCLIENTE = " "
Me.txtCOD = " "
Me.txtEND = " "
CBCadastro.Enabled = False

End Sub
Private Sub txtCOD_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

KeyAscii = SoNumeros(KeyAscii)

End Sub
Private Function SoNumeros(X As IReturnInteger)

Select Case X
Case Asc("0") To Asc("9")
SoNumeros = X
Case Else
SoNumeros = 0
MsgBox "Favor inserir somente números", vbCritical, "Campo tipo numérico"
txtCOD.SetFocus
End Select

End Function
Private Sub UserForm_Initialize()

CBCadastro.Enabled = False

Ordenar2

End Sub
Public Sub Valida()

If txtCOD.Text = Empty Or txtCLIENTE.Text = Empty Or txtEND.Text = Empty Then
CBCadastro.Enabled = False

ElseIf txtCOD.Text <> Empty Or txtCLIENTE.Text <> Empty Or txtEND.Text <> Empty Then
CBCadastro.Enabled = True

End If

End Sub
Private Sub txtCLIENTE_Change()

Valida

End Sub
Private Sub txtCOD_Change()

Valida

End Sub
Private Sub txtEND_Change()

Valida

End Sub
