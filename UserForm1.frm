VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16050
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Planilha1.ExportAsFixedFormat


End Sub


Private Sub CommandButton3_Click()
    If MsgBox("Deseja mesmo sair?", vbYesNo + vbQuestion, "Sair") = vbYes Then
        End
    Else
       Cancel = True
    End If
End Sub

Private Sub CommandButton4_Click()
ListBox1.RowSource = Range("a1").CurrentRegion.Address
End Sub

Private Sub exportar_Click()
Application.ScreenUpdating = False
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:="C:\PDF\Export.pdf", _
            OpenAfterPublish:=False
    Application.ScreenUpdating = True

End Sub


Private Sub CommandButton6_Click()

End Sub

Private Sub CommandButton7_Click()

End Sub



Private Sub EXPORTARPDF_Click()
Application.ScreenUpdating = False
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:="C:\PDF\Export.pdf", _
            OpenAfterPublish:=False
    Application.ScreenUpdating = True
End Sub

Private Sub EXPORTARTXT_Click()
Application.ScreenUpdating = False
    ActiveSheet.ExportAsFixedFormat Type:=xlTypeTxt, _
            Filename:="C:\txt\Export.txt", _
            OpenAfterPublish:=False
    Application.ScreenUpdating = True
End Sub

Private Sub EXPORTARXLS_Click()
Application.ScreenUpdating = False
    ActiveSheet.ExportAsFixedFormat Type:=xlTypeXls, _
            Filename:="C:\xls\Export.xls", _
            OpenAfterPublish:=False
    Application.ScreenUpdating = True
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
caminhoArquivo = Application.GetOpenFilename(FileFilter:="Image Files(*.jpg), *.jpg")
    Me.Image1.Picture = download1.jpg
End Sub

Private Sub LIMPAR_Click()
TextBox2.Text = ""

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub TextBox10_Change()
Planilha1.Range("Z2") = TextBox10
Call filtro
End Sub

Private Sub TextBox11_Change()
Planilha1.Range("Y2") = TextBox11
Call filtro
End Sub

Private Sub TextBox12_Change()
Planilha1.Range("AD2") = TextBox12
Call filtro
End Sub

Private Sub TextBox13_Change()
Planilha1.Range("AF2") = TextBox13
Call filtro
End Sub

Private Sub TextBox14_Change()
Planilha1.Range("AE2") = TextBox14
Call filtro
End Sub

Private Sub TextBox15_Change()
Planilha1.Range("AC2") = TextBox15
Call filtro
End Sub

Private Sub TextBox16_Change()
Planilha1.Range("AG2") = TextBox16
Call filtro
End Sub

Private Sub TextBox17_Change()
Planilha1.Range("AH2") = TextBox17
Call filtro
End Sub

Private Sub TextBox18_Change()
Planilha1.Range("AI2") = TextBox18
Call filtro
End Sub

Private Sub TextBox19_Change()
Planilha1.Range("AJ2") = TextBox19
Call filtro
End Sub

Private Sub TextBox2_Change()
Planilha1.Range("U2") = TextBox2
Call filtro
End Sub

Private Sub TextBox5_Change()
Planilha1.Range("X2") = TextBox5
Call filtro
End Sub

Private Sub TextBox6_Change()
Planilha1.Range("W2") = TextBox6
Call filtro
End Sub

Private Sub TextBox7_Change()
Planilha1.Range("V2") = TextBox7
Call filtro
End Sub

Private Sub TextBox8_Change()
Planilha1.Range("AB2") = TextBox8
Call filtro
End Sub

Private Sub TextBox9_Change()
Planilha1.Range("AA2") = TextBox9
Call filtro
End Sub

Private Sub UserForm_Click()
Dim base As Range
Dim nome As String

Dim l As Long
l = Planilha1.Range("A2").CurrentRegion.Rows.Count
Set base = _
Planilha1.Range(Planilha1.Cells(3000, 1), Planilha1.Cells(l, 16))


nome = "'" & Planilha.Name & "'!"

ListBox1.RowSource = nome & base.Address
ListBox1.ColumnCount = 5
Planilha1.Range("U2:AG2").ClearContents
ListBox1.ColumnHeads = True


End Sub
