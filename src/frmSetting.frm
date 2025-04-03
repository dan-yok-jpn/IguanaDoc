VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetting 
   Caption         =   "Setting"
   ClientHeight    =   2580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5250
   OleObjectBlob   =   "frmSetting.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Update As Boolean

Private Sub CommandButton1_Click()

    Dim file_path As String

    With Application.FileDialog(msoFileDialogFilePicker)
        With .Filters
            .Clear
            .Add "excutable", "*.exe"
        End With
        If .Show = -1 Then
            file_path = .SelectedItems(1)
            TextBox1.Text = file_path
            .Filters.Clear
        Else
            MsgBox "Install TeX2imgc first, and try again."
            .Filters.Clear
            Exit Sub
        End If
    End With

End Sub

Private Sub CommandButton2_Click()

    Dim file_path As String

    file_path = TextBox1.Text
    If InStr(TextBox1.Text, "TeX2imgc") > 0 Then
        If Me.Update Then
            ThisDocument.CustomDocumentProperties("tex2img").Value = file_path
        Else
            With ThisDocument.CustomDocumentProperties
                .Add _
                    Name:="tex2img", _
                    LinkToContent:=False, _
                    Type:=msoPropertyTypeString, _
                    Value:=file_path
            End With
        End If

        ThisDocument.Save

    Else

        MsgBox "Not selected TeX2imgc.exe"

    End If

    Unload Me

End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub
