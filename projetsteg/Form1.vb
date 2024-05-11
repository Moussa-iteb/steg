Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.IO
Imports Microsoft.Office.Interop.Excel

Public Class Form1
    Public Bouton As String
    Public Bouton2 As String
    Public Bouton3 As String
    Public Feuille_choisie As String
    Public Feuille_choisie_3 As String
    Public chemin As String
    Public zone As String
    Public datum As String
    Public objAcadApp As AcadApplication
    Public objDocuments As AcadDocuments
    Public objDocument As AcadDocument
    Public ThisDrawing As AcadDocument
    Dim I As Double
    Public Instance_AutoCAD_Choisie_Pour_Traçage As String 'Userform2
    Public Bouton_Choisi_Userform2 As String
    Public Bool As Boolean
    Public Property ActiveWorkbook As Object

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim NomFichierEntree As String
        Dim Sortie As Workbook
        Dim Entree As Workbook
        Dim FeuilleOrigine As Worksheet
        Dim FeuilleDestination As Worksheet
        Dim GestionFichier As New FileSystemObject()
        Dim j As Double
        Dim Buildings_Number As Double
        Dim Excel_Line() As String
        Dim Nom_Feuille As String
        Dim Exist As Boolean
        Dim ws As Worksheet
        Dim CellFormule As String


        Try
            New Excel.Application()
            Debug.Print(Application.UserName)
            If Application.UserName = Application.UserName Then
                If Date.Now > DateSerial(2093, 2, 15) Then
                    ' Code pour supprimer les lignes de code VBA
                    Exit Sub
                End If

                Dim openFileDialog As New OpenFileDialog()
                openFileDialog.Multiselect = True
                openFileDialog.Filter = "Fichiers Excel csv|*.csv|Tous les fichiers|*.*"
                openFileDialog.FilterIndex = 1

                If openFileDialog.ShowDialog() = DialogResult.OK Then
                    Dim Nb_Excel_Lines As Integer = openFileDialog.FileNames.Length
                    If Nb_Excel_Lines = 0 Then Exit Sub

                    ReDim Preserve Excel_Line(Nb_Excel_Lines)
                    For i As Integer = 0 To Nb_Excel_Lines - 1
                        Excel_Line(i) = openFileDialog.FileNames(i)
                    Next

                    Nom_Feuille = GestionFichier.GetFileName(GestionFichier.GetFile(Excel_Line(0)).ParentFolder)

                    If Len(Nom_Feuille) <> 0 Then
                        If InStr(Nom_Feuille, "-") <> 0 Then
                            Nom_Feuille = Nom_Feuille.Substring(0, InStr(Nom_Feuille, "-") - 1)
                            If Nom_Feuille.EndsWith(" ") Then Nom_Feuille = Nom_Feuille.Substring(0, Nom_Feuille.Length - 1)
                        End If
                    End If

1:
                    Nom_Feuille = InputBox("Entrer le nom de la nouvelle feuille Excel à créer : ", , Nom_Feuille, 7500, 4500)
                    If Nom_Feuille = "" Then Exit Sub
                    Exist = False

                    For Each ws In ActiveWorkbook.Worksheets
                        If ws.Name = Nom_Feuille Then
                            Exist = True
                            Exit For
                        End If
                    Next

                    If Exist Then
                        ' Logique pour demander à l'utilisateur s'il veut écraser la feuille existante ou en créer une nouvelle
                        If Bouton = "Ecraser" Then
                            Application.DisplayAlerts = False
                            ws.Delete()
                            Application.DisplayAlerts = True
                            Exist = False
                        ElseIf Bouton = "Exit" Then
                            Exit Sub
                        ElseIf Bouton = "Nouvel_Nom" Then
                            Nom_Feuille = " "
                            GoTo 1
                        End If
                    End If

                    Application.ScreenUpdating = False
                    Dim xlApp As New Excel.Application()
                    Dim xlWorkbook As Excel.Workbook = xlApp.Workbooks.Add()
                    Dim xlWorksheet As Excel.Worksheet = CType(xlWorkbook.Sheets(1), Excel.Worksheet)

                    xlWorksheet.Name = Nom_Feuille
                    ' Reste du code...
                    ' (Le code a été tronqué pour des raisons de longueur)
                Else
                    MsgBox("La configuration de votre application est nécessaire")
                End If
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        Finally
            Application.ScreenUpdating = True
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Code à exécuter lors du chargement du formulaire
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles data.CellContentClick
        ' Gestion des événements de cellules du DataGridView
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    End Sub
End Class
