Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports AutoCADLibrary
Imports Windows.Win32.System


Public Class Form1
    Private Const V As Integer = 1
    Private Const V1 As Boolean = False
    Private Const V2 As Integer = 1
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
    Public Property ActiveSheet As Object
    Public Property Selection As Object
    Public Property ActiveWindow As Object
    Public Property ThisWorkbook As Object
    Public Property UserForm2_Choix_Instance_ACAD As Object
    Public Property YourWorksheet As Object
    Public Property UserForm3 As Object
    Public Property AcWindowState As Object
    Public Property AcColor As Object
    Public Property UserForm4 As Object
    Public Property ActiveCell As Object
    Private Application As Excel.Application
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim NomFichierEntree As String
        Dim Sortie As Workbook
        Dim Entree As Workbook
        Dim form2 As New UserForm2()

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
        Dim I As Double
        Dim L_Fin_Terrain_Topographie As Double
        Dim Nb_Excel_Lines As Double
        Dim k As Double
        Dim h As Double
        Dim Total_Length_3x50 As Double
        Dim Cable3x50_Feeders As Double
        Dim Occurence As Double
        Dim Feeder_Length As Double
        Dim Feeder_Length3Q As Double
        Dim Feeder_Length1Q As Double
        Dim y As Double
        Dim Length_Longest_Line As Double
        Dim Longest_Line As Double
        Dim TRF_Feeders As Double
        Dim Total_Length As Double
        Dim Actual_Pourcentage_3x95 As Double
        Dim Actual_Pourcentage_3x50 As Double
        Dim Total_Length_TRF As Double
        Dim Gravité As Double
        Dim xlApp As New Excel.Application
        Dim Partial3Q As Double
        Dim Partial1Q As Double
        Dim MinSpan As Double
        Dim MaxSpan As Double
        Dim MinMinSpan As Double
        Dim MaxMaxSpan As Double
        Dim AverageSpan As Double
        Dim Ex_MV_Pole As Double
        Dim ThisWorkbook As Excel.Workbook




        Dim First_Pole() As String
        Dim Length_Line() As String
        Dim TRF_Feeders_Numbers() As String
        Dim Cable3x50_Feeders_Numbers() As String
        Dim Liste_Excel_Lines As String

        Dim FirstPole As String

        Dim Bool As Boolean




        On Error GoTo errHandler
        Application = New Excel.Application()
        Debug.Print(Application.UserName)
        If Application.UserName = Environment.UserName Then
            If Date.Now > New Date(2093, 2, 15) Then

                Exit Sub
            End If
            ' File Dialog to select multiple CSV files
            Using dialog As New OpenFileDialog()
                dialog.Multiselect = True
                dialog.Filter = "Fichiers Excel csv|*.csv|Tous les fichiers|*.*"
                If dialog.ShowDialog() = DialogResult.OK Then
                    Dim Index As Integer ' Renommer la variable I à l'intérieur de Using

                    Nb_Excel_Lines = dialog.FileNames.Length
                    If Nb_Excel_Lines = 0 Then Exit Sub
                    ReDim Preserve Excel_Line(Nb_Excel_Lines)
                    For Index = 0 To Nb_Excel_Lines - 1
                        Excel_Line(Index) = dialog.FileNames(Index)
                    Next
                End If
            End Using


            ' Get the name of the first selected file
            Nom_Feuille = Path.GetFileName(Path.GetDirectoryName(Excel_Line(0)))

            If Nom_Feuille.Length <> 0 Then
                If InStr(Nom_Feuille, "-") <> 0 Then
                    Nom_Feuille = Nom_Feuille.Substring(0, InStr(Nom_Feuille, "-") - 1)
                    If Nom_Feuille.EndsWith(" ") Then Nom_Feuille = Nom_Feuille.Substring(0, Nom_Feuille.Length - 1)
                End If
            End If
            Nom_Feuille = InputBox("Entrer le nom de la nouvelle feuille Excel à créer :", "Nom de la feuille", Nom_Feuille)
            If Nom_Feuille = "" Then Exit Sub

            Exist = False
            For Each worksheet As Worksheet In ActiveWorkbook.Worksheets
                If worksheet.Name = Nom_Feuille Then
                    Exist = True
                    Exit For
                End If


            Next
Label1:
            ' Votre code ici

            If Exist Then
                Dim response As DialogResult
                response = MessageBox.Show("Une feuille Excel de nom " & Nom_Feuille & " existe déjà." & vbCrLf & "Voulez-vous l'écraser ou créer une nouvelle feuille ?", "Confirmation", MessageBoxButtons.YesNoCancel)

                If response = DialogResult.Yes Then
                    Application.DisplayAlerts = False
                    ws.Delete()
                    Application.DisplayAlerts = True
                    Exist = False
                ElseIf response = DialogResult.Cancel Then
                    Exit Sub
                ElseIf response = DialogResult.No Then
                    Nom_Feuille = ""
                    GoTo Label1
                End If
            End If

            Application.ScreenUpdating = False
            Dim newSheet As Worksheet = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
            newSheet.Name = Nom_Feuille

            ThisWorkbook.Sheets("Template").Activate()
            ThisWorkbook.Sheets("Template").Range("A1:DD2").Copy()
            ThisWorkbook.Sheets(Nom_Feuille).Activate()
            ThisWorkbook.Sheets(Nom_Feuille).Range("A1").PasteSpecial(Paste:=XlPasteType.xlPasteAll)
            ThisWorkbook.Sheets(Nom_Feuille).Range("P2").Value = "LV POLES AND ASSEMBLIES SCHEDULE - " & UCase(Nom_Feuille) & " VILLAGE"
            For index As Integer = 0 To Nb_Excel_Lines - 1
                ThisWorkbook.Sheets("Template").Activate()
                ThisWorkbook.Sheets("Template").Range("A3:DD58").Copy()
                ThisWorkbook.Sheets(Nom_Feuille).Activate()
                ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * index + 3, 1).PasteSpecial(Paste:=XlPasteType.xlPasteAll)
                ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * index + 4, 8).Value = "LV-LINE-" & (index + 1)
            Next

            For index As Integer = 0 To Nb_Excel_Lines - 1
                Dim ExcelApp As New Excel.Application
                Dim Workbooks As Excel.Workbooks = ExcelApp.Workbooks
                Sortie = ThisWorkbook
                NomFichierEntree = Excel_Line(index + 1)
                If NomFichierEntree <> False Then
                    Entree = Workbooks.Open(NomFichierEntree)
                    FeuilleOrigine = Entree.Sheets(1)
                    FeuilleDestination = Sortie.Sheets(1)

                    While Not String.IsNullOrEmpty(Entree.Sheets(1).Cells(j, 1).Value)
                        j = j + 1
                    End While



                    Buildings_Number = j - 1
                    FeuilleDestination.Range("A1:E" & Buildings_Number).Value = FeuilleOrigine.Range("A1:E" & Buildings_Number).Value
                    Entree.Close()
                End If
                ThisWorkbook.Sheets("Commands ").Activate()
                If ThisWorkbook.Sheets("Commands ").Cells(1, 1).Value = "Points" Or ThisWorkbook.Sheets("Commands ").Cells(1, 1).Value = "Point" Then
                    ThisWorkbook.Sheets("Commands ").Rows("1:1").Delete(Shift:=XlDeleteShiftDirection.xlShiftUp)
                End If
                j = 1
                While Not String.IsNullOrEmpty(ThisWorkbook.Sheets("Commands ").Cells(j, 1).Value)
                    j = j + 1
                End While


                L_Fin_Terrain_Topographie = j - 1

                Dim commandsSheet As Worksheet = ThisWorkbook.Sheets("Commands ") ' Référence à la feuille de calcul "Commands "
                Dim destinationRange As Range = commandsSheet.Range(commandsSheet.Cells(1, 1), commandsSheet.Cells(L_Fin_Terrain_Topographie, 5)) ' Plage de destination

                destinationRange.Copy()

                ' Assurez-vous que Nom_Feuille est déclaré et défini correctement


                ThisWorkbook.Sheets(Nom_Feuille).Activate()
                ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7, 1).Select()

                ActiveSheet.Paste()
                ThisWorkbook.Sheets(Nom_Feuille).Activate()
                ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 4, 1).Value = GestionFichier.GetFileName(NomFichierEntree)
                ThisWorkbook.Sheets("Commands ").Activate()
                ThisWorkbook.Sheets("Commands ").Range("A1:E100").ClearContents()
            Next

            ThisWorkbook.Sheets("Template").Activate()
            ThisWorkbook.Sheets("Template").Range("A60:DD62").Select()
            Selection.Copy()
            ThisWorkbook.Sheets(Nom_Feuille).Activate()
            ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * Nb_Excel_Lines + 5, 1).Select()
            ActiveSheet.Paste()

            Dim Tabl(62) As Integer
            Tabl(0) = 12 : Tabl(1) = 13 : Tabl(2) = 14 : Tabl(3) = 15 : Tabl(4) = 16 : Tabl(5) = 17 : Tabl(6) = 19 : Tabl(7) = 20 : Tabl(8) = 21 : Tabl(9) = 22 : Tabl(10) = 23 : Tabl(11) = 24 : Tabl(12) = 25 : Tabl(13) = 26 : Tabl(14) = 27 : Tabl(15) = 28 : Tabl(16) = 30 : Tabl(17) = 31 : Tabl(18) = 32 : Tabl(19) = 33 : Tabl(20) = 34 : Tabl(21) = 35 : Tabl(22) = 36 : Tabl(23) = 37 : Tabl(24) = 38 : Tabl(25) = 39 : Tabl(26) = 40 : Tabl(27) = 41 : Tabl(28) = 42 : Tabl(29) = 43 : Tabl(30) = 44 : Tabl(31) = 45 : Tabl(32) = 46 : Tabl(33) = 47 : Tabl(34) = 48 : Tabl(35) = 49 : Tabl(36) = 50 : Tabl(37) = 51 : Tabl(38) = 52 : Tabl(39) = 53 : Tabl(40) = 80 : Tabl(41) = 81 : Tabl(42) = 82 : Tabl(43) = 83 : Tabl(44) = 84 : Tabl(45) = 85 : Tabl(46) = 86 : Tabl(47) = 87 : Tabl(48) = 88 : Tabl(49) = 89 : Tabl(50) = 90 : Tabl(51) = 91 : Tabl(52) = 92 : Tabl(53) = 93 : Tabl(54) = 94 : Tabl(55) = 95 : Tabl(56) = 96 : Tabl(57) = 97 : Tabl(58) = 98 : Tabl(59) = 99 : Tabl(60) = 100 : Tabl(61) = 101 : Tabl(62) = 102

            For c As Integer = 0 To 62
                CellFormule = "="
                For index As Integer = 0 To Nb_Excel_Lines - 1
                    CellFormule &= "+" & ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * index + 57, Tabl(c)).Address
                Next



                ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * Nb_Excel_Lines + 7, Tabl(c)).Formula = CellFormule
            Next c
            ThisWorkbook.Sheets("Template").Activate()
            ThisWorkbook.Sheets("Template").Range("A64:DD106").Copy()
            ThisWorkbook.Sheets(Nom_Feuille).Activate()
            ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * Nb_Excel_Lines + 9, 1).PasteSpecial(Paste:=XlPasteType.xlPasteAll)
            ThisWorkbook.Sheets(Nom_Feuille).Activate()
            ThisWorkbook.Sheets(Nom_Feuille).Columns("A:G").Hidden = True
            ActiveWindow.ScrollRow = ThisWorkbook.Sheets(ActiveSheet.Name).Cells(1, 8).Row
            ActiveWindow.ScrollColumn = ThisWorkbook.Sheets(ActiveSheet.Name).Cells(1, 8).Column

            ThisWorkbook.Sheets(Nom_Feuille).Columns("M:M").Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Columns("O:O").Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Columns("Q:Q").Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Columns("AD:AD").Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Columns("AF:AF").Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Columns("AH:AI").Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Columns("AK:AK").Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Columns("AM:AM").Hidden = True
            'ThisWorkbook.Sheets(Nom_Feuille).Columns("AQ:AQ").Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Columns("AS:AT").Hidden = True
            'ThisWorkbook.Sheets(Nom_Feuille).Columns("AU:BA").Hidden = True
            Nb_Excel_Lines = 0

            While Not String.IsNullOrEmpty(ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7, 1).Value)
                I = I + 1
            End While

            Nb_Excel_Lines = I
            ReDim Preserve Excel_Line(Nb_Excel_Lines)
            For I = 0 To Nb_Excel_Lines - 1
                L_Fin_Terrain_Topographie = 0
                j = 0
                While Not String.IsNullOrEmpty(ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 1).Value)
                    j = j + 1
                End While

                L_Fin_Terrain_Topographie = j
                Excel_Line(I + 1) = L_Fin_Terrain_Topographie
            Next I

            '########################################################## Assemblies
            ReDim Preserve First_Pole(Nb_Excel_Lines)
            For p As Integer = 0 To Nb_Excel_Lines - V
                First_Pole(p) = ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * p + 7, 8).Value
            Next p

            For f As Integer = 0 To Nb_Excel_Lines - 1
                For s As Integer = 1 To Excel_Line(f + 1) - 1
                    Dim Occ As Integer = 0
                    For a As Integer = 0 To Nb_Excel_Lines - 1
                        If ThisWorkbook.Sheets(Nom_Feuille).Cells(s + 56 * I + 7, 8).Value = First_Pole(a + 1) Then
                            Occ += 1
                        End If
                    Next a
                    If Occ >= 1 Then
                        For q As Integer = 0 To 6
                            'ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 47 + k).Interior.ColorIndex = 43
                        Next q
                        ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 53).Value = Occ

                        'If ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 47).Value = 1 Or ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 48).Value = 1 Or ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 49).Value = 1 Then
                        '    ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 47).ClearContents()
                        '    ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 48).ClearContents()
                        '    ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 49).ClearContents()
                        '    ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 53).Value = 1
                        '    ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 52).Value = Occurence - 1
                        '    If ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 52).Value = 0 Then ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 52).ClearContents()
                        'ElseIf ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 50).Value = 1 Or ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 51).Value = 1 Then
                        '    ThisWorkbook.Sheets(Nom_Feuille).Cells(j + 56 * I + 7, 52).Value = Occurence
                        'End If
                    End If
                Next
            Next
            '########################################################## Assemblies
            '########################################################## Cables
            ReDim Preserve Length_Line(Nb_Excel_Lines)
            Total_Length = 0
            j = -1
            Total_Length_TRF = 0
            For m As Integer = 0 To Nb_Excel_Lines - 1
                Length_Line(m) = ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * m + 7 + 50, 12).Value
                Total_Length += Length_Line(m)
                If Microsoft.VisualBasic.Left(ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7, 8).Value, 3) = "TRF" Then
                    j += 1
                    ReDim Preserve TRF_Feeders_Numbers(j)
                    TRF_Feeders_Numbers(j) = I
                    Total_Length_TRF += Length_Line(I)
                End If
            Next
            TRF_Feeders = j
            Actual_Pourcentage_3x95 = Total_Length_TRF / Total_Length

            If Actual_Pourcentage_3x95 < ThisWorkbook.Sheets("Assumptions").Range("Pourcentage_3x95").Value Then
                Longest_Line = 0
                Length_Longest_Line = 0
                For I = 0 To Nb_Excel_Lines - 1

                    Bool = False
                    For j = 0 To TRF_Feeders
                        If I = TRF_Feeders_Numbers(j) Then
                            Bool = True
                        End If
                    Next j
                    If Not Bool Then
                        If ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7 + 50, 12).Value > Length_Longest_Line Then
                            Length_Longest_Line = ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7 + 50, 12).Value
                            Longest_Line = I
                        End If
                    End If
                Next I
                TRF_Feeders = TRF_Feeders + 1
                ReDim Preserve TRF_Feeders_Numbers(TRF_Feeders)
                TRF_Feeders_Numbers(TRF_Feeders) = Longest_Line
            End If
            Total_Length_TRF = Total_Length_TRF + Length_Longest_Line
            Actual_Pourcentage_3x95 = Total_Length_TRF / Total_Length
            ' Définition de l'étiquette Label777
Label777:
            If Actual_Pourcentage_3x95 < ThisWorkbook.Sheets("Assumptions").Range("Pourcentage_3x95").Value Then
                GoTo Label777
            End If

            For b As Integer = 0 To Nb_Excel_Lines - 1
                Bool = False
                For t As Integer = 0 To TRF_Feeders
                    If b = TRF_Feeders_Numbers(t) Then
                        Bool = True
                    End If
                Next t
                If Not Bool Then
                    For J_ As Integer = 0 To Excel_Line(I + 1) - 1
                        ThisWorkbook.Sheets(Nom_Feuille).Cells(J_ + 56 * I + 7, 17).ClearContents()
                    Next J_

                End If

            Next

            Cable3x50_Feeders = -1
            Total_Length_3x50 = 0
Label888:
            Length_Longest_Line = 0
            Longest_Line = 0

            For x As Integer = 0 To Nb_Excel_Lines - 1
                Bool = False
                For j_ As Integer = 0 To TRF_Feeders
                    If I = TRF_Feeders_Numbers(j_) Then
                        Bool = True
                    End If
                Next j_

                For u As Integer = 0 To Cable3x50_Feeders
                    If I = Cable3x50_Feeders_Numbers(u) Then
                        Bool = True
                    End If
                Next u
                If Not Bool Then
                    If ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7 + 50, 12).Value > Length_Longest_Line Then
                        Length_Longest_Line = ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7 + 50, 12).Value
                        Longest_Line = I
                    End If
                End If
            Next
            Cable3x50_Feeders = Cable3x50_Feeders + 1
            ReDim Preserve Cable3x50_Feeders_Numbers(Cable3x50_Feeders)
            Cable3x50_Feeders_Numbers(Cable3x50_Feeders) = Longest_Line
            Total_Length_3x50 = Total_Length_3x50 + Length_Longest_Line
            Actual_Pourcentage_3x50 = Total_Length_3x50 / Total_Length

            If Actual_Pourcentage_3x50 < ThisWorkbook.Sheets("Assumptions").Range("Pourcentage_3x50").Value - Actual_Pourcentage_3x95 + ThisWorkbook.Sheets("Assumptions").Range("Pourcentage_3x95").Value Then
                GoTo Label888
            End If

            For index As Integer = 0 To Nb_Excel_Lines - 1
                Bool = False
                For j_ As Integer = 0 To TRF_Feeders
                    If index = TRF_Feeders_Numbers(j_) Then
                        Bool = True
                    End If
                Next j_

                For j_ As Integer = 0 To Cable3x50_Feeders
                    If index = Cable3x50_Feeders_Numbers(j_) Then
                        Bool = True
                    End If
                Next j_

                If Not Bool Then
                    For i_ As Integer = 0 To Excel_Line(index + 1) - 1
                        ThisWorkbook.Sheets(Nom_Feuille).Cells(i_ + 56 * index + 7, 16).ClearContents()
                    Next i_


                End If
            Next

            MinMinSpan = ThisWorkbook.Sheets(Nom_Feuille).Cells(8, 12).Value
            MaxMaxSpan = ThisWorkbook.Sheets(Nom_Feuille).Cells(8, 12).Value

            For I_ As Integer = 0 To Nb_Excel_Lines - 1
                MinSpan = ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I_ + 7 + 1, 12).Value
                MaxSpan = ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I_ + 7 + 1, 12).Value


                For k_ As Integer = 1 To Excel_Line(I + 1) - 1
                    If ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7 + k_, 12).Value < MinSpan Then
                        MinSpan = ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7 + k_, 12).Value
                    End If
                    If ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7 + k_, 12).Value > MaxSpan Then
                        MaxSpan = ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * I + 7 + k_, 12).Value
                    End If
                    If MinSpan < MinMinSpan Then MinMinSpan = MinSpan
                    If MaxSpan > MaxMaxSpan Then MaxMaxSpan = MaxSpan
                Next k_


            Next I_

            ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * Nb_Excel_Lines + 39, 19).Value = MinMinSpan
            ThisWorkbook.Sheets(Nom_Feuille).Cells(56 * Nb_Excel_Lines + 40, 19).Value = MaxMaxSpan

            For I_ As Integer = 1 To Nb_Excel_Lines
                Dim j_ As Integer = 50 - Excel_Line(I_)
                For k_ As Integer = 1 To j_
                    ThisWorkbook.Sheets(Nom_Feuille).Rows(-6 - k_ + 56 * I_ + 7).Hidden = True
                Next k_

                ThisWorkbook.Sheets(Nom_Feuille).Rows(-5 + 56 * I_ + 7).Hidden = True
            Next I_

            ThisWorkbook.Sheets(Nom_Feuille).Rows(56 * (I - 1) + 4).Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Rows(56 * (I - 1) + 5).Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Rows(56 * (I - 1) + 6).Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Rows(56 * (I - 1) + 7).Hidden = True
            ThisWorkbook.Sheets(Nom_Feuille).Columns("I:K").EntireColumn.AutoFit()
            ThisWorkbook.Sheets(Nom_Feuille).Columns("AC:AC").EntireColumn.AutoFit()
            ThisWorkbook.Sheets(Nom_Feuille).Columns("BE:CX").EntireColumn.AutoFit()
            ThisWorkbook.Sheets(Nom_Feuille).Columns("AC:AC").Copy()
            ThisWorkbook.Sheets(Nom_Feuille).Columns("H:H").PasteSpecial(8) ' 8 correspond à xlPasteColumnWidths


            With ActiveWindow
                .SplitColumn = 1
                .SplitRow = 0
            End With

            ActiveWindow.FreezePanes = True
            Application.ScreenUpdating = True
            xlApp.Run("NomDuModule.Macro1")

            Exit Sub

errHandler:
            MsgBox("Error: " & Err.Description & " " & Erl())
        Else
            MsgBox("La configuration de votre application est nécessaire")
        End If
    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Code à exécuter lors du chargement du formulaire
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        ' Gestion des événements de cellules du DataGridView
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim I As Double
        Dim j As Double
        Dim Pt_Start(2) As Double
        Dim Pt_End(2) As Double
        Dim Pt_Start_text(2) As Double
        Dim Pt_End_text(2) As Double
        Dim Ligne As AcadLine
        Dim Ligne1 As AcadLine
        Dim Ligne2 As AcadLine
        Dim Ligne3 As AcadLine
        Dim Ligne4 As AcadLine
        Dim strListe As String
        Dim Numdoc As Double
        Dim DocExist As Boolean
        Dim feuille As Object
        Dim Excel_Line() As String
        Dim objAcadApp As AcadApplication
        Dim objDocuments As AcadDocuments
        Dim objDocument As AcadDocument
        Dim ThisDrawing As AcadDocument

        On Error GoTo errHandler
        Debug.WriteLine(Application.UserName)

        If Application.UserName = Application.UserName Then
            If Date.Now > New Date(2093, 2, 28) Then
                Dim codeModule As EnvDTE.CodeModule = CType(ActiveWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule, EnvDTE.CodeModule)
                codeModule.DeleteLines(1, codeModule.CountOfLines)
                Exit Sub
            End If
        End If
        For j = 1 To UserForm2.ComboBox1.Items.Count
            UserForm2.ComboBox1.Items.RemoveAt(0)
        Next j



        Dim Bool1 As Boolean
        For Each feuille_ As Worksheet In ThisWorkbook.Sheets
            If feuille_.Name <> "Commands " AndAlso feuille_.Name <> "Template" AndAlso feuille_.Name <> "Assumptions" AndAlso feuille_.Name <> "Detailed Network Summary" AndAlso feuille_.Name <> "To improve" AndAlso feuille_.Name <> "Conversion Clark1880" Then
                UserForm2.ComboBox1.Items.Add(feuille_.Name)
                Bool1 = True
            End If
        Next feuille_


        If Bool1 = False Then
            MsgBox("Aucun Réseau Basse Tension Traité!")
            Exit Sub
        End If

        Dim bouton2 As String = ""
        Dim feuille_choisie As String = ""
        UserForm2.ShowDialog()
        If bouton2 = "Exit" Then Exit Sub
        If bouton2 <> "Confirmer" Then Exit Sub

        Dim Bool As Boolean
        For Each feuille_ As Worksheet In ThisWorkbook.Sheets
            If feuille_choisie = feuille_.Name Then Bool = True
        Next feuille_

5:
        If Bool = False Then GoTo 5
        While Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(56 * I + 7, 1).Value)
            I += 1
        End While


        ReDim Preserve Excel_Line(I)
        For iy As Integer = 0 To I - 1
            Dim jt As Integer = 0
            While Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(jt + 56 * iy + 7, 1).Value)
                jt += 1
            End While
        Next iy
        Excel_Line(I + 1) = j

        On Error Resume Next

        Dim ThisDrawingObj As Object

        objAcadApp = GetObject(, "AutoCAD.Application")
        If Err.Number <> 0 Then
            Err.Clear()
            objAcadApp = CreateObject("AutoCAD.Application")
            If Err.Number <> 0 Then
                MsgBox(Err.Description)
                Err.Clear()
                Exit Sub
            End If
            ThisDrawing = objAcadApp.ActiveDocument
            If Err.Number <> 0 Then
                MsgBox(Err.Description)
                Err.Clear()
                Exit Sub
            End If
            GoTo AutoCAD_OK
        End If

        ' Si hay una instancia de AutoCAD abierta, se procede aquí

AutoCAD_OK:
        ' Código a ejecutar si se detecta una instancia de AutoCAD abierta

        ' Limpiar ComboBox en UserForm2_Choix_Instance_ACAD
        For I = 1 To UserForm2_Choix_Instance_ACAD.ComboBox1.ListCount
            UserForm2_Choix_Instance_ACAD.ComboBox1.RemoveItem(0)
        Next I
        ' Remplir la ComboBox du Userform2
        ThisDrawing = objAcadApp.ActiveDocument
        For Each objDocument In ThisDrawing.Application.Documents
            UserForm2_Choix_Instance_ACAD.ComboBox1.Items.Add(objDocument.Name)
        Next
        UserForm2_Choix_Instance_ACAD.ComboBox1.Items.Add("Créer une nouvelle instance")
        UserForm2_Choix_Instance_ACAD.ComboBox1.SelectedIndex = 0
        Instance_AutoCAD_Choisie_Pour_Traçage = UserForm2_Choix_Instance_ACAD.ComboBox1.Items(0)

        ' Au cas où il y a une une seule instance AutoCAD ouverte, inutile de demander le choix de l'utilisateur
        'If UserForm2_Choix_Instance_ACAD.ComboBox1.Items.Count = 1 Then
        '    ThisDrawing = objAcadApp.ActiveDocument
        '    GoTo AutoCAD_OK
        'End If

        ' Afficher le Userform2 pourque l'utilisateur choisisse l'instance AutoCAD
        Bouton_Choisi_Userform2 = ""
        UserForm2_Choix_Instance_ACAD.Confirmer.Focus()
        UserForm2_Choix_Instance_ACAD.ShowDialog()
        If Err.Number <> 0 Then
            MsgBox("Err : " & Err.Description)
            Err.Clear()
            Exit Sub
        End If
        If Bouton_Choisi_Userform2 <> "Confirmer" Then Exit Sub

        ' Vérification du choix de l'utilisateur
        Dim myBool As Boolean = False
        If Instance_AutoCAD_Choisie_Pour_Traçage = "Créer une nouvelle instance" Then myBool = True
        For Each objDocument In ThisDrawing.Application.Documents
            If Instance_AutoCAD_Choisie_Pour_Traçage = objDocument.Name Then myBool = True
        Next

        If Bool = False Then
            MsgBox("Choix incorrect")
        End If

        ' Validación de la elección del usuario
        Dim objDocumentsNew As Object = ThisDrawing.Application.Documents

        If Instance_AutoCAD_Choisie_Pour_Traçage = "Créer une nouvelle instance" Then
            objDocument = objDocumentsNew.Add("acad.dwt")
            ThisDrawing = objAcadApp.ActiveDocument
            If Err.Number <> 0 Then
                MsgBox("Err : " & Err.Description)
                Err.Clear()
            Else
                ThisDrawing = objDocuments.Item(Instance_AutoCAD_Choisie_Pour_Traçage)
                Exit Sub
            End If
        End If






        If Err.Number <> 0 Then
            MsgBox("Err : " & Err.Description)
            Err.Clear()
            Exit Sub
        End If

        ' Abrir y activar la instancia elegida por el usuario
        objAcadApp.Visible = True
        ThisDrawing.SummaryInfo.Author = "Wael Bagga"
        ThisDrawing.Application.WindowState = acMin() ' Agrandir ou minimiser la ventana de AutoCAD, es solo para mostrarla
        ThisDrawing.Application.WindowState = acMax()
        ThisDrawing.WindowState = acMax()

        If Err.Number <> 0 Then
            MsgBox("Err : " & Err.Description)
            Err.Clear()
            Exit Sub
        End If
        ' Call Open_AutoCAD
        If Bouton_Choisi_Userform2 <> "Confirmer" Then Exit Sub

        Dim text As Object
        Dim Circ As Object
        Dim AcExcelLV As Object
        Dim AcExcelPoles As Object
        Dim AcExcelCotes As Object
        Dim AcExcelMalt As Object
        Dim AcExcelCartouche As Object
        Dim AcExcelPolesTypes As Object
        Dim AcExcelStays As Object
        Dim AcExcelLineNumber As Object

        Dim objDimAligned As Object
        Dim Côtes(2) As Double
        Dim Bool3 As Boolean
        ' Excel LV
        AcExcelLV = objAcadApp.ActiveDocument.Layers.Add("01 Excel_LV")
        objAcadApp.ActiveDocument.ActiveLayer = AcExcelLV
        Bool3 = False
        For Ii As Integer = 0 To I - 1
            Dim Pt_Start2(1) As Double ' Renommer la variable pour éviter le conflit de noms
            Pt_Start(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(56 * Ii + 7, 9).Value
            Pt_Start(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(56 * Ii + 7, 10).Value
            Dim Pt_Start_textt(1) As Double
            Pt_Start_text(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(56 * Ii + 7, 9).Value + 3.64
            Pt_Start_text(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(56 * Ii + 7, 10).Value
            If Left(ThisWorkbook.Sheets(feuille_choisie).Cells(56 * Ii + 7, 8).Value, 3) = "TRF" AndAlso Not Bool3 Then
                Bool3 = True
                text = objAcadApp.ActiveDocument.ModelSpace.AddText(ThisWorkbook.Sheets(feuille_choisie).Cells(56 * I + 7, 8).Value, Pt_Start_text, 6)
                text.Color = acByLayer()
            End If

            For _j As Integer = 1 To Excel_Line(Ii + 1)
                If Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(_j + 56 * I + 7, 10).Value) Then
                    Pt_End(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(_j + 56 * I + 7, 9).Value
                    Pt_End(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(_j + 56 * I + 7, 10).Value
                    Pt_End_text(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(_j + 56 * I + 7, 9).Value + 3.64
                    Pt_End_text(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(_j + 56 * I + 7, 10).Value
                    Dim Lignee As Object = objAcadApp.ActiveDocument.ModelSpace.AddLine(Pt_Start, Pt_End)
                End If
            Next _j
        Next Ii
        If Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 17).Value) Then
            If Not String.IsNullOrEmpty(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 17).Value) Then
                Ligne.Color = ThisWorkbook.Sheets("Assumptions").Range("Color_3_95").Value
                Ligne.LineWeight = ThisWorkbook.Sheets("Assumptions").Range("Weight_3_95").Value
            ElseIf Not String.IsNullOrEmpty(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 16).Value) Then
                Ligne.Color = ThisWorkbook.Sheets("Assumptions").Range("Color_3_50").Value
                Ligne.LineWeight = ThisWorkbook.Sheets("Assumptions").Range("Weight_3_50").Value
            ElseIf Not String.IsNullOrEmpty(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 15).Value) Then
                Ligne.Color = ThisWorkbook.Sheets("Assumptions").Range("Color_3_25").Value
                Ligne.LineWeight = ThisWorkbook.Sheets("Assumptions").Range("Weight_3_25").Value
            ElseIf Not String.IsNullOrEmpty(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 14).Value) Then
                Ligne.Color = ThisWorkbook.Sheets("Assumptions").Range("Color_1_50").Value
                Ligne.LineWeight = ThisWorkbook.Sheets("Assumptions").Range("Weight_1_50").Value
            Else
                Ligne.Color = ThisWorkbook.Sheets("Assumptions").Range("Color_1_25").Value
                Ligne.LineWeight = ThisWorkbook.Sheets("Assumptions").Range("Weight_1_25").Value
            End If

            ThisDrawing.Application.ZoomExtents()
            text = objAcadApp.ActiveDocument.ModelSpace.AddText(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 8).Value, Pt_End_text, 6)
            text.Color = acByLayer()
            Pt_Start(0) = Pt_End(0)
            Pt_Start(1) = Pt_End(1)
        End If
        Dim CircColor As Long
        Dim HatchColor As Long
        Dim LineColor As Long
        Dim ObjetFrontiere(0 To 0) As Object
        Dim ObjetHachures As Object
        Dim L As Long

        '###########################################################################
        ' Excel Poles
        AcExcelPoles = objAcadApp.ActiveDocument.Layers.Add("02 Excel_Poles")
        objAcadApp.ActiveDocument.ActiveLayer = AcExcelPoles
        AcExcelPoles.Color = acWhite()

        For Ii As Integer = 0 To I - 1
            For jj As Integer = 1 To Excel_Line(I + 1)
                If Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 19).Value) Then
                    Pt_End(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 9).Value
                    Pt_End(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 10).Value
                    Dim Circ2 As Object = objAcadApp.ActiveDocument.ModelSpace.AddCircle(Pt_End, 2)
                    Circ2.Color = acWhite()
                    Circ.LineWeight = 60
                    ObjetFrontiere(0) = Circ
                    ObjetHachures = ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True)
                    ObjetHachures.AppendOuterLoop(ObjetFrontiere)
                    ObjetHachures.Evaluate
                    ObjetHachures.Color = acWhite()
                End If
            Next jj
        Next Ii
        If Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 20).Value) Then
            Pt_End(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 9).Value
            Pt_End(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 10).Value
            Dim Circ3 As Object = objAcadApp.ActiveDocument.ModelSpace.AddCircle(Pt_End, 2)
            Circ.Color = acBlue()
            Circ.LineWeight = 60
            Dim ObjetFrontierer(0 To 0) As Object
            ObjetFrontiere(0) = Circ3
            ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True).AppendOuterLoop(ObjetFrontiere)
            ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True).Evaluate
            CObj(ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True)).Color = acBlue()
        End If

        If Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 21).Value) Then
            Pt_End(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 9).Value
            Pt_End(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 10).Value
            Dim Circ6 As Object = objAcadApp.ActiveDocument.ModelSpace.AddCircle(Pt_End, 2)
            Circ.Color = acRed()
            Circ.LineWeight = 60
            Dim ObjetFrontieree(0 To 0) As Object
            ObjetFrontieree(0) = Circ6

            ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True).AppendOuterLoop(ObjetFrontiere)
            ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True).Evaluate
            CObj(ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True)).Color = acRed()
        End If
        If Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 22).Value) Then
            Pt_End(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 9).Value
            Pt_End(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 10).Value
            Dim Circc As Object = objAcadApp.ActiveDocument.ModelSpace.AddCircle(Pt_End, 2)
            Circ.Color = acYellow()
            Circ.LineWeight = 60
            Dim ObjetFrontieree(0 To 0) As Object
            ObjetFrontiere(0) = Circ
            ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True).AppendOuterLoop(ObjetFrontiere)
            ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True).Evaluate
            CObj(ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True)).Color = acYellow()
        End If

        If Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 23).Value) Then
            Pt_End(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 9).Value
            Pt_End(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 10).Value
            Dim Circs As Object = objAcadApp.ActiveDocument.ModelSpace.AddCircle(Pt_End, 2)
            Circ.Color = acGreen()
            Dim ObjetFrontierez(0 To 0) As Object
            ObjetFrontiere(0) = Circs
            ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True).AppendOuterLoop(ObjetFrontiere)
            ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True).Evaluate
            CObj(ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True)).Color = acGreen()
        End If
        If Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 24).Value) Then
            Pt_End(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 9).Value
            Pt_End(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(j + 56 * I + 7, 10).Value
            Dim Circa As Object = objAcadApp.ActiveDocument.ModelSpace.AddCircle(Pt_End, 2)
            Circa.Color = acMagenta()
            Dim ObjetFrontieree(0 To 0) As Object
            ObjetFrontiere(0) = Circ
            Dim NouvelObjetHachures As Object = ThisDrawing.ModelSpace.AddHatch(acHatchPatternTypePreDefined, "SOLID", True)
            ObjetHachures.AppendOuterLoop(ObjetFrontieree)
            ObjetHachures.Evaluate
            ObjetHachures.Color = acMagenta()
        End If

        ' Excel Côtes
        AcExcelCotes = objAcadApp.ActiveDocument.Layers.Add("03 Excel_Côtes")
        objAcadApp.ActiveDocument.ActiveLayer = AcExcelCotes
        AcExcelCotes.Color = acMagenta()

        For iLoop As Integer = 0 To CDbl(I) - 1
            Pt_Start(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(56 * iLoop + 7, 9).Value
            Pt_Start(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(56 * iLoop + 7, 10).Value
            For j_ As Integer = 1 To Excel_Line(iLoop + 1)
                If Not IsNothing(ThisWorkbook.Sheets(feuille_choisie).Cells(j_ + 56 * iLoop + 7, 10).Value) Then
                    Dim Pt_Endd(1) As Double
                    Pt_Endd(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(j_ + 56 * iLoop + 7, 9).Value
                    Pt_Endd(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(j_ + 56 * iLoop + 7, 10).Value
                    Dim NouvellesCôtes(2) As Double
                    NouvellesCôtes(0) = Pt_Start(0) / 2 + Pt_Endd(0) / 2
                    NouvellesCôtes(1) = 2 + Pt_Start(1) / 2 + Pt_Endd(1) / 2
                    NouvellesCôtes(2) = 0#
                    Dim NouvelObjDimAligned As Object = ThisDrawing.ModelSpace.AddDimAligned(Pt_Start, Pt_Endd, NouvellesCôtes)
                End If
            Next j_
        Next iLoop

        ' Excel Malt
        AcExcelMalt = objAcadApp.ActiveDocument.Layers.Add("04 Excel_Malt")
        objAcadApp.ActiveDocument.ActiveLayer = AcExcelMalt
        AcExcelMalt.Color = acRed()
        Dim objBloc As Object
        Dim objInsert As Object


        Pt_Start(0) = 0
        Pt_Start(1) = 0
        objBloc = ThisDrawing.Blocks.Add(Pt_Start, "Malt")
        objBloc.AddCircle(Pt_Start, 6)
        Pt_Start(0) = 0
        Pt_Start(1) = -1.33 * 1.3
        Pt_End(0) = Pt_Start(0)
        Pt_End(1) = Pt_Start(1) + 1.33 * 1.3 + 3.31 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = -2.24 * 1.3
        Pt_Start(1) = -1.33 * 1.3
        Pt_End(0) = Pt_Start(0) + 2 * 2.24 * 1.3
        Pt_End(1) = Pt_Start(1)
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = -1.12 * 1.3
        Pt_Start(1) = -2.19 * 1.3
        Pt_End(0) = Pt_Start(0) + 2 * 1.12 * 1.3
        Pt_End(1) = Pt_Start(1)
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = -0.56 * 1.3
        Pt_Start(1) = -3.05 * 1.3
        Pt_End(0) = Pt_Start(0) + 2 * 0.56 * 1.3
        Pt_End(1) = Pt_Start(1)
        objBloc.AddLine(Pt_Start, Pt_End)

        For iLoop As Integer = 0 To CDbl(I) - 1
            For jl As Integer = 1 To Excel_Line(iLoop + 1)
                If ThisWorkbook.Sheets(feuille_choisie).Cells(jl + 56 * iLoop + 7, 28).Value = 1 AndAlso Not String.IsNullOrEmpty(ThisWorkbook.Sheets(feuille_choisie).Cells(jl + 56 * iLoop + 7, 2).Value) Then
                    Pt_Start(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(jl + 56 * iLoop + 7, 9).Value + 13.8
                    Pt_Start(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(jl + 56 * iLoop + 7, 10).Value - 6.15
                    objInsert = ThisDrawing.ModelSpace.InsertBlock(Pt_Start, "Malt", 1, 1, 1, 0)
                End If
            Next jl
        Next iLoop

        ThisDrawing.Application.ZoomExtents()
        ' Excel Cartouche
        AcExcelCartouche = objAcadApp.ActiveDocument.Layers.Add("05 Excel_Cartouche")
        objAcadApp.ActiveDocument.ActiveLayer = AcExcelCartouche
        AcExcelCartouche.Color = acWhite()


        Pt_Start(0) = 0
        Pt_Start(1) = 0
        objBloc = ThisDrawing.Blocks.Add(Pt_Start, "Transformateur")
        Pt_Start(0) = -4.53
        Pt_Start(1) = 0
        objBloc.AddCircle(Pt_Start, 7.28)
        Pt_Start(0) = 4.53
        Pt_Start(1) = 0
        objBloc.AddCircle(Pt_Start, 7.28)

        For Ih As Integer = 0 To I - 1
            For jh As Integer = 0 To Excel_Line(I + 1) - 1
                If Left(ThisWorkbook.Sheets(feuille_choisie).Cells(56 * I + 7, 8).Value, 3) = "TRF" Then
                    Pt_Start(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(jh + 56 * I + 7, 9).Value
                    Pt_Start(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(jh + 56 * I + 7, 10).Value
                    objInsert = ThisDrawing.ModelSpace.InsertBlock(Pt_Start, "Transformateur", 1, 1, 1, 0)
                    GoTo ZoomExtents
                End If
            Next jh
        Next Ih

ZoomExtents:
        ThisDrawing.Application.ZoomExtents()
        ' Poles Type
        'AcExcelPolesTypes = objAcadApp.ActiveDocument.Layers.Add("06 Excel_PolesTypes")
        'objAcadApp.ActiveDocument.ActiveLayer = AcExcelPolesTypes
        'AcExcelPolesTypes.Color = 104
        'Dim Pt_Start(1) As Double
        'Pt_Start(0) = 0
        'Pt_Start(1) = 0
        'Dim objBloc As Object = ThisDrawing.Blocks.Add(Pt_Start, "9StoutPole")
        'objBloc.AddCircle(Pt_Start, 6)
        'Pt_Start(0) = -2.52
        'Pt_Start(1) = -2.88
        'objBloc.AddText("S", Pt_Start, 5.5)
        'Pt_Start(0) = 0
        'Pt_Start(1) = 0
        'objBloc = ThisDrawing.Blocks.Add(Pt_Start, "10MediumPole")
        'Dim Points(14) As Double
        'Points(0) = -8.23 * 1.3: Points(1) = -3.79 * 1.3: Points(2) = 0
        'Points(3) = 8.23 * 1.3: Points(4) = -3.79 * 1.3: Points(5) = 0
        'Points(6) = 8.23 * 1.3: Points(7) = 3.79 * 1.3: Points(8) = 0
        'Points(9) = -8.23 * 1.3: Points(10) = 3.79 * 1.3: Points(11) = 0
        'Points(12) = -8.23 * 1.3: Points(13) = -3.79 * 1.3: Points(14) = 0
        'objBloc.AddPolyline(Points)
        'Pt_Start(0) = -8.53
        'Pt_Start(1) = -3
        'objBloc.AddText("10M", Pt_Start, 6)
        'Pt_Start(0) = 0
        'Pt_Start(1) = 0
        'objBloc = ThisDrawing.Blocks.Add(Pt_Start, "10StoutPole")
        'Points(0) = -8.23 * 1.3: Points(1) = -3.79 * 1.3: Points(2) = 0
        'Points(3) = 8.23 * 1.3: Points(4) = -3.79 * 1.3: Points(5) = 0
        'Points(6) = 8.23 * 1.3: Points(7) = 3.79 * 1.3: Points(8) = 0
        'Points(9) = -8.23 * 1.3: Points(10) = 3.79 * 1.3: Points(11) = 0
        'Points(12) = -8.23 * 1.3: Points(13) = -3.79 * 1.3: Points(14) = 0
        'objBloc.AddPolyline(Points)
        'Pt_Start(0) = -8.53 + 0.91
        'Pt_Start(1) = -3
        'objBloc.AddText("10S", Pt_Start, 6)
        ' Stays
        AcExcelStays = objAcadApp.ActiveDocument.Layers.Add("07 Excel_Stays")
        objAcadApp.ActiveDocument.ActiveLayer = AcExcelStays
        AcExcelStays.Color = acBlue()


        Pt_Start(0) = 0
        Pt_Start(1) = 0
        objBloc = ThisDrawing.Blocks.Add(Pt_Start, "1Stay")
        objBloc.AddCircle(Pt_Start, 6)
        Pt_Start(0) = 0
        Pt_Start(1) = -2.5 * 1.3
        Pt_End(0) = 0
        Pt_End(1) = 2.5 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = 0
        Pt_Start(1) = 2.5 * 1.3
        Pt_End(0) = 1.34 * 1.3
        Pt_End(1) = 1.35 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = 0
        Pt_Start(1) = 2.5 * 1.3
        Pt_End(0) = -1.34 * 1.3
        Pt_End(1) = 1.35 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = 0
        Pt_Start(1) = 0
        objBloc = ThisDrawing.Blocks.Add(Pt_Start, "2Stay")
        objBloc.AddCircle(Pt_Start, 6)
        Pt_Start(0) = 1.7 * 1.3
        Pt_Start(1) = -2.5 * 1.3
        Pt_End(0) = 1.7 * 1.3
        Pt_End(1) = 2.5 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = 1.7 * 1.3
        Pt_Start(1) = 2.5 * 1.3
        Pt_End(0) = 1.7 * 1.3 + 1.34 * 1.3
        Pt_End(1) = 1.35 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = 1.7 * 1.3
        Pt_Start(1) = 2.5 * 1.3
        Pt_End(0) = 1.7 * 1.3 - 1.34 * 1.3
        Pt_End(1) = 1.35 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = -1.7 * 1.3
        Pt_Start(1) = -2.5 * 1.3
        Pt_End(0) = -1.7 * 1.3
        Pt_End(1) = 2.5 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = -1.7 * 1.3
        Pt_Start(1) = 2.5 * 1.3
        Pt_End(0) = -1.7 * 1.3 + 1.34 * 1.3
        Pt_End(1) = 1.35 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = -1.7 * 1.3
        Pt_Start(1) = -2.5 * 1.3
        Pt_End(0) = -1.7 * 1.3
        Pt_End(1) = 2.5 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = -1.7 * 1.3
        Pt_Start(1) = 2.5 * 1.3
        Pt_End(0) = -1.7 * 1.3 + 1.34 * 1.3
        Pt_End(1) = 1.35 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        Pt_Start(0) = -1.7 * 1.3
        Pt_Start(1) = 2.5 * 1.3
        Pt_End(0) = -1.7 * 1.3 - 1.34 * 1.3
        Pt_End(1) = 1.35 * 1.3
        objBloc.AddLine(Pt_Start, Pt_End)
        For Ib As Integer = 0 To I - 1
            For jb As Integer = 0 To Excel_Line(I + 1) - 1
                If ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 25).Value = 1 AndAlso Not String.IsNullOrEmpty(ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 2).Value) Then
                    Pt_Start(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 9).Value - 8.32
                    Pt_Start(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 10).Value
                    If jb = 0 Then
                        Pt_Start(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 9).Value - 8.32
                        Pt_Start(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 10).Value + 8.77
                    End If
                    Dim objInsertt As Object = ThisDrawing.ModelSpace.InsertBlock(Pt_Start, "1Stay", 1, 1, 1, 0)
                End If
                If ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 25).Value = 2 AndAlso Not String.IsNullOrEmpty(ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 2).Value) Then
                    Pt_Start(0) = ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 9).Value - 8.32
                    Pt_Start(1) = ThisWorkbook.Sheets(feuille_choisie).Cells(jb + 56 * I + 7, 10).Value
                    Dim objInsertt As Object = ThisDrawing.ModelSpace.InsertBlock(Pt_Start, "2Stay", 1, 1, 1, 0)
                End If
            Next jb
        Next Ib

        ThisDrawing.Application.ZoomExtents()
        ' Line Number
        AcExcelLineNumber = objAcadApp.ActiveDocument.Layers.Add("08 Excel_LineNumber")
        objAcadApp.ActiveDocument.ActiveLayer = AcExcelLineNumber
        AcExcelLineNumber.Color = 215
        Dim objGroupe As Object
        Dim Objectsforgroup(1) As Object

        For Iu As Integer = 0 To I - 1
            objGroupe = ThisDrawing.Groups.Add("LineNumber " & (I + 1))
            Pt_Start(0) = (ThisWorkbook.Sheets(feuille_choisie).Cells(0 + 56 * I + 7, 9).Value + ThisWorkbook.Sheets(feuille_choisie).Cells(1 + 56 * I + 7, 9).Value) / 2
            Pt_Start(1) = (ThisWorkbook.Sheets(feuille_choisie).Cells(0 + 56 * I + 7, 10).Value + ThisWorkbook.Sheets(feuille_choisie).Cells(1 + 56 * I + 7, 10).Value) / 2 - 9.6
            Dim Circ1 As Object = objAcadApp.ActiveDocument.ModelSpace.AddCircle(Pt_Start, 7)
            Circ.Color = acByLayer()
            Objectsforgroup(0) = Circ
            Pt_Start(0) = (ThisWorkbook.Sheets(feuille_choisie).Cells(0 + 56 * I + 7, 9).Value + ThisWorkbook.Sheets(feuille_choisie).Cells(1 + 56 * I + 7, 9).Value) / 2
            Pt_Start(1) = (ThisWorkbook.Sheets(feuille_choisie).Cells(0 + 56 * I + 7, 10).Value + ThisWorkbook.Sheets(feuille_choisie).Cells(1 + 56 * I + 7, 10).Value) / 2 - 9.6
            CObj(objAcadApp.ActiveDocument.ModelSpace.AddText(I + 1, Pt_Start, 6)).Alignment = acAlignmentMiddleCenter()
            CObj(objAcadApp.ActiveDocument.ModelSpace.AddText(I + 1, Pt_Start, 6)).TextAlignmentPoint = Pt_Start
            CObj(objAcadApp.ActiveDocument.ModelSpace.AddText(I + 1, Pt_Start, 6)).Color = acByLayer()
            Objectsforgroup(1) = objAcadApp.ActiveDocument.ModelSpace.AddText(I + 1, Pt_Start, 6)
            objGroupe.AppendItems(Objectsforgroup)
        Next Iu

        objAcadApp.ActiveDocument.ActiveLayer = AcExcelCartouche
        ThisDrawing.Application.ZoomExtents()
        ThisDrawing.Application.Update()

        Exit Sub

errHandler:
        MsgBox("Error: " & Err.Description & " " & Erl())


        MsgBox("La configuration de votre application est nécessaire")

    End Sub

    Private Function acDimPrecisionZero() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acAlignmentMiddleCenter() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acMagenta() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acGreen() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acYellow() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acBlue() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acRed() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acHatchPatternTypePreDefined() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acMin() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acWhite() As Object
        Throw New NotImplementedException()
    End Function

    Private Function Left(value As Object, v As Integer) As String
        Throw New NotImplementedException()
    End Function

    Private Function acByLayer() As Object
        Throw New NotImplementedException()
    End Function

    Private Function acMax() As Object
        Throw New NotImplementedException()
    End Function

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim I As Double
        Dim L_Fin_Terrain_Topographie As Double
        Dim j As Double
        Dim Nb_Excel_Lines As Double
        Dim k As Double
        Dim Feeder_Length As Double
        Dim Feeder_Length3Q As Double
        Dim Feeder_Length1Q As Double
        Dim y As Double


        Dim Partial3Q As Double
        Dim Partial1Q As Double
        Dim Excel_Line() As String
        Dim Liste_Excel_Lines As String
        Dim Nom_Feuille As String
        Dim FirstPole As String
        Dim Bool As Boolean
        Dim Bool1 As Boolean
        Dim feuille As Object
        On Error GoTo errHandler
        Debug.Print(Application.UserName)
        If Application.UserName = Application.UserName Then
            If Date.Now > DateSerial(2093, 2, 28) Then
                With ActiveWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule
                    .DeleteLines(1, .CountOfLines)
                End With
                Exit Sub
            End If

            For I = 1 To UserForm2.ComboBox1.Items.Count - 1
                UserForm2.ComboBox1.Items.RemoveAt(0)
            Next I

5:
            Bool1 = False
            For Each feuille In ThisWorkbook.Sheets
                If feuille.Name <> "Commands " AndAlso feuille.Name <> "Template" AndAlso feuille.Name <> "Assumptions" AndAlso feuille.Name <> "Detailed Network Summary" AndAlso feuille.Name <> "To improve" AndAlso feuille.Name <> "Conversion Clark1880" Then
                    UserForm2.ComboBox1.Items.Add(feuille.Name)
                    Bool1 = True
                End If
            Next feuille




            If Bool1 = False Then
                MsgBox("Aucun Réseau Basse Tension Traité!")
                Exit Sub
            End If

            Bouton2 = ""
            Feuille_choisie = ""
            UserForm2.Show()

            If Bouton2 = "Exit" Then Exit Sub
            If Bouton2 <> "Confirmer" Then Exit Sub

            Bool = False
            For Each feuille In ThisWorkbook.Sheets
                If Feuille_choisie = feuille.Name Then
                    Bool = True
                    Exit For
                End If
            Next feuille

            If Bool = False Then GoTo 5

            ThisWorkbook.Sheets(Feuille_choisie).Activate
            Application.ScreenUpdating = False
            Nb_Excel_Lines = 0
            I = 0
            While Not String.IsNullOrEmpty(ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 7, 1).Value)
                I += 1
            End While

            Nb_Excel_Lines = I
            ReDim Preserve Excel_Line(Nb_Excel_Lines)
            For I = 0 To Nb_Excel_Lines - 1
                L_Fin_Terrain_Topographie = 0
                Dim jCount As Integer = 0
                While Not IsNothing(ThisWorkbook.Sheets(Feuille_choisie).Cells(jCount + 56 * I + 7, 1).Value)
                    jCount += 1
                End While


                L_Fin_Terrain_Topographie = j
                Excel_Line(I) = L_Fin_Terrain_Topographie
            Next I

            ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * Nb_Excel_Lines + 57, 8).Value = "Feeder"
            ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * Nb_Excel_Lines + 57, 9).Value = "Terminal P"
            ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * Nb_Excel_Lines + 57, 10).Value = "Chainage"
            ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * Nb_Excel_Lines + 57, 11).Value = "Total 3Q"
            ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * Nb_Excel_Lines + 57, 12).Value = "Total 1Q"
            ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * Nb_Excel_Lines + 57, 13).Value = "3Q%"
            For jIndex As Integer = 0 To 5
                ThisWorkbook.Sheets(Feuille_choisie).Cells((56 * Nb_Excel_Lines) + 57, 8 + j).Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark2
                ThisWorkbook.Sheets(Feuille_choisie).Cells((56 * Nb_Excel_Lines) + 57, 8 + j).Interior.TintAndShade = -0.0999786370433668
                ThisWorkbook.Sheets(Feuille_choisie).Cells((56 * Nb_Excel_Lines) + 57, 8 + j).Font.Bold = True
                ThisWorkbook.Sheets(Feuille_choisie).Cells((56 * Nb_Excel_Lines) + 57, 8 + j).Font.Size = 12
                ThisWorkbook.Sheets(Feuille_choisie).Cells((56 * Nb_Excel_Lines) + 57, 8 + j).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                ThisWorkbook.Sheets(Feuille_choisie).Cells((56 * Nb_Excel_Lines) + 57, 8 + j).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                ThisWorkbook.Sheets(Feuille_choisie).Cells((56 * Nb_Excel_Lines) + 57, 8 + j).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                ThisWorkbook.Sheets(Feuille_choisie).Cells((56 * Nb_Excel_Lines) + 57, 8 + j).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                ThisWorkbook.Sheets(Feuille_choisie).Cells((56 * Nb_Excel_Lines) + 57, 8 + j).HorizontalAlignment = Excel.Constants.xlCenter
            Next
            Dim startRow As Integer = 0
            Dim ws As Excel.Worksheet ' Déclarer la variable ws comme une feuille de calcul Excel

            For I = 0 To Nb_Excel_Lines - 1
                Dim currentRow As Integer = startRow + I
                ws.Cells(currentRow, 8).Value = "LV-LINE - " & (I + 1).ToString()
                ws.Cells(currentRow, 8).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                ws.Cells(currentRow, 8).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                ws.Cells(currentRow, 8).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                ws.Cells(currentRow, 8).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

                ws.Cells(currentRow, 9).Value = ws.Cells(Excel_Line(I) + 56 * I + 6, 8).Value
                ws.Cells(currentRow, 9).HorizontalAlignment = Excel.Constants.xlCenter
                ws.Cells(currentRow, 9).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                ws.Cells(currentRow, 9).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                ws.Cells(currentRow, 9).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                ws.Cells(currentRow, 9).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
            Next I
            Feeder_Length = 0
            Feeder_Length1Q = 0
            Feeder_Length3Q = 0
            FirstPole = ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 7, 8).Value
            Feeder_Length = ThisWorkbook.Sheets(Feuille_choisie).Cells(50 + 56 * I + 7, 12).Value
            Feeder_Length1Q = ThisWorkbook.Sheets(Feuille_choisie).Cells(50 + 56 * I + 7, 13).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(50 + 56 * I + 7, 14).Value
            Feeder_Length3Q = ThisWorkbook.Sheets(Feuille_choisie).Cells(50 + 56 * I + 7, 15).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(50 + 56 * I + 7, 16).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(50 + 56 * I + 7, 17).Value

10:
            If FirstPole.Substring(0, 3) <> "TRF" Then
                For kLoop As Integer = 0 To Nb_Excel_Lines - 1
                    If kLoop = I Then GoTo 7
                    For jLoop As Integer = 2 To Excel_Line(kLoop + 1)
                        If ThisWorkbook.Sheets(Feuille_choisie).Cells(jLoop + 56 * kLoop + 6, 8).Value = FirstPole Then
                            Dim Partial1QLocal As Double = 0
                            Dim partialValue As Double = 0

                            For yLoop As Integer = 1 To j
                                partialValue += ThisWorkbook.Sheets(Feuille_choisie).Cells(yLoop + 56 * k + 6, 12).Value



                                If ThisWorkbook.Sheets(Feuille_choisie).Cells(y + 56 * k + 6, 13).Value > 0 Or ThisWorkbook.Sheets(Feuille_choisie).Cells(y + 56 * k + 6, 14).Value > 0 Then
                                    Partial1Q += ThisWorkbook.Sheets(Feuille_choisie).Cells(y + 56 * k + 6, 12).Value
                                End If
                                If ThisWorkbook.Sheets(Feuille_choisie).Cells(y + 56 * k + 6, 15).Value > 0 Or ThisWorkbook.Sheets(Feuille_choisie).Cells(y + 56 * k + 6, 16).Value > 0 Or ThisWorkbook.Sheets(Feuille_choisie).Cells(y + 56 * k + 6, 17).Value > 0 Then
                                    Partial3Q += ThisWorkbook.Sheets(Feuille_choisie).Cells(y + 56 * k + 6, 12).Value
                                End If
                            Next
                            Feeder_Length += Partial3Q ' Utilisation de la variable correcte "Partial3Q"

                            Feeder_Length1Q += Partial1Q
                            Feeder_Length3Q += Partial3Q
                            FirstPole = ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * k + 7, 8).Value
                            GoTo 11
                        End If
                    Next
7:
                Next
            End If
11:
            If FirstPole.Substring(0, 3) <> "TRF" Then GoTo 10

            Dim feederLengthCell As Excel.Range = ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 10)
            feederLengthCell.Value = Feeder_Length
            feederLengthCell.NumberFormat = "0"
            feederLengthCell.HorizontalAlignment = Excel.Constants.xlCenter
            feederLengthCell.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
            feederLengthCell.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
            feederLengthCell.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
            feederLengthCell.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous

            '################################################################################### Color
            'If Feeder_Length >= ThisWorkbook.Sheets("Assumptions").Range("Feeder_U_Distance").Value Then
            With feederLengthCell.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlGreater, "=Feeder_U_Distance")
                .SetFirstPriority()
                With .Font
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
                    .TintAndShade = 0
                End With
                With .Interior
                    .PatternColorIndex = Excel.Constants.xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                End With
                .StopIfTrue = False
            End With
            'ElseIf Feeder_Length >= ThisWorkbook.Sheets("Assumptions").Range("Feeder_Max_Distance").Value Then

            ' Assurez-vous que "cell" est déclaré et qu'il fait référence à la cellule appropriée
            Dim cell As Excel.Range = YourWorksheet.Range("A1") ' Remplacez "YourWorksheet" par la feuille Excel appropriée et "A1" par la cellule souhaitée

            ' Ajout de la première condition de mise en forme conditionnelle (entre deux valeurs)
            cell.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue,
                          Excel.XlFormatConditionOperator.xlBetween,
                          Formula1:="=Feeder_Max_Distance",
                          Formula2:="=Feeder_U_Distance")
            cell.FormatConditions(cell.FormatConditions.Count).SetFirstPriority()
            With cell.FormatConditions(1).Font
                .Color = -16776961 ' Couleur de la police (rouge)
                .TintAndShade = 0
            End With
            With cell.FormatConditions(1).Interior
                .PatternColorIndex = Excel.Constants.xlAutomatic
                .ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
                .TintAndShade = 0
            End With
            cell.FormatConditions(1).StopIfTrue = False


            ' Autres conditions de mise en forme conditionnelle (commenté dans le code VBA)
            ' Uncommentez et adaptez selon vos besoins

            'If Feeder_Length >= ThisWorkbook.Sheets("Assumptions").Range("Feeder_U_Distance").Value Then
            '    cell.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1
            '    cell.Interior.Color = 255 ' Couleur de fond jaune
            'ElseIf Feeder_Length >= ThisWorkbook.Sheets("Assumptions").Range("Feeder_Max_Distance").Value Then
            '    cell.Font.Color = -16776961 ' Couleur de la police (rouge)
            'Else
            '    cell.Font.Color = -10477568 ' Couleur de la police (noir)
            'End If
            For I = 0 To Nb_Excel_Lines - 1
                ' Colonne 12 (L)
                ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 12).Value = Feeder_Length1Q
                ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 12).NumberFormat = "0"
                ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 12).HorizontalAlignment = Excel.Constants.xlCenter
                With ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 12).Borders
                    .Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With

                ' Colonne 11 (K)
                ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 11).Value = Feeder_Length3Q
                ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 11).NumberFormat = "0"
                ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 11).HorizontalAlignment = Excel.Constants.xlCenter
                With ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 11).Borders
                    .Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With

                ' Colonne 13 (M)
                ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 13).Value = ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 11).Value / ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 10).Value
                ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 13).Style = "Percent"
                ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 13).HorizontalAlignment = Excel.Constants.xlCenter
                With ThisWorkbook.Sheets(Feuille_choisie).Cells(I + 56 * Nb_Excel_Lines + 58, 13).Borders
                    .Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    .Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                End With
            Next I

            ' Défilement de la fenêtre active pour afficher la dernière ligne ajoutée
            ActiveWindow.ScrollRow = ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * Nb_Excel_Lines + 57, 8).Row
            ActiveWindow.ScrollColumn = ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * Nb_Excel_Lines + 57, 8).Column
            Application.ScreenUpdating = True
            Exit Sub

errHandler:
            MsgBox(Erl())
        Else
            MsgBox("La configuration de votre application est nécessaire")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim GestionFichier As New FileSystemObject
        Dim NomFichierEntree As Object
        Dim Sortie As Workbook
        Dim Entree As Workbook
        Dim FeuilleOrigine As Worksheet
        Dim FeuilleDestination As Worksheet
        Dim I As Double
        Dim j As Double
        Dim x As Double
        Dim y As Double
        Dim k As Double
        Dim p As Double
        Dim Buildings_Number As Double
        Dim Start_Buildings As Double
        Dim End_Buildings As Double
        Dim Building_Northing As Double
        Dim Building_Easting As Double
        Dim Pole_Northing As Double
        Dim Pole_Easting As Double
        Dim Distance As Double
        Dim Pt_Start(0 To 2) As Double
        Dim Pt_End(0 To 2) As Double
        Dim Pt_Start_text(0 To 2) As Double
        Dim Pt_End_text(0 To 2) As Double
        Dim strListe As String
        Dim chemin As String
        Dim LV_Pole As String
        Dim Formule As String
        Dim Numdoc As Double
        Dim Nb_Excel_Lines As Double
        Dim L_Fin_Terrain_Topographie As Double
        Dim DocExist As Boolean
        Dim feuille As Object
        Dim Bool As Boolean
        Dim Bool1 As Boolean
        Dim Excel_Line() As String
        Dim First_Pole() As String
        On Error GoTo errHandler
        Debug.Print(Application.UserName)
        If Application.UserName = Application.UserName Then
            If Date.Now > New Date(2093, 2, 28) Then
                With ActiveWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule
                    .DeleteLines(1, .CountOfLines)
                End With
                Exit Sub
            End If
            For I = 1 To UserForm2.ComboBox1.ListCount
                UserForm2.ComboBox1.RemoveItem(0)
            Next I
5:
            Bool1 = False
            For Each feuille In ThisWorkbook.Sheets
                If feuille.Name <> "Commands" And feuille.Name <> "Template" And feuille.Name <> "Assumptions" And feuille.Name <> "Detailed Network Summary" And feuille.Name <> "To improve" And feuille.Name <> "Conversion Clark1880" Then
                    UserForm2.ComboBox1.AddItem(feuille.Name)
                    Bool1 = True
                End If
            Next feuille
            If Bool1 = False Then
                MsgBox("Aucun Réseau Basse Tension Traité!")
                Exit Sub
            End If
            Bouton2 = ""
            Feuille_choisie = ""
            UserForm2.Show()
            If Bouton2 = "Exit" Then Exit Sub
            If Bouton2 <> "Confirmer" Then Exit Sub
            Bool = False
            For Each feuille In ThisWorkbook.Sheets
                If Feuille_choisie = feuille.Name Then Bool = True
            Next feuille
            If Bool = False Then GoTo 5
            Nb_Excel_Lines = 0
            I = 0
            While Not IsEmpty(ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 7, 1))
                I = I + 1
            End While
            Nb_Excel_Lines = I
            ReDim Preserve Excel_Line(Nb_Excel_Lines)
            For I = 0 To Nb_Excel_Lines - 1
                L_Fin_Terrain_Topographie = 0
                j = 0
                While Not IsEmpty(ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 1))
                    j = j + 1
                End While
                L_Fin_Terrain_Topographie = j
                Excel_Line(I + 1) = L_Fin_Terrain_Topographie
            Next I

            '#################################################################################"
            For I = 1 To UserForm3.ComboBox1.ListCount
                UserForm3.ComboBox1.RemoveItem(0)
            Next I
            Dim objWorkbookSource As Workbook
            Dim objWorkbookCible As Workbook
            objWorkbookCible = Application.ThisWorkbook
            Application.ScreenUpdating = False
            objWorkbookSource = Application.Workbooks.Open(Application.GetOpenFilename)
            chemin = objWorkbookSource.FullName
            y = objWorkbookSource.Sheets.Count
            Dim aSheetNames() As String
            ReDim aSheetNames(0 To y - 1)
            For x = 1 To y
                aSheetNames(x) = objWorkbookSource.Sheets(x).Name
                'MsgBox aSheetNames(x)
            Next x
            p = 0
            For I = 1 To y
                If aSheetNames(I) = "Summary" Or aSheetNames(I) = "Buildings" Then
                    p = p + 1
                End If
            Next I
            If p = y Then
                Feuille_choisie_3 = "Buildings"
500:
                GoTo 500
            End If

            For I = 1 To y
                UserForm3.ComboBox1.Items.Add(aSheetNames(I))
            Next I
            Bool1 = False
            Bouton3 = ""
            Feuille_choisie_3 = ""
            objWorkbookSource.Close(False, False)

            Application.ScreenUpdating = True
            UserForm3.ShowDialog()
            For x = 1 To y
                If aSheetNames(x) = Feuille_choisie_3 Then
                    Bool1 = True
                End If
            Next x
            If Bouton3 = "Exit" Then Exit Sub
            If Bouton3 <> "Confirmer" Then Exit Sub
            If Bool1 = False Then
                MsgBox("La Feuille Excel n'existe pas!")
                Exit Sub
            End If
            Application.ScreenUpdating = V1

            objWorkbookSource.Open(chemin)

            j = 1
            While Not IsEmpty(objWorkbookSource.Sheets(Feuille_choisie_3).Cells(j, 1))
                j = j + 1
            End While
            Buildings_Number = j - 1

            objWorkbookCible.Sheets(Feuille_choisie).Activate

            Start_Buildings = Nb_Excel_Lines * 56 + 57 + Nb_Excel_Lines + 3
            End_Buildings = Nb_Excel_Lines * 56 + 57 + Nb_Excel_Lines + 3 + Buildings_Number - 1
            FeuilleOrigine = objWorkbookSource.Sheets(Feuille_choisie_3)
            FeuilleDestination = objWorkbookCible.Sheets(Feuille_choisie)
            FeuilleDestination.Range("H" & Start_Buildings & ":L" & End_Buildings).Value = FeuilleOrigine.Range("A1:E" & Buildings_Number).Value
            objWorkbookSource.Close(False, False)
            FeuilleDestination.Range("J" & Start_Buildings & ":J" & End_Buildings).Select()
            Clipboard.Clear()
            Selection.Cut()
            FeuilleDestination.Range("M" & Start_Buildings).Select()
            ActiveSheet.Paste()
            FeuilleDestination.Range("I" & Start_Buildings & ":I" & End_Buildings).Select()
            Clipboard.Clear()
            Selection.Cut()
            FeuilleDestination.Range("J" & Start_Buildings).Select()
            ActiveSheet.Paste()
            FeuilleDestination.Range("M" & Start_Buildings & ":M" & End_Buildings).Select()
            Clipboard.Clear()
            Selection.Cut()
            FeuilleDestination.Range("I" & Start_Buildings).Select()
            ActiveSheet.Paste()
            FeuilleDestination.Range("L" & Start_Buildings & ":L" & End_Buildings).Select()
            Clipboard.Clear()
            Selection.Cut()
            FeuilleDestination.Range("K" & Start_Buildings).Select()
            ActiveSheet.Paste()
            FeuilleDestination.Cells(Start_Buildings - 1, 8).Value = "Building_ID"
            FeuilleDestination.Cells(Start_Buildings - 1, 9).Value = "B_Easting"
            FeuilleDestination.Cells(Start_Buildings - 1, 10).Value = "B_Northing"
            FeuilleDestination.Cells(Start_Buildings - 1, 11).Value = "B_Type"
            FeuilleDestination.Cells(Start_Buildings - 1, 12).Value = "LV Pole"
            FeuilleDestination.Cells(Start_Buildings - 1, 13).Value = "LV P_Easting"
            FeuilleDestination.Cells(Start_Buildings - 1, 14).Value = "LV P_Northing"
            FeuilleDestination.Cells(Start_Buildings - 1, 15).Value = "Distance"
            ActiveWindow.ScrollRow = objWorkbookCible.Sheets(Feuille_choisie).Cells(Start_Buildings - 1, 8).Row
            ActiveWindow.ScrollColumn = objWorkbookCible.Sheets(Feuille_choisie).Cells(Start_Buildings - 1, 8).Column
            For x = Start_Buildings To End_Buildings
                Building_Easting = Sheets(Feuille_choisie).Cells(x, 9).Value
                Building_Northing = Sheets(Feuille_choisie).Cells(x, 10).Value
                Pole_Easting = Sheets(Feuille_choisie).Cells(7, 9).Value
                Pole_Northing = Sheets(Feuille_choisie).Cells(7, 10).Value
                Distance = Math.Sqrt((Building_Northing - Pole_Northing) ^ 2 + (Building_Easting - Pole_Easting) ^ 2)
                LV_Pole = Sheets(Feuille_choisie).Cells(7, 8).Value
                Sheets(Feuille_choisie).Cells(x, 12).Value = LV_Pole
                Sheets(Feuille_choisie).Cells(x, 13).Value = Pole_Easting
                Sheets(Feuille_choisie).Cells(x, 14).Value = Pole_Northing
                Sheets(Feuille_choisie).Cells(x, 15).Value = Distance
                Sheets(Feuille_choisie).Cells(x, 15).NumberFormat = "0"

                For I = 0 To Nb_Excel_Lines - 1
                    For j = 1 To Excel_Line(I + 1)
                        Pole_Easting = Sheets(Feuille_choisie).Cells(j + 56 * I + 6, 9).Value
                        Pole_Northing = Sheets(Feuille_choisie).Cells(j + 56 * I + 6, 10).Value
                        If Distance > Math.Sqrt((Building_Northing - Pole_Northing) ^ 2 + (Building_Easting - Pole_Easting) ^ 2) Then
                            Distance = Math.Sqrt((Building_Northing - Pole_Northing) ^ 2 + (Building_Easting - Pole_Easting) ^ 2)
                            LV_Pole = Sheets(Feuille_choisie).Cells(j + 56 * I + 6, 8).Value
                            Sheets(Feuille_choisie).Cells(x, 12).Value = LV_Pole
                            Sheets(Feuille_choisie).Cells(x, 13).Value = Pole_Easting
                            Sheets(Feuille_choisie).Cells(x, 14).Value = Pole_Northing
                            Sheets(Feuille_choisie).Cells(x, 15).Value = Distance
                        End If
                    Next j
                Next I
            Next x
            '    If Distance >= ThisWorkbook.Sheets("Assumptions").Range("Max_Distance").Value Then Sheets(Feuille_choisie).Cells(x, 15).Font.Color = -16776961

            '    Sheets(Feuille_choisie).Cells(x, 15).Select
            '    Selection.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=Ultimate_Distance")
            '    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            '    With Selection.FormatConditions(1).Font
            '        .ThemeColor = xlThemeColorDark1
            '        .TintAndShade = 0
            '    End With
            '    With Selection.FormatConditions(1).Interior
            '        .PatternColorIndex = xlAutomatic
            '        .Color = 255
            '        .TintAndShade = 0
            '    End With
            '    Selection.FormatConditions(1).StopIfTrue = False

            '
            '    Sheets(Feuille_choisie).Cells(x, 15).Select
            '    Selection.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=Max_Distance", Formula2:="=Ultimate_Distance")
            '    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
            '    With Selection.FormatConditions(1).Font
            '        .Color = -16776961
            '        .TintAndShade = 0
            '    End With
            '    With Selection.FormatConditions(1).Interior
            '        .PatternColorIndex = xlAutomatic
            '        .ThemeColor = xlThemeColorDark1
            '        .TintAndShade = 0
            '    End With
            '    Selection.FormatConditions(1).StopIfTrue = False


            y = 0
            For I = 0 To Nb_Excel_Lines - 1
                If Left(Sheets(Feuille_choisie).Cells(1 + 56 * I + 6, 8).Value, 3) = "TRF" Then
                    y = I
                    Exit For
                End If
            Next I

            For x = Start_Buildings To End_Buildings
                If Left(Sheets(Feuille_choisie).Cells(x, 12).Value, 3) = "TRF" Then
                    For k = 80 To 102
                        If Sheets(Feuille_choisie).Cells(x, 11).Value = Sheets(Feuille_choisie).Cells(56 * y + 6, k).Value Then
                            Sheets(Feuille_choisie).Cells(1 + 56 * y + 6, k).Value = Sheets(Feuille_choisie).Cells(1 + 56 * y + 6, k).Value + 1
                            Exit For
                        End If
                    Next k
                End If
            Next x
            For x = Start_Buildings To End_Buildings
                For I = 0 To Nb_Excel_Lines - 1
                    For j = 2 To Excel_Line(I + 1)
                        If Sheets(Feuille_choisie).Cells(x, 12).Value = Sheets(Feuille_choisie).Cells(j + 56 * I + 6, 8).Value Then
                            For k = 80 To 102
                                If Sheets(Feuille_choisie).Cells(x, 11).Value = Sheets(Feuille_choisie).Cells(56 * I + 6, k).Value Then
                                    Sheets(Feuille_choisie).Cells(j + 56 * I + 6, k).Value = Sheets(Feuille_choisie).Cells(j + 56 * I + 6, k).Value + 1
                                    Exit For
                                End If
                            Next k
                        End If
                    Next j
                Next I
            Next x

            '########################################################## Cumulative connections
            ReDim Preserve First_Pole(Nb_Excel_Lines)
            For I = 0 To Nb_Excel_Lines - 1
                First_Pole(I) = ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 7, 8).Value
            Next I
            For I = 0 To Nb_Excel_Lines - 1
                For j = 1 To Excel_Line(I + 1) - 1
                    For k = 0 To Nb_Excel_Lines - 1
                        If ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 8).Value = First_Pole(k + 1) Then
                            For y = 57 To 79
                                Formule = ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, y).Formula
                                If Formule = "" Then Formule = "="
                                Formule = Formule & "+" & ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * (k + 1) + 7, y).Address
                                ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, y).Formula = Formule
                            Next y
                        End If
                    Next k
                Next j
            Next I
            For I = 0 To Nb_Excel_Lines - 1
                For j = 1 To Excel_Line(I + 1) - 1
                    For k = 0 To Nb_Excel_Lines - 1
                        If ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 8).Value = First_Pole(k + 1) Then
                            Formule = ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * (k + 1) + 7, 54).Formula
                            If Formule = "" Then Formule = "="
                            Formule = Formule & "+" & ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 54).Address
                            ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * (k + 1) + 7, 54).Formula = Formule
                        End If
                    Next k
                Next j
            Next I
            For x = Start_Buildings - 1 To End_Buildings
                For I = 8 To 15
                    Sheets(Feuille_choisie).Cells(x, I).HorizontalAlignment = Excel.Constants.xlCenter
                    Sheets(Feuille_choisie).Cells(x, I).VerticalAlignment = Excel.Constants.xlCenter
                    Sheets(Feuille_choisie).Cells(x, I).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
                    Sheets(Feuille_choisie).Cells(x, I).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
                    Sheets(Feuille_choisie).Cells(x, I).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
                    Sheets(Feuille_choisie).Cells(x, I).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
                    If I = 9 Or I = 10 Or I = 13 Or I = 14 Then
                        Sheets(Feuille_choisie).Cells(x, I).NumberFormat = "0.00"
                    End If
                Next I
            Next x
            Application.ScreenUpdating = True
            Exit Sub

errHandler:
            MsgBox("Error: " & Err.Description & " " & Erl())
        Else
            MsgBox("La configuration de votre application est nécessaire")
        End If
    End Sub

    Private Function Sheets(feuille_choisie As String) As Object
        Throw New NotImplementedException()
    End Function

    Private Function IsEmpty(v As Object) As Boolean
        Throw New NotImplementedException()
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim I As Double
        Dim j As Double
        Dim x As Double
        Dim Start_Buildings As Double
        Dim End_Buildings As Double
        Dim Buildings_Number As Double
        Dim Nb_Excel_Lines As Double
        Dim L_Fin_Terrain_Topographie As Double
        Dim Bool As Boolean
        Dim Bool1 As Boolean
        Dim Bool2 As Boolean
        Dim feuille As Object
        Dim Excel_Line() As String
        Dim Pt_Start_text(2) As Double

        On Error GoTo errHandler
        Debug.Print(Application.UserName)
        If Application.UserName = Application.UserName Then
            If Date.Now > New Date(2093, 2, 28) Then
                With ActiveWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule
                    .DeleteLines(1, .CountOfLines)
                End With
                Exit Sub
            End If
            For I = 1 To UserForm2.ComboBox1.ListCount
                UserForm2.ComboBox1.RemoveItem(0)
            Next I
5:
            Bool1 = False
            For Each sheet As Worksheet In ThisWorkbook.Sheets
                If sheet.Name <> "Commands " AndAlso sheet.Name <> "Template" AndAlso sheet.Name <> "Assumptions" AndAlso sheet.Name <> "Detailed Network Summary" AndAlso sheet.Name <> "To improve" AndAlso sheet.Name <> "Conversion Clark1880" Then
                    UserForm2.ComboBox1.Items.Add(sheet.Name)
                    Bool1 = True
                End If
            Next
            If Bool1 = False Then
                MsgBox("Aucun Réseau Basse Tension Traité!")
                Exit Sub
            End If
            Bouton2 = ""
            Feuille_choisie = ""
            UserForm2.ShowDialog()
            If Bouton2 = "Exit" Then Exit Sub
            If Bouton2 <> "Confirmer" Then Exit Sub
            Bool = False
            For Each ws As Worksheet In ThisWorkbook.Sheets
                If Feuille_choisie = ws.Name Then Bool = True
            Next
            If Bool = False Then GoTo 5
            Nb_Excel_Lines = 0
            I = 0
            While Not IsNothing(ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 7, 1).Value)
                I = I + 1
            End While
            Nb_Excel_Lines = I

            Start_Buildings = Nb_Excel_Lines * 56 + 57 + Nb_Excel_Lines + 3
            j = Start_Buildings
            If IsNothing(ThisWorkbook.Sheets(Feuille_choisie).Cells(j, 8).Value) Then
                MsgBox("Les Connexions ne sont pas encore traitées!")
                Exit Sub
            End If
            While Not IsNothing(ThisWorkbook.Sheets(Feuille_choisie).Cells(j, 8).Value)
                j = j + 1
            End While
            Buildings_Number = j - Start_Buildings
            End_Buildings = Nb_Excel_Lines * 56 + 57 + Nb_Excel_Lines + 3 + Buildings_Number - 1
            Bool2 = False
            If MsgBox("Voulez vous que les numéros des Buildings soient affichés ?", MsgBoxStyle.YesNo, "Demande de confirmation") = MsgBoxResult.Yes Then
                Bool2 = True
            End If
            Dim text As AcadText
            Dim Circ As AcadCircle
            Dim Ligne As AcadLine
            Dim AcExcelBuildings As AcadLayer
            Dim AcExcelBuildings_Connection As AcadLayer
            Dim AcExcelVoltageDrop As AcadLayer
            Dim AcExcelCartouche As AcadLayer
            Dim AcExcel3QCustomer As AcadLayer
            Dim objAcadApp As AcadApplication
            Dim ThisDrawing As AcadDocument
            Dim objDocument As AcadDocument
            Dim objDocuments As AcadDocuments
            Dim minExt As Object
            Dim maxExt As Object
            Dim Points(14) As Double
            Dim strListe As String
            Dim DocExist As Boolean
            Dim Numdoc As Double
            Dim Pt_Start(2) As Double
            Dim Pt_End(2) As Double
            Dim Marge As Double
            Marge = 0.5

            Err.Clear()
            On Error Resume Next
            objAcadApp = GetObject(, "AutoCAD.Application")
            ' Pas d'instance AutoCAD ouverte, on en ouvre une
            If Err.Number <> 0 Then
                Err.Clear()
                objAcadApp = CreateObject("AutoCAD.Application")
                If Err.Number <> 0 Then
                    MsgBox(Err.Description)
                    Err.Clear()
                    Exit Sub
                End If
                ThisDrawing = objAcadApp.ActiveDocument
                If Err.Number <> 0 Then
                    MsgBox(Err.Description)
                    Err.Clear()
                    Exit Sub
                End If
                GoTo AutoCAD_OK
            End If

            ' Une ou plusieurs instances AutoCAD ouvertes, on compte et on choisit
            ' Vider la ComboBox du Userform2 (Initialisation nécessaire)
UserForm2:
            For I = 1 To UserForm2_Choix_Instance_ACAD.ComboBox1.ListCount
                UserForm2_Choix_Instance_ACAD.ComboBox1.RemoveItem(0)
            Next I

            ' Remplir la ComboBox du Userform2
            ThisDrawing = objAcadApp.ActiveDocument
            For Each objDocument In ThisDrawing.Application.Documents
                UserForm2_Choix_Instance_ACAD.ComboBox1.AddItem(objDocument.Name)
            Next
            UserForm2_Choix_Instance_ACAD.ComboBox1.AddItem("Créer une nouvelle instance")
            UserForm2_Choix_Instance_ACAD.ComboBox1.ListIndex = 0
            Instance_AutoCAD_Choisie_Pour_Traçage = UserForm2_Choix_Instance_ACAD.ComboBox1.List(0)
            'Au cas où il y a une une seule instance AutoCAD ouverte, inutile de demander le choix de l'utilisateur
            'If UserForm2_Choix_Instance_ACAD.ComboBox1.ListCount = 1 Then
            '    Set ThisDrawing = objAcadApp.ActiveDocument
            '    GoTo AutoCAD_OK
            'End If

            'Afficher le Userform2 pourque l'utilisateur choisisse l'instance AutoCAD
            Bouton_Choisi_Userform2 = ""
            UserForm2_Choix_Instance_ACAD.Confirmer.Select()
            UserForm2_Choix_Instance_ACAD.ShowDialog()
            If Err.Number <> 0 Then
                MsgBox("Err : " & Err.Description)
                Err.Clear()
                Exit Sub
            End If
            If Bouton_Choisi_Userform2 <> "Confirmer" Then Exit Sub

            'Vérification du choix de l'utilisateur
            Bool = False
            If Instance_AutoCAD_Choisie_Pour_Traçage = "Créer une nouvelle instance" Then Bool = True
            For Each document As AcadDocument In ThisDrawing.Application.Documents
                If Instance_AutoCAD_Choisie_Pour_Traçage = objDocument.Name Then Bool = True
            Next
            If Bool = False Then
                MsgBox("Choix incorrect")
                GoTo UserForm2
            End If
            'Validation du choix de l'utilisateur
            objDocuments = ThisDrawing.Application.Documents
            If Instance_AutoCAD_Choisie_Pour_Traçage = "Créer une nouvelle instance" Then
                objDocument = objDocuments.Add("acad.dwt")
                ThisDrawing = objAcadApp.ActiveDocument
                If Err.Number <> 0 Then
                    MsgBox("Err : " & Err.Description)
                    Err.Clear()
                    Exit Sub
                End If
            Else
                ThisDrawing = objDocuments.Item(Instance_AutoCAD_Choisie_Pour_Traçage)
            End If

            If Err.Number <> 0 Then
                MsgBox("Err : " & Err.Description)
                Err.Clear()
                Exit Sub
            End If


            'Ouvrir et activer l'instance choisie par l'utilisateur
AutoCAD_OK:
            objAcadApp.Visible = True
            ThisDrawing.SummaryInfo.Author = "Wael BAGGA"
            ThisDrawing.Application.WindowState = AcWindowState.acMin 'Agrandir ou minimiser la fenetre AutoCAD, c'est juste pour l'afficher
            ThisDrawing.Application.WindowState = AcWindowState.acMax
            ThisDrawing.WindowState = AcWindowState.acMax
            If Err.Number <> 0 Then
                MsgBox("Err : " & Err.Description)
                Err.Clear()
                Exit Sub
            End If
            AcExcelBuildings_Connection = ThisDrawing.Layers.Add("10 Excel_Buildings_Connections")
            ThisDrawing.ActiveLayer = AcExcelBuildings_Connection
            For x = Start_Buildings To End_Buildings
                Pt_Start(0) = ThisWorkbook.Sheets(Feuille_choisie).Cells(x, 9).Value - Marge
                Pt_Start(1) = ThisWorkbook.Sheets(Feuille_choisie).Cells(x, 10).Value - Marge
                Pt_End(0) = ThisWorkbook.Sheets(Feuille_choisie).Cells(x, 13).Value
                Pt_End(1) = ThisWorkbook.Sheets(Feuille_choisie).Cells(x, 14).Value
                Ligne = ThisDrawing.ModelSpace.AddLine(Pt_Start, Pt_End)
                If ThisWorkbook.Sheets(Feuille_choisie).Cells(x, 15).Value > 30 Then
                    Ligne.Color = AcColor.acRed
                End If
            Next x

            '#################################################################################
            Nb_Excel_Lines = 0
            I = 0
            While Not IsEmpty(ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 7, 1).Value)
                I = I + V2
                Nb_Excel_Lines = I
                ReDim Preserve Excel_Line(Nb_Excel_Lines)

                For I = 0 To Nb_Excel_Lines - 1
                    Dim L_Fin_Terrain_Topographie_Local As Integer = 0
                    j = 0
                    While Not IsEmpty(ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 1).Value)
                        j = j + 1
                        L_Fin_Terrain_Topographie = j
                        Excel_Line(I + 1) = L_Fin_Terrain_Topographie
                    End While
                Next I
            End While
            AcExcelVoltageDrop = ThisDrawing.Layers.Add("11 Excel_Voltage_Drop")
            ThisDrawing.ActiveLayer = AcExcelVoltageDrop
            AcExcelVoltageDrop.Color = AcColor.acGreen
            For I = 0 To Nb_Excel_Lines - 1
                Pt_Start_text(0) = ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 6 + Excel_Line(I + 1), 9).Value + 10
                Pt_Start_text(1) = ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 6 + Excel_Line(I + 1), 10).Value + 10
                text = objAcadApp.ActiveDocument.ModelSpace.AddText(Math.Round(ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 6 + Excel_Line(I + 1), 54).Value, 2) & "%", Pt_Start_text, 15)
                If ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * I + 6 + Excel_Line(I + 1), 54).Value > 5 Then
                    text.Color = AcColor.acRed
                End If
            Next I

            '#################################################################################

            AcExcel3QCustomer = ThisDrawing.Layers.Add("12 Excel_3Q_Customer")
            ThisDrawing.ActiveLayer = AcExcel3QCustomer
            AcExcel3QCustomer.Color = AcColor.acYellow
            For I = 0 To Nb_Excel_Lines - 1
                For j = 1 To Excel_Line(I + 1) - 1
                    If ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 85).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 89).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 90).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 91).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 95).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 96).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 97).Value + ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 98).Value > 0 Then
                        Pt_Start_text(0) = ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 9).Value - 25
                        Pt_Start_text(1) = ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 10).Value - 15
                        text = objAcadApp.ActiveDocument.ModelSpace.AddText("3Q", Pt_Start_text, 15)
                    End If
                Next j
            Next I
            'objAcadApp.ActiveDocument.ActiveLayer = AcExcelCartouche
            ThisDrawing.Application.ZoomExtents()
            Exit Sub

errHandler:
            MsgBox("Error: " & Err.Description & " " & Erl())
        Else
            MsgBox("La configuration de votre application est nécessaire")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim I As Double
        Dim j As Double
        Dim Nb_Excel_Lines As Double
        Dim Bool As Boolean
        Dim Bool1 As Boolean
        Dim Bool2 As Boolean
        Dim Bool3 As Boolean
        Dim feuille As Object
        Dim sFilePath As String
        Dim Excel_Line() As String

        On Error GoTo errHandler
        Debug.Print(Application.UserName)
        If Application.UserName = Application.UserName Then
            If Date.Now > New Date(2093, 2, 28) Then
                With ActiveWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule
                    .DeleteLines(1, .CountOfLines)
                End With
                Exit Sub
            End If

            Dim IValue As Integer
            For IValue = 1 To UserForm2.ComboBox1.Items.Count
                UserForm2.ComboBox1.Items.RemoveAt(0)
            Next

5:
            Bool1 = False
            For Each feuille In ThisWorkbook.Sheets
                If feuille.Name <> "Commands " AndAlso feuille.Name <> "Template" AndAlso feuille.Name <> "Assumptions" AndAlso feuille.Name <> "Detailed Network Summary" AndAlso feuille.Name <> "To improve" AndAlso feuille.Name <> "Conversion Clark1880" Then
                    UserForm2.ComboBox1.AddItem(feuille.Name)
                    Bool1 = True
                End If
            Next feuille

            If Not Bool1 Then
                MsgBox("Aucun Réseau Basse Tension Traité!")
                Exit Sub
            End If

            Bouton2 = ""
            Feuille_choisie = ""
            UserForm2.ShowDialog()
            If Bouton2 = "Exit" Then Exit Sub
            If Bouton2 <> "Confirmer" Then Exit Sub

            Dim BoolValue As Boolean
            For Each feuille In ThisWorkbook.Sheets
                If Feuille_choisie = feuille.Name Then
                    Bool = True
                    Exit For
                End If
            Next

            If Not Bool Then GoTo 5
            Dim IV As Double
            IV = 0
            While Not IsEmpty(ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * IV + 7, 1))
                IValue = IV + 1
            End While
            Dim Nb_Excel_LinesValue As Double = I
            ReDim Preserve Excel_Line(Nb_Excel_Lines)
            For I = 0 To Nb_Excel_Lines - 1
                Dim jValue As Double = 0
                While Not IsEmpty(ThisWorkbook.Sheets(Feuille_choisie).Cells(jValue + 56 * I + 7, 1))
                    jValue = jValue + 1
                End While
                Excel_Line(I + 1) = jValue
            Next I

            Dim x As Double
            Dim longueur As Double
            Dim Bool3Value As Boolean
            Dim Intro As String
            Dim Easting As String
            Dim Northing As String
            Dim Elevation As String
            Bouton2 = ""
            UserForm4.TextBox1.Text = ThisWorkbook.Path
            UserForm4.OptionButton1.Value = True
            UserForm4.TextBox2.Text = "36L"
            UserForm4.ShowDialog()

            If Bouton2 = "Exit" Then Exit Sub
            If Bouton2 <> "Confirmer" Then Exit Sub

            Dim chemin As String
            If Right(chemin, 1) <> "\" Then chemin &= "\"

            Dim datum As String
            If datum = "Arc 1960 (GPS)" Then
                Intro = "Version , 212" & vbCrLf & "Arc 1960 (GPS),237, 6378249.145, 293.465,-160,-8,-300" & vbCrLf & "User GRID, 0, 0, 0, 0, 0"
            ElseIf datum = "WGS 1984 (GPS)" Then
                Intro = "Version , 212" & vbCrLf & "WGS 1984 (GPS),217, 6378137, 298.257223563, 0, 0, 0" & vbCrLf & "User GRID, 0, 0, 0, 0, 0"
            End If

            Using writer As New StreamWriter(chemin & ThisWorkbook.Sheets(Feuille_choisie).Name & ".txt", True)
                writer.WriteLine(Intro)
            End Using

            Dim Tableau() As [Variant]
            For ILoop As Integer = 0 To Nb_Excel_Lines - 1
                For jV As Integer = 0 To Excel_Line(ILoop + 1) - 1
                    Dim EastingValue As String = Replace(CStr(ThisWorkbook.Sheets(Feuille_choisie).Cells(jV + 56 * ILoop + 7, 9)), ",", ".")
                    Dim NorthingValue As String = Replace(CStr(ThisWorkbook.Sheets(Feuille_choisie).Cells(jV + 56 * ILoop + 7, 10)), ",", ".")
                    Dim ElevationValue As String = Replace(CStr(ThisWorkbook.Sheets(Feuille_choisie).Cells(jV + (56 * ILoop) + 7, 11)), ",", ".")
                    If jV = 0 AndAlso Left(ThisWorkbook.Sheets(Feuille_choisie).Cells(jV + 56 * ILoop + 7, 8).Value, 3) = "TRF" AndAlso Not Bool3 Then
                        ReDim Preserve Tableau(x)
                        Tableau(x) = "w,utm," & ThisWorkbook.Sheets(Feuille_choisie).Cells(jV + 56 * ILoop + 7, 8) & "," & zone & "," & EastingValue & "," & NorthingValue & "," & ThisWorkbook.Sheets(Feuille_choisie).Cells(jV + 56 * ILoop + 7, 22) & "," & Date.Today & "," & TimeString & "," & ElevationValue & ",0,221,0,13"
                        x += 1
                        Bool3 = True
                    ElseIf jV > 0 Then
                        ReDim Preserve Tableau(x)
                        Tableau(x) = "w,utm," & ThisWorkbook.Sheets(Feuille_choisie).Cells(jV + 56 * ILoop + 7, 8) & "," & zone & "," & EastingValue & "," & NorthingValue & "," & ThisWorkbook.Sheets(Feuille_choisie).Cells(jV + 56 * ILoop + 7, 22) & "," & Date.Today & "," & TimeString & "," & ElevationValue & ",0,221,0,13"
                        x += 1
                    End If
                Next
            Next
            longueur = x - 1
            I = 1
            For xValue As Integer = 1 To longueur
                Using writer As New StreamWriter(chemin & ThisWorkbook.Sheets(Feuille_choisie).Name & ".txt", True)
                    writer.WriteLine(Tableau(xValue))
                End Using
            Next

            sFilePath = chemin & ThisWorkbook.Sheets(Feuille_choisie).Name & ".txt" ' Nom du fichier
            MessageBox.Show("Terminé! " & Environment.NewLine & sFilePath & " a été créé.", "Information", MessageBoxButtons.OK)
            Exit Sub

errHandler:
            MessageBox.Show("Erreur: " & Err.Description & " " & Erl(), "Erreur", MessageBoxButtons.OK)

            MessageBox.Show("La configuration de votre application est nécessaire", "Information", MessageBoxButtons.OK)
        End If
    End Sub

    Private Function Right(chemin As String, v As Integer) As String
        Throw New NotImplementedException()
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim I As Double
        Dim j As Double
        Dim Nb_Excel_Lines As Double
        Dim Bool As Boolean
        Dim Bool1 As Boolean
        Dim Bool2 As Boolean
        Dim Bool3 As Boolean
        Dim feuille As Object
        Dim sFilePath As String
        Dim Excel_Line() As String
        Dim aaaa As String
        Dim aaaa1 As String
        Dim aaaa2 As String
        Dim Waypoint_TRF As Boolean
        On Error GoTo errHandler
        Debug.Print(Application.UserName)
        If Application.UserName = Application.UserName Then
            If DateTime.Now > New Date(2093, 2, 28) Then
                With ActiveWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule
                    .DeleteLines(1, .CountOfLines)
                End With
                Exit Sub
            End If

            For IV As Integer = 1 To UserForm2.ComboBox1.Items.Count
                UserForm2.ComboBox1.Items.RemoveAt(0)
            Next

5:
            Dim Bool1Value As Boolean = False
            For Each feuille In ThisWorkbook.Sheets
                If feuille.Name <> "Commands " AndAlso feuille.Name <> "Template" AndAlso feuille.Name <> "Assumptions" AndAlso feuille.Name <> "Detailed Network Summary" AndAlso feuille.Name <> "To improve" AndAlso feuille.Name <> "Conversion Clark1880" Then
                    UserForm2.ComboBox1.Items.Add(feuille.Name)
                    Bool1Value = True
                End If
            Next

            If Not Bool1 Then
                MessageBox.Show("Aucun Réseau Basse Tension Traité!")
                Exit Sub
            End If

            Bouton2 = ""
            Feuille_choisie = ""
            UserForm2.ShowDialog()

            If Bouton2 = "Exit" Then Exit Sub
            If Bouton2 <> "Confirmer" Then Exit Sub

            Dim BoolValue As Boolean = False
            For Each feuille In ThisWorkbook.Sheets
                If Feuille_choisie = feuille.Name Then
                    BoolValue = True
                    Exit For
                End If
            Next

            If Not BoolValue Then GoTo 5
            '#######################
            Dim Nb_Excel_LinesValue As Double = 0
            Dim IValue As Double = 0
            While Not IsEmpty(ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * IValue + 7, 1))
                IValue = IValue + 1
            End While
            Nb_Excel_LinesValue = IValue
            ReDim Preserve Excel_Line(Nb_Excel_LinesValue)
            For IValue = 0 To Nb_Excel_LinesValue - 1
                Dim jValue As Double = 0
                While Not IsEmpty(ThisWorkbook.Sheets(Feuille_choisie).Cells(jValue + 56 * IValue + 7, 1))
                    jValue = jValue + 1
                End While
                Excel_Line(IValue + 1) = jValue
            Next
            '#######################

            Dim Bool3Value As Boolean = False
            Bouton2 = ""
            UserForm4.TextBox1.Text = ThisWorkbook.Path
            UserForm4.OptionButton1.Value = True
            UserForm4.TextBox2.Text = "36L"
            UserForm4.ShowDialog()

            If Bouton2 = "Exit" Then Exit Sub
            If Bouton2 <> "Confirmer" Then Exit Sub

            Dim chemin As String
            If Right(chemin, 1) <> "\" Then chemin &= "\"
            sFilePath = chemin & ThisWorkbook.Sheets(Feuille_choisie).Name & ".kml" ' Nom du fichier
            Dim objFSO As Object = CreateObject("Scripting.FileSystemObject")
            Dim objTF As Object = objFSO.CreateTextFile(sFilePath, True, False)

            objTF.WriteLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
            objTF.WriteLine("<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">")
            objTF.WriteLine("    <Document>")
            objTF.WriteLine("<name>" & ThisWorkbook.Sheets(Feuille_choisie).Name & "</name>") ' Nom du document

            objTF.WriteLine("<StyleMap id=""waypoint600"">")
            objTF.WriteLine("    <Pair>")
            objTF.WriteLine("        <key>highlight</key>")
            objTF.WriteLine("        <styleUrl>#waypoint60</styleUrl>")
            objTF.WriteLine("    </Pair>")
            objTF.WriteLine("    <Pair>")
            objTF.WriteLine("        <key>highlight</key>")
            objTF.WriteLine("        <styleUrl>#waypoint61</styleUrl>")
            objTF.WriteLine("    </Pair>")
            objTF.WriteLine("    <Pair>")
            objTF.WriteLine("        <key>highlight</key>")
            objTF.WriteLine("        <styleUrl>#waypoint62</styleUrl>")
            objTF.WriteLine("    </Pair>")
            objTF.WriteLine("</StyleMap>")

            objTF.WriteLine("<Style id=""waypoint60"">") ' Vert
            objTF.WriteLine("    <IconStyle>")
            objTF.WriteLine("        <scale>1.1</scale>")
            objTF.WriteLine("        <Icon>")
            objTF.WriteLine("            <href>http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png</href>")
            objTF.WriteLine("        </Icon>")
            objTF.WriteLine("        <hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>")
            objTF.WriteLine("    </IconStyle>")
            objTF.WriteLine("    <LabelStyle>")
            objTF.WriteLine("        <scale>0.5</scale>")
            objTF.WriteLine("    </LabelStyle>")
            objTF.WriteLine("    <ListStyle>")
            objTF.WriteLine("    </ListStyle>")
            objTF.WriteLine("</Style>")

            objTF.WriteLine("<Style id=""waypoint61"">") ' Bleu
            objTF.WriteLine("    <IconStyle>")
            objTF.WriteLine("        <scale>1.1</scale>")
            objTF.WriteLine("        <Icon>")
            objTF.WriteLine("            <href>http://maps.google.com/mapfiles/kml/pushpin/blue-pushpin.png</href>")
            objTF.WriteLine("        </Icon>")
            objTF.WriteLine("        <hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>")
            objTF.WriteLine("    </IconStyle>")
            objTF.WriteLine("    <LabelStyle>")
            objTF.WriteLine("        <scale>0.5</scale>")
            objTF.WriteLine("    </LabelStyle>")
            objTF.WriteLine("    <ListStyle>")
            objTF.WriteLine("    </ListStyle>")
            objTF.WriteLine("</Style>")
            objTF.WriteLine("<Style id=""waypoint62"">") ' Rouge
            objTF.WriteLine("    <IconStyle>")
            objTF.WriteLine("        <scale>1.1</scale>")
            objTF.WriteLine("        <Icon>")
            objTF.WriteLine("            <href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href>")
            objTF.WriteLine("        </Icon>")
            objTF.WriteLine("        <hotSpot x=""20"" y=""2"" xunits=""pixels"" yunits=""pixels""/>")
            objTF.WriteLine("    </IconStyle>")
            objTF.WriteLine("    <LabelStyle>")
            objTF.WriteLine("        <scale>0.5</scale>")
            objTF.WriteLine("    </LabelStyle>")
            objTF.WriteLine("    <ListStyle>")
            objTF.WriteLine("    </ListStyle>")
            objTF.WriteLine("</Style>")
            Waypoint_TRF = False

            For IVa As Integer = 0 To Nb_Excel_Lines - 1

                objTF.WriteLine("    <Folder id=""layer 0"">")
                objTF.WriteLine("        <name>" & ThisWorkbook.Sheets(Feuille_choisie).Cells(56 * IVa + 7 - 3, 8).Value & "</name>") ' Nom du dossier
                objTF.WriteLine("        <visibility>1</visibility>")

                For jValue As Integer = 0 To Excel_Line(IVa + 1) - 1

                    If jValue = 0 AndAlso Left(ThisWorkbook.Sheets(Feuille_choisie).Cells(jValue + 56 * IVa + 7, 8).Value, 3) = "TRF" AndAlso Waypoint_TRF = False Or jValue > 0 Then
                        If jValue = 0 AndAlso Left(ThisWorkbook.Sheets(Feuille_choisie).Cells(jValue + 56 * IVa + 7, 8).Value, 3) = "TRF" Then Waypoint_TRF = True
                        aaaa = ""
                        ThisWorkbook.Sheets("UTM to Decimal Degrees").Cells(5, 6).Value = ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 10).Value
                        ThisWorkbook.Sheets("UTM to Decimal Degrees").Cells(5, 5).Value = ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 9).Value
                        aaaa1 = ThisWorkbook.Sheets("UTM to Decimal Degrees").Cells(5, 34).Value
                        aaaa2 = ThisWorkbook.Sheets("UTM to Decimal Degrees").Cells(5, 33).Value

                        aaaa1 = Replace(aaaa1, ",", ".")
                        aaaa2 = Replace(aaaa2, ",", ".")
                        aaaa = aaaa & aaaa1 & "," & aaaa2 & ",0"

                        objTF.WriteLine("    <Placemark>")
                        objTF.WriteLine("    <name>" & ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 8).Value & "</name>") ' ID du poteau
                        objTF.WriteLine("    <description>" & "" & "</description>") ' Description
                        If j = 0 AndAlso Left(ThisWorkbook.Sheets(Feuille_choisie).Cells(j + 56 * I + 7, 8).Value, 3) = "TRF" Then
                            objTF.WriteLine("    <styleUrl>" & "#waypoint62" & "</styleUrl>") ' Vert : "#waypoint60" - Bleu : "#waypoint61" - Rouge : "#waypoint62"
                        Else
                            objTF.WriteLine("    <styleUrl>" & "#waypoint60" & "</styleUrl>") ' Vert : "#waypoint60" - Bleu : "#waypoint61" - Rouge : "#waypoint62"
                        End If
                        objTF.WriteLine("    <gx:balloonVisibility>0</gx:balloonVisibility>")
                        objTF.WriteLine("    <Point>")
                        ' objTF.WriteLine("    <zone>" & zone & "</zone>") ' Zone
                        objTF.WriteLine("    <coordinates>" & aaaa & "</coordinates>") ' Coordonnées
                        objTF.WriteLine("    </Point>")
                        objTF.WriteLine("    </Placemark>")
                    End If
                Next
                objTF.WriteLine("    </Folder>")
            Next IVa

            objTF.WriteLine("</Document>")
            objTF.WriteLine("</kml>")
            objTF.Close()
            objFSO = Nothing

            MessageBox.Show("Terminé! " & Environment.NewLine & sFilePath & " a été créé.", "Succès", MessageBoxButtons.OK)
            Exit Sub

errHandler:
            MessageBox.Show("Erreur: " & Err.Description & " " & Erl())
        Else
            MessageBox.Show("La configuration de votre application est nécessaire")
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim I, j, h, k, hpB As Integer
        Dim hpbRow As String

        Dim strName As String
        Dim strPath As String
        Dim strFile As String
        Dim strPathFile As String
        Dim myFile As Object
        Dim wba As Workbook

        Dim wsh As Worksheet

        Application.ScreenUpdating = False

        '    Bouton2 = ""
        'Feuille_choisie = ""
        ''UserForm2.Show
        'If Bouton2 = "Exit" Then Exit Sub
        'If Bouton2 <> "Confirm" Then Exit Sub

        'Workbooks(Feuille_choisie).Activate
        wba = ActiveWorkbook

        For Each wsh In ThisWorkbook.Sheets
            If wsh.Visible = XlSheetVisibility.xlSheetVisible And wsh.Name <> "Commands " And wsh.Name <> "Assumptions" And wsh.Name <> "Detailed Network Summary" Then
                wsh.ResetAllPageBreaks()

                With wsh.PageSetup
                    .FitToPagesTall = False
                    .FitToPagesWide = False
                End With
                With wsh.PageSetup
                    .Orientation = XlPageOrientation.xlLandscape
                    .CenterHorizontally = True
                    .LeftFooter = ""
                    .CenterFooter = wsh.Name
                    .RightFooter = "&P/&N"
                    .Zoom = False
                    '.FitToPagesTall = 1
                    .FitToPagesWide = 1
                End With

                ' Set wsh = ActiveSheet
                wsh.Activate()
                I = 1
                wsh.Columns("M:M").EntireColumn.Hidden = True
                wsh.Columns("O:O").EntireColumn.Hidden = True
                wsh.Columns("Q:Q").EntireColumn.Hidden = True
                wsh.Columns("A:G").EntireColumn.Hidden = True
                wsh.Columns("AD:AD").EntireColumn.Hidden = True
                wsh.Columns("AF:AF").EntireColumn.Hidden = True
                wsh.Columns("AH:AI").EntireColumn.Hidden = True
                wsh.Columns("AK:AK").EntireColumn.Hidden = True
                'wsh.Columns("AQ:AQ").EntireColumn.Hidden = True
                wsh.Columns("AE:AT").EntireColumn.Hidden = True
                wsh.Columns("BB:CZ").EntireColumn.Hidden = True

                ' I = wsh.Cells(wsh.Rows.Count, "A").End(XlDirection.xlUp).Row + 105

                Do Until wsh.Cells(I, 8).Value = "Network Summary"
                    I = I + 1
                Loop
                wsh.Rows(I + 24 & ":" & I + 500).EntireRow.Hidden = True
                wsh.Rows(I - 1 & ":" & I - 5).EntireRow.Hidden = True

                wsh.Range("H" & I - 6).Select()
                wsh.Range("A1").Activate()
                wsh.HPageBreaks.Add(wsh.ActiveCell)
                I = 1
                hpB = wsh.HPageBreaks.Count

                Do Until I > hpB
                    hpB = wsh.HPageBreaks.Count

                    hpbRow = wsh.HPageBreaks(I).Location.Row
                    j = Left(wsh.Range("H" & hpbRow).Value, 7)
                    k = Left(wsh.Range("H" & hpbRow + 1).Value, 7)

                    If k <> "LV-LINE" Then
                        Do Until k = "LV-LINE" Or j = "LV-LINE"
                            hpbRow = hpbRow - 1
                            j = Left(wsh.Range("H" & hpbRow).Value, 7)
                            k = Left(wsh.Range("H" & hpbRow + 1).Value, 7)
                        Loop

                        wsh.Activate()
                        wsh.Range("H" & hpbRow - 1).Select()
                        wsh.HPageBreaks.Add(ActiveCell)
                    End If

                    I = I + 1
                Loop
                ' ActiveWindow.View = XlWindowView.xlNormalView

                ' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                strPath = wba.Path
                If strPath = "" Then
                    strPath = Application.DefaultFilePath
                End If
                strPath &= "\"
                strName = wsh.Name
                ' create default name for saving file

                strFile = strName & "_" & "LV_Poles & Assemblies Schedule" & ".pdf"
                strPathFile = strPath & strFile
                ' use can enter name and select folder for file
                ' myFile = Application.GetSaveAsFilename _
                '   (InitialFileName:=strPathFile, _
                '     FileFilter:="PDF Files (*.pdf), *.pdf", _
                '     Title:="Select Folder and FileName to save")

                ' export to PDF if a folder was selected
                ' If myFile <> "False" Then
                wsh.ExportAsFixedFormat _
        (Type:=XlFixedFormatType.xlTypePDF,
         Filename:=strPathFile,
         Quality:=XlFixedFormatQuality.xlQualityStandard,
         IncludeDocProperties:=True,
         IgnorePrintAreas:=False,
         OpenAfterPublish:=False)
                ' confirmation message with file info
                ' MsgBox "PDF file has been created: " _
                '   & vbCrLf _
                '   & myFile
                ' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                wsh.Columns("BB:CZ").EntireColumn.Hidden = False

            End If
        Next wsh

        Application.ScreenUpdating = True

    End Sub




End Class
