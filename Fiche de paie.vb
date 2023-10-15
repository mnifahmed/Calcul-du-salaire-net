Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class FicheDePaie
    Dim Brut As Double = 0
    Dim BrutImposable As Double = 0
    Dim IRPP As Double = 0
    Dim SalAnnuel As Double = 0
    Dim Impot As Double = 0
    Dim abattementEnfant As Integer = 0
    Dim abattementChef As Integer = 0
    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs) Handles TextBox_Base.TextChanged, TextBox_Primes.TextChanged, TextBox_NbHeure.TextChanged, TextBox_MtHeure.TextChanged
        If String.IsNullOrEmpty(TextBox_Base.Text) = False And String.IsNullOrEmpty(TextBox_Primes.Text) = False And String.IsNullOrEmpty(TextBox_NbHeure.Text) = False And String.IsNullOrEmpty(TextBox_MtHeure.Text) = False Then
            Button_Brut.Enabled = True
        Else
            Button_Brut.Enabled = False
        End If
    End Sub

    Private Sub TextBox_TextChanged2(sender As Object, e As EventArgs) Handles TextBox_PrimeBilan.TextChanged, TextBox_PrimeRendement.TextChanged, TextBox_PrimeAnc.TextChanged
        If String.IsNullOrEmpty(TextBox_PrimeBilan.Text) = False And String.IsNullOrEmpty(TextBox_PrimeRendement.Text) = False And String.IsNullOrEmpty(TextBox_PrimeAnc.Text) = False Then
            SalaireBrutToolStripMenuItem.Enabled = True
        Else
            SalaireBrutToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub TextBox_TextChanged3(sender As Object, e As EventArgs) Handles TextBox_M.TextChanged, TextBox_N.TextChanged, TextBox_P.TextChanged, TextBox_Func.TextChanged, RadioButton_H.CheckedChanged, RadioButton_F.CheckedChanged
        If String.IsNullOrEmpty(TextBox_M.Text) = False And String.IsNullOrEmpty(TextBox_N.Text) = False And String.IsNullOrEmpty(TextBox_P.Text) = False And (RadioButton_H.Checked = True Or RadioButton_F.Checked = True) And String.IsNullOrEmpty(TextBox_Func.Text) = False Then
            PrimesToolStripMenuItem.Enabled = True
        Else
            PrimesToolStripMenuItem.Enabled = False
        End If
    End Sub

    Private Sub Button_Brut_Click(sender As Object, e As EventArgs) Handles Button_Brut.Click
        If CheckBox_Chef.Checked = True Then
            abattementChef = 300
        Else
            abattementChef = 0
        End If
        If NumericUpDown_Enfant.Value > 0 Then
            abattementEnfant = (NumericUpDown_Enfant.Value * 100)
        Else
            abattementEnfant = 0
        End If
        If IsNumeric(TextBox_Base.Text) = False Or IsNumeric(TextBox_Primes.Text) = False Or IsNumeric(TextBox_NbHeure.Text) = False Or IsNumeric(TextBox_MtHeure.Text) = False Then
            TextBox_Base.Clear()
            TextBox_Primes.Clear()
            TextBox_NbHeure.Clear()
            TextBox_MtHeure.Clear()
            MessageBox.Show("Veuillez saisir des valeurs numériques.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox_Base.Focus()
        Else
            Brut = CDbl(TextBox_Base.Text) + CDbl(TextBox_Primes.Text) + (CInt(TextBox_NbHeure.Text) * CDbl(TextBox_MtHeure.Text))
            TextBox_Brut.Text = Math.Round(Brut, 3)
            TextBox_CNSS.Text = Math.Round(Brut * 0.0918, 3)
            BrutImposable = Math.Round(Brut * 0.9082, 3)
            TextBox_Imposable.Text = BrutImposable
            SalAnnuel = BrutImposable * 12
            If (SalAnnuel * 0.1) <= 2000 Then
                SalAnnuel *= 0.9
            Else
                SalAnnuel -= 2000
            End If
            SalAnnuel = SalAnnuel - abattementEnfant - abattementChef

            If SalAnnuel < 5000 Then
                Impot = SalAnnuel * 0
            ElseIf SalAnnuel < 20000 Then
                Impot = ((SalAnnuel - 5000) * 0.26) + 50
            ElseIf SalAnnuel < 30000 Then
                Impot = ((SalAnnuel - 20000) * 0.28) + 4100
            ElseIf SalAnnuel < 50000 Then
                Impot = ((SalAnnuel - 30000) * 0.32) + 7000
            ElseIf SalAnnuel >= 50000 Then
                Impot = ((SalAnnuel - 50000) * 0.35) + 13600
            End If
            IRPP = (Impot / 12) - ((SalAnnuel * 0.01 / 12) * 1000) / 1000
            TextBox_IRPP.Text = Math.Round(IRPP, 3)
        End If
        TextBox_Net.Text = Math.Round(BrutImposable - IRPP, 3)
        SalaireNetToolStripMenuItem.Enabled = True
        CréerToolStripMenuItem.Enabled = True
    End Sub

    Private Sub DonnéesPersonnellesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DonnéesPersonnellesToolStripMenuItem.Click
        ImporterToolStripMenuItem.Enabled = False
        ExporterToolStripMenuItem.Enabled = False
        ImprimerToolStripMenuItem.Enabled = False
        GroupBox2.Visible = False
        GroupBox6.Visible = False
        GroupBox3.Visible = False
        ListView.Visible = False
        Me.Width = 465
        Me.Height = 275
        GroupBox1.Visible = True
    End Sub

    Private Sub FicheDePaie_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Application.CurrentCulture = New CultureInfo("ar-TN")
        Me.Width = 465
        Me.Height = 275
        GroupBox1.Visible = True
    End Sub

    Private Sub PrimesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrimesToolStripMenuItem.Click
        ImporterToolStripMenuItem.Enabled = False
        ExporterToolStripMenuItem.Enabled = False
        ImprimerToolStripMenuItem.Enabled = False
        GroupBox2.Visible = False
        GroupBox1.Visible = False
        GroupBox3.Visible = False
        ListView.Visible = False
        Me.Width = 465
        Me.Height = 185
        GroupBox6.Visible = True
    End Sub

    Private Sub SalaireBrutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalaireBrutToolStripMenuItem.Click
        If IsNumeric(TextBox_PrimeBilan.Text) = False Or IsNumeric(TextBox_PrimeRendement.Text) = False Or IsNumeric(TextBox_PrimeAnc.Text) = False Then
            TextBox_PrimeBilan.Clear()
            TextBox_PrimeRendement.Clear()
            TextBox_PrimeAnc.Clear()
            MessageBox.Show("Le montant des primes doit être numérique.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox_PrimeBilan.Focus()
        Else
            ImporterToolStripMenuItem.Enabled = False
            ExporterToolStripMenuItem.Enabled = False
            ImprimerToolStripMenuItem.Enabled = False
            GroupBox6.Visible = False
            GroupBox1.Visible = False
            GroupBox3.Visible = False
            ListView.Visible = False
            Me.Width = 465
            Me.Height = 315
            TextBox_Primes.Text = CDbl(TextBox_PrimeBilan.Text) + CDbl(TextBox_PrimeRendement.Text) + CDbl(TextBox_PrimeAnc.Text)
            GroupBox2.Visible = True
        End If
    End Sub

    Private Sub SalaireNetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalaireNetToolStripMenuItem.Click
        ImporterToolStripMenuItem.Enabled = False
        ExporterToolStripMenuItem.Enabled = False
        ImprimerToolStripMenuItem.Enabled = False
        GroupBox6.Visible = False
        GroupBox1.Visible = False
        GroupBox2.Visible = False
        ListView.Visible = False
        Me.Width = 465
        Me.Height = 220
        GroupBox3.Visible = True
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitterToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ImprimerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImprimerToolStripMenuItem.Click
        Dim printDialog As New PrintDialog With {
            .Document = PrintDocument
        }
        Dim result As DialogResult = printDialog.ShowDialog
        If (result = DialogResult.OK) Then
            PrintDocument.Print()
        End If
    End Sub

    Private Sub CréerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CréerToolStripMenuItem.Click
        If IsNumeric(TextBox_M.Text) = False Then
            TextBox_M.Clear()
            MessageBox.Show("Le matricule doit être numérique.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox_M.Focus()
        ElseIf Char.IsLetter(TextBox_N.Text) = False Then
            TextBox_N.Clear()
            MessageBox.Show("Le nom doit être alphabétique.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox_N.Focus()
        ElseIf Char.IsLetter(TextBox_P.Text) = False Then
            TextBox_P.Clear()
            MessageBox.Show("Le prénom doit être alphabétique.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox_P.Focus()
        ElseIf Char.IsLetter(TextBox_Func.Text) = False Then
            TextBox_Func.Clear()
            MessageBox.Show("La fonction doit être alphabétique.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox_Func.Focus()
        Else
            Dim lvi As New ListViewItem With {
        .Text = ""
        }
            If RadioButton_H.Checked Then
                lvi.SubItems.AddRange(New String() {TextBox_M.Text, TextBox_N.Text, TextBox_P.Text, DateTimePicker.Text, RadioButton_H.Text, TextBox_Func.Text, TextBox_Base.Text, TextBox_Primes.Text, TextBox_Brut.Text, TextBox_CNSS.Text, TextBox_IRPP.Text, TextBox_Net.Text})
            Else
                lvi.SubItems.AddRange(New String() {TextBox_M.Text, TextBox_N.Text, TextBox_P.Text, DateTimePicker.Text, RadioButton_F.Text, TextBox_Func.Text, TextBox_Base.Text, TextBox_Primes.Text, TextBox_Brut.Text, TextBox_CNSS.Text, TextBox_IRPP.Text, TextBox_Net.Text})
            End If
            ListView.Items.Add(lvi)
            TextBox_M.Clear()
            TextBox_N.Clear()
            TextBox_P.Clear()
            RadioButton_H.Checked = False
            RadioButton_F.Checked = False
            TextBox_Func.Clear()
            TextBox_Base.Clear()
            TextBox_PrimeBilan.Clear()
            TextBox_PrimeRendement.Clear()
            TextBox_PrimeAnc.Clear()
            TextBox_Primes.Clear()
            TextBox_NbHeure.Clear()
            TextBox_MtHeure.Clear()
            CheckBox_Chef.Checked = False
            NumericUpDown_Enfant.Value = 0
            TextBox_Brut.Clear()
            TextBox_CNSS.Clear()
            TextBox_Imposable.Clear()
            TextBox_IRPP.Clear()
            TextBox_Net.Clear()
            Button_Modifier.Enabled = False
            SalaireNetToolStripMenuItem.Enabled = False
            CréerToolStripMenuItem.Enabled = False
            GroupBox6.Visible = False
            GroupBox1.Visible = False
            GroupBox2.Visible = False
            GroupBox3.Visible = False
            Me.Width = 910
            Me.Height = 335
            ListView.Visible = True
            Button_Modifier.Visible = True
            Button_Supprimer.Visible = True
            ImporterToolStripMenuItem.Enabled = True
            ExporterToolStripMenuItem.Enabled = True
            ImprimerToolStripMenuItem.Enabled = True
            MessageBox.Show("Fiche créée avec succès.", "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ImporterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImporterToolStripMenuItem.Click
        Dim sfile As New OpenFileDialog
        With sfile
            .Title = "Veuillez indiquer le chemin d'ouverture"
            .InitialDirectory = "C:\"
            .Filter = (".txt | *.txt")
        End With

        If sfile.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim MyStream As New StreamReader(sfile.FileName)
            Dim strTemp() As String

            Using MyStream

                Do While MyStream.Peek <> -1

                    strTemp = MyStream.ReadLine.Split(",")

                    Dim LVItem As New ListViewItem
                    ListView.Items.Add(LVItem)

                    LVItem.Text = ""
                    LVItem.SubItems.Add((strTemp(0).ToString).Substring(11, (strTemp(0).ToString).Length - 11))
                    LVItem.SubItems.Add((strTemp(1).ToString).Substring(6, (strTemp(1).ToString).Length - 6))
                    LVItem.SubItems.Add((strTemp(2).ToString).Substring(9, (strTemp(2).ToString).Length - 9))
                    LVItem.SubItems.Add((strTemp(3).ToString).Substring(20, (strTemp(3).ToString).Length - 20))
                    LVItem.SubItems.Add((strTemp(4).ToString).Substring(7, (strTemp(4).ToString).Length - 7))
                    LVItem.SubItems.Add((strTemp(5).ToString).Substring(11, (strTemp(5).ToString).Length - 11))
                    LVItem.SubItems.Add((strTemp(6).ToString).Substring(18, (strTemp(6).ToString).Length - 18))
                    LVItem.SubItems.Add((strTemp(7).ToString).Substring(9, (strTemp(7).ToString).Length - 9))
                    LVItem.SubItems.Add((strTemp(8).ToString).Substring(15, (strTemp(8).ToString).Length - 15))
                    LVItem.SubItems.Add((strTemp(9).ToString).Substring(7, (strTemp(9).ToString).Length - 7))
                    LVItem.SubItems.Add((strTemp(10).ToString).Substring(7, (strTemp(10).ToString).Length - 7))
                    LVItem.SubItems.Add((strTemp(11).ToString).Substring(14, (strTemp(11).ToString).Length - 14))


                Loop

            End Using
            MessageBox.Show("Fiche importée avec succès.", "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ExporterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExporterToolStripMenuItem.Click
        Dim sfile As New SaveFileDialog
        With sfile
            .Title = "Veuillez indiquer le chemin d'enregistrement"
            .InitialDirectory = "C:\"
            .Filter = (".txt | *.txt")
        End With

        If sfile.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Dim Write As New IO.StreamWriter(sfile.FileName)
            Dim k As Windows.Forms.ListView.ColumnHeaderCollection = ListView.Columns
            For Each x As ListViewItem In ListView.Items
                Dim StrLn As String = ""
                For i = 1 To x.SubItems.Count - 1
                    StrLn += k(i).Text + ": " + x.SubItems(i).Text + ", "
                Next
                StrLn = StrLn.Substring(0, StrLn.Length - 2)
                Write.WriteLine(StrLn)
            Next
            Write.Close()
            MessageBox.Show("Fiche exportée avec succès.", "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ConsulterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsulterToolStripMenuItem.Click
        GroupBox6.Visible = False
        GroupBox1.Visible = False
        GroupBox2.Visible = False
        GroupBox3.Visible = False
        Me.Width = 910
        Me.Height = 335
        ListView.Visible = True
        Button_Modifier.Visible = True
        Button_Supprimer.Visible = True
        ImporterToolStripMenuItem.Enabled = True
        ExporterToolStripMenuItem.Enabled = True
        ImprimerToolStripMenuItem.Enabled = True
    End Sub

    Private Sub RéinitialiserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RéinitialiserToolStripMenuItem.Click
        TextBox_M.Clear()
        TextBox_N.Clear()
        TextBox_P.Clear()
        RadioButton_H.Checked = False
        RadioButton_F.Checked = False
        TextBox_Func.Clear()
        TextBox_Base.Clear()
        TextBox_PrimeBilan.Clear()
        TextBox_PrimeRendement.Clear()
        TextBox_PrimeAnc.Clear()
        TextBox_Primes.Clear()
        TextBox_NbHeure.Clear()
        TextBox_MtHeure.Clear()
        CheckBox_Chef.Checked = False
        NumericUpDown_Enfant.Value = 0
        TextBox_Brut.Clear()
        TextBox_CNSS.Clear()
        TextBox_Imposable.Clear()
        TextBox_IRPP.Clear()
        TextBox_Net.Clear()
        CréerToolStripMenuItem.Enabled = False
        Button_Modifier.Enabled = False
        Button_Supprimer.Enabled = False
        For Each i As ListViewItem In ListView.Items
            ListView.Items.Remove(i)
        Next
    End Sub

    Private Sub Button_Modifier_Click(sender As Object, e As EventArgs) Handles Button_Modifier.Click
        If String.IsNullOrEmpty(TextBox_M.Text) Or String.IsNullOrEmpty(TextBox_N.Text) Or String.IsNullOrEmpty(TextBox_P.Text) Or (RadioButton_H.Checked = False And RadioButton_F.Checked = False) Or String.IsNullOrEmpty(TextBox_Func.Text) Or String.IsNullOrEmpty(TextBox_Base.Text) Or String.IsNullOrEmpty(TextBox_Primes.Text) Or String.IsNullOrEmpty(TextBox_Brut.Text) Or String.IsNullOrEmpty(TextBox_CNSS.Text) Or String.IsNullOrEmpty(TextBox_IRPP.Text) Or String.IsNullOrEmpty(TextBox_Net.Text) Then
            MessageBox.Show("Un ou plusieurs champs sont vides.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            For Each item As ListViewItem In ListView.SelectedItems

                item.Text = ""
                item.SubItems(1).Text = TextBox_M.Text
                item.SubItems(2).Text = TextBox_N.Text
                item.SubItems(3).Text = TextBox_P.Text
                item.SubItems(4).Text = DateTimePicker.Text
                If (RadioButton_H.Checked = True) Then
                    item.SubItems(5).Text = "Homme"
                Else
                    item.SubItems(5).Text = "Femme"
                End If
                item.SubItems(6).Text = TextBox_Func.Text
                item.SubItems(7).Text = TextBox_Base.Text
                item.SubItems(8).Text = TextBox_Primes.Text
                item.SubItems(9).Text = TextBox_Brut.Text
                item.SubItems(10).Text = TextBox_CNSS.Text
                item.SubItems(11).Text = TextBox_IRPP.Text
                item.SubItems(12).Text = TextBox_Net.Text

            Next

            TextBox_M.Clear()
            TextBox_N.Clear()
            TextBox_P.Clear()
            RadioButton_H.Checked = False
            RadioButton_F.Checked = False
            TextBox_Func.Clear()
            TextBox_Base.Clear()
            TextBox_PrimeBilan.Clear()
            TextBox_PrimeRendement.Clear()
            TextBox_PrimeAnc.Clear()
            TextBox_Primes.Clear()
            TextBox_NbHeure.Clear()
            TextBox_MtHeure.Clear()
            CheckBox_Chef.Checked = False
            NumericUpDown_Enfant.Value = 0
            TextBox_Brut.Clear()
            TextBox_CNSS.Clear()
            TextBox_Imposable.Clear()
            TextBox_IRPP.Clear()
            TextBox_Net.Clear()
            CréerToolStripMenuItem.Enabled = False
            Button_Modifier.Enabled = False
            SalaireNetToolStripMenuItem.Enabled = False
            MessageBox.Show("Fiche modifiée avec succès.", "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ListView1_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles ListView.ItemSelectionChanged
        Button_Modifier.Enabled = True
        Button_Supprimer.Enabled = True
        For I As Integer = 0 To ListView.Items.Count - 1

            For Each item As ListViewItem In ListView.SelectedItems

                TextBox_M.Text = item.SubItems(1).Text
                TextBox_N.Text = item.SubItems(2).Text
                TextBox_P.Text = item.SubItems(3).Text
                DateTimePicker.Text = item.SubItems(4).Text
                If (item.SubItems(5).Text = "Homme") Then
                    RadioButton_H.Checked = True
                Else
                    RadioButton_F.Checked = True
                End If
                TextBox_Func.Text = item.SubItems(6).Text
            Next

        Next
    End Sub

    Private Sub Button_Reset_Click(sender As Object, e As EventArgs) Handles Button_Supprimer.Click
        If MessageBox.Show("Êtes-vous sûr de vouloir supprimer cette fiche?", "Avertissement", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
        Else
            For Each i As ListViewItem In ListView.SelectedItems
                ListView.Items.Remove(i)
            Next
            MessageBox.Show("Fiche supprimée avec succès.", "Avertissement", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub ÀProposToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ÀProposToolStripMenuItem.Click
        MessageBox.Show("Réalisé par Mnif Ahmed")
    End Sub
End Class
