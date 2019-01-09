Public Class Form1
    Dim daerah(10, 10) As Integer
    Dim rute(100, 100) As Integer
    Dim ambiljalur() As String
    Dim jalur(100) As String
    Dim jarak As Double
    Dim dari, ke, s, jmljalur As Integer
    Dim tanda As Boolean
    Dim topjalur As Integer

    Function Number(ByVal kode As String) As Integer
        Select Case kode
            Case "A"
                Number = 1
            Case "B"
                Number = 2
            Case "C"
                Number = 3
            Case "D"
                Number = 4
            Case "E"
                Number = 5
            Case "F"
                Number = 6
            Case "G"
                Number = 7
            Case "H"
                Number = 8
        End Select
    End Function

    Function Code(ByVal num As Integer) As Char
        Select Case num
            Case 1
                Code = "A"
            Case 2
                Code = "B"
            Case 3
                Code = "C"
            Case 4
                Code = "D"
            Case 5
                Code = "E"
            Case 6
                Code = "F"
            Case 7
                Code = "G"
            Case 8
                Code = "H"
        End Select
    End Function

    Sub bestrute()
        For i As Integer = 1 To 8
            s = 0
            For j As Integer = 1 To 8
                If daerah(i, j) = 1 Then
                    s += 1
                    rute(i, s) = j
                End If
            Next
        Next
    End Sub

    Sub refresh()
        cicaheum.BackColor = Color.LightSalmon
        samsat.BackColor = Color.LightSalmon
        lw_panjang.BackColor = Color.LightSalmon
        bubat.BackColor = Color.LightSalmon
        cikutra.BackColor = Color.LightSalmon
        cicadas.BackColor = Color.LightSalmon
        kosambi.BackColor = Color.LightSalmon
        alun.BackColor = Color.LightSalmon
        lbla.BackColor = Color.Azure
        lblb.BackColor = Color.Azure
        lblc.BackColor = Color.Azure
        lbld.BackColor = Color.Azure
        lblf.BackColor = Color.Azure
        lblg.BackColor = Color.Azure
        lblh.BackColor = Color.Azure
        ab.BorderWidth = 2
        ab.BorderColor = Color.Black
        ac.BorderWidth = 2
        ac.BorderColor = Color.Black
        ad.BorderWidth = 2
        ad.BorderColor = Color.Black
        bc.BorderWidth = 2
        bc.BorderColor = Color.Black
        be.BorderWidth = 2
        be.BorderColor = Color.Black
        cd.BorderWidth = 2
        cd.BorderColor = Color.Black
        ce.BorderWidth = 2
        ce.BorderColor = Color.Black
        cf.BorderWidth = 2
        cf.BorderColor = Color.Black
        df.BorderWidth = 2
        df.BorderColor = Color.Black
        eg.BorderWidth = 2
        eg.BorderColor = Color.Black
        fh.BorderWidth = 2
        fh.BorderColor = Color.Black
        gh.BorderWidth = 2
        gh.BorderColor = Color.Black
    End Sub

    Sub gambarjalur()
        topjalur = 0
        jarak = 0
        ambiljalur = txthasil.Text.Split("-")
        jmljalur = UBound(ambiljalur)

        For i As Integer = 0 To jmljalur - 1
            jalur(i) = ambiljalur(i) & ambiljalur(i + 1)
            topjalur += 1
        Next

        For i As Integer = 0 To topjalur - 1
            If jalur(i) = "AB" Or jalur(i) = "BA" Then
                ab.BorderWidth = 7
                ab.BorderColor = Color.OrangeRed
                cicaheum.BackColor = Color.Turquoise
                samsat.BackColor = Color.Turquoise
                lbla.BackColor = Color.Azure
                lblb.BackColor = Color.Azure
                jarak += 4.7
            ElseIf jalur(i) = "AC" Or jalur(i) = "CA" Then
                ac.BorderWidth = 7
                ac.BorderColor = Color.OrangeRed
                cicaheum.BackColor = Color.Turquoise
                cicadas.BackColor = Color.Turquoise
                lbla.BackColor = Color.Azure
                lblc.BackColor = Color.Azure
                jarak += 3.0
            ElseIf jalur(i) = "AD" Or jalur(i) = "DA" Then
                ad.BorderWidth = 7
                ad.BorderColor = Color.OrangeRed
                cicaheum.BackColor = Color.Turquoise
                cikutra.BackColor = Color.Turquoise
                lbla.BackColor = Color.Azure
                lbld.BackColor = Color.Azure
                jarak += 3.2
            ElseIf jalur(i) = "BE" Or jalur(i) = "EB" Then
                be.BorderWidth = 7
                be.BorderColor = Color.OrangeRed
                samsat.BackColor = Color.Turquoise
                bubat.BackColor = Color.Turquoise
                lblb.BackColor = Color.Azure
                lble.BackColor = Color.Azure
                jarak += 3.2
            ElseIf jalur(i) = "BC" Or jalur(i) = "CB" Then
                bc.BorderWidth = 7
                bc.BorderColor = Color.OrangeRed
                samsat.BackColor = Color.Turquoise
                cicadas.BackColor = Color.Turquoise
                lblb.BackColor = Color.Azure
                lblc.BackColor = Color.Azure
                jarak += 5.3
            ElseIf jalur(i) = "CD" Or jalur(i) = "DC" Then
                cd.BorderWidth = 7
                cd.BorderColor = Color.OrangeRed
                cicadas.BackColor = Color.Turquoise
                cikutra.BackColor = Color.Turquoise
                lblc.BackColor = Color.Azure
                lbld.BackColor = Color.Azure
                jarak += 8.2
            ElseIf jalur(i) = "CE" Or jalur(i) = "EC" Then
                ce.BorderWidth = 7
                ce.BorderColor = Color.OrangeRed
                cicadas.BackColor = Color.Turquoise
                bubat.BackColor = Color.Turquoise
                lblc.BackColor = Color.Azure
                lble.BackColor = Color.Azure
                jarak += 0.7
            ElseIf jalur(i) = "CF" Or jalur(i) = "FC" Then
                cf.BorderWidth = 7
                cf.BorderColor = Color.OrangeRed
                cicadas.BackColor = Color.Turquoise
                kosambi.BackColor = Color.Turquoise
                lblb.BackColor = Color.Azure
                lble.BackColor = Color.Azure
                jarak += 3.0
            ElseIf jalur(i) = "DF" Or jalur(i) = "FD" Then
                df.BorderWidth = 7
                df.BorderColor = Color.OrangeRed
                cikutra.BackColor = Color.Turquoise
                kosambi.BackColor = Color.Turquoise
                lbld.BackColor = Color.Azure
                lblf.BackColor = Color.Azure
                jarak += 3.5
            ElseIf jalur(i) = "EG" Or jalur(i) = "GE" Then
                eg.BorderWidth = 7
                eg.BorderColor = Color.OrangeRed
                bubat.BackColor = Color.Turquoise
                lw_panjang.BackColor = Color.Turquoise
                lble.BackColor = Color.Azure
                lblg.BackColor = Color.Azure
                jarak += 8.4
            ElseIf jalur(i) = "FH" Or jalur(i) = "HF" Then
                fh.BorderWidth = 7
                fh.BorderColor = Color.OrangeRed
                kosambi.BackColor = Color.Turquoise
                alun.BackColor = Color.Turquoise
                lblf.BackColor = Color.Azure
                lblh.BackColor = Color.Azure
                jarak += 1.8
            ElseIf jalur(i) = "GH" Or jalur(i) = "HG" Then
                gh.BorderWidth = 7
                gh.BorderColor = Color.OrangeRed
                lw_panjang.BackColor = Color.Turquoise
                alun.BackColor = Color.Turquoise
                lblg.BackColor = Color.Azure
                lblh.BackColor = Color.Azure
                jarak += 2.8
            End If
        Next
        txtjarak.Text = jarak & " Km"
    End Sub

    Private Sub cmdtampilkan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdtampilkan.Click
        dari = Number(txtdari.Text)
        ke = Number(txtke.Text)
        If dari = ke Then
            MessageBox.Show("Tempat Tidak Boleh Sama...", "Instruksi!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        bestrute()
        For i As Integer = 1 To 8
            If rute(dari, i) = ke Then
                txthasil.Text = Code(dari) & "-" & Code(rute(dari, i))
                refresh()
                gambarjalur()
                Exit Sub
            End If
        Next
        For j As Integer = 1 To 8
            tanda = False
            If rute(dari, j) <> 0 Then
                For k As Integer = 1 To 8
                    If rute(rute(dari, j), k) = ke Then
                        tanda = True
                        txthasil.Text = Code(dari) & "-" & Code(rute(dari, j)) & "-" & Code(rute(rute(dari, j), k))
                        refresh()
                        gambarjalur()
                        Exit Sub
                    End If
                Next
            End If
        Next
        For j As Integer = 1 To 8
            For l As Integer = 1 To 8
                If rute(rute(dari, j), l) <> 0 Then
                    For m As Integer = 1 To 8
                        If rute(rute(rute(dari, j), l), m) = ke Then
                            txthasil.Text = Code(dari) & "-" & Code(rute(dari, j)) & "-" & Code(rute(rute(dari, j), l)) & "-" & Code(rute(rute(rute(dari, j), l), m))
                            refresh()
                            gambarjalur()
                            Exit Sub

                        End If
                    Next

                End If
            Next
        Next

    End Sub

    Sub data()
        daerah(1, 1) = 0
        daerah(1, 2) = 1
        daerah(1, 3) = 1
        daerah(1, 4) = 1
        daerah(1, 5) = 0
        daerah(1, 6) = 0
        daerah(1, 7) = 0
        daerah(1, 8) = 0
        daerah(2, 1) = 1
        daerah(2, 2) = 0
        daerah(2, 3) = 1
        daerah(2, 4) = 0
        daerah(2, 5) = 1
        daerah(2, 6) = 0
        daerah(2, 7) = 0
        daerah(2, 8) = 0
        daerah(3, 1) = 1
        daerah(3, 2) = 1
        daerah(3, 3) = 0
        daerah(3, 4) = 1
        daerah(3, 5) = 1
        daerah(3, 6) = 1
        daerah(3, 7) = 0
        daerah(3, 8) = 0
        daerah(4, 1) = 1
        daerah(4, 2) = 0
        daerah(4, 3) = 1
        daerah(4, 4) = 0
        daerah(4, 5) = 0
        daerah(4, 6) = 1
        daerah(4, 7) = 0
        daerah(4, 8) = 0
        daerah(5, 1) = 0
        daerah(5, 2) = 1
        daerah(5, 3) = 1
        daerah(5, 4) = 0
        daerah(5, 5) = 0
        daerah(5, 6) = 0
        daerah(5, 7) = 1
        daerah(5, 8) = 0
        daerah(6, 1) = 0
        daerah(6, 2) = 0
        daerah(6, 3) = 1
        daerah(6, 4) = 1
        daerah(6, 5) = 0
        daerah(6, 6) = 0
        daerah(6, 7) = 0
        daerah(6, 8) = 1
        daerah(7, 1) = 0
        daerah(7, 2) = 0
        daerah(7, 3) = 0
        daerah(7, 4) = 0
        daerah(7, 5) = 1
        daerah(7, 6) = 0
        daerah(7, 7) = 0
        daerah(7, 8) = 1
        daerah(8, 1) = 0
        daerah(8, 2) = 0
        daerah(8, 3) = 0
        daerah(8, 4) = 0
        daerah(8, 5) = 0
        daerah(8, 6) = 1
        daerah(8, 7) = 1
        daerah(8, 8) = 0
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        data()

    End Sub



    Private Sub lbla_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbla.Click

    End Sub

    Private Sub cf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cf.Click, df.Click, ce.Click

    End Sub

    Private Sub lbld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lble.Click

    End Sub

    Private Sub nodef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cicadas.Click, kosambi.Click, alun.Click

    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub lblb_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblb.Click

    End Sub

    Private Sub ab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ac.Click, bc.Click, ab.Click

    End Sub

    Private Sub nodea_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cicaheum.Click

    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

    End Sub

    Private Sub lblc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblg.Click

    End Sub

    Private Sub lble_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbld.Click

    End Sub

    Private Sub txtdari_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtdari.TextChanged

    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click

    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub

    Private Sub txtke_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtke.TextChanged

    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click

    End Sub

    Private Sub txthasil_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txthasil.TextChanged

    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click

    End Sub

    Private Sub txtjarak_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtjarak.TextChanged

    End Sub

    Private Sub GroupBox2_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub lblf_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblc.Click

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click

    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click

    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click

    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click

    End Sub

    Private Sub Label25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label25.Click

    End Sub
End Class
