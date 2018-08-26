Public Class Form1

    Dim RTVP As Double
    Dim RVIPI As Double
    Dim RVBICMS As Double
    Dim RBCST As Double
    Dim RCVICMSST As Double
    Dim RTDNCST As Double
    Dim IB As Double
    Dim RTVPS As Double
    Dim RVIPIS As Double
    Dim RVBICMSS As Double
    Dim RBCSTS As Double
    Dim RCVICMSSTS As Double
    Dim RTDNCSTS As Double
    Dim IBS As Double


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        End

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        txtQtde.Clear()
        txtVU.Clear()
        txtRB.Clear()
        txtAI.Clear()
        txtIB.Clear()
        txtAINF.Clear()
        txtAIST.Clear()
        txtIB.Clear()
        txtIST.Clear()
        txtQtde.Focus()
        labelBCST.Text = ""
        labelVBICMS.Text = ""
        labelVICMS.Text = ""
        labelVICMSST.Text = ""
        labelVIPI.Text = ""
        labelVTNFCST.Text = ""
        labelVTP.Text = ""
        labelBCSTS.Text = ""
        labelVBICMSS.Text = ""
        labelVICMSS.Text = ""
        labelVICMSSTS.Text = ""
        labelVIPIS.Text = ""
        labelVTNFCSTS.Text = ""
        labelVTPS.Text = ""

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Q, VU, RB, AI, AINF, AIST, IST As Double

        Q = Val(txtQtde.Text)
        VU = Val(txtVU.Text)
        RB = Val(txtRB.Text)
        AI = Val(txtAI.Text)
        IB = Val(txtIB.Text)
        AINF = Val(txtAINF.Text)
        AIST = Val(txtAIST.Text)
        IST = Val(txtIST.Text)


        'Estrutura de validação de campo

        If txtQtde.Text = "" Or txtQtde.Text = "0" Then
            MsgBox("Preencha o campo quantidade! Ex: 1", MsgBoxStyle.Critical)
            txtQtde.Focus()

        ElseIf txtVU.Text = "" Or txtVU.Text = "0" Then
            MsgBox("Preencha o campo valor unitário! Ex: R$ 1.25", MsgBoxStyle.Critical)
            txtVU.Focus()

        ElseIf txtAINF.Text = "" Or txtAIST.Text = "0" Then
            MsgBox("Preencha o ICMS de origem! Ex: 18%", MsgBoxStyle.Critical)
            txtAINF.Focus()

        ElseIf txtIST.Text = "" And txtAIST.Text <> "" Then  'condição verdadeira
            MsgBox("Preencha o campo Índice ST. Ex: 1.63 para 63%", MsgBoxStyle.Information)
            txtIST.Focus()

        ElseIf txtAIST.Text = "" And txtIST.Text <> "" Then  ' condição verdadeira
            MsgBox("Preencha o campo ICMS de destino. Ex: 17%", MsgBoxStyle.Information)
            txtAIST.Focus()

        ElseIf txtIB.Text = "" And txtAI.Text <> "" Then  'condição verdadeira
            MsgBox("Preencha o campo IPI BASE Ex: 90 para 90%", MsgBoxStyle.Information)
            txtIB.Focus()

        ElseIf txtAI.Text = "" And txtIB.Text <> "" Then  ' condição verdadeira
            MsgBox("Preencha o campo ALIC IPI Ex: 15 para 15%", MsgBoxStyle.Information)
            txtAI.Focus()
        Else


            ' calculo  valor total de produtos
            RTVP = Q * VU
            labelVTP.Text = "V. Total Produto " & FormatCurrency(RTVP)

            ' valor ipi
            RVIPI = (RTVP / 100) * AI
            labelVIPI.Text = "V. Total IPI " & FormatCurrency(RVIPI)

            'calculo valor base icms
            If (txtIB.Text <> "") Then
                IB = (RTVP + RVIPI) / (100) * (100 - RB)
                labelVBICMS.Text = "Valor Base ICMS " & FormatCurrency(IB)

            Else
                IB = (RTVP / (100) * (100 - RB))
                labelVBICMS.Text = "Valor Base ICMS " & FormatCurrency(IB)
            End If

            ' calculo valor icms
            RVBICMS = (IB / 100) * AINF
            labelVICMS.Text = "Valor ICMS " & FormatCurrency(RVBICMS)

            'calculo base st
            RBCST = (RTVP + RVIPI) * IST
            labelBCST.Text = "Base Cálculo ST " & FormatCurrency(RBCST)

            'calculo valor icms st
            If txtIST.Text = "" Then
                labelVICMSST.Text = "Valor ICMS ST R$ 0 "
            Else
                RCVICMSST = ((RBCST / 100) * AIST) - RVBICMS
                labelVICMSST.Text = "Valor ICMS ST " & FormatCurrency(RCVICMSST)

            End If

            ' Total da Nota com st
            RTDNCST = RTVP + RCVICMSST
            labelVTNFCST.Text = "V. Total NF C/ ICMS ST " & FormatCurrency(RTDNCST)

        End If
    End Sub

    Private Sub SobreToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SobreToolStripMenuItem.Click
        Dim frm2 As New Sobre()
        frm2.Show()
    End Sub

    Public Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        RTVPS = RTVPS + RTVP
        labelVTPS.Text = "V. Total Produto " & FormatCurrency(RTVPS)

        IBS = IBS + IB
        labelVBICMSS.Text = "Valor Base ICMS " & FormatCurrency(IBS)

        RVIPIS = RVIPIS + RVIPI
        labelVIPIS.Text = "V. Total IPI " & FormatCurrency(RVIPIS)

        RVBICMSS = RVBICMSS + RVBICMS
        labelVICMSS.Text = "Valor ICMS " & FormatCurrency(RVBICMSS)

        RBCSTS = RBCSTS + RBCST
        labelBCSTS.Text = "Base Cálculo ST " & FormatCurrency(RBCSTS)

        If labelVICMSST.Text = "0" Then
            labelVICMSSTS.Text = "Valor ICMS ST R$ 0 "
        Else
            RCVICMSSTS = RCVICMSSTS + RCVICMSST
            labelVICMSSTS.Text = "Valor ICMS ST " & FormatCurrency(RCVICMSSTS)
        End If

        RTDNCSTS = RTDNCSTS + RTDNCST
        labelVTNFCSTS.Text = "V. Total NF C/ ICMS ST " & FormatCurrency(RTDNCSTS)

        txtQtde.Clear()
        txtVU.Clear()
        txtRB.Clear()
        txtAI.Clear()
        txtIB.Clear()
        txtAINF.Clear()
        txtAIST.Clear()
        txtIB.Clear()
        txtIST.Clear()
        txtQtde.Focus()
        labelBCST.Text = ""
        labelVBICMS.Text = ""
        labelVICMS.Text = ""
        labelVICMSST.Text = ""
        labelVIPI.Text = ""
        labelVTNFCST.Text = ""
        labelVTP.Text = ""

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
End Class
