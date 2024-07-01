Imports SG_ECC_MASTER_DLL.SGECCMASTER
Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim res As New List(Of SG_ECC_MASTER_DLL.SGECCMASTER.ECC_MASTER)
        ' res = GETECCMASTER(TextBox1.Text, TextBox2.Text, 0, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, True)
        ' res = GETECCMASTER(615930, 0, 0, 14270.0, 60, "2022-06-30", "2022-07-31", True, True, True)
        res = GETECCMASTERBUTTOM(25000.0, 29512.0, 10000.0, 813.0, 24, "2022-07-01", "2022-07-01", True, True)
        'res = GETECCMASTERBUTTOM(913790, 0, 0, 21170.0, 60, "2023-06-29", "2023-07-31", True, True)


        If res IsNot Nothing AndAlso res.Count > 0 Then
            For Each r As ECC_MASTER In res
                Dim str As String = String.Format("{0},{1},{2},{3},{4},{5},{6},{7} ", r.MONTHDATE.ToString("yyyy-MM-dd"), r.NumberOfDate, r.CUR_PUR_INTEREST, r.CUR_PUR_PRINCIPLE, r.Instamemt_AMT, r.PUR_PRINCIPLE_BAL, r.PUR_UCC_BAL, r.DiscountAMT)
                RichTextBox1.AppendText(vbCrLf & str)
            Next
        End If


        'res = GETECCMASTERTOP(208535, 289863.65, 0, 4832, 60, "2022-12-09", "2023-01-31", True, True, False)

        'If res IsNot Nothing AndAlso res.Count > 0 Then
        '    For Each r As ECC_MASTER In res
        '        Dim str As String = String.Format("{0},{1},{2},{3},{4},{5},{6},{7} ", r.MONTHDATE.ToString("yyyy-MM-dd"), r.NumberOfDate, r.CUR_PUR_INTEREST, r.CUR_PUR_PRINCIPLE, r.Instamemt_AMT, r.PUR_PRINCIPLE_BAL, r.PUR_UCC_BAL, r.DiscountAMT)
        '        RichTextBox1.AppendText(vbCrLf & str)
        '    Next
        'End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim res As New List(Of SG_ECC_MASTER_DLL.SGECCMASTER.ECC_MASTER)
        'res = GETECCMASTERBUTTOM(TextBox1.Text, TextBox2.Text, 0, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox6.Text, True, False)
        'res = GETECCMASTERTOP(10000.25, 11824.0, 4000.0, 326.0, 24, "2024-05-01", "2024-07-01", False, False)

        'res = GETECCMASTERTOP(10000.25, 11824.0, 4000.0, 326.0, 24, CDate("2024-05-31"), CDate("2024-07-01"), False, False, False)

        res = GETECCMASTERBUTTOM(8999.0, 11160.0, 1800.0, 390.0, 24, "2024-07-31", "2024-07-31", False, False)
        'res = GETECCMASTERBUTTOM(10990.0, 12328.0, 2200.0, 422.0, 24, "2022-06-05", "2022-07-31", False, False)
        If res IsNot Nothing AndAlso res.Count > 0 Then
            For Each r As ECC_MASTER In res
                Dim str As String = String.Format("{0},{1},{2},{3},{4} ", r.MONTHDATE.ToString("yyyy-MM-dd"), r.CUR_PUR_INTEREST, r.CUR_PUR_PRINCIPLE, r.Instamemt_AMT, r.PUR_PRINCIPLE_BAL)
                RichTextBox1.AppendText(vbCrLf & str)
            Next
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim res As New List(Of SG_ECC_MASTER_DLL.SGECCMASTER.ECC_MASTER)
        res = GETECCMASTERBUTTOM(8840.0, 0, 0, 797.0, 12, "2023-09-19", "2023-10-19", True, True)
        If res IsNot Nothing AndAlso res.Count > 0 Then
            For Each r As ECC_MASTER In res
                Dim str As String = String.Format("{0},{1},{2},{3},{4},{5} ", r.MONTHDATE.ToString("yyyy-MM-dd"), r.NumberOfDate, r.CUR_PUR_INTEREST, r.CUR_PUR_PRINCIPLE, r.Instamemt_AMT, r.PUR_PRINCIPLE_BAL)
                RichTextBox1.AppendText(vbCrLf & str)
            Next
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim res As New List(Of SG_ECC_MASTER_DLL.SGECCMASTER.ECC_MASTER)
        res = GETECCMASTERCURRENT(165000.0, 170000.0, 236400.0, 0, 3940.0, 60, "2020-02-29", "2020-07-31", "2020-07-31", True)
        If res IsNot Nothing AndAlso res.Count > 0 Then
            For Each r As ECC_MASTER In res
                Dim str As String = String.Format("{0},{1},{2},{3},{4} ", r.MONTHDATE.ToString("yyyy-MM-dd"), r.CUR_PUR_INTEREST, r.CUR_PUR_PRINCIPLE, r.Instamemt_AMT, r.PUR_PRINCIPLE_BAL)
                RichTextBox1.AppendText(vbCrLf & str)
            Next
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim res As New List(Of SG_ECC_MASTER_DLL.SGECCMASTER.ECC_MASTER)
        Dim ins As New List(Of Decimal)
        ins.Add(1000)
        ins.Add(1000)
        ins.Add(1000)
        ins.Add(1000)
        ins.Add(1000)
        ins.Add(1000)
        ins.Add(1000)
        ins.Add(1000)
        ins.Add(1000)
        ins.Add(1000)

        ins.Add(2000)
        ins.Add(2000)
        ins.Add(2000)
        ins.Add(2000)
        ins.Add(2000)
        ins.Add(2000)
        ins.Add(2000)
        ins.Add(2000)
        ins.Add(2000)
        ins.Add(2000)

        res = GETECCMASTERTOP(10000.25, 11824.0, 4000.0, 326.0, 24, CDate("2024-05-31"), CDate("2024-07-01"), False, False, False)
        If res IsNot Nothing AndAlso res.Count > 0 Then
            For Each r As ECC_MASTER In res
                Dim str As String = String.Format("{0},{1},{2},{3},{4},{5} ", r.MONTHDATE.ToString("yyyy-MM-dd"), r.CUR_PUR_INTEREST, r.CUR_PUR_PRINCIPLE, r.CUR_VAT, r.Instamemt_AMT, r.PUR_PRINCIPLE_BAL)
                RichTextBox1.AppendText(vbCrLf & str)
            Next
        End If
    End Sub
End Class
