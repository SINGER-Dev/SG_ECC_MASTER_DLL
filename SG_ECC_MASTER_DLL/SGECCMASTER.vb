Public Class SGECCMASTER


    Public Structure ECC_MASTER
        Dim Period As Integer
        Dim MONTHDATE As Date
        Dim PUR_UCC_BAL As Decimal
        Dim PUR_PRINCIPLE_BAL As Decimal

        Dim CUR_PUR_INTEREST As Decimal
        Dim CUR_PUR_PRINCIPLE As Decimal
        Dim CUR_VAT As Decimal
        Dim CUR_TOTAL As Decimal
        Dim VAT_BAL As Decimal
        Dim AR_BAL As Decimal
        Dim AccureBegin As Decimal
        Dim CurrAccure As Decimal
        Dim AccureEnd As Decimal
        Dim Instamemt_AMT As Decimal
        Dim InterestRate As Decimal
        Dim NumberOfDate As Integer
        Dim CalFromDate As Date
        Dim CalToDate As Date
        Dim DiscountAMT As Decimal

    End Structure

    'Public Shared Function GETECCMASTERBUTTOM(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As List(Of Decimal), ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True) As List(Of ECC_MASTER)
    '    Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
    '    Try
    '        Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
    '        Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
    '        Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
    '        Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
    '        Dim M_DATE As Date = FirstDueDate
    '        Dim lastdate As Date = SaleDate
    '        Dim Ecc As Decimal = 0
    '        Dim ECC_ACCU As Decimal = 0
    '        Dim ACCU_ECC_EXC_VAT As Decimal = 0
    '        Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
    '        Dim PincipleBal As Decimal = Math.Round(Principle, 2)
    '        Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
    '        Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

    '        Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
    '        Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
    '        If ISLOAN = True Then VAT_BAL = 0

    '        Dim NumOfDate As Integer = 30
    '        If ISLOAN = True Then
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '            CONTRACT_UCC = 0
    '            CONTRACT_UCC_EXC_VAT = 0
    '            UCC_EXC_VAT = 0
    '            UCCBAL = 0
    '            PUR_Principle_Bal = PincipleBal
    '        End If

    '        Dim ACCURE_ECC As Decimal = 0
    '        For m As Integer = 1 To ARM_CONTRACT_TERM

    '            If ISLOAN = True Then
    '                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '                Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * NumOfDate, 2)
    '            Else
    '                Ecc = Math.Round(PincipleBal * RATE, 2)
    '            End If
    '            Ecc += ACCURE_ECC

    '            If Ecc > ARM_INSTALLMENT_AMT(m - 1) Then
    '                ACCURE_ECC = Ecc - ARM_INSTALLMENT_AMT(m - 1)
    '                Ecc = ARM_INSTALLMENT_AMT(m - 1)
    '            Else
    '                ACCURE_ECC = 0
    '            End If

    '            Dim install As Decimal = 0
    '            If m >= ARM_CONTRACT_TERM Then
    '                install = PincipleBal + Ecc
    '                If ISLOAN Then
    '                    ARM_INSTALLMENT_AMT(m - 1) = install
    '                End If

    '            End If

    '            PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT(m - 1) - Ecc), 2)
    '            ECC_ACCU += Ecc
    '            If Math.Floor(PincipleBal) <= 0 Then
    '                PincipleBal = 0
    '            End If
    '            If Math.Floor(Ecc) <= 0 Then
    '                Ecc = 0
    '            End If




    '            Dim ECC_EXC_VAT As Decimal = 0
    '            Dim PUR_PRINCIPLE As Decimal = 0
    '            If ISLOAN = True Then
    '                ECC_EXC_VAT = Ecc
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = PincipleBal
    '            Else
    '                ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
    '            End If

    '            If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                PUR_PRINCIPLE = 0
    '                ECC_EXC_VAT -= PUR_PRINCIPLE
    '            End If

    '            If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                ECC_EXC_VAT += UCC_EXC_VAT
    '                UCC_EXC_VAT = 0
    '            End If



    '            Dim ECC_DATA As New ECC_MASTER
    '            With ECC_DATA
    '                .Period = m
    '                .MONTHDATE = M_DATE
    '                .Instamemt_AMT = ARM_INSTALLMENT_AMT(m - 1)
    '                .InterestRate = RATE

    '                .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
    '                If ISLOAN = False Then
    '                    .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT(m - 1) * 7.0 / 107, 2)
    '                Else
    '                    .CUR_VAT = 0
    '                End If

    '                .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT(m - 1) - (.CUR_PUR_PRINCIPLE + .CUR_VAT)

    '                '.CUR_PUR_INTEREST = ECC_EXC_VAT
    '                ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

    '                If ISLOAN = True Then
    '                    .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
    '                Else
    '                    .CUR_TOTAL = 0
    '                End If

    '                .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
    '                PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
    '                .PUR_UCC_BAL = UCC_EXC_VAT


    '                AR_BAL = AR_BAL - .Instamemt_AMT
    '                VAT_BAL = VAT_BAL - .CUR_VAT

    '                .AR_BAL = AR_BAL
    '                .VAT_BAL = VAT_BAL
    '                .NumberOfDate = NumOfDate
    '            End With

    '            LIST_ECC_MASTER.Add(ECC_DATA)
    '            lastdate = M_DATE
    '            If EndofMonth = True Then
    '                M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
    '            Else
    '                M_DATE = M_DATE.AddMonths(1)
    '            End If
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





    '            ECC_DATA = Nothing
    '        Next

    '    Catch ex As Exception
    '        ex = Nothing
    '    End Try
    '    Return LIST_ECC_MASTER
    'End Function

    'Public Shared Function GETECCMASTERBUTTOM(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As Decimal, ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True) As List(Of ECC_MASTER)
    '    Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
    '    Try
    '        Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
    '        Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
    '        Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
    '        Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
    '        Dim M_DATE As Date = FirstDueDate
    '        Dim lastdate As Date = SaleDate
    '        Dim Ecc As Decimal = 0
    '        Dim ECC_ACCU As Decimal = 0
    '        Dim ACCU_ECC_EXC_VAT As Decimal = 0
    '        Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
    '        Dim PincipleBal As Decimal = Math.Round(Principle, 2)
    '        Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
    '        Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

    '        Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
    '        Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
    '        If ISLOAN = True Then VAT_BAL = 0

    '        Dim NumOfDate As Integer = 30
    '        If ISLOAN = True Then
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '            CONTRACT_UCC = 0
    '            CONTRACT_UCC_EXC_VAT = 0
    '            UCC_EXC_VAT = 0
    '            UCCBAL = 0
    '            PUR_Principle_Bal = PincipleBal
    '        End If

    '        Dim ACCURE_ECC As Decimal = 0
    '        For m As Integer = 1 To ARM_CONTRACT_TERM

    '            If ISLOAN = True Then
    '                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '                Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * NumOfDate, 2)
    '            Else
    '                Ecc = Math.Round(PincipleBal * RATE, 2)
    '            End If
    '            Ecc += ACCURE_ECC

    '            If Ecc > ARM_INSTALLMENT_AMT Then
    '                ACCURE_ECC = Ecc - ARM_INSTALLMENT_AMT
    '                Ecc = ARM_INSTALLMENT_AMT
    '            Else
    '                ACCURE_ECC = 0
    '            End If

    '            Dim install As Decimal = 0
    '            If m >= ARM_CONTRACT_TERM Then
    '                install = PincipleBal + Ecc
    '                If ISLOAN Then
    '                    ARM_INSTALLMENT_AMT = install
    '                End If

    '            End If

    '            PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT - Ecc), 2)
    '            ECC_ACCU += Ecc
    '            If Math.Floor(PincipleBal) <= 0 Then
    '                PincipleBal = 0
    '            End If
    '            If Math.Floor(Ecc) <= 0 Then
    '                Ecc = 0
    '            End If




    '            Dim ECC_EXC_VAT As Decimal = 0
    '            Dim PUR_PRINCIPLE As Decimal = 0
    '            If ISLOAN = True Then
    '                ECC_EXC_VAT = Ecc
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = PincipleBal
    '            Else
    '                ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
    '            End If

    '            If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                PUR_PRINCIPLE = 0
    '                ECC_EXC_VAT -= PUR_PRINCIPLE
    '            End If

    '            If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                ECC_EXC_VAT += UCC_EXC_VAT
    '                UCC_EXC_VAT = 0
    '            End If



    '            Dim ECC_DATA As New ECC_MASTER
    '            With ECC_DATA
    '                .Period = m
    '                .MONTHDATE = M_DATE
    '                .Instamemt_AMT = ARM_INSTALLMENT_AMT
    '                .InterestRate = RATE

    '                .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
    '                If ISLOAN = False Then
    '                    .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT * 7.0 / 107, 2)
    '                Else
    '                    .CUR_VAT = 0
    '                End If

    '                .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_VAT)

    '                '.CUR_PUR_INTEREST = ECC_EXC_VAT
    '                ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

    '                If ISLOAN = True Then
    '                    .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
    '                Else
    '                    .CUR_TOTAL = 0
    '                End If

    '                .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
    '                PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
    '                .PUR_UCC_BAL = UCC_EXC_VAT


    '                AR_BAL = AR_BAL - .Instamemt_AMT
    '                VAT_BAL = VAT_BAL - .CUR_VAT

    '                .AR_BAL = AR_BAL
    '                .VAT_BAL = VAT_BAL
    '                .NumberOfDate = NumOfDate
    '            End With

    '            LIST_ECC_MASTER.Add(ECC_DATA)
    '            lastdate = M_DATE
    '            If EndofMonth = True Then
    '                M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
    '            Else
    '                M_DATE = M_DATE.AddMonths(1)

    '                Dim tmpdate As Date = M_DATE
    '                Dim SaleDay As String = SaleDate.ToString("dd")
    '                If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
    '                    M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
    '                Else
    '                    M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
    '                End If


    '            End If
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





    '            ECC_DATA = Nothing
    '        Next

    '    Catch ex As Exception
    '        ex = Nothing
    '    End Try
    '    Return LIST_ECC_MASTER
    'End Function

    'Public Shared Function GETECCMASTERCURRENT(ByVal PrincipleBal As Decimal, ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As Decimal, ByVal ARM_CONTRACT_TERM As Integer, ByVal FirstDueDate As Date, ByVal CurrentDate As Date, ByVal lastCutPrincDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True) As List(Of ECC_MASTER)
    '    Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
    '    Try
    '        Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
    '        Dim Principle As Decimal = PrincipleBal
    '        Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
    '        Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
    '        Dim M_DATE As Date = CurrentDate
    '        Dim lastdate As Date = lastCutPrincDate
    '        Dim Ecc As Decimal = 0
    '        Dim ECC_ACCU As Decimal = 0
    '        Dim ACCU_ECC_EXC_VAT As Decimal = 0
    '        Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
    '        Dim PincipleBal As Decimal = Math.Round(Principle, 2)
    '        Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
    '        Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

    '        Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
    '        Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
    '        If ISLOAN = True Then VAT_BAL = 0

    '        Dim NumOfDate As Integer = 30
    '        If ISLOAN = True Then
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '            CONTRACT_UCC = 0
    '            CONTRACT_UCC_EXC_VAT = 0
    '            UCC_EXC_VAT = 0
    '            UCCBAL = 0
    '            PUR_Principle_Bal = PincipleBal
    '        End If
    '        Dim AccureBegin As Decimal = 0
    '        Dim ACCURE_ECC As Decimal = 0
    '        Dim N As Integer = DateDiff(DateInterval.Month, CurrentDate, DateAdd(DateInterval.Month, ARM_CONTRACT_TERM, FirstDueDate))
    '        For m As Integer = 1 To N
    '            AccureBegin = ACCURE_ECC
    '            If ISLOAN = True Then
    '                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '                Ecc = Math.Round(PincipleBal * (RATE * 12 / 365) * NumOfDate, 2)
    '            Else
    '                Ecc = Math.Round(PincipleBal * RATE, 2)
    '            End If
    '            Ecc += ACCURE_ECC

    '            If Ecc > ARM_INSTALLMENT_AMT Then
    '                ACCURE_ECC = Ecc - ARM_INSTALLMENT_AMT
    '                Ecc = ARM_INSTALLMENT_AMT

    '            Else
    '                ACCURE_ECC = 0
    '            End If

    '            Dim install As Decimal = 0
    '            If m >= N Then
    '                install = PincipleBal + Ecc
    '                If ISLOAN Then
    '                    ARM_INSTALLMENT_AMT = install
    '                End If

    '            End If

    '            PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT - Ecc), 2)
    '            ECC_ACCU += Ecc
    '            If Math.Floor(PincipleBal) <= 0 Then
    '                PincipleBal = 0
    '            End If
    '            If Math.Floor(Ecc) <= 0 Then
    '                Ecc = 0
    '            End If




    '            Dim ECC_EXC_VAT As Decimal = 0
    '            Dim PUR_PRINCIPLE As Decimal = 0
    '            If ISLOAN = True Then
    '                ECC_EXC_VAT = Ecc
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = PincipleBal
    '            Else
    '                ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
    '            End If

    '            If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                PUR_PRINCIPLE = 0
    '                ECC_EXC_VAT -= PUR_PRINCIPLE
    '            End If

    '            If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                ECC_EXC_VAT += UCC_EXC_VAT
    '                UCC_EXC_VAT = 0
    '            End If



    '            Dim ECC_DATA As New ECC_MASTER
    '            With ECC_DATA
    '                .Period = m
    '                .MONTHDATE = M_DATE
    '                .Instamemt_AMT = ARM_INSTALLMENT_AMT
    '                .InterestRate = RATE
    '                .CalFromDate = lastdate
    '                .CalToDate = M_DATE
    '                .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
    '                If ISLOAN = False Then
    '                    .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT * 7.0 / 107, 2)
    '                Else
    '                    .CUR_VAT = 0
    '                End If

    '                .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_VAT)
    '                .CurrAccure = ACCURE_ECC - AccureBegin
    '                .AccureEnd = ACCURE_ECC
    '                .AccureBegin = AccureBegin
    '                '.CUR_PUR_INTEREST = ECC_EXC_VAT
    '                ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

    '                If ISLOAN = True Then
    '                    .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
    '                Else
    '                    .CUR_TOTAL = 0
    '                End If

    '                .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
    '                PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
    '                .PUR_UCC_BAL = UCC_EXC_VAT


    '                AR_BAL = AR_BAL - .Instamemt_AMT
    '                VAT_BAL = VAT_BAL - .CUR_VAT

    '                .AR_BAL = AR_BAL
    '                .VAT_BAL = VAT_BAL
    '                .NumberOfDate = NumOfDate
    '            End With

    '            LIST_ECC_MASTER.Add(ECC_DATA)
    '            lastdate = M_DATE
    '            If EndofMonth = True Then
    '                M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
    '            Else
    '                M_DATE = M_DATE.AddMonths(1)
    '                Dim tmpdate As Date = M_DATE
    '                Dim SaleDay As String = CurrentDate.ToString("dd")
    '                If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
    '                    M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
    '                Else
    '                    M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
    '                End If
    '            End If
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





    '            ECC_DATA = Nothing
    '        Next

    '    Catch ex As Exception
    '        ex = Nothing
    '    End Try
    '    Return LIST_ECC_MASTER
    'End Function


    'Public Shared Function GETECCMASTERTOP(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As List(Of Decimal), ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True) As List(Of ECC_MASTER)
    '    Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
    '    Try
    '        Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
    '        Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
    '        Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
    '        Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
    '        Dim M_DATE As Date = FirstDueDate
    '        Dim lastdate As Date = SaleDate
    '        Dim EndMonthOfSaleDate As Date = CDate(SaleDate.AddMonths(1).ToString("yyyy-MM-01")).AddDays(-1)

    '        Dim Ecc As Decimal = 0
    '        Dim ECC_ACCU As Decimal = 0
    '        Dim ACCU_ECC_EXC_VAT As Decimal = 0
    '        Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
    '        Dim PincipleBal As Decimal = Math.Round(Principle, 2)
    '        Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
    '        Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

    '        Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
    '        Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
    '        If ISLOAN = True Then VAT_BAL = 0

    '        Dim NumOfDate As Integer = 30
    '        Dim FirstNum As Integer = 0
    '        Dim FirstInterest As Decimal = 0.0

    '        If ISLOAN = True Then
    '            CONTRACT_UCC = 0
    '            CONTRACT_UCC_EXC_VAT = 0
    '            UCC_EXC_VAT = 0
    '            UCCBAL = 0
    '            PUR_Principle_Bal = PincipleBal

    '            FirstNum = DateDiff(DateInterval.Day, SaleDate, EndMonthOfSaleDate)
    '            FirstInterest = Math.Round(PincipleBal * (Math.Round(RATE * 12, 2) / 365) * FirstNum, 2)
    '        End If

    '        Dim ACCURE_ECC As Decimal = 0
    '        For m As Integer = 1 To ARM_CONTRACT_TERM

    '            If ISLOAN = True Then
    '                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '                Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 2) / 365) * NumOfDate, 2)
    '            Else
    '                Ecc = Math.Round(PincipleBal * RATE, 2)
    '            End If
    '            Ecc += ACCURE_ECC



    '            If Ecc > ARM_INSTALLMENT_AMT(m - 1) + FirstInterest Then
    '                ACCURE_ECC = Ecc - (ARM_INSTALLMENT_AMT(m - 1) + FirstInterest)
    '                Ecc = ARM_INSTALLMENT_AMT(m - 1) + FirstInterest
    '            Else
    '                ACCURE_ECC = 0
    '            End If

    '            Dim install As Decimal = 0
    '            If m >= ARM_CONTRACT_TERM Then
    '                install = PincipleBal + Ecc
    '                If ISLOAN Then
    '                    ARM_INSTALLMENT_AMT(m - 1) = install
    '                End If

    '            End If

    '            PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT(m - 1) + FirstInterest - Ecc), 2)
    '            ECC_ACCU += Ecc
    '            If Math.Floor(PincipleBal) <= 0 Then
    '                PincipleBal = 0
    '            End If
    '            If Math.Floor(Ecc) <= 0 Then
    '                Ecc = 0
    '            End If




    '            Dim ECC_EXC_VAT As Decimal = 0
    '            Dim PUR_PRINCIPLE As Decimal = 0
    '            If ISLOAN = True Then
    '                ECC_EXC_VAT = Ecc
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = PincipleBal
    '            Else
    '                ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
    '            End If

    '            If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                PUR_PRINCIPLE = 0
    '                ECC_EXC_VAT -= PUR_PRINCIPLE
    '            End If

    '            If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                ECC_EXC_VAT += UCC_EXC_VAT
    '                UCC_EXC_VAT = 0
    '            End If



    '            Dim ECC_DATA As New ECC_MASTER
    '            With ECC_DATA
    '                .Period = m
    '                .MONTHDATE = M_DATE
    '                .Instamemt_AMT = ARM_INSTALLMENT_AMT(m - 1) + FirstInterest

    '                .InterestRate = RATE

    '                .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
    '                If ISLOAN = False Then
    '                    .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT(m - 1) * 7.0 / 107, 2)
    '                Else
    '                    .CUR_VAT = 0
    '                End If

    '                .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT(m - 1) + FirstInterest - (.CUR_PUR_PRINCIPLE + .CUR_VAT)
    '                FirstInterest = 0
    '                '.CUR_PUR_INTEREST = ECC_EXC_VAT
    '                ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

    '                If ISLOAN = True Then
    '                    .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
    '                Else
    '                    .CUR_TOTAL = 0
    '                End If

    '                .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
    '                PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
    '                .PUR_UCC_BAL = UCC_EXC_VAT


    '                AR_BAL = AR_BAL - .Instamemt_AMT
    '                VAT_BAL = VAT_BAL - .CUR_VAT

    '                .AR_BAL = AR_BAL
    '                .VAT_BAL = VAT_BAL
    '                .NumberOfDate = NumOfDate
    '            End With

    '            LIST_ECC_MASTER.Add(ECC_DATA)
    '            lastdate = M_DATE
    '            If EndofMonth = True Then
    '                M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
    '            Else
    '                M_DATE = M_DATE.AddMonths(1)

    '                Dim tmpdate As Date = M_DATE
    '                Dim SaleDay As String = SaleDate.ToString("dd")
    '                If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
    '                    M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
    '                Else
    '                    M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
    '                End If
    '            End If
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





    '            ECC_DATA = Nothing
    '        Next

    '    Catch ex As Exception
    '        ex = Nothing
    '    End Try
    '    Return LIST_ECC_MASTER
    'End Function


    'Public Shared Function GETECCMASTERTOP(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As Decimal, ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True) As List(Of ECC_MASTER)
    '    Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
    '    Try
    '        Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
    '        Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
    '        Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
    '        Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
    '        Dim M_DATE As Date = FirstDueDate
    '        Dim lastdate As Date = SaleDate
    '        Dim EndMonthOfSaleDate As Date = CDate(SaleDate.AddMonths(1).ToString("yyyy-MM-01")).AddDays(-1)

    '        Dim Ecc As Decimal = 0
    '        Dim ECC_ACCU As Decimal = 0
    '        Dim ACCU_ECC_EXC_VAT As Decimal = 0
    '        Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
    '        Dim PincipleBal As Decimal = Math.Round(Principle, 2)
    '        Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
    '        Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

    '        Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
    '        Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
    '        If ISLOAN = True Then VAT_BAL = 0

    '        Dim NumOfDate As Integer = 30
    '        Dim FirstNum As Integer = 0
    '        Dim FirstInterest As Decimal = 0.0

    '        If ISLOAN = True Then
    '            CONTRACT_UCC = 0
    '            CONTRACT_UCC_EXC_VAT = 0
    '            UCC_EXC_VAT = 0
    '            UCCBAL = 0
    '            PUR_Principle_Bal = PincipleBal

    '            FirstNum = DateDiff(DateInterval.Day, SaleDate, EndMonthOfSaleDate)
    '            FirstInterest = Math.Round(PincipleBal * (Math.Round(RATE * 12, 2) / 365) * FirstNum, 2)
    '        End If

    '        Dim ACCURE_ECC As Decimal = 0
    '        For m As Integer = 1 To ARM_CONTRACT_TERM

    '            If ISLOAN = True Then
    '                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '                Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 2) / 365) * NumOfDate, 2)
    '            Else
    '                Ecc = Math.Round(PincipleBal * RATE, 2)
    '            End If
    '            Ecc += ACCURE_ECC



    '            If Ecc > ARM_INSTALLMENT_AMT + FirstInterest Then
    '                ACCURE_ECC = Ecc - (ARM_INSTALLMENT_AMT + FirstInterest)
    '                Ecc = ARM_INSTALLMENT_AMT + FirstInterest
    '            Else
    '                ACCURE_ECC = 0
    '            End If

    '            Dim install As Decimal = 0
    '            If m >= ARM_CONTRACT_TERM Then
    '                install = PincipleBal + Ecc
    '                If ISLOAN Then
    '                    ARM_INSTALLMENT_AMT = install
    '                End If

    '            End If

    '            PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT + FirstInterest - Ecc), 2)
    '            ECC_ACCU += Ecc
    '            If Math.Floor(PincipleBal) <= 0 Then
    '                PincipleBal = 0
    '            End If
    '            If Math.Floor(Ecc) <= 0 Then
    '                Ecc = 0
    '            End If




    '            Dim ECC_EXC_VAT As Decimal = 0
    '            Dim PUR_PRINCIPLE As Decimal = 0
    '            If ISLOAN = True Then
    '                ECC_EXC_VAT = Ecc
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = PincipleBal
    '            Else
    '                ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
    '            End If

    '            If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                PUR_PRINCIPLE = 0
    '                ECC_EXC_VAT -= PUR_PRINCIPLE
    '            End If

    '            If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                ECC_EXC_VAT += UCC_EXC_VAT
    '                UCC_EXC_VAT = 0
    '            End If



    '            Dim ECC_DATA As New ECC_MASTER
    '            With ECC_DATA
    '                .Period = m
    '                .MONTHDATE = M_DATE
    '                .Instamemt_AMT = ARM_INSTALLMENT_AMT + FirstInterest

    '                .InterestRate = RATE

    '                .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
    '                If ISLOAN = False Then
    '                    .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT * 7.0 / 107, 2)
    '                Else
    '                    .CUR_VAT = 0
    '                End If

    '                .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT + FirstInterest - (.CUR_PUR_PRINCIPLE + .CUR_VAT)
    '                FirstInterest = 0
    '                '.CUR_PUR_INTEREST = ECC_EXC_VAT
    '                ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

    '                If ISLOAN = True Then
    '                    .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
    '                Else
    '                    .CUR_TOTAL = 0
    '                End If

    '                .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
    '                PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
    '                .PUR_UCC_BAL = UCC_EXC_VAT


    '                AR_BAL = AR_BAL - .Instamemt_AMT
    '                VAT_BAL = VAT_BAL - .CUR_VAT

    '                .AR_BAL = AR_BAL
    '                .VAT_BAL = VAT_BAL
    '                .NumberOfDate = NumOfDate
    '            End With

    '            LIST_ECC_MASTER.Add(ECC_DATA)
    '            lastdate = M_DATE
    '            If EndofMonth = True Then
    '                M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
    '            Else
    '                M_DATE = M_DATE.AddMonths(1)

    '                Dim tmpdate As Date = M_DATE
    '                Dim SaleDay As String = SaleDate.ToString("dd")
    '                If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
    '                    M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
    '                Else
    '                    M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
    '                End If
    '            End If
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





    '            ECC_DATA = Nothing
    '        Next

    '    Catch ex As Exception
    '        ex = Nothing
    '    End Try
    '    Return LIST_ECC_MASTER
    'End Function

    'Public Shared Function GETECCMASTER(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As Decimal, ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True) As List(Of ECC_MASTER)
    '    Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
    '    Try
    '        Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
    '        Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
    '        Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
    '        Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
    '        Dim M_DATE As Date = FirstDueDate
    '        Dim lastdate As Date = SaleDate
    '        Dim Ecc As Decimal = 0
    '        Dim ECC_ACCU As Decimal = 0
    '        Dim ACCU_ECC_EXC_VAT As Decimal = 0
    '        Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
    '        Dim PincipleBal As Decimal = Math.Round(Principle, 2)
    '        Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
    '        Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

    '        Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
    '        Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
    '        If ISLOAN = True Then VAT_BAL = 0

    '        Dim NumOfDate As Integer = 30
    '        If ISLOAN = True Then
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '            CONTRACT_UCC = 0
    '            CONTRACT_UCC_EXC_VAT = 0
    '            UCC_EXC_VAT = 0
    '            UCCBAL = 0
    '            PUR_Principle_Bal = PincipleBal
    '        End If

    '        Dim ACCURE_ECC As Decimal = 0
    '        For m As Integer = 1 To ARM_CONTRACT_TERM

    '            If ISLOAN = True Then
    '                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
    '                Ecc = Math.Round(PincipleBal * (RATE * 12 / 365) * NumOfDate, 2)
    '            Else
    '                Ecc = Math.Round(PincipleBal * RATE, 2)
    '            End If
    '            Ecc += ACCURE_ECC

    '            If Ecc > ARM_INSTALLMENT_AMT Then
    '                ACCURE_ECC = Ecc - ARM_INSTALLMENT_AMT
    '                Ecc = ARM_INSTALLMENT_AMT
    '            Else
    '                ACCURE_ECC = 0
    '            End If

    '            Dim install As Decimal = 0
    '            If m >= ARM_CONTRACT_TERM Then
    '                install = PincipleBal + Ecc
    '                If ISLOAN Then
    '                    ARM_INSTALLMENT_AMT = install
    '                End If

    '            End If

    '            PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT - Ecc), 2)
    '            ECC_ACCU += Ecc
    '            If Math.Floor(PincipleBal) <= 0 Then
    '                PincipleBal = 0
    '            End If
    '            If Math.Floor(Ecc) <= 0 Then
    '                Ecc = 0
    '            End If




    '            Dim ECC_EXC_VAT As Decimal = 0
    '            Dim PUR_PRINCIPLE As Decimal = 0
    '            If ISLOAN = True Then
    '                ECC_EXC_VAT = Ecc
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = PincipleBal
    '            Else
    '                ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
    '                UCC_EXC_VAT -= ECC_EXC_VAT
    '                ACCU_ECC_EXC_VAT += ECC_EXC_VAT
    '                PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
    '            End If

    '            If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                PUR_PRINCIPLE = 0
    '                ECC_EXC_VAT -= PUR_PRINCIPLE
    '            End If

    '            If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
    '                ECC_EXC_VAT += UCC_EXC_VAT
    '                UCC_EXC_VAT = 0
    '            End If



    '            Dim ECC_DATA As New ECC_MASTER
    '            With ECC_DATA
    '                .Period = m
    '                .MONTHDATE = M_DATE
    '                .Instamemt_AMT = ARM_INSTALLMENT_AMT
    '                .InterestRate = RATE

    '                .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
    '                If ISLOAN = False Then
    '                    .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT * 7.0 / 107, 2)
    '                Else
    '                    .CUR_VAT = 0
    '                End If

    '                .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_VAT)

    '                '.CUR_PUR_INTEREST = ECC_EXC_VAT
    '                ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

    '                If ISLOAN = True Then
    '                    .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
    '                Else
    '                    .CUR_TOTAL = 0
    '                End If

    '                .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
    '                PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
    '                .PUR_UCC_BAL = UCC_EXC_VAT


    '                AR_BAL = AR_BAL - .Instamemt_AMT
    '                VAT_BAL = VAT_BAL - .CUR_VAT

    '                .AR_BAL = AR_BAL
    '                .VAT_BAL = VAT_BAL
    '                .NumberOfDate = NumOfDate
    '            End With

    '            LIST_ECC_MASTER.Add(ECC_DATA)
    '            lastdate = M_DATE
    '            If EndofMonth = True Then
    '                M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
    '            Else
    '                M_DATE = M_DATE.AddMonths(1)

    '                Dim tmpdate As Date = M_DATE
    '                Dim SaleDay As String = SaleDate.ToString("dd")
    '                If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
    '                    M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
    '                Else
    '                    M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
    '                End If
    '            End If
    '            NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





    '            ECC_DATA = Nothing
    '        Next

    '    Catch ex As Exception
    '        ex = Nothing
    '    End Try
    '    Return LIST_ECC_MASTER
    'End Function







    Public Shared Function CalInterest(ByVal Principle As Decimal, ByVal InstallmentAMT As List(Of Decimal), ByVal ContracTerm As Integer) As Decimal
        Dim Out As Decimal = 0
        Try
            Dim data(ContracTerm) As Double
            data(0) = Math.Abs(Principle) * -1
            For i As Integer = 1 To ContracTerm
                data(i) = InstallmentAMT(i - 1)
            Next
            Out = Math.Round(Financial.IRR(data, 0.01), 6)
        Catch ex As Exception
            ex = Nothing
        End Try
        Return Out
    End Function

    Public Shared Function CalInterest(ByVal Principle As Decimal, ByVal InstallmentAMT As Decimal, ByVal ContracTerm As Integer) As Decimal
        Dim Out As Decimal = 0
        Try
            Dim data(ContracTerm) As Double
            data(0) = Math.Abs(Principle) * -1
            For i As Integer = 1 To ContracTerm
                data(i) = InstallmentAMT
            Next
            Out = Math.Round(Financial.IRR(data, 0.01), 6)
        Catch ex As Exception
            ex = Nothing
        End Try
        Return Out
    End Function






    Public Shared Function GETECCMASTERBUTTOM(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As List(Of Decimal), ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True, Optional ByVal IsDiscountClosedStep As Boolean = False) As List(Of ECC_MASTER)
        Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
        Try
            Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
            Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
            Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
            Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
            Dim M_DATE As Date = FirstDueDate
            Dim lastdate As Date = SaleDate
            Dim Ecc As Decimal = 0
            Dim ECC_ACCU As Decimal = 0
            Dim ACCU_ECC_EXC_VAT As Decimal = 0
            Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
            Dim PincipleBal As Decimal = Math.Round(Principle, 2)
            Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
            Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

            Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
            Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
            If ISLOAN = True Then VAT_BAL = 0

            Dim NumOfDate As Integer = 30
            If ISLOAN = True Then
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                CONTRACT_UCC = 0
                CONTRACT_UCC_EXC_VAT = 0
                UCC_EXC_VAT = 0
                UCCBAL = 0
                PUR_Principle_Bal = PincipleBal
            End If

            Dim ACCURE_ECC As Decimal = 0
            For m As Integer = 1 To ARM_CONTRACT_TERM

                If ISLOAN = True Then
                    NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                    Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * NumOfDate, 2)
                Else
                    Ecc = Math.Round(PincipleBal * RATE, 2)
                End If
                Ecc += ACCURE_ECC

                If Ecc > ARM_INSTALLMENT_AMT(m - 1) Then
                    ACCURE_ECC = Ecc - ARM_INSTALLMENT_AMT(m - 1)
                    Ecc = ARM_INSTALLMENT_AMT(m - 1)
                Else
                    ACCURE_ECC = 0
                End If

                Dim install As Decimal = 0
                If m >= ARM_CONTRACT_TERM Then
                    install = PincipleBal + Ecc
                    If ISLOAN Then
                        ARM_INSTALLMENT_AMT(m - 1) = install
                    End If

                End If

                PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT(m - 1) - Ecc), 2)
                ECC_ACCU += Ecc
                If Math.Floor(PincipleBal) <= 0 Then
                    PincipleBal = 0
                End If
                If Math.Floor(Ecc) <= 0 Then
                    Ecc = 0
                End If




                Dim ECC_EXC_VAT As Decimal = 0
                Dim PUR_PRINCIPLE As Decimal = 0
                If ISLOAN = True Then
                    ECC_EXC_VAT = Ecc
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = PincipleBal
                Else
                    ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
                End If

                If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    PUR_PRINCIPLE = 0
                    ECC_EXC_VAT -= PUR_PRINCIPLE
                End If

                If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    ECC_EXC_VAT += UCC_EXC_VAT
                    UCC_EXC_VAT = 0
                End If



                Dim ECC_DATA As New ECC_MASTER
                With ECC_DATA
                    .Period = m
                    .MONTHDATE = M_DATE
                    .Instamemt_AMT = ARM_INSTALLMENT_AMT(m - 1)
                    .InterestRate = RATE

                    .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
                    If ISLOAN = False Then
                        .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT(m - 1) * 7.0 / 107, 2)
                    Else
                        .CUR_VAT = 0
                    End If

                    .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT(m - 1) - (.CUR_PUR_PRINCIPLE + .CUR_VAT)

                    '.CUR_PUR_INTEREST = ECC_EXC_VAT
                    ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

                    If ISLOAN = True Then
                        .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
                    Else
                        .CUR_TOTAL = 0
                    End If

                    .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
                    PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
                    .PUR_UCC_BAL = UCC_EXC_VAT


                    AR_BAL = AR_BAL - .Instamemt_AMT
                    VAT_BAL = VAT_BAL - .CUR_VAT

                    .AR_BAL = AR_BAL
                    .VAT_BAL = VAT_BAL
                    .NumberOfDate = NumOfDate


                    If IsDiscountClosedStep = True Then
                        Dim CK As Decimal = Math.Round(m / ARM_CONTRACT_TERM, 2)
                        If CK <= 0.33 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 60 / 100, 2)

                        ElseIf CK > 0.33 AndAlso CK <= 0.67 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 70 / 100, 2)
                        Else
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 100 / 100, 2)
                        End If
                    Else
                        .DiscountAMT = UCC_EXC_VAT / 2
                    End If
                End With


                LIST_ECC_MASTER.Add(ECC_DATA)
                lastdate = M_DATE
                If EndofMonth = True Then
                    M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
                Else
                    M_DATE = M_DATE.AddMonths(1)
                End If
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)




                ECC_DATA = Nothing
            Next

        Catch ex As Exception
            ex = Nothing
        End Try
        Return LIST_ECC_MASTER
    End Function

    Public Shared Function GETECCMASTERBUTTOM(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As Decimal, ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True, Optional ByVal IsDiscountClosedStep As Boolean = False) As List(Of ECC_MASTER)
        Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
        Try
            Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
            Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
            Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
            Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
            Dim M_DATE As Date = FirstDueDate
            Dim lastdate As Date = SaleDate
            Dim Ecc As Decimal = 0
            Dim ECC_ACCU As Decimal = 0
            Dim ACCU_ECC_EXC_VAT As Decimal = 0
            Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
            Dim PincipleBal As Decimal = Math.Round(Principle, 2)
            Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
            Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

            Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
            Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
            If ISLOAN = True Then VAT_BAL = 0

            Dim NumOfDate As Integer = 30
            If ISLOAN = True Then
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                CONTRACT_UCC = 0
                CONTRACT_UCC_EXC_VAT = 0
                UCC_EXC_VAT = 0
                UCCBAL = 0
                PUR_Principle_Bal = PincipleBal
            End If

            Dim ACCURE_ECC As Decimal = 0
            For m As Integer = 1 To ARM_CONTRACT_TERM

                If ISLOAN = True Then
                    NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                    Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * NumOfDate, 2)
                Else
                    Ecc = Math.Round(PincipleBal * RATE, 2)
                End If
                Ecc += ACCURE_ECC

                If Ecc > ARM_INSTALLMENT_AMT Then
                    ACCURE_ECC = Ecc - ARM_INSTALLMENT_AMT
                    Ecc = ARM_INSTALLMENT_AMT
                Else
                    ACCURE_ECC = 0
                End If

                Dim install As Decimal = 0
                If m >= ARM_CONTRACT_TERM Then
                    install = PincipleBal + Ecc
                    If ISLOAN Then
                        ARM_INSTALLMENT_AMT = install
                    End If

                End If

                PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT - Ecc), 2)
                ECC_ACCU += Ecc
                If Math.Floor(PincipleBal) <= 0 Then
                    PincipleBal = 0
                End If
                If Math.Floor(Ecc) <= 0 Then
                    Ecc = 0
                End If




                Dim ECC_EXC_VAT As Decimal = 0
                Dim PUR_PRINCIPLE As Decimal = 0
                If ISLOAN = True Then
                    ECC_EXC_VAT = Ecc
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = PincipleBal
                Else
                    ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
                End If

                If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    PUR_PRINCIPLE = 0
                    ECC_EXC_VAT -= PUR_PRINCIPLE
                End If

                If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    ECC_EXC_VAT += UCC_EXC_VAT
                    UCC_EXC_VAT = 0
                End If



                Dim ECC_DATA As New ECC_MASTER
                With ECC_DATA
                    .Period = m
                    .MONTHDATE = M_DATE
                    .Instamemt_AMT = ARM_INSTALLMENT_AMT
                    .InterestRate = RATE

                    .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
                    If ISLOAN = False Then
                        .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT * 7.0 / 107, 2)
                    Else
                        .CUR_VAT = 0
                    End If

                    .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_VAT)

                    '.CUR_PUR_INTEREST = ECC_EXC_VAT
                    ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

                    If ISLOAN = True Then
                        .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
                    Else
                        .CUR_TOTAL = 0
                    End If

                    .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
                    PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
                    .PUR_UCC_BAL = UCC_EXC_VAT


                    AR_BAL = AR_BAL - .Instamemt_AMT
                    VAT_BAL = VAT_BAL - .CUR_VAT

                    .AR_BAL = AR_BAL
                    .VAT_BAL = VAT_BAL
                    .NumberOfDate = NumOfDate


                    If IsDiscountClosedStep = True Then
                        Dim CK As Decimal = Math.Round(m / ARM_CONTRACT_TERM, 2)
                        If CK <= 0.33 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 60 / 100, 2)

                        ElseIf CK > 0.33 AndAlso CK <= 0.67 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 70 / 100, 2)
                        Else
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 100 / 100, 2)
                        End If
                    Else
                        .DiscountAMT = UCC_EXC_VAT / 2
                    End If

                End With

                LIST_ECC_MASTER.Add(ECC_DATA)
                lastdate = M_DATE
                If EndofMonth = True Then
                    M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
                Else
                    M_DATE = M_DATE.AddMonths(1)

                    Dim tmpdate As Date = M_DATE
                    Dim SaleDay As String = SaleDate.ToString("dd")
                    If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
                        M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
                    Else
                        M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
                    End If


                End If
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





                ECC_DATA = Nothing
            Next

        Catch ex As Exception
            ex = Nothing
        End Try
        Return LIST_ECC_MASTER
    End Function

    Public Shared Function GETECCMASTERCURRENT(ByVal PrincipleBal As Decimal, ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As Decimal, ByVal ARM_CONTRACT_TERM As Integer, ByVal FirstDueDate As Date, ByVal CurrentDate As Date, ByVal lastCutPrincDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True, Optional ByVal IsDiscountClosedStep As Boolean = False) As List(Of ECC_MASTER)
        Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
        Try
            Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
            Dim Principle As Decimal = PrincipleBal
            Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
            Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
            Dim M_DATE As Date = CurrentDate
            Dim lastdate As Date = lastCutPrincDate
            Dim Ecc As Decimal = 0
            Dim ECC_ACCU As Decimal = 0
            Dim ACCU_ECC_EXC_VAT As Decimal = 0
            Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
            Dim PincipleBal As Decimal = Math.Round(Principle, 2)
            Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
            Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

            Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
            Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
            If ISLOAN = True Then VAT_BAL = 0

            Dim NumOfDate As Integer = 30
            If ISLOAN = True Then
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                CONTRACT_UCC = 0
                CONTRACT_UCC_EXC_VAT = 0
                UCC_EXC_VAT = 0
                UCCBAL = 0
                PUR_Principle_Bal = PincipleBal
            End If
            Dim AccureBegin As Decimal = 0
            Dim ACCURE_ECC As Decimal = 0
            Dim N As Integer = DateDiff(DateInterval.Month, CurrentDate, DateAdd(DateInterval.Month, ARM_CONTRACT_TERM, FirstDueDate))
            For m As Integer = 1 To N
                AccureBegin = ACCURE_ECC
                If ISLOAN = True Then
                    NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                    Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * NumOfDate, 2)
                Else
                    Ecc = Math.Round(PincipleBal * RATE, 2)
                End If
                Ecc += ACCURE_ECC

                If Ecc > ARM_INSTALLMENT_AMT Then
                    ACCURE_ECC = Ecc - ARM_INSTALLMENT_AMT
                    Ecc = ARM_INSTALLMENT_AMT

                Else
                    ACCURE_ECC = 0
                End If

                Dim install As Decimal = 0
                If m >= N Then
                    install = PincipleBal + Ecc
                    If ISLOAN Then
                        ARM_INSTALLMENT_AMT = install
                    End If

                End If

                PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT - Ecc), 2)
                ECC_ACCU += Ecc
                If Math.Floor(PincipleBal) <= 0 Then
                    PincipleBal = 0
                End If
                If Math.Floor(Ecc) <= 0 Then
                    Ecc = 0
                End If




                Dim ECC_EXC_VAT As Decimal = 0
                Dim PUR_PRINCIPLE As Decimal = 0
                If ISLOAN = True Then
                    ECC_EXC_VAT = Ecc
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = PincipleBal
                Else
                    ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
                End If

                If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    PUR_PRINCIPLE = 0
                    ECC_EXC_VAT -= PUR_PRINCIPLE
                End If

                If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    ECC_EXC_VAT += UCC_EXC_VAT
                    UCC_EXC_VAT = 0
                End If



                Dim ECC_DATA As New ECC_MASTER
                With ECC_DATA
                    .Period = m
                    .MONTHDATE = M_DATE
                    .Instamemt_AMT = ARM_INSTALLMENT_AMT
                    .InterestRate = RATE
                    .CalFromDate = lastdate
                    .CalToDate = M_DATE
                    .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
                    If ISLOAN = False Then
                        .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT * 7.0 / 107, 2)
                    Else
                        .CUR_VAT = 0
                    End If

                    .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_VAT)
                    .CurrAccure = ACCURE_ECC - AccureBegin
                    .AccureEnd = ACCURE_ECC
                    .AccureBegin = AccureBegin
                    '.CUR_PUR_INTEREST = ECC_EXC_VAT
                    ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

                    If ISLOAN = True Then
                        .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
                    Else
                        .CUR_TOTAL = 0
                    End If

                    .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
                    PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
                    .PUR_UCC_BAL = UCC_EXC_VAT


                    AR_BAL = AR_BAL - .Instamemt_AMT
                    VAT_BAL = VAT_BAL - .CUR_VAT

                    .AR_BAL = AR_BAL
                    .VAT_BAL = VAT_BAL
                    .NumberOfDate = NumOfDate

                    If IsDiscountClosedStep = True Then
                        Dim CK As Decimal = Math.Round(m / ARM_CONTRACT_TERM, 2)
                        If CK <= 0.33 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 60 / 100, 2)

                        ElseIf CK > 0.33 AndAlso CK <= 0.67 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 70 / 100, 2)
                        Else
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 100 / 100, 2)
                        End If
                    Else
                        .DiscountAMT = UCC_EXC_VAT / 2
                    End If
                End With

                LIST_ECC_MASTER.Add(ECC_DATA)
                lastdate = M_DATE
                If EndofMonth = True Then
                    M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
                Else
                    M_DATE = M_DATE.AddMonths(1)
                    Dim tmpdate As Date = M_DATE
                    Dim SaleDay As String = CurrentDate.ToString("dd")
                    If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
                        M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
                    Else
                        M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
                    End If
                End If
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





                ECC_DATA = Nothing
            Next

        Catch ex As Exception
            ex = Nothing
        End Try
        Return LIST_ECC_MASTER
    End Function


    Public Shared Function GETECCMASTERTOP(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As List(Of Decimal), ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True, Optional ByVal IsDiscountClosedStep As Boolean = False) As List(Of ECC_MASTER)
        Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
        Try
            Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
            Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
            Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
            Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
            Dim M_DATE As Date = FirstDueDate
            Dim lastdate As Date = SaleDate
            Dim EndMonthOfSaleDate As Date = CDate(SaleDate.AddMonths(1).ToString("yyyy-MM-01")).AddDays(-1)

            Dim Ecc As Decimal = 0
            Dim ECC_ACCU As Decimal = 0
            Dim ACCU_ECC_EXC_VAT As Decimal = 0
            Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
            Dim PincipleBal As Decimal = Math.Round(Principle, 2)
            Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
            Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

            Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
            Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
            If ISLOAN = True Then VAT_BAL = 0

            Dim NumOfDate As Integer = 30
            Dim FirstNum As Integer = 0
            Dim FirstInterest As Decimal = 0.0

            If ISLOAN = True Then
                CONTRACT_UCC = 0
                CONTRACT_UCC_EXC_VAT = 0
                UCC_EXC_VAT = 0
                UCCBAL = 0
                PUR_Principle_Bal = PincipleBal

                FirstNum = DateDiff(DateInterval.Day, SaleDate, EndMonthOfSaleDate)
                FirstInterest = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * FirstNum, 2)
            End If

            Dim ACCURE_ECC As Decimal = 0
            For m As Integer = 1 To ARM_CONTRACT_TERM

                If ISLOAN = True Then
                    NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                    Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * NumOfDate, 2)
                Else
                    Ecc = Math.Round(PincipleBal * RATE, 2)
                End If
                Ecc += ACCURE_ECC



                If Ecc > ARM_INSTALLMENT_AMT(m - 1) + FirstInterest Then
                    ACCURE_ECC = Ecc - (ARM_INSTALLMENT_AMT(m - 1) + FirstInterest)
                    Ecc = ARM_INSTALLMENT_AMT(m - 1) + FirstInterest
                Else
                    ACCURE_ECC = 0
                End If

                Dim install As Decimal = 0
                If m >= ARM_CONTRACT_TERM Then
                    install = PincipleBal + Ecc
                    If ISLOAN Then
                        ARM_INSTALLMENT_AMT(m - 1) = install
                    End If

                End If

                PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT(m - 1) + FirstInterest - Ecc), 2)
                ECC_ACCU += Ecc
                If Math.Floor(PincipleBal) <= 0 Then
                    PincipleBal = 0
                End If
                If Math.Floor(Ecc) <= 0 Then
                    Ecc = 0
                End If




                Dim ECC_EXC_VAT As Decimal = 0
                Dim PUR_PRINCIPLE As Decimal = 0
                If ISLOAN = True Then
                    ECC_EXC_VAT = Ecc
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = PincipleBal
                Else
                    ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
                End If

                If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    PUR_PRINCIPLE = 0
                    ECC_EXC_VAT -= PUR_PRINCIPLE
                End If

                If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    ECC_EXC_VAT += UCC_EXC_VAT
                    UCC_EXC_VAT = 0
                End If



                Dim ECC_DATA As New ECC_MASTER
                With ECC_DATA
                    .Period = m
                    .MONTHDATE = M_DATE
                    .Instamemt_AMT = ARM_INSTALLMENT_AMT(m - 1) + FirstInterest

                    .InterestRate = RATE

                    .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
                    If ISLOAN = False Then
                        .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT(m - 1) * 7.0 / 107, 2)
                    Else
                        .CUR_VAT = 0
                    End If

                    .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT(m - 1) + FirstInterest - (.CUR_PUR_PRINCIPLE + .CUR_VAT)
                    FirstInterest = 0
                    '.CUR_PUR_INTEREST = ECC_EXC_VAT
                    ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

                    If ISLOAN = True Then
                        .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
                    Else
                        .CUR_TOTAL = 0
                    End If

                    .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
                    PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
                    .PUR_UCC_BAL = UCC_EXC_VAT


                    AR_BAL = AR_BAL - .Instamemt_AMT
                    VAT_BAL = VAT_BAL - .CUR_VAT

                    .AR_BAL = AR_BAL
                    .VAT_BAL = VAT_BAL
                    .NumberOfDate = NumOfDate


                    If IsDiscountClosedStep = True Then
                        Dim CK As Decimal = Math.Round(m / ARM_CONTRACT_TERM, 2)
                        If CK <= 0.33 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 60 / 100, 2)

                        ElseIf CK > 0.33 AndAlso CK <= 0.67 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 70 / 100, 2)
                        Else
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 100 / 100, 2)
                        End If
                    Else
                        .DiscountAMT = UCC_EXC_VAT / 2
                    End If
                End With

                LIST_ECC_MASTER.Add(ECC_DATA)
                lastdate = M_DATE
                If EndofMonth = True Then
                    M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
                Else
                    M_DATE = M_DATE.AddMonths(1)

                    Dim tmpdate As Date = M_DATE
                    Dim SaleDay As String = SaleDate.ToString("dd")
                    If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
                        M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
                    Else
                        M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
                    End If
                End If
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





                ECC_DATA = Nothing
            Next

        Catch ex As Exception
            ex = Nothing
        End Try
        Return LIST_ECC_MASTER
    End Function


    Public Shared Function GETECCMASTERTOP(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As Decimal, ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True, Optional ByVal IsDiscountClosedStep As Boolean = False) As List(Of ECC_MASTER)
        Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
        Try
            Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
            Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
            Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
            Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
            Dim M_DATE As Date = FirstDueDate
            Dim lastdate As Date = SaleDate
            Dim EndMonthOfSaleDate As Date = CDate(SaleDate.AddMonths(1).ToString("yyyy-MM-01")).AddDays(-1)

            Dim Ecc As Decimal = 0
            Dim ECC_ACCU As Decimal = 0
            Dim ACCU_ECC_EXC_VAT As Decimal = 0
            Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
            Dim PincipleBal As Decimal = Math.Round(Principle, 2)
            Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
            Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

            Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
            Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
            If ISLOAN = True Then VAT_BAL = 0

            Dim NumOfDate As Integer = 30
            Dim FirstNum As Integer = 0
            Dim FirstInterest As Decimal = 0.0

            If ISLOAN = True Then
                CONTRACT_UCC = 0
                CONTRACT_UCC_EXC_VAT = 0
                UCC_EXC_VAT = 0
                UCCBAL = 0
                PUR_Principle_Bal = PincipleBal

                FirstNum = DateDiff(DateInterval.Day, SaleDate, EndMonthOfSaleDate)
                FirstInterest = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * FirstNum, 2)
            End If

            Dim ACCURE_ECC As Decimal = 0
            For m As Integer = 1 To ARM_CONTRACT_TERM

                If ISLOAN = True Then
                    NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                    Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * NumOfDate, 2)
                Else
                    Ecc = Math.Round(PincipleBal * RATE, 2)
                End If
                Ecc += ACCURE_ECC



                If Ecc > ARM_INSTALLMENT_AMT + FirstInterest Then
                    ACCURE_ECC = Ecc - (ARM_INSTALLMENT_AMT + FirstInterest)
                    Ecc = ARM_INSTALLMENT_AMT + FirstInterest
                Else
                    ACCURE_ECC = 0
                End If

                Dim install As Decimal = 0
                If m >= ARM_CONTRACT_TERM Then
                    install = PincipleBal + Ecc
                    If ISLOAN Then
                        ARM_INSTALLMENT_AMT = install
                    End If

                End If

                PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT + FirstInterest - Ecc), 2)
                ECC_ACCU += Ecc
                If Math.Floor(PincipleBal) <= 0 Then
                    PincipleBal = 0
                End If
                If Math.Floor(Ecc) <= 0 Then
                    Ecc = 0
                End If




                Dim ECC_EXC_VAT As Decimal = 0
                Dim PUR_PRINCIPLE As Decimal = 0
                If ISLOAN = True Then
                    ECC_EXC_VAT = Ecc
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = PincipleBal
                Else
                    ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
                End If

                If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    PUR_PRINCIPLE = 0
                    ECC_EXC_VAT -= PUR_PRINCIPLE
                End If

                If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    ECC_EXC_VAT += UCC_EXC_VAT
                    UCC_EXC_VAT = 0
                End If



                Dim ECC_DATA As New ECC_MASTER
                With ECC_DATA
                    .Period = m
                    .MONTHDATE = M_DATE
                    .Instamemt_AMT = ARM_INSTALLMENT_AMT + FirstInterest

                    .InterestRate = RATE

                    .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
                    If ISLOAN = False Then
                        .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT * 7.0 / 107, 2)
                    Else
                        .CUR_VAT = 0
                    End If

                    .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT + FirstInterest - (.CUR_PUR_PRINCIPLE + .CUR_VAT)
                    FirstInterest = 0
                    '.CUR_PUR_INTEREST = ECC_EXC_VAT
                    ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

                    If ISLOAN = True Then
                        .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
                    Else
                        .CUR_TOTAL = 0
                    End If

                    .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
                    PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
                    .PUR_UCC_BAL = UCC_EXC_VAT


                    AR_BAL = AR_BAL - .Instamemt_AMT
                    VAT_BAL = VAT_BAL - .CUR_VAT

                    .AR_BAL = AR_BAL
                    .VAT_BAL = VAT_BAL
                    .NumberOfDate = NumOfDate

                    If IsDiscountClosedStep = True Then
                        Dim CK As Decimal = Math.Round(m / ARM_CONTRACT_TERM, 2)
                        If CK <= 0.33 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 60 / 100, 2)

                        ElseIf CK > 0.33 AndAlso CK <= 0.67 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 70 / 100, 2)
                        Else
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 100 / 100, 2)
                        End If
                    Else
                        .DiscountAMT = UCC_EXC_VAT / 2
                    End If
                End With

                LIST_ECC_MASTER.Add(ECC_DATA)
                lastdate = M_DATE
                If EndofMonth = True Then
                    M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
                Else
                    M_DATE = M_DATE.AddMonths(1)

                    Dim tmpdate As Date = M_DATE
                    Dim SaleDay As String = SaleDate.ToString("dd")
                    If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
                        M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
                    Else
                        M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
                    End If
                End If
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





                ECC_DATA = Nothing
            Next

        Catch ex As Exception
            ex = Nothing
        End Try
        Return LIST_ECC_MASTER
    End Function

    Public Shared Function GETECCMASTER(ByVal ARM_CASH_PRICE As Decimal, ByVal ARM_HP_PRICE As Decimal, ByVal ARM_DOWN_AMT As Decimal, ByVal ARM_INSTALLMENT_AMT As Decimal, ByVal ARM_CONTRACT_TERM As Integer, ByVal SaleDate As Date, ByVal FirstDueDate As Date, Optional ByVal ISLOAN As Boolean = False, Optional ByVal EndofMonth As Boolean = True, Optional ByVal IsDiscountClosedStep As Boolean = False) As List(Of ECC_MASTER)
        Dim LIST_ECC_MASTER As New List(Of ECC_MASTER)
        Try
            Dim RATE = CalInterest(ARM_CASH_PRICE - ARM_DOWN_AMT, ARM_INSTALLMENT_AMT, ARM_CONTRACT_TERM)
            Dim Principle As Decimal = ARM_CASH_PRICE - ARM_DOWN_AMT
            Dim CONTRACT_UCC As Decimal = ARM_HP_PRICE - ARM_CASH_PRICE
            Dim CONTRACT_UCC_EXC_VAT As Decimal = Math.Round(CONTRACT_UCC * 100 / 107, 2)
            Dim M_DATE As Date = FirstDueDate
            Dim lastdate As Date = SaleDate
            Dim Ecc As Decimal = 0
            Dim ECC_ACCU As Decimal = 0
            Dim ACCU_ECC_EXC_VAT As Decimal = 0
            Dim UCC_EXC_VAT As Decimal = CONTRACT_UCC_EXC_VAT
            Dim PincipleBal As Decimal = Math.Round(Principle, 2)
            Dim PUR_Principle_Bal As Decimal = Math.Round(PincipleBal * 100 / 107, 2)
            Dim UCCBAL As Decimal = Math.Round(CONTRACT_UCC_EXC_VAT, 2)

            Dim AR_BAL As Decimal = ARM_HP_PRICE - ARM_DOWN_AMT
            Dim VAT_BAL As Decimal = AR_BAL - (AR_BAL / 1.07)
            If ISLOAN = True Then VAT_BAL = 0

            Dim NumOfDate As Integer = 30
            If ISLOAN = True Then
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                CONTRACT_UCC = 0
                CONTRACT_UCC_EXC_VAT = 0
                UCC_EXC_VAT = 0
                UCCBAL = 0
                PUR_Principle_Bal = PincipleBal
            End If

            Dim ACCURE_ECC As Decimal = 0
            For m As Integer = 1 To ARM_CONTRACT_TERM

                If ISLOAN = True Then
                    NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)
                    Ecc = Math.Round(PincipleBal * (Math.Round(RATE * 12, 6) / 365) * NumOfDate, 2)
                Else
                    Ecc = Math.Round(PincipleBal * RATE, 2)
                End If
                Ecc += ACCURE_ECC

                If Ecc > ARM_INSTALLMENT_AMT Then
                    ACCURE_ECC = Ecc - ARM_INSTALLMENT_AMT
                    Ecc = ARM_INSTALLMENT_AMT
                Else
                    ACCURE_ECC = 0
                End If

                Dim install As Decimal = 0
                If m >= ARM_CONTRACT_TERM Then
                    install = PincipleBal + Ecc
                    If ISLOAN Then
                        ARM_INSTALLMENT_AMT = install
                    End If

                End If

                PincipleBal = PincipleBal - Math.Round((ARM_INSTALLMENT_AMT - Ecc), 2)
                ECC_ACCU += Ecc
                If Math.Floor(PincipleBal) <= 0 Then
                    PincipleBal = 0
                End If
                If Math.Floor(Ecc) <= 0 Then
                    Ecc = 0
                End If




                Dim ECC_EXC_VAT As Decimal = 0
                Dim PUR_PRINCIPLE As Decimal = 0
                If ISLOAN = True Then
                    ECC_EXC_VAT = Ecc
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = PincipleBal
                Else
                    ECC_EXC_VAT = Math.Round(Ecc * 100 / 107, 2)
                    UCC_EXC_VAT -= ECC_EXC_VAT
                    ACCU_ECC_EXC_VAT += ECC_EXC_VAT
                    PUR_PRINCIPLE = Math.Round(PincipleBal * 100 / 107, 2)
                End If

                If Math.Floor(PUR_PRINCIPLE) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    PUR_PRINCIPLE = 0
                    ECC_EXC_VAT -= PUR_PRINCIPLE
                End If

                If Math.Floor(UCC_EXC_VAT) = 0 OrElse m = ARM_CONTRACT_TERM Then
                    ECC_EXC_VAT += UCC_EXC_VAT
                    UCC_EXC_VAT = 0
                End If



                Dim ECC_DATA As New ECC_MASTER
                With ECC_DATA
                    .Period = m
                    .MONTHDATE = M_DATE
                    .Instamemt_AMT = ARM_INSTALLMENT_AMT
                    .InterestRate = RATE

                    .CUR_PUR_PRINCIPLE = PUR_Principle_Bal - PUR_PRINCIPLE
                    If ISLOAN = False Then
                        .CUR_VAT = Math.Round(ARM_INSTALLMENT_AMT * 7.0 / 107, 2)
                    Else
                        .CUR_VAT = 0
                    End If

                    .CUR_PUR_INTEREST = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_VAT)

                    '.CUR_PUR_INTEREST = ECC_EXC_VAT
                    ' .CUR_VAT = ARM_INSTALLMENT_AMT - (.CUR_PUR_PRINCIPLE + .CUR_PUR_INTEREST)

                    If ISLOAN = True Then
                        .CUR_TOTAL = .CUR_PUR_INTEREST + .CUR_PUR_PRINCIPLE + .CUR_VAT
                    Else
                        .CUR_TOTAL = 0
                    End If

                    .PUR_PRINCIPLE_BAL = PUR_PRINCIPLE
                    PUR_Principle_Bal = .PUR_PRINCIPLE_BAL
                    .PUR_UCC_BAL = UCC_EXC_VAT

                    AR_BAL = AR_BAL - .Instamemt_AMT
                    VAT_BAL = VAT_BAL - .CUR_VAT

                    .AR_BAL = AR_BAL
                    .VAT_BAL = VAT_BAL
                    .NumberOfDate = NumOfDate

                    If IsDiscountClosedStep = True Then
                        Dim CK As Decimal = Math.Round(m / ARM_CONTRACT_TERM, 2)
                        If CK <= 0.33 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 60 / 100, 2)

                        ElseIf CK > 0.33 AndAlso CK <= 0.67 Then
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 70 / 100, 2)
                        Else
                            .DiscountAMT = Math.Round(UCC_EXC_VAT * 100 / 100, 2)
                        End If
                    Else
                        .DiscountAMT = UCC_EXC_VAT / 2
                    End If
                End With

                LIST_ECC_MASTER.Add(ECC_DATA)
                lastdate = M_DATE
                If EndofMonth = True Then
                    M_DATE = CDate(String.Format("{0}-{1}-01", (M_DATE.AddMonths(2).Year), (M_DATE.AddMonths(2).Month))).AddDays(-1)
                Else
                    M_DATE = M_DATE.AddMonths(1)

                    Dim tmpdate As Date = M_DATE
                    Dim SaleDay As String = SaleDate.ToString("dd")
                    If IsDate((String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))) = True Then
                        M_DATE = CDate(String.Format("{0}-{1}", tmpdate.ToString("yyyy-MM"), SaleDay))
                    Else
                        M_DATE = CDate(String.Format("{0}-01", tmpdate.AddMonths(1).ToString("yyyy-MM"))).AddDays(-1)
                    End If
                End If
                NumOfDate = DateDiff(DateInterval.Day, lastdate, M_DATE)





                ECC_DATA = Nothing
            Next

        Catch ex As Exception
            ex = Nothing
        End Try
        Return LIST_ECC_MASTER
    End Function








End Class
