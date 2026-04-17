Sub ExportPeriodData()
    '
    ' Roda sPricing para um bond CDI+ e exporta dados per-periodo
    ' para um arquivo texto que o Python consegue ler e comparar.
    '
    ' Basta rodar este macro (F5) com o Excel aberto.
    '

    Dim sCetip As String
    Dim dtDay As Date
    Dim dPU As Double
    Dim sFile As String
    Dim f As Integer
    Dim i As Integer
    Dim n As Integer

    ' === CONFIGURACAO (mude aqui se quiser outro ativo) ===
    sCetip = "6581326SR1"
    dtDay = CDate("04/14/2026")
    dPU = 1008.640978
    sFile = "X:\BDM\CRI\vba_period_dump_" & sCetip & ".txt"
    ' =====================================================

    ' Rodar pricing completo (taxa + duration + spread)
    Call sPricing(sCetip, dtDay, dPU, 2, 8)

    n = tBondInfo.iPeriods

    ' Abrir arquivo para escrita
    f = FreeFile
    Open sFile For Output As #f

    ' Header com resultados
    Print #f, "BOND=" & tBondInfo.sCETIP
    Print #f, "INDEX=" & tBondInfo.sIndex
    Print #f, "YIELD=" & CStr(tBondResults.dYield)
    Print #f, "PRICE=" & CStr(tBondResults.dPrice)
    Print #f, "DURATION=" & CStr(tBondResults.dDuration)
    Print #f, "SPREAD=" & CStr(tBondResults.dSpread)
    Print #f, "PAR=" & CStr(tBondResults.dPar)
    Print #f, "PERIODS=" & CStr(n)
    Print #f, ""

    ' Header colunas
    Print #f, "i" & vbTab & "dtDay" & vbTab & "dPVpmtCalc" & vbTab & "dPVfactorCalc" & vbTab & _
              "dPMTTotal" & vbTab & "dYdi1" & vbTab & "dYcdi" & vbTab & "dYspread" & vbTab & _
              "dYtotal" & vbTab & "dSN" & vbTab & "dSNA" & vbTab & "dPMTJuros" & vbTab & _
              "dPMTAmort" & vbTab & "dFatAmAcc" & vbTab & "dPVpmtPar"

    ' Dados per-periodo
    For i = 1 To n
        Print #f, CStr(i) & vbTab & _
                  CStr(tPeriodInfo(i).dtDay) & vbTab & _
                  CStr(tPeriodBond(i).dPVpmtCalc) & vbTab & _
                  CStr(tPeriodBond(i).dPVfactorCalc) & vbTab & _
                  CStr(tPeriodBond(i).dPMTTotal) & vbTab & _
                  CStr(tPeriodBond(i).dYdi1) & vbTab & _
                  CStr(tPeriodBond(i).dYcdi) & vbTab & _
                  CStr(tPeriodBond(i).dYspread) & vbTab & _
                  CStr(tPeriodBond(i).dYtotal) & vbTab & _
                  CStr(tPeriodBond(i).dSN) & vbTab & _
                  CStr(tPeriodBond(i).dSNA) & vbTab & _
                  CStr(tPeriodBond(i).dPMTJuros) & vbTab & _
                  CStr(tPeriodBond(i).dPMTAmort) & vbTab & _
                  CStr(tPeriodBond(i).dFatAmAcc) & vbTab & _
                  CStr(tPeriodBond(i).dPVpmtPar)
    Next

    Close #f

    MsgBox "Exportado " & n & " periodos para:" & vbCrLf & sFile, vbInformation, "OK"

End Sub
