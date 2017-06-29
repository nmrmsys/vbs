@if(0)==(0)echo off
  @echo off
  cscript //nologo "%~f0?.wsf" %*
  goto :EOF
@end
<job><script language="VBScript">
  Dim oArgs, oWSH, oFSO, oXLS, oWbk, oSht, arFiles(), sTmp
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oWSH = CreateObject("WScript.Shell")

  With WScript
    Redim arFiles(0)
    For i = 0 To .Arguments.Count - 1
      n = Ubound(arFiles) + 1
      Redim Preserve arFiles(n)
      arFiles(n) = .Arguments(i)
    Next
    If .Arguments.Count = 0 Then
      sTmp = oFSO.GetSpecialFolder(2) & "\Tix.tmp"
      Set oTmp = oFSO.CreateTextFile(sTmp, True)
      oTmp.Write .StdIn.ReadAll
      oTmp.Close
      n = Ubound(arFiles) + 1
      Redim Preserve arFiles(n)
      arFiles(n) = sTmp
    End If
    sCurDir = oWSH.CurrentDirectory
    For i = 1 To Ubound(arFiles)
      'MsgBox arFiles(i)
      If IsEmpty(oXLS) Then
          Set oXLS = CreateObject("Excel.Application")
          oXLS.Visible = True
          oXLS.WindowState = -4137 'xlMaximized
          oWSH.AppActivate oXLS.Caption
      End If
      If InStr(arFiles(i),"\") = 0 Then
        arFiles(i) = sCurDir & "\" & arFiles(i)
      End If
      'MsgBox arFiles(i)
      Call CsvImpExcel(arFiles(i))
    Next
    If .Arguments.Count = 0 Then
      oFSO.DeleteFile(sTmp)
    End If
  End With
  
  Sub CsvImpExcel(argFileName)
    'MsgBox argFileName
    oXLS.Workbooks.Add
    Set oWbk = oXLS.ActiveWorkbook
    Set oSht = oWbk.ActiveSheet
    With oSht.QueryTables.Add("TEXT;" & argFileName, oSht.Range("$A$1"))
      '.CommandType = 0
      '.Name = "sample"
      '.FieldNames = True
      '.RowNumbers = False
      '.FillAdjacentFormulas = False
      '.PreserveFormatting = True
      '.RefreshOnFileOpen = False
      '.RefreshStyle = 1 'xlInsertDeleteCells
      '.RefreshPeriod = 0
      '.TextFilePromptOnRefresh = False
      '.SavePassword = False
      '.SaveData = True
      
      .TextFilePlatform = 932
      .TextFileStartRow = 1
      .TextFileTextQualifier = 1 'xlTextQualifierDoubleQuote
      .TextFileTrailingMinusNumbers = True
      .AdjustColumnWidth = True
      
      .TextFileParseType = 1 'xlDelimited
      .TextFileCommaDelimiter = True
      .TextFileTabDelimiter = True
      .TextFileSemicolonDelimiter = False
      .TextFileSpaceDelimiter = False
      .TextFileConsecutiveDelimiter = False
      '.TextFileOtherDelimiter = ":"
      nCols = GetCsvCols(argFileName) 'Interim
      Redim arCDTs(nCols - 1)
      For j = 0 To ubound(arCDTs)
        arCDTs(j) = 2
      Next
      
      '.TextFileParseType = 2 'xlFixedWidth
      '.TextFileFixedColumnWidths = Array(1, 1, 1)
      
      '.TextFileColumnDataTypes = Array(2, 2, 2)
      .TextFileColumnDataTypes = arCDTs
      
      .Refresh
      .Delete
      
    End With
    oWbk.Saved = True
  End Sub
  
  Function GetCsvCols(argFileName)
    Set oCsv = oFSO.OpenTextFile(argFileName)
    If Not oCsv.AtEndOfLine Then
      sLine = oCsv.ReadLine
      arLine = Split(sLine,",")
      GetCsvCols = ubound(arLine) + 1
    Else
      GetCsvCols = 1
    End If
  End Function
</script></job>
