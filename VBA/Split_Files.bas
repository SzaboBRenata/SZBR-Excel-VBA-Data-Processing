Attribute VB_Name = "Split_Files"
Option Explicit

Sub SplitFiles()

    Dim sourceWorkbook As Workbook           ' Forrásmunkafüzet – az a munkafüzet, amelybõl az adatokat beolvassuk.
    Dim sourceWorksheet As Worksheet         ' Forrásmunkalap – a forrásmunkafüzet azon munkalapja, amely tartalmazza a feldolgozandó adatokat.
    Dim outputWorkbook As Workbook           ' Kimeneti munkafüzet – az egyes megyék adatainak külön mentett Excel-fájlja.
    Dim outputWorksheet As Worksheet         ' Kimeneti munkalap – az új munkafüzet munkalapja, amelyre az adott megye adatai kerülnek.
    
    Dim countyDictionary As Object           ' Megyék szótára – a feldolgozás során elõforduló megyék egyedi azonosítására és nyilvántartására szolgál.
    Dim countyRows As Object                 ' Megyék sorainak szótára – megyénként tárolja az adott megyéhez tartozó adatsorokat.
    
    Dim sourceFilePath As String             ' Forrásfájl elérési útja – a feldolgozandó Excel-fájl teljes elérési útja.
    Dim outputFolderPath As String           ' Kimeneti mappa elérési útja – az elkészült fájlok mentési helye.
    
    Dim lastRow As Long                      ' Utolsó sor – a forrásmunkalapon található utolsó adatot tartalmazó sor száma.
    Dim lastColumn As Long                   ' Utolsó oszlop – a forrásmunkalapon található utolsó adatot tartalmazó oszlop száma.
    
    Dim sourceData As Variant                ' Forrásadatok – a munkalapról tömbbe beolvasott teljes adattartomány.
    Dim countyName As String                 ' Megyenév – az aktuálisan feldolgozott megye neve.
    Dim safeCountyName As String             ' Biztonságos megyenév – a fájlnévként is használható, tiltott karakterektõl megtisztított megyenév.
    
    Dim rowIndex As Long                     ' Sorszámláló – a forrásadatok sorainak bejárására szolgál.
    Dim columnIndex As Long                  ' Oszlopszámláló – a forrásadatok oszlopainak bejárására szolgál.
    Dim outputRowCount As Long               ' Kimeneti sorok száma – az adott megyéhez tartozó adatsorok darabszáma.
    
    Dim countyKey As Variant                 ' Megyekulcs – a megyék szótárában szereplõ aktuális megye azonosítója.
    Dim rowCollection As Collection          ' Sorok gyûjteménye – az aktuális megyéhez tartozó forrássorok sorszámait tartalmazza.
    Dim outputData() As Variant              ' Kimeneti adattömb – az adott megyéhez tartozó, mentésre elõkészített adatokat tárolja.
    
    Dim previousCalculation As XlCalculation ' Korábbi számítási mód – az Excel eredeti számítási beállítását tárolja a visszaállításhoz.
    Dim previousScreenUpdating As Boolean    ' Korábbi képernyõfrissítési állapot – az Excel eredeti képernyõfrissítési beállítását tárolja.
    Dim previousEnableEvents As Boolean      ' Korábbi eseménykezelési állapot – az Excel eredeti eseménykezelési beállítását tárolja.
    Dim previousDisplayAlerts As Boolean     ' Korábbi figyelmeztetési állapot – az Excel eredeti figyelmeztetési beállítását tárolja.
    
    Dim startTime As Double                  ' Kezdési idõ – a program futásának kezdõ idõpontját tárolja.
    Dim processedCounties As Long            ' Feldolgozott megyék száma – az eddig elkészült megyei fájlok darabszáma.
    Dim totalCounties As Long                ' Összes megye száma – a feldolgozásra váró megyék teljes darabszáma.
    
    Dim sourceRow As Long                    ' Forrássor – az aktuálisan feldolgozott forrássor sorszáma.
    Dim cellValue As Variant                 ' Cellaérték – az aktuálisan vizsgált vagy módosított cella tartalmát tárolja.
    
    On Error GoTo ErrorHandler

    startTime = Timer


    ' Excel beállítások mentése és gyorsítás
    previousCalculation = Application.Calculation
    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Export elõkészítése"


    ' Forrásfájl megnyitása
    sourceFilePath = ThisWorkbook.Path & "\forrasadatok.xlsx"

    If Dir(sourceFilePath) = "" Then
        MsgBox "Az forrasadatok.xlsx fájl nem található:" & vbCrLf & sourceFilePath, vbExclamation
        GoTo CleanExit
    End If

    Set sourceWorkbook = Workbooks.Open(fileName:=sourceFilePath, ReadOnly:=True)


    ' A feldolgozandó munkalap neve
    Set sourceWorksheet = Nothing
    On Error Resume Next
    Set sourceWorksheet = sourceWorkbook.Worksheets("megye")
    On Error GoTo ErrorHandler

    If sourceWorksheet Is Nothing Then
        MsgBox "A megye nevû munkalap nem található az forrasadatok.xlsx fájlban.", vbExclamation
        GoTo CleanExit
    End If
    
    
    ' Az adattartomány meghatározása és beolvasása
    lastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "B").End(xlUp).Row
    lastColumn = sourceWorksheet.Cells(1, sourceWorksheet.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "Nincs exportálható adat a munkalapon.", vbInformation
        GoTo CleanExit
    End If

    Application.StatusBar = "Adatok beolvasása"

    sourceData = sourceWorksheet.Range( _
        sourceWorksheet.Cells(1, 1), _
        sourceWorksheet.Cells(lastRow, lastColumn) _
    ).Value2

    
    ' Megyék szerinti csoportosítás a "B" oszlop alapján
    Set countyDictionary = CreateObject("Scripting.Dictionary")
    Set countyRows = CreateObject("Scripting.Dictionary")

    countyDictionary.CompareMode = vbTextCompare
    countyRows.CompareMode = vbTextCompare

    Application.StatusBar = "Csoportosítás"

    For rowIndex = 2 To lastRow

        If Not IsError(sourceData(rowIndex, 2)) Then
            countyName = Trim$(CStr(sourceData(rowIndex, 2)))

            If Len(countyName) > 0 Then
                If Not countyDictionary.Exists(countyName) Then
                    countyDictionary.Add countyName, True
                    Set rowCollection = New Collection
                    countyRows.Add countyName, rowCollection
                End If

                countyRows(countyName).Add rowIndex
            End If
        End If

    Next rowIndex

    totalCounties = countyDictionary.Count

    If totalCounties = 0 Then
        MsgBox "Nem található egyetlen megye sem a B oszlopban.", vbExclamation
        GoTo CleanExit
    End If

    outputFolderPath = sourceWorkbook.Path & "\"
    processedCounties = 0

    
    ' Megyénkénti exportálás
    For Each countyKey In countyDictionary.Keys

        processedCounties = processedCounties + 1
        Application.StatusBar = "Exportálás: " & processedCounties & "/" & _
                                totalCounties & " - " & CStr(countyKey)

        Set rowCollection = countyRows(countyKey)
        outputRowCount = rowCollection.Count

        ' Új munkafüzet létrehozása
        Set outputWorkbook = Workbooks.Add(xlWBATWorksheet)
        Set outputWorksheet = outputWorkbook.Worksheets(1)
        outputWorksheet.Name = "adatok"

        ' Fejléc és oszlopszélességek másolása
        sourceWorksheet.Range(sourceWorksheet.Cells(1, 1), _
                              sourceWorksheet.Cells(1, lastColumn)).Copy
        outputWorksheet.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
        outputWorksheet.Rows(1).RowHeight = sourceWorksheet.Rows(1).RowHeight

        For columnIndex = 1 To lastColumn
            outputWorksheet.Columns(columnIndex).ColumnWidth = _
                sourceWorksheet.Columns(columnIndex).ColumnWidth
        Next columnIndex

        ' Kimeneti adattömb összeállítása
        ReDim outputData(1 To outputRowCount, 1 To lastColumn)

        For rowIndex = 1 To outputRowCount
            sourceRow = rowCollection(rowIndex)

            For columnIndex = 1 To lastColumn
                cellValue = sourceData(sourceRow, columnIndex)

                ' Cellán belüli sortörések és többszörös szóközök eltávolítása
                If VarType(cellValue) = vbString Then
                    cellValue = Replace(cellValue, vbCrLf, " ")
                    cellValue = Replace(cellValue, vbCr, " ")
                    cellValue = Replace(cellValue, vbLf, " ")

                    Do While InStr(cellValue, "  ") > 0
                        cellValue = Replace(cellValue, "  ", " ")
                    Loop

                    cellValue = Trim$(cellValue)
                End If

                outputData(rowIndex, columnIndex) = cellValue
            Next columnIndex
        Next rowIndex

        ' Eredeti adatsorok formátumának átvétele
        For rowIndex = 1 To outputRowCount
            sourceRow = rowCollection(rowIndex)

            sourceWorksheet.Range(sourceWorksheet.Cells(sourceRow, 1), _
                                  sourceWorksheet.Cells(sourceRow, lastColumn)).Copy
            outputWorksheet.Cells(rowIndex + 1, 1).PasteSpecial Paste:=xlPasteFormats
        Next rowIndex

        Application.CutCopyMode = False

        ' Adatok kiírása
        outputWorksheet.Cells(2, 1).Resize(outputRowCount, lastColumn).Value = outputData
        Erase outputData

        ' Egységes betûtípus és sortörés-beállítások
        With outputWorksheet.Range(outputWorksheet.Cells(1, 1), _
                                   outputWorksheet.Cells(outputRowCount + 1, lastColumn))
            .Font.Name = "Segoe UI"
            .Font.Size = 9
        End With

        With outputWorksheet.Range(outputWorksheet.Cells(2, 1), _
                                   outputWorksheet.Cells(outputRowCount + 1, lastColumn))
            .WrapText = False
            .ShrinkToFit = False
        End With

        outputWorksheet.Range("F1").WrapText = False
        outputWorksheet.Range("I1").WrapText = False
        outputWorksheet.Rows("2:" & outputRowCount + 1).RowHeight = 18

        ' Fájlnév elõkészítése és mentés
        safeCountyName = CleanFileName(CStr(countyKey))
        If Len(safeCountyName) = 0 Then safeCountyName = "Ismeretlen_megye"

        outputWorksheet.Activate
        outputWorksheet.Range("A1").Select

        outputWorkbook.SaveAs _
            fileName:=GetUniqueFileName(outputFolderPath, safeCountyName), _
            FileFormat:=xlOpenXMLWorkbook

        outputWorkbook.Close SaveChanges:=False
        Set outputWorkbook = Nothing
        Set outputWorksheet = Nothing

    Next countyKey

    MsgBox "Az exportálás sikeresen befejezõdött!" & vbCrLf & vbCrLf & _
           "Feldolgozott sorok: " & Format(lastRow - 1, "#,##0") & vbCrLf & _
           "Létrehozott megyék: " & Format(totalCounties, "#,##0") & vbCrLf & _
           "Idõ: " & Format(Timer - startTime, "0.0") & " másodperc", vbInformation

CleanExit:

    On Error Resume Next

    If Not outputWorkbook Is Nothing Then outputWorkbook.Close SaveChanges:=False
    If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close SaveChanges:=False

    Application.CutCopyMode = False
    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
    Application.DisplayAlerts = previousDisplayAlerts
    Application.Calculation = previousCalculation
    Application.StatusBar = False

    On Error GoTo 0
    Exit Sub

ErrorHandler:

    MsgBox "Hiba történt!" & vbCrLf & vbCrLf & _
           "Hibakód: " & Err.Number & vbCrLf & _
           "Leírás: " & Err.Description & vbCrLf & vbCrLf & _
           "Feldolgozott megyék: " & processedCounties & "/" & totalCounties, vbCritical

    Resume CleanExit

End Sub


' Tiltott karakterek eltávolítása a fájlnévbõl
Private Function CleanFileName(ByVal fileName As String) As String

    Dim invalidCharacters As Variant
    Dim currentCharacter As Variant

    invalidCharacters = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    CleanFileName = Trim$(fileName)

    For Each currentCharacter In invalidCharacters
        CleanFileName = Replace(CleanFileName, CStr(currentCharacter), "_")
    Next currentCharacter

    Do While Len(CleanFileName) > 0
        If Right$(CleanFileName, 1) <> "." Then Exit Do
        CleanFileName = Left$(CleanFileName, Len(CleanFileName) - 1)
    Loop

End Function


' Egyedi fájlnév képzése, ha a fájl már létezik
Private Function GetUniqueFileName( _
    ByVal folderPath As String, _
    ByVal baseName As String _
) As String

    Dim fullPath As String
    Dim fileNumber As Long

    fullPath = folderPath & baseName & ".xlsx"

    If Dir(fullPath) = "" Then
        GetUniqueFileName = fullPath
        Exit Function
    End If

    fileNumber = 1

    Do
        fullPath = folderPath & baseName & "_" & fileNumber & ".xlsx"
        fileNumber = fileNumber + 1
    Loop While Dir(fullPath) <> ""

    GetUniqueFileName = fullPath

End Function


