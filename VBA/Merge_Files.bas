Attribute VB_Name = "Merge_Files"
Option Explicit

Public Sub MergeFiles()

    Dim targetWorkbook As Workbook      ' Célmunkafüzet – az a munkafüzet, amelybe az adatokat összegyûjtjük.
    Dim targetWorksheet As Worksheet    ' Célmunkalap – a munkalap, amelyre az adatokat bemásoljuk.
    Dim sourceWorkbook As Workbook      ' Forrásmunkafüzet – az éppen megnyitott fájl, amelybõl az adatokat másoljuk.
    Dim sourceWorksheet As Worksheet    ' Forrásmunkalap – a forrásmunkafüzet azon munkalapja, amelyrõl másolunk.
    
    Dim folderPath As String            ' Mappa elérési útja – az ADAT mappa teljes elérési útja.
    Dim fileName As String              ' Fájlnév – az éppen feldolgozás alatt álló fájl neve.
    Dim filePath As String              ' Fájl elérési útja – az aktuális fájl teljes elérési útja.
    
    Dim targetLastRow As Long           ' Cél utolsó sora – az utolsó adatot tartalmazó sor a célmunkalapon.
    Dim sourceLastRow As Long           ' Forrás utolsó sora – az utolsó adatot tartalmazó sor a forrásmunkalapon.
    Dim sourceLastColumn As Long        ' Forrás utolsó oszlopa – az utolsó adatot tartalmazó oszlop a forrásmunkalapon.
    Dim firstDataRow As Long            ' Elsõ adatsor – megadja, hogy a forrásmunkalap melyik sorától kezdõdjön az adatok másolása.
    
    Dim isTargetEmpty As Boolean        ' Üres célmunkalap – jelzi, hogy a célmunkalap üres-e.
    Dim columnIndex As Long             ' Oszlopindex – az oszlopok szélességének beállításakor használt számláló.
    
    Dim startTime As Double             ' Kezdési idõ – a makró futásának kezdõ idõpontja.
    Dim processedFiles As Long          ' Feldolgozott fájlok száma – a sikeresen beolvasott fájlok darabszáma.
    Dim processedRows As Long           ' Feldolgozott sorok száma – az összefûzött adatsorok darabszáma.
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    processedFiles = 0
    processedRows = 0

    
    ' Az ADAT mappában lévõ fájlokat egyesíti
    Set targetWorkbook = ThisWorkbook
    Set targetWorksheet = targetWorkbook.Worksheets(1)

    folderPath = targetWorkbook.Path & "\ADAT\"

    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Az ADAT mappa nem található:" & vbCrLf & folderPath, vbExclamation
        Exit Sub
    End If

    
    ' A cél munkalapot ellenõrzi
    isTargetEmpty = (Application.WorksheetFunction.CountA(targetWorksheet.Cells) = 0)

    If isTargetEmpty Then
        targetLastRow = 0
    Else
        targetLastRow = targetWorksheet.Cells(targetWorksheet.Rows.Count, "A").End(xlUp).Row
    End If

 
    ' Az excel-fájlok beolvasása
    fileName = Dir(folderPath & "*.xls*")

    Do While Len(fileName) > 0

        ' Az Excel által létrehozott ideiglenes fájlok kihagyása
        If Left$(fileName, 2) <> "~$" Then

            filePath = folderPath & fileName
            Set sourceWorkbook = Nothing

            ' Forrásfájl megnyitása
            On Error Resume Next
            Set sourceWorkbook = Workbooks.Open(fileName:=filePath, ReadOnly:=True)
            On Error GoTo ErrorHandler

            If Not sourceWorkbook Is Nothing Then
            
                processedFiles = processedFiles + 1

                Set sourceWorksheet = sourceWorkbook.Worksheets(1)

                ' Utolsó használt oszlop a fejléc alapján
                sourceLastColumn = sourceWorksheet.Cells(1, sourceWorksheet.Columns.Count).End(xlToLeft).Column

                ' Utolsó használt sor az A oszlop alapján
                sourceLastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "A").End(xlUp).Row

                If sourceLastRow >= 1 And sourceLastColumn >= 1 Then

                    ' Üres célmunkalap esetén a fejlécet is átmásoljuk
                    If isTargetEmpty Then
                        sourceWorksheet.Range(sourceWorksheet.Cells(1, 1), _
                                              sourceWorksheet.Cells(1, sourceLastColumn)).Copy _
                            Destination:=targetWorksheet.Cells(1, 1)
                                                       
                        ' A fejléc sormagasságának átvétele
                        targetWorksheet.Rows(1).RowHeight = sourceWorksheet.Rows(1).RowHeight
                    
                        ' Az oszlopszélességek átvétele
                        For columnIndex = 1 To sourceLastColumn
                            targetWorksheet.Columns(columnIndex).ColumnWidth = _
                                sourceWorksheet.Columns(columnIndex).ColumnWidth
                        Next columnIndex

                        targetLastRow = 1
                        isTargetEmpty = False
                        firstDataRow = 2
                    Else
                        firstDataRow = 2
                    End If

                    ' Adatsorok hozzáfûzése a célmunkalaphoz
                    If sourceLastRow >= firstDataRow Then
                        targetLastRow = targetLastRow + 1

                        sourceWorksheet.Range(sourceWorksheet.Cells(firstDataRow, 1), _
                                              sourceWorksheet.Cells(sourceLastRow, sourceLastColumn)).Copy _
                            Destination:=targetWorksheet.Cells(targetLastRow, 1)
                        
                        ' A sikeresen átmásolt adatsorok számlálása
                        processedRows = processedRows + (sourceLastRow - firstDataRow + 1)

                        targetLastRow = targetWorksheet.Cells(targetWorksheet.Rows.Count, "A").End(xlUp).Row
                    End If

                End If

                sourceWorkbook.Close SaveChanges:=False
                Set sourceWorkbook = Nothing
            End If

        End If

        ' Következõ fájl – pontosan egyszer léptetjük a Dir listát
        fileName = Dir

    Loop

    MsgBox "Az összefûzés sikeresen befejezõdött!" & vbCrLf & vbCrLf & _
       "Feldolgozott fájlok: " & Format(processedFiles, "#,##0") & vbCrLf & _
       "Összefûzött adatsorok: " & Format(processedRows, "#,##0") & vbCrLf & _
       "Idõ: " & Format(Timer - startTime, "0.0") & " másodperc", vbInformation
    
    Exit Sub

ErrorHandler:

    On Error Resume Next
    If Not sourceWorkbook Is Nothing Then sourceWorkbook.Close SaveChanges:=False
    On Error GoTo 0

    MsgBox "Hiba történt az összefûzés során:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical

End Sub

