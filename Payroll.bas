Attribute VB_Name = "Module1"

Sub Msg_Box()
'
    MsgBox Prompt:="Este Modulo no esta Trabajando Todavia", _
        Buttons:=vbOK, _
            Title:="Araujo's Mexican Grill"
                Sheets("Macros").Select
                    Range("A1").Select
End Sub

Sub Print_Sobres()
'

    'Mensaje antes de correr el macro.
    If MsgBox("Prepare the Envelopes, (aprox. 32 sobres)" & Chr(13) _
    & Chr(13) & "Desea Continuar?", vbOKCancel + vbQuestion, "Araujo's Mexican Grill Payroll 2019 ©") = vbCancel Then
    Exit Sub
    End If


    Sheets("Sobres").Select
        Range("C2").Select
            'Application.ScreenUpdating = False
                Range(Selection, Selection.End(xlDown)).Select
                    j = 1  ' contador array real
                        empleados = ActiveSheet.Range(Range("C2"), Range("C2").End(xlDown)).rows.Count

    Dim ListEmpleadosArray() As String
        ReDim ListEmpleadosArray(1 To empleados) As String  'sabiendo el numero de elementos a incluir, redimensionamos nuestra Matriz
            On Error Resume Next    'En caso de error, que continue
                'anadimos cada valor del rango definido de la Hoja1 como elemento de la Matriz
    
For i = 2 To empleados + 1  ' ****
   If Worksheets("Sobres").Cells(i, 3).Value <> 0 Then
        ListEmpleadosArray(j) = Worksheets("Sobres").Cells(i, 3).Value
            j = j + 1
                End If
                    Next i
    
For i = 1 To j - 1   ' ***
     Range("H28").Value = ListEmpleadosArray(i)
        ActiveSheet.PageSetup.PrintArea = "E22:l32"
            ActiveSheet.PrintOut
Next i
                        
Sheets("Macros").Select
Range("A1").Select

End Sub

Sub Senter_Pay_Stubs()
'
        Sheets("Senter Pay Stubs").Select
            ActiveSheet.PrintOut
                Application.Wait Now + TimeSerial(0, 0, 1)
    
    Sheets("Macros").Select
        Range("A1").Select
        
End Sub
Sub SenterStoryPayStubs_CopyFile()
'
    'Print Senter & Story Pay Stubs & Copy Payroll Sheet
        Call Senter_Pay_Stubs
            Call Story_Pay_Stubs
                Call Copy_Sheet
                
    'MsgBox Prompt:="Please don't forget to rename 43417.1948032407 Sheet", _
        Buttons:=vbOK, _
            Title:="Araujo's Mexcian Grill"
                Sheets("Macros").Select
                    Range("A1").Select
                                 
End Sub

Sub Story_Pay_Stubs()
'
' Story_Pay_Stubs Macro
'
'Story_Pay_Stubs Macro to print Story Pay Stubs
        Sheets("Story Pay Stubs").Select
            ActiveSheet.PrintOut
                Application.Wait Now + TimeSerial(0, 0, 1)
                
    Sheets("Macros").Select
        Range("A1").Select
        
End Sub

Sub Macro_Forma()
'
' Macro1 Macro
    
       
    'Mensaje antes de correr el macro.
    If MsgBox("Prepare Las Medias Hojas 8.5 x5.5, Por El Total de Empleados que Tengamos" & Chr(13) _
    & Chr(13) & "Desea Continuar?", vbOKCancel + vbQuestion, "Araujo's Mexican Grill Payroll 2019 ©") = vbCancel Then
    Exit Sub
    End If
     
        
    Sheets("Database").Select
    Range("C3").Select
    Range(Selection, Selection.End(xlDown)).Select
    j = 1  ' contador array real
    empleados = ActiveSheet.Range(Range("C3"), Range("C3").End(xlDown)).rows.Count

    Dim EmpleadosArray() As String
    'sabiendo el numero de elementos a incluir, redimensionamos nuestra Matriz
    ReDim EmpleadosArray(1 To 4, 1 To empleados) As String
    'En caso de error, que continue
    On Error Resume Next
    'anadimos cada valor del rango definido de la Hoja1 como elemento de la Matriz
    For i = 3 To empleados + 3
        If Worksheets("Database").Cells(i, 3).Value = "Active" Then
           
           EmpleadosArray(1, j) = Worksheets("Database").Cells(i, 7).Value
           EmpleadosArray(2, j) = Worksheets("Database").Cells(i, 8).Value
           EmpleadosArray(3, j) = Worksheets("Database").Cells(i, 9).Value
           EmpleadosArray(4, j) = Worksheets("Database").Cells(i, 10).Value
           
           j = j + 1
        End If
    Next i
    
    Sheets("Forma").Select
    'Application.ScreenUpdating = False
    For i = 1 To j - 1
        Range("C11").Value = EmpleadosArray(1, i)
        Range("E11").Value = EmpleadosArray(2, i)
        Range("K11").Value = EmpleadosArray(3, i)
        Range("L11").Value = EmpleadosArray(4, i)
        ActiveSheet.PageSetup.PrintArea = "B2:N33"
        ActiveSheet.PrintOut
        'ActiveSheet.preview
    Next i
    

    Sheets("Macros").Select
        Range("A1").Select
        
End Sub

Sub Print_ALL()
    'Call Msg_Box
        Call Print_Sobres
            Call Macro_Forma
                Call Senter_Pay_Stubs
                    Call Story_Pay_Stubs
                        Call Copy_Sheet

    MsgBox Prompt:="Please don't forget to rename 43417.1948032407 Sheet", _
        Buttons:=vbOK, _
            Title:="Araujo's Mexcian Grill"
                Sheets("Macros").Select
                    Range("A1").Select
                        
End Sub

Sub Test_forma_2()
MsgBox Prompt:="Prepare Las Medias Hojas 8.5 x5.5, Por El Total de Empleados que Tengamos", _
        Buttons:=vbOKCancel, _
            Title:="Araujo's Mexican Grill"
        
    Sheets("Database").Select
    Range("C3").Select
    Range(Selection, Selection.End(xlDown)).Select
    j = 1  ' contador array real
    empleados = ActiveSheet.Range(Range("C3"), Range("C3").End(xlDown)).rows.Count

    Dim EmpleadosArray() As String
    'sabiendo el numero de elementos a incluir, redimensionamos nuestra Matriz
    ReDim EmpleadosArray(1 To 4, 1 To empleados) As String
    'En caso de error, que continue
    On Error Resume Next
    'anadimos cada valor del rango definido de la Hoja1 como elemento de la Matriz
    For i = 3 To empleados + 3
        If Worksheets("Database").Cells(i, 3).Value = "Active" Then
           
           EmpleadosArray(1, j) = Worksheets("Database").Cells(i, 7).Value
           EmpleadosArray(2, j) = Worksheets("Database").Cells(i, 8).Value
           EmpleadosArray(3, j) = Worksheets("Database").Cells(i, 9).Value
           EmpleadosArray(4, j) = Worksheets("Database").Cells(i, 10).Value
           
           j = j + 1
        End If
    Next i
    
    Sheets("Forma").Select
        Range("B2:N33").Select
        
    ActiveSheet.PageSetup.PrintArea = "$B$2:$N$33"
        With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    ActiveSheet.PageSetup.PrintArea = "$B$2:$N$33"
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0)
        .RightMargin = Application.InchesToPoints(0)
        .TopMargin = Application.InchesToPoints(0)
        .BottomMargin = Application.InchesToPoints(0)
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperStatement     ' To change the papar size or sobre. (= xlPaperForm)
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    For i = 1 To j - 1
        Range("C11").Value = EmpleadosArray(1, i)
        Range("E11").Value = EmpleadosArray(2, i)
        Range("K11").Value = EmpleadosArray(3, i)
        Range("L11").Value = EmpleadosArray(4, i)
        ActiveSheet.PageSetup.PrintArea = "$B$2:$N$33"
        ActiveSheet.PrintOut
        'ActiveSheet.preview
    Next i
       
       Sheets("Macros").Select
        Range("A1").Select
        
End Sub


Sub AppClose()

    Application.Quit

End Sub

Sub Copy_Sheet()
'
    Dim name_date As String
    name_date = Format(Date, "mm-dd")
    
    For i = 1 To Worksheets.Count
    If Worksheets(i).name = name_date Then
        exists = True
        MsgBox "La Hoja " & name_date & " ya existe. Se borrara y creara una copia nueva. Gracias,bye!", vbCritical, "Araujo's Mexican Grill"
        
        On Error Resume Next
        Application.DisplayAlerts = False
        Worksheets(name_date).Delete
        
        
            Sheets("Payroll").Select
        Sheets("Payroll").Copy Before:=Sheets(18)
            Sheets("Payroll (2)").Select
                'Sheets("Payroll (2)").Name = Now() * 1
                Sheets("Payroll (2)").name = name_date
                    Range("J13").Select
    End If
Next i

If Not exists Then
        Sheets("Payroll").Select
        Sheets("Payroll").Copy Before:=Sheets(18)
            Sheets("Payroll (2)").Select
                'Sheets("Payroll (2)").Name = Now() * 1
                Sheets("Payroll (2)").name = name_date
                    Range("J13").Select
End If
    
    
    
    'MsgBox Prompt:="Please don't forget to rename 43417.1948032407 Sheet", _
     '   Buttons:=vbOK, _
      '      Title:="Araujo's Mexcian Grill"
                    
        Sheets("Macros").Select
            Range("A1").Select
    
End Sub
Sub Macro_changeName()
Attribute Macro_changeName.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro_changeName Macro
'

'
    Sheets("Forma").Select
    Range("K7:L8").Select
    Selection.Copy
    Range("T5").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
    Range("T5").Select
    
End Sub
Sub test_renamesheet()
Attribute test_renamesheet.VB_ProcData.VB_Invoke_Func = " \n14"
'
' macro23232 Macro
'

'
    Call Copy_Sheet
    
    Sheets("Forma").Select
    Range("T5").Select
    ActiveCell.FormulaR1C1 = "11/16/2018"
    Range("T5").Select
    Selection.Copy
    Range("T7").Select
    ActiveWindow.SmallScroll Down:=2
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("T7").Select
    Selection.Copy
    Sheets("Payroll (2)").Select
    Application.CutCopyMode = False
    Sheets("Payroll (2)").name = Now() * 1
    Sheets("Forma").Select
    Range("T4:T8").Select
    Selection.ClearContents
    Range("V18").Select
    Sheets("Macros").Select
    Range("A1").Select
End Sub


Sub Process_Payroll()

' EL PROPOSITO DE ESTE MACRO ES LLENAR LOS VALORES DE LAS HORAS A PAGAR AUTOMATICAMENTE. AL MISMO TIEMPO SE SACARA UN TOTAL A PAGAR Y CUANTO RETIRAR DE LA CADA CAJA.
    
    'Seleccionar el archivo. este codigo lo baje de internet, funciona. Eso es lo importante. Usa la funcion de hasta abajo, NO BORRAR
    'Select files in Mac Excel with the format that you want
    'Working in Mac Excel 2011 and 2016x
    
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String
    Dim OneFile As Boolean
    Dim FileFormat As String
    Dim r As Range, rows As Long, i As Long
    Dim lrow As Long
    Dim FolderPath As String
    Dim RootFolder As String
    Dim scriptstr As String
    Dim dir As String
    On Error Resume Next
    
    'Selecciona el tipo de archivos y deja seleccionar solamente un archivo
    FileFormat = "{""public.plain-text""}"
    OneFile = True
    On Error Resume Next
    MyPath = MacScript("return (path to desktop folder) as String")
    dir = ActiveWorkbook.Path

    'Building the applescript string, do not change this
    If Val(Application.Version) < 15 Then
        'This is Mac Excel 2011
        If OneFile = True Then
            MyScript = _
                "set theFile to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
                "return theFile"
        Else
            MyScript = _
                "set applescript's text item delimiters to {ASCII character 10} " & vbNewLine & _
                "set theFiles to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file or files"" default location alias """ & _
                MyPath & """ with multiple selections allowed) as string" & vbNewLine & _
                "set applescript's text item delimiters to """" " & vbNewLine & _
                "return theFiles"
        End If
    Else
        'This is Mac Excel 2016
        If OneFile = True Then
            MyScript = _
                "set theFile to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
                "return posix path of theFile"
        Else
            MyScript = _
                "set theFiles to (choose file of type" & _
                " " & FileFormat & " " & _
                "with prompt ""Please select a file or files"" default location alias """ & _
                MyPath & """ with multiple selections allowed)" & vbNewLine & _
                "set thePOSIXFiles to {}" & vbNewLine & _
                "repeat with aFile in theFiles" & vbNewLine & _
                "set end of thePOSIXFiles to POSIX path of aFile" & vbNewLine & _
                "end repeat" & vbNewLine & _
                "set {TID, text item delimiters} to {text item delimiters, ASCII character 10}" & vbNewLine & _
                "set thePOSIXFiles to thePOSIXFiles as text" & vbNewLine & _
                "set text item delimiters to TID" & vbNewLine & _
                "return thePOSIXFiles"
        End If
    End If
    
    'IF STATEMENT PARA VER QUE COMPUTADORA ESTA CORRIENDO EL MACRO
    
        'Corriendo desde la Compu de la Oficina
    If dir = "Mancintosh HD:Users:ElPaisaMac:Dropbox:Araujos Mexican Grill Files" Then
        MyFiles = "Mancintosh HD:Users:ElPaisaMac:Dropbox:Payroll:Activity Detail Export.txt"
        FolderPath = "Mancintosh HD:Users:ElPaisaMac:Dropbox:Payroll:ArchivePayroll"
        'MsgBox "Seleccione el archivo Activity Detail Export mas reciente de la carpeta Payroll", vbCritical, "Araujo's Mexican Grill Payroll 2019 ©"
        
        'Corriendo de la computadora de Frank
    ElseIf dir = "/Users/fjaraujo/Dropbox/!!!  !!2019 Araujos Mexican Grill/Cash Activities" Then
        MyFiles = "/Users/fjaraujo/Dropbox/Payroll/Activity Detail Export.txt"
        FolderPath = "/Users/fjaraujo/Dropbox/Payroll/Archive Payroll"
        
        'Corriendo de la computadora de Julio
    ElseIf dir = "/Users/J.Alvarez/Dropbox/Julio Alvarez" Then
        MyFiles = "/Users/J.Alvarez/Dropbox/Julio Alvarez/Activity Detail Export.txt"
        FolderPath = "/Users/J.Alvarez/Dropbox/Julio Alvaez/latest files"
        'Corriendo desde otra computadora
    Else
        'MyPath = MacScript("return (path to desktop folder) as String")
        'MsgBox "Seleccione el archivo Activity Detail Export mas reciente de la carpeta Payroll", vbCritical, "Araujo's Mexican Grill Payroll 2019 ©"
        ' si no es ninguna Computadora va dejar que seleccione el archivo.
        
        
    MyFiles = MacScript(MyScript)
       ' If MsgBox("Se ha seleccionado el archivo: " & Chr(13) & MyFiles & Chr(13) & "Es este el archivo correcto?", vbOKCancel + vbQuestion, "Araujo's Mexican Grill Payroll 2019 ©") = vbCancel Then
       '     MsgBox "Por Favor Volver a correr el Macro con el archivo correspondiente"
       '     Exit Sub
       ' End If
    End If
    
    On Error GoTo 0
    If FileExists(MyFiles) = False Then
        MsgBox "Please Provide the Correct Verification File in the Payroll Folder", vbCritical, "Araujo's Mexican Grill Payroll 2019 ©"
        Exit Sub
    End If
    
    
        'Mensaje antes de correr el macro.
    If MsgBox("Ya ha llenado los siguientes requisitos:  " & Chr(13) & Chr(13) _
    & Chr(200) & "Exportar horas mas recientes de Virtual Time Clock" & Chr(13) _
    & Chr(200) & "Actualizar las fechas de pago en la hoja de Payroll" & Chr(13) & Chr(13) _
    & "Desea Continuar?", vbOKCancel + vbQuestion, "Araujo's Mexican Grill Payroll 2019 ©") = vbCancel Then
    
    MsgBox "Por Favor Asegurarse de llenar los requisitos antes de procesar las horas", vbCritical, "Araujo's Mexican Grill Payroll 2019 ©"
    Exit Sub
    End If
    
    'DISPLAY MESSAGE BOX HERE TO TELL THE USER TO WAIT WHLE THERE IS A LOOP RUNNING.
      'MESSAGE BOX IN HERE
    ActiveWorkbook.Sheets("MSG").Visible = xlSheetVisible
    Worksheets("MSG").Activate
    
        On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Worksheets("PivotTable").Delete
    Worksheets("Data").Delete
    Worksheets("Sick Leave").Activate
    Sheets.Add(After:=Sheets(3)).name = "Data"
    
    'ABRIR EL DOCUMENTO PONER ALGO QUE SI FALLA QUE NO SE ABRIO EL DOCUMENTO...
     Workbooks.OpenText Filename:= _
        MyFiles, Origin:= _
        xlMacintosh, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
    
    'Cambiar el nombre y destruir el otro archivo & CHANGE NAME OF ACTIVESHEET TO DATA FOR THE PIVOTTABLE Y PREPARAR EL ARCHIVO
        'DELETE ROWS AND PREPARE FOR PIVOTTABLE
    ActiveWorkbook.SaveAs "Payroll.txt"
    ActiveSheet.name = "Data"
    
    'If FolderPath <> "" Then Kill (MyFiles)
    'Windows("Payroll.txt").Activate
    Set r = ActiveSheet.Range("A1:J2000") ' antes estaba hasta la J
    rows = r.rows.Count
    For i = rows To 1 Step (-1)
        If WorksheetFunction.CountA(r.rows(i)) = 0 Then r.rows(i).Delete
    Next
  
  'INSERT COLUMNS FOR THE EMPLOYEE NAMES
   Columns("D:D").Insert Shift:=xlToRight, _
      CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D1").Value = "Employee Name"
    
    'FIND LAST ROW AND CONCATENATE NAMES STILL NEED TO MAKE SURE NAMES MATCH TO THE DATABASE
    lrow = Range("A1").End(xlDown).Row
    For i = 2 To lrow
        Cells(i, 4) = Cells(i, 3) & ", " & Cells(i, 1)
    Next i
    
    Range("A:C").Delete
    LastRow = lrow
    Range("E:E").Delete
    Range("F:F").Delete
    Columns("F:F").Insert Shift:=xlToRight, _
      CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F1").Value = "Total Overtime Hours"
    For i = 2 To lrow
        Cells(i, 6) = Cells(i, 5).Value + Cells(i, 4).Value
    Next i
    'Copy Paste the data
    'Windows("Payroll.xlsm").Activate
    'Windows("Payroll.txt").Activate
    Range("A1:H10000").Select
    Selection.Copy
    
    If Val(Application.Version) < 15 Then Windows("Payroll.xlsm").Activate
    If Val(Application.Version) > 15 Then Stop
    
    Worksheets("Data").Activate
    Range("A1").Select
    ActiveSheet.Paste
    
    'CAMBIAR LAS LETRAS!!! – por ENE
    
    Columns("A:A").Select
    Selection.Replace What:="–", Replacement:="Ð", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
    
    'Guardar y Cerrar el Documento. Se Guarda como un excel spreadsheet con la fecha de hoy y se borra el payroll.txt original
    If Val(Application.Version) < 15 Then Windows("Payroll.txt").Activate
    If Val(Application.Version) > 15 Then Stop
    
    'Select a folder in which the Files will be saved.
       'SCRIPT PARA SELECCIONAR EL FOLDER
  '   'If Val(Application.Version) < 15 Then
  '     scriptstr = "(choose folder with prompt ""Select the folder""" & _
  '          " default location alias """ & RootFolder & """) as string"
  '  Else
  '       scriptstr = "return posix path of (choose folder with prompt ""Select the folder""" & _
  '           " default location alias """ & RootFolder & """) as string"
  '  End If
    
    'Si no se ha selecccionado un FolderPath desde antes se debe de seleccionar a mano.
  '  If FolderPath = "" Then
  '       RootFolder = MacScript("return (path to desktop folder) as String")
  '      MsgBox "Seleccione el Folder donde desea guardar el archivo con las horas procesadas", vbCritical, "Araujo's Mexican Grill Payroll 2019 ©"
  '      FolderPath = MacScript(scriptstr)
           'Double check del folder se
  '      If MsgBox("Se ha seleccionado el Folder: " & Chr(13) & FolderPath & Chr(13) & "Es este el folder correcto?", vbOKCancel + vbQuestion, "Araujo's Mexican Grill Payroll 2019 ©") = vbCancel Then
  '      MsgBox "Por Favor Volver a correr el Macro con el folder correspondiente"
        'ActiveWindow.Close
  '      Exit Sub
  '      End If
        'MsgBox "Navegue y Seleccione el Folder donde desea guardar el archivo con las horas procesadas", vbCritical, "Araujo's Mexican Grill Payroll 2019 ©"
  '  End If
    
    'Guardar los archivos.. no es 100% necesario
  '  On Error GoTo 0
  '  ActiveWorkbook.SaveAs (FolderPath & "Payroll " & Format(Now(), "DD-MMM-YYYY hh mm AMPM") & ".txt")
  '  ActiveWorkbook.SaveAs Filename:= _
  '      FolderPath & "Payroll " & Format(Now(), "DD-MMM-YYYY hh mm AMPM") _
  '      , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    'ActiveWindow.Close
    
    'Confirmacion de donde se Guardo el Archivo no es 100% necesario
    'If FolderPath <> "" Then
    '    MsgBox "Se ha guardado una copia del documento en el siguiente folder" & Chr(13) & FolderPath & Chr(13) & "Gracias, Bye!", vbCritical, "Araujo's Mexican Grill Payroll 2019 ©"
   ' End If
    
    'Windows("Payroll 2018.xlsm").Activate
    '
    'Continuar con la Pivot table para horas regulares
    'Windows("Data").Activate
    ActiveWindow.Close
    Kill "Payroll.txt"
    
    
    'CREATE PIVOTTABLE AND GET IT READY.. STORE BY STORE
    Windows("Payroll.xlsm").Activate
    Dim PSheet As Worksheet
    Dim DSheet As Worksheet
    Dim PCache As PivotCache
    Dim PTable As PivotTable
    Dim PRange As Range
    Dim LastCol As Long

    'Insert a New Blank Worksheet and define Varaibles for PivotTable

On Error Resume Next
Worksheets("PivotTable").Delete
Worksheets("Data").Activate
Sheets.Add Before:=ActiveSheet
ActiveSheet.name = "PivotTable"
Set PSheet = Worksheets("PivotTable")
Set DSheet = Worksheets("Data")

'Define Data Range
Worksheets("Data").Activate
LastRow = Range("A1").End(xlDown).Row
LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Define Pivot Cache
Worksheets("PivotTable").Activate
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet.Cells(2, 2), _
TableName:="PayrollPivotTable")

'Insert Blank Pivot Table
Set PTable = PCache.CreatePivotTable _
(TableDestination:=PSheet.Cells(1, 1), TableName:="PayrollPivotTable")

'Insert Row Fields
With ActiveSheet.PivotTables("PayrollPivotTable").PivotFields("Employee Name")
.Orientation = xlRowField
.Position = 1
.name = "Employee Name"
End With


'IF YOU WANT TO ADD A SECOND ROW
'With ActiveSheet.PivotTables("PayrollPivotTable").PivotFields("Total")
'.Orientation = xlRowField
'.Position = 2
'End With

'Insert Column Fields
With ActiveSheet.PivotTables("PayrollPivotTable").PivotFields("Start Date")
.Orientation = xlColumnField
.Position = 1
.name = "Day of The Week"
End With

'Insert Data Field
With ActiveSheet.PivotTables("PayrollPivotTable").PivotFields("Total Paid Hours")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
'.NumberFormat = "#.##" THIS FORMAT DOES NOT WORK WELL. :(
.name = "WEEKLY HOURS WORKED "
End With


'Format Pivot

'ActiveSheet.PivotTables("PayrollPivotTable").TableStyle2 = "PivotStyleMedium9"
    ExecuteExcel4Macro "(1,""PayrollPivotTable"",4,TRUE)"
    ActiveSheet.PivotTables("PayrollPivotTable").TableStyle2 = "PivotStyleMedium2"

' VLOOKUP MONSTER!!!!!!

    'ExecuteExcel4Macro "(""PayrollPivotTable"","""",0,FALSE,TRUE)"
    'ExecuteExcel4Macro "(1,""PayrollPivotTable"",4,TRUE)"
   ''ActiveSheet.PivotTables("PayrollPivotTable").TableStyle2 = "PivotStyleLight2"
    'ActiveWindow.SmallScroll Down:=-41

' Popoulate Senter

Worksheets("Pre-Payroll").Activate
    Range("E4:R26").Select
    Selection.ClearContents
    Range("U4:U26").Select
    Selection.ClearContents
    Range("E31:R47").Select
    Selection.ClearContents
    Range("U31:U51").Select
    Selection.ClearContents
    
    Range("E4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],PivotTable!R3C2:R43C17,2,FALSE),0)"
    Range("F4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],PivotTable!R3C2:R43C17,3,FALSE),0)"
    Range("G4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],PivotTable!R3C2:R43C17,4,FALSE),0)"
    Range("H4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],PivotTable!R3C2:R43C17,5,FALSE),0)"
    Range("I4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],PivotTable!R3C2:R43C17,6,FALSE),0)"
    Range("J4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],PivotTable!R3C2:R43C17,7,FALSE),0)"
    Range("K4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-8],PivotTable!R3C2:R43C17,8,FALSE),0)"
    Range("L4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],PivotTable!R3C2:R43C17,9,FALSE),0)"
    Range("M4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-10],PivotTable!R3C2:R43C17,10,FALSE),0)"
    Range("N4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],PivotTable!R3C2:R43C17,11,FALSE),0)"
    Range("O4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],PivotTable!R3C2:R43C17,12,FALSE),0)"
    Range("P4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-13],PivotTable!R3C2:R43C17,13,FALSE),0)"
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-14],PivotTable!R3C2:R43C17,14,FALSE),0)"
    Range("R4").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],PivotTable!R3C2:R43C17,15,FALSE),0)"
    Range("E4:R4").Select
    Selection.AutoFill Destination:=Range("E4:R26"), Type:=xlFillValues

'Story
Worksheets("Pre-Payroll").Activate
    Range("E31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],PivotTable!R3C2:R43C17,2,FALSE),0)"
    Range("F31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],PivotTable!R3C2:R43C17,3,FALSE),0)"
    Range("G31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-4],PivotTable!R3C2:R43C17,4,FALSE),0)"
    Range("H31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5],PivotTable!R3C2:R43C17,5,FALSE),0)"
    Range("I31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],PivotTable!R3C2:R43C17,6,FALSE),0)"
    Range("J31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],PivotTable!R3C2:R43C17,7,FALSE),0)"
    Range("K31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-8],PivotTable!R3C2:R43C17,8,FALSE),0)"
    Range("L31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],PivotTable!R3C2:R43C17,9,FALSE),0)"
    Range("M31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-10],PivotTable!R3C2:R43C17,10,FALSE),0)"
    Range("N31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],PivotTable!R3C2:R43C17,11,FALSE),0)"
    Range("O31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],PivotTable!R3C2:R43C17,12,FALSE),0)"
    Range("P31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-13],PivotTable!R3C2:R43C17,13,FALSE),0)"
    Range("Q31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-14],PivotTable!R3C2:R43C17,14,FALSE),0)"
    Range("R31").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-15],PivotTable!R3C2:R43C17,15,FALSE),0)"
    Range("E31:R31").Select
    Selection.AutoFill Destination:=Range("E31:R47"), Type:=xlFillValues


'CREATE PIVOTTABLE AND GET IT READY.. STORE BY STORE
    Dim PSheet_OT As Worksheet
    Dim DSheet_OT As Worksheet
    Dim PCache_OT As PivotCache
    Dim PTable_OT As PivotTable
    Dim PRange_OT As Range
    'Dim LastRow As Long

    'Insert a New Blank Worksheet
On Error Resume Next
Worksheets("PivotTable_OT").Delete
Worksheets("PivotTable").Activate
Sheets.Add Before:=ActiveSheet
ActiveSheet.name = "PivotTable_OT"
Set PSheet_OT = Worksheets("PivotTable_OT")

'Define Data Range, but this should be the same as the last one
'Worksheets("Data").Activate
'LastRow = Range("A1").End(xlDown).Row
'LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column
'Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

'Define Pivot Cache
Set PCache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=PRange). _
CreatePivotTable(TableDestination:=PSheet_OT.Cells(2, 2), _
TableName:="PayrollPivotTable_OT")

'Insert Blank Pivot Table
Worksheets("PivotTable_OT").Activate
Set PTable_OT = PCache.CreatePivotTable _
(TableDestination:=PSheet_OT.Cells(1, 1), TableName:="PayrollPivotTable_OT")

'Insert Row Fields
With ActiveSheet.PivotTables("PayrollPivotTable_OT").PivotFields("Employee Name")
.Orientation = xlRowField
.Position = 1
.name = "Employee Name"
End With


'IF YOU WANT TO ADD A SECOND ROW
'With ActiveSheet.PivotTables("PayrollPivotTable").PivotFields("Total")
'.Orientation = xlRowField
'.Position = 2
'End With

'Insert Column Fields
'With ActiveSheet.PivotTables("PayrollPivotTable").PivotFields("Start Date")
'.Orientation = xlColumnField
'.Position = 1
'.Name = "Day of The Week"
'End With

'Insert Data Field
With ActiveSheet.PivotTables("PayrollPivotTable_OT").PivotFields("Total Overtime Hours")
.Orientation = xlDataField
.Position = 1
.Function = xlSum
'.NumberFormat = "#.##" THIS FORMAT DOES NOT WORK WELL. :(
.name = "TOTAL OVERTIME HOURS"
End With

'Format Pivot

'ActiveSheet.PivotTables("PayrollPivotTable").TableStyle2 = "PivotStyleMedium9"

' VLOOKUP MONSTER X 2

    'ExecuteExcel4Macro "(1,""PayrollPivotTable"",4,TRUE)"
    'ActiveSheet.PivotTables("PayrollPivotTable").TableStyle2 = "PivotStyleLight2"
    'ActiveWindow.SmallScroll Down:=-41

' Popoulate Senter

'ACA VA EL LOOKUP!!!

Worksheets("Pre-Payroll").Activate

' REHACER EL OT Y PONERLE el VLOOK BIEN CHIDO

    'SENTER
    Range("U4").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-18],PivotTable_OT!C[-19]:C[-18],2,FALSE),0)"
    Selection.AutoFill Destination:=Range("U4:U26"), Type:=xlFillValues
    
    'STORY
    Range("U31").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-18],PivotTable_OT!C[-19]:C[-18],2,FALSE),0)"
    Selection.AutoFill Destination:=Range("U31:U47"), Type:=xlFillValues
    
   
    'Poner los abonos de prestamos
    'Senter
    Range("Z4").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-23],Prestamos!R2C1:R50C4,3,FALSE),0)"
    Range("Z4").Select
    Selection.AutoFill Destination:=Range("Z4:Z29"), Type:=xlFillValues
    
    'Story
    Range("Z31").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-23],Prestamos!R1C1:R50C4,3,FALSE),0)"
    Range("Z32").Select
    Selection.AutoFill Destination:=Range("Z31:Z51"), Type:=xlFillValues
    Range("A1").Select
    
    'Hide Sheets
    ActiveWorkbook.Sheets("PivotTable_OT").Visible = xlSheetHidden
    ActiveWorkbook.Sheets("PivotTable").Visible = xlSheetHidden
    ActiveWorkbook.Sheets("Data").Visible = xlSheetHidden
    ActiveWorkbook.Sheets("MSG").Visible = xlSheetHidden
    
    
    'End of non Visibility
    Application.DisplayAlerts = True
    Worksheets("Pre-Payroll").Activate
    Range("A1").Select
    ActiveWindow.ScrollRow = 1
    
    
    'MsgBox "Revisar las horas de los que se les paga Salario", vbCritical

End Sub

Sub Verify()
   
    Dim MyPath As String
    Dim MyScript As String
    Dim MyFiles As String
    Dim OneFile As Boolean
    Dim FileFormat As String
    Dim r As Range, rows As Long, i As Long
    Dim lrow As Long
    Application.DisplayAlerts = False
    
    'prepare if Direct Lookup Fails
    FileFormat = "{""org.openxmlformats.spreadsheetml.sheet""}"
    OneFile = True
    On Error Resume Next
    Dim dir As String
    dir = ActiveWorkbook.Path
    
        'IF STATEMENT PARA VER QUE COMPUTADORA ESTA CORRIENDO EL MACRO
    
        'Corriendo desde la Compu de la Oficina
    If dir = "Mancintosh HD:Users:ElPaisaMac:Dropbox:Araujos Mexican Grill Files" Then
        MyPath = "Mancintosh HD:Users:ElPaisaMac:Dropbox:Payroll"
        MyFiles = "Mancintosh HD:Users:ElPaisaMac:Dropbox:Payroll:Payroll Summary.xlsx"
        'MyFiles = "Mancintosh HD:Users:ElPaisaMac:Dropbox:Payroll:Payroll Summary" & "filename*esy" & ".xlsx"
        
        'Corriendo de la computadora de Frank
    ElseIf dir = "/Users/fjaraujo/Dropbox/!!!  !!2019 Araujos Mexican Grill/Cash Activities" Then
        MyFiles = "/Users/fjaraujo/Dropbox/Payroll/Activity Detail Export.txt"
        FolderPath = "/Users/fjaraujo/Dropbox/Payroll/Archive Payroll"
        
                'Corriendo desde la Compu de Julio
   ElseIf dir = "/Users/J.Alvarez/Dropbox/Julio Alvarez/Nathan Payroll" Then
        MyPath = "/Users/J.Alvarez/Dropbox/Julio Alvarez/Nathan Payroll"
        MyFiles = "/Users/J.Alvarez/Dropbox/Julio Alvarez/Nathan Payroll/Payroll Summary.xlsx"
        'MyFiles = "Mancintosh HD:Users:ElPaisaMac:Dropbox:Payroll:Payroll Summary" & "filename*esy" & ".xlsx"
        
        'Corriendo desde otra computadora
    Else
       MsgBox "Please Run this macro from Frank or the Office Computer", vbCritical, "Araujo's Mexican Grill Payroll 2019 ©"
       Exit Sub
    End If
    
    On Error GoTo 0
    If FileExists(MyFiles) = False Then
        MsgBox "Please Provide the Correct Verification File in the Payroll Folder", vbCritical, "Araujo's Mexican Grill Payroll 2019 ©"
        Exit Sub
    End If
    
    'LImpiar donde va  a llegar la informacion
    Worksheets("Summary").Activate
    Range("A1:W10000").Select
    Selection.ClearContents
    
    
    'Abrir el Archivo y darle para adelante
     Workbooks.OpenText Filename:= _
        MyFiles, Origin:= _
        xlMacintosh, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
    'ActiveWorkbook.SaveAs "Verificacion.xlsx"
    
    'Sheets("Table 1").Select
    'Range("1:5").Select
    'Selection.Delete Shift:=xlUp
    'Cells.Select
    'Selection.UnMerge
    
    'paste rest of it
    'Range("B:B,C:C,D:D,E:E,F:F,H:H,I:I").Select
    'Selection.Delete Shift:=xlToLeft
    'Sheets("Table 2").Select
    'Range("1:2").Select
    'Selection.Delete Shift:=xlUp
    'Range("E:E").Select
    'Selection.UnMerge
    'Range("F:F").Select
    'Selection.Delete Shift:=xlToLeft
    'Range("A1:E50").Select
    'Selection.Copy
    'Sheets("Table 1").Select
    'Range("A1").End(xlDown).Offset(1).Select
    'Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
    '    False, Transpose:=False
      
      
  'CHANGE NAME OF ACTIVESHEET TO DATA FOR THE PIVOTTABLE
    Range("A1:E10000").Select
    Selection.Copy
    
    If Val(Application.Version) < 15 Then Windows("Payroll.xlsm").Activate
    If Val(Application.Version) > 15 Then Stop

    Range("A1").Select
    ActiveSheet.Paste
  
    'Guardar y Cerrar el Documento. Se Guarda como un excel spreadsheet con la fecha de hoy y se borra el payroll.txt original
    If Val(Application.Version) < 15 Then Windows("Payroll Summary.xlsx").Activate
    If Val(Application.Version) > 15 Then Stop
    
     '   On Error GoTo 0
    'foldepath = "Mancintosh HD:Users:ElPaisaMac:Dropbox:Payroll:Archive Payroll"
    
   ' ActiveWorkbook.SaveAs (FolderPath & "Verification " & Format(Now(), "DD-MMM-YYYY hh mm AMPM") & ".txt")
    'ActiveWorkbook.SaveAs Filename:= _
     '   FolderPath & "Payroll " & Format(Now(), "DD-MMM-YYYY hh mm AMPM") _
      '  , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    ActiveWindow.Close
    'Kill "Mancintosh HD:Users:ElPaisaMac:Dropbox:Payroll:Payroll Summary.xlsx"
    
    'Stop
    
    'Lookup hours in the Verify Sheet
    Worksheets("verify").Activate
    
        'Reg
    Range("E4").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],Summary!R1C1:R55C5,2,FALSE),-100)"
           
    'Overtime
    Range("F4").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-4],Summary!R1C1:R60C5,3,FALSE),-100)"
        
    'leave (sick)
    Range("G4").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-5],Summary!R1C1:R60C5,4,FALSE),-100)"
    
    'fill
    Range("E4:G4").Select
    Selection.AutoFill Destination:=Range("E4:G51"), Type:=xlFillValues
    Range("E4:G51").Select
    
    
    'clear contents on the Payroll sheet
    
    Sheets("Payroll").Select
    Range("E4:R23").Select
    Selection.ClearContents
    Range("U4:AA26").Select
    Selection.ClearContents
    Range("E32:R47").Select
    Selection.ClearContents
    Range("U32:AA51").Select
    Selection.ClearContents
    Sheets("verify").Select
    
    'start loop for regular
    Dim linea As Integer
    linea = 4
    
    
    For Each ncell In Range("H4:H51")
    If ncell.Value < 0.25 Then
    'paste the values over to Payroll
        Sheets("Pre-Payroll").Select
        Range("E" & linea & ":R" & linea & "").Select
        Selection.Copy
        Sheets("Payroll").Select
        Range("E" & linea).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
            
        linea = linea + 1
        Sheets("verify").Select
        
    Else
        'mark the lines so they wont be messy!
        linea = linea + 1
    End If
Next ncell

    'OT VALUES
    linea = 4

    For Each ncell In Range("I4:I51")
    If ncell.Value < 0.25 Then
    'paste the OT values
        Sheets("Pre-Payroll").Select
        Range("U" & linea).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Payroll").Select
        Range("U" & linea).Select
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=False
        
        linea = linea + 1
        Sheets("verify").Select
        
    Else
        'mark the lines so they wont be messy!
        linea = linea + 1
    End If
Next ncell

    ' Stop
    
    
    'Sick Days
    Dim linea2 As String
    linea2 = 1
    linea = 4
    Dim name As String
    
    For Each ncell In Range("G4:G51")
    If ncell.Value > 0 Then
    'paste the OT values
        Sheets("Payroll").Select
        Range("V" & linea).Value = ncell.Value
        Sheets("verify").Select
        name = Range("B" & linea).Value
        Sheets("sick").Select
        Range("A" & linea2).Value = name
        Range("B" & linea2).Value = ncell.Value
        Range("C" & linea2).Value = Now()
        linea = linea + 1
        linea2 = linea2 + 1
        Sheets("verify").Select
        
    Else
        'mark the lines so they wont be messy!
        linea = linea + 1
    End If
Next ncell

    'Poner las Fechas de los dias que se tomaron en el Database. Recordar de Mover a los empleados que ya no trabajan a la parte de abajo!!!!
    Sheets("Sick Leave").Select
        Range("R3").Select
    ActiveCell.FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-11],Sick!R1C1:R50C3,3,FALSE),"""")"
    Range("S3").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],Sick!R1C1:R50C3,2,FALSE),0)"
    Range("R3:S3").Select
    Selection.AutoFill Destination:=Range("R3:S36"), Type:=xlFillValues
    Range("R3:S36").Select
   
    For i = 3 To 36
    If Cells(i, 19).Value > 0 Then
        finalcol = 1 + Cells(i, 3).End(xlToRight).Column
        Cells(i, finalcol).Value = Cells(i, 18).Value
        Cells(i, finalcol + 1).Value = Cells(i, 19).Value
    End If
Next i


    'Poner y Actualizar los abonos de los trabajadores.
    
    If MsgBox("Process Employee Loans?" & Chr(13) _
    & Chr(13) & "Procesar los prestamos de empleados", vbOKCancel + vbQuestion, "Araujo's Mexican Grill Payroll 2019 ©") = vbCancel Then
        Exit Sub
    End If
    
          
    'Abono VALUES
    linea = 4
    Sheets("Pre-Payroll").Select
    For Each ncell In Range("Z4:Z51")
    If ncell.Value > 0 Then
    'paste the Prestamos values
        'Range("Z" & linea).Select
       ' Application.CutCopyMode = False
        'Selection.Copy
        'Sheets("Payroll").Select
        'Range("Z" & linea).Select
        'Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
         '   False, Transpose:=False
        'UPDATE the Prestamos Sheet
        name = Range("C" & linea).Value
        Sheets("Prestamos").Select
            'loop para encontrar a los nombres de las personas y ver si ya terminaron de pagar o cuanto les falta
                For i = 2 To 60
                    If Range("A" & i).Value = name Then
                        If Range("D" & i).Value + Range("C" & i).Value <= Range("B" & i) Then
                            Range("D" & i).Value = Range("D" & i).Value + Range("C" & i).Value
                           Exit For
                        Else: MsgBox name & " Ya termino de pagar su adelanto"
                        Range("B" & i).Value = ""
                        Range("D" & i).Value = ""
                        Exit For
                        End If
                    End If
                Next i
        linea = linea + 1
        Sheets("Pre-Payroll").Select
    Else
        'mark the lines so they wont be messy!
        linea = linea + 1
    End If
Next ncell


    ActiveWorkbook.Sheets("Summary").Visible = xlSheetHidden
    
End Sub




Function FileExists(FilePath As String) As Boolean
Dim TestStr As String
    TestStr = ""
    On Error Resume Next
    TestStr = dir(FilePath)
    On Error GoTo 0
    If TestStr = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'
Dim linea As Integer
Dim name As String

          
    'Abono VALUES
    linea = 4
    Sheets("Pre-Payroll").Select
    For Each ncell In Range("Z4:Z51")
    If ncell.Value > 0 Then
    'paste the Prestamos values
        'Range("Z" & linea).Select
       ' Application.CutCopyMode = False
        'Selection.Copy
        'Sheets("Payroll").Select
        'Range("Z" & linea).Select
        'Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
         '   False, Transpose:=False
        'UPDATE the Prestamos Sheet
        name = Range("C" & linea).Value
        Sheets("Prestamos").Select
            'loop para encontrar a los nombres de las personas y ver si ya terminaron de pagar o cuanto les falta
                For i = 2 To 60
                    If Range("A" & i).Value = name Then
                        If Range("D" & i).Value + Range("C" & i).Value <= Range("B" & i) Then
                            Range("D" & i).Value = Range("D" & i).Value + Range("C" & i).Value
                           Exit For
                        Else: MsgBox name & " Ya termino de pagar su adelanto"
                        Range("B" & i).Value = ""
                        Range("D" & i).Value = ""
                        Exit For
                        End If
                    End If
                Next i
        linea = linea + 1
        Sheets("Pre-Payroll").Select
    Else
        'mark the lines so they wont be messy!
        linea = linea + 1
    End If
Next ncell
End Sub


Sub try()


    Range("B:B,C:c,D:D,E:E,F:F,H:H,I:I").Select
    Selection.Delete Shift:=xlToLeft


End Sub
    Range("O59").Select
    Selection.ClearContents
    Range("F18").Select
    Selection.ClearContents


