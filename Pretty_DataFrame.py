import pandas as pd
import numpy as np
import random
import os
import tempfile
import win32com.client


#Principal
def generate_excel_table(
        df_display,
        group_col=None,
        col_widths=30,
        synchronized_panels=False,
        output_name=r"./styled_table.csv"
    ):
    _, ext = os.path.splitext(output_name)
    if ext.lower() != '.csv':
        output_name = os.path.splitext(output_name)[0] + '.csv'
    for col in df_display.select_dtypes(include=['object']).columns:
        if df_display[col].apply(lambda x: isinstance(x, str) and x.isnumeric()).all():
            df_display[col] = df_display[col].astype(int)
    while True: 
        try: 
            # df_display.to_excel(output_name, engine='openpyxl')
            df_display.to_csv(output_name)
            print(f"Excel file saved to: {os.path.abspath(output_name)}")
            break 
        except: 
            input(f"Cierraa {os.path.abspath(output_name)} y Enter para reintentar...")
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    workbook = excel.Workbooks.Open(os.path.abspath(output_name))
    posiciones = sorted([df_display.columns.get_loc(col) for col in group_col])
    posiciones = [x + 2 for x in posiciones]

    timedelta_columns = df_display.dtypes[df_display.dtypes == 'datetime64[ns]'].index
    timedelta_positions = [df_display.columns.get_loc(col) for col in timedelta_columns]
    timedelta_positions = [x + 2 for x in timedelta_positions]
    for i in df_display.columns:
        if "_id" in i:
            try:
                df_display[i] = df_display[i].astype(int)
            except:
                print("error en", i)
    int_positions = df_display.dtypes[df_display.dtypes == 'int64'].index
    int_positions = [df_display.columns.get_loc(col) for col in int_positions]
    int_positions = [x + 2 for x in int_positions]
    vba_code = f"""
        Sub ColorearDuplicadosConMismoColor()
            Dim rng As Range
            Dim cell As Range
            Dim color As Long
            Dim dict As Object
            Dim headerRow As Long
            Dim columnas As Variant
            Dim columna As Variant
            Set dict = CreateObject("Scripting.Dictionary")
            
            ' Definir el número de la fila de encabezados (asumimos que es la primera fila)
            headerRow = 1
            
            ' Definir las columnas que queremos analizar ejemplo (columna 4 es D y columna 32 es AF)
            columnas = Array({str(posiciones)[1:-1]})
            
            ' Recorrer las columnas definidas en el arreglo
            For Each columna In columnas
                ' Definir el rango que contiene los datos en cada columna (comienza en la fila 2 para ignorar los encabezados)
                Set rng = Range(Cells(2, columna), Cells(Rows.Count, columna).End(xlUp))
                
                ' Recorrer todas las celdas del rango y asignar un color único a cada valor distinto
                Randomize
                For Each cell In rng
                    If Not dict.exists(cell.Value) And cell.Value <> "" Then
                        ' Asignar un color aleatorio para cada grupo de valores
                        color = RGB(Int((255 + 1) * Rnd), Int((255 + 1) * Rnd), Int((255 + 1) * Rnd))
                        dict.Add cell.Value, color ' Guardar el valor y el color asignado
                    End If
                Next cell
                
                ' Volver a recorrer el rango y aplicar el color del grupo a todas las celdas con el mismo valor
                For Each cell In rng
                    If dict.exists(cell.Value) Then
                        ' Asignar el color previamente asignado al grupo de valores
                        cell.Interior.color = dict(cell.Value)
                        ' Calcular el brillo del color de fondo
                        brightness = 0.299 * (dict(cell.Value) Mod 256) + 0.587 * ((dict(cell.Value) \ 256) Mod 256) + 0.114 * ((dict(cell.Value) \ 65536) Mod 256)
                        
                        ' Determinar el color del texto según el brillo del fondo
                        If brightness > 128 Then
                            textColor = RGB(0, 0, 0) ' Texto negro para fondos claros
                        Else
                            textColor = RGB(255, 255, 255) ' Texto blanco para fondos oscuros
                        End If
                        
                        ' Aplicar el color del texto
                        cell.Font.Color = textColor
                    End If
                Next cell
                
                ' Limpiar el diccionario para la siguiente columna
                dict.RemoveAll
            Next columna
            ' ---------------------------------------------------------------------------------
            ' ---------------------------------------------------------------------------------
            ' Copiar formato
            Dim i As Integer, ultimaColumna As Integer
            Dim rangoOrigen As Range, rangoDestino As Range
            
            ' Encontrar la última columna con datos en la hoja activa
            ultimaColumna = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
            
            ' 1. Copiar formato de la posición 0 del array a la columna 1 hasta la posición 0
            Set rangoOrigen = ActiveSheet.Columns(columnas(0))
            If columnas(0) > 1 Then
                Set rangoDestino = ActiveSheet.Range( _
                    ActiveSheet.Cells(1, 1), _
                    ActiveSheet.Cells(1, columnas(0) - 1)).EntireColumn
                rangoOrigen.Copy
                rangoDestino.PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            End If
            
            ' 2. Copiar formato de la última posición del array hasta la última columna con datos
            Set rangoOrigen = ActiveSheet.Columns(columnas(UBound(columnas)))
            If columnas(UBound(columnas)) < ultimaColumna Then
                Set rangoDestino = ActiveSheet.Range( _
                    ActiveSheet.Cells(1, columnas(UBound(columnas)) + 1), _
                    ActiveSheet.Cells(1, ultimaColumna)).EntireColumn
                rangoOrigen.Copy
                rangoDestino.PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            End If
            
            ' 3. Copiar formato desde columnas(n) hasta columnas(n+1)-1
            For i = 0 To UBound(columnas) - 1
                Set rangoOrigen = ActiveSheet.Columns(columnas(i))
                
                ' Solo si hay un espacio entre las columnas del array
                If columnas(i + 1) - columnas(i) > 1 Then
                    Set rangoDestino = ActiveSheet.Range( _
                        ActiveSheet.Cells(1, columnas(i) + 1), _
                        ActiveSheet.Cells(1, columnas(i + 1) - 1)).EntireColumn
                    rangoOrigen.Copy
                    rangoDestino.PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                End If
            Next i
            ' ---------------------------------------------------------------------------------
            ' ---------------------------------------------------------------------------------

            'Filtro en encabezados
            ActiveWindow.Panes(1).Activate
            Range("A1").Select
            Selection.AutoFilter
            ' ---------------------------------------------------------------------------------
            ' ---------------------------------------------------------------------------------
            Dim ws As Worksheet
            Dim y As Integer
            Dim col As Integer
            ' Definir el array con las columnas a las que aplicar el formato (por ejemplo, columna 1 = A, columna 2 = B, etc.)
            columnas = Array({str(timedelta_positions)[1:-1]}) ' Cambia estos números por los índices de las columnas que deseas formatear
            Set ws = ActiveSheet
            ' Recorrer las columnas en el array
            For y = LBound(columnas) To UBound(columnas)
                col = columnas(y)
                
                ' Aplicar el formato de fecha con hora a toda la columna
                ws.Columns(col).NumberFormat = "dd/mm/yyyy HH:mm:ss"
            Next y

            Dim y2 As Integer
            Dim col2 As Integer
            Dim columnas2 As Variant
            columnas2 = Array({str(int_positions)[1:-1]})
            For y2 = LBound(columnas2) To UBound(columnas2)
                col2 = columnas2(y2)
                ws.Columns(col2).NumberFormat = "0"
            Next y2

            ' ---------------------------------------------------------------------------------
            ' ---------------------------------------------------------------------------------
            Dim rango As Range
            Dim celda As Range
            Set rango = ActiveSheet.Range("A1").CurrentRegion
            For Each celda In rango
                If IsEmpty(celda.Value) Then
                    celda.Value = " "
                End If
            Next celda

    """

    if synchronized_panels:
        vba_code += f"""
            Dim primeraColumnaVisible As Integer
            Dim ultimaColumnaVisible As Integer
            Dim columnasVisibles As Integer
            Dim mitadColumnaVisible As Integer
            
            ' Obtener las columnas visibles en el rango de la ventana activa
            primeraColumnaVisible = ActiveWindow.VisibleRange.Columns(1).Column
            ultimaColumnaVisible = ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column
            
            ' Calcular la cantidad de columnas visibles
            columnasVisibles = ultimaColumnaVisible - primeraColumnaVisible + 1
            
            ' Calcular la columna en la mitad de las visibles
            mitadColumnaVisible = primeraColumnaVisible + (columnasVisibles \\ 2) - 1
            
            ' Seleccionamos la celda en la mitad de las columnas visibles en la fila 2
            Cells(2, mitadColumnaVisible).Select
            
            ' Configuramos el Split para que se divida en la columna de la mitad
            With ActiveWindow
                .SplitColumn = mitadColumnaVisible - 1 ' Columna donde se quiere hacer el split (0-indexed)
                .SplitRow = 1 ' Fila donde se hace el split
            End With
        """
    
    vba_code += """
    End Sub
    """
    """
    vba_code +=
    Option Explicit

    Private PrevRow As Long
    Private CellBackColors() As Long
    Private CellFontColors() As Long
    Private HighlightColor As Long
    Private HighlightFontColor As Long

    Sub Workbook_Open()
        ' Establecer el color de resaltado (puedes cambiarlo)
        HighlightColor = RGB(0, 0, 0) ' Color negro
        HighlightFontColor = RGB(255, 255, 255) ' Color azul oscuro para la letra
        
        ' Inicializar valores
        PrevRow = 0
        
        ' Configurar eventos
        Application.EnableEvents = True
    End Sub

    Private Sub Worksheet_SelectionChange(ByVal Target As Range)
        Application.EnableEvents = False
        
        ' Obtener la fila actual
        Dim CurrentRow As Long
        CurrentRow = Target.Row
        
        ' Si la fila anterior es diferente a la actual
        If PrevRow <> CurrentRow Then
            ' Restaurar colores de la fila anterior
            If PrevRow > 0 Then
                RestoreRowFormat (PrevRow)
            End If
            
            ' Guardar formato de la fila actual
            SaveRowFormat (CurrentRow)
            
            ' Resaltar la fila actual
            HighlightRow (CurrentRow)
            
            ' Actualizar la fila anterior
            PrevRow = CurrentRow
        End If
        
        Application.EnableEvents = True
    End Sub

    Private Sub SaveRowFormat(RowNum As Long)
        Dim LastCol As Long
        LastCol = ActiveSheet.UsedRange.Columns.Count
        
        ' Redimensionar los arrays
        ReDim CellBackColors(1 To LastCol)
        ReDim CellFontColors(1 To LastCol)
        
        ' Guardar los colores originales
        Dim i As Long
        For i = 1 To LastCol
            CellBackColors(i) = ActiveSheet.Cells(RowNum, i).Interior.color
            CellFontColors(i) = ActiveSheet.Cells(RowNum, i).Font.color
        Next i
    End Sub

    Private Sub RestoreRowFormat(RowNum As Long)
        Dim LastCol As Long
        LastCol = UBound(CellBackColors)
        
        ' Restaurar los colores originales
        Dim i As Long
        For i = 1 To LastCol
            ActiveSheet.Cells(RowNum, i).Interior.color = CellBackColors(i)
            ActiveSheet.Cells(RowNum, i).Font.color = CellFontColors(i)
        Next i
    End Sub

    Private Sub HighlightRow(RowNum As Long)
        Dim LastCol As Long
        LastCol = ActiveSheet.UsedRange.Columns.Count
        
        ' Resaltar la fila con el color elegido
        Dim i As Long
        For i = 1 To LastCol
            ActiveSheet.Cells(RowNum, i).Interior.color = HighlightColor
            ' Cambiar el color de la letra
            ActiveSheet.Cells(RowNum, i).Font.color = HighlightFontColor
        Next i
    End Sub
    """
    vba_module = workbook.VBProject.VBComponents.Add(1)
    vba_module.CodeModule.AddFromString(vba_code)
    excel.Application.Run("ColorearDuplicadosConMismoColor")
    workbook.VBProject.VBComponents.Remove(vba_module)

    # worksheet_module = workbook.VBProject.VBComponents("Sheet1")  # Asume que estás trabajando con Sheet1
    # worksheet_module.CodeModule.AddFromString(worksheet_code)
    # workbook.Save()
    # workbook.Close(SaveChanges=True)
    # excel.Quit()
    return



def print_colored_df(
        df,
        group_col=None,
        max_column_width=30,
        output_format="terminal",
        synchronized_panels=False,
        output_name = r"./styled_table.csv"
    ):
    df_display = df.copy()
    if len(df_display) == 0:
        print("Len df 0")
        return
    if isinstance(group_col, str):
        group_col = [group_col]
    group_col = [col for col in group_col if col in df_display.columns]
    if (group_col == None) or (len(group_col) == 0):
        group_col = df_display.columns[:1]

     # Handle excel output mode
    if output_format == "excel":
        excel_table = generate_excel_table(
            df_display,
            group_col,
            max_column_width,
            synchronized_panels,
            output_name
        )


# Modified example function to demonstrate synchronized panels
def example():
    # Create a test dataframe with some NaN values
    df = pd.DataFrame({
        'Group': ['A'] * 6 + ['B'] * 6 + ['C'] * 6,
        'col1': np.random.randint(0, 100, 18),
        'col2': [np.nan if i % 7 == 0 else np.random.randint(0, 100) for i in range(18)],
        'col3': ["This is a very long text that should be truncated" if i % 5 == 0 else i for i in range(18)],
        'special_col': ['X', 'Y', np.nan, 'V', 'Z', 'V', 'X', 'U', 'W', 'Y', 'Z', 'X', 'V', 'Y', 'Z', 'Z', 'Y', 'X'],
        'col4': np.random.randint(0, 100, 18),
        'col5': np.random.randint(0, 100, 18)
    })
    
    print_colored_df(df, group_col=["Group", 'special_col'], output_format='excel', synchronized_panels=True)


if __name__ == "__main__":
    example()

    # file_path = os.path.abspath(r'./styled_table.xlsx')
    # excel = win32com.client.Dispatch("Excel.Application")
    # excel.Visible = True
    # workbook = excel.Workbooks.Open(file_path)
    # vba_code = """
    # Sub MiMacro()
    #     MsgBox "¡Hola desde Python!"
    # End Sub
    # """
    # vba_module = workbook.VBProject.VBComponents.Add(1)
