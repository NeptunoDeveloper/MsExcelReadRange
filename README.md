# MsExcelReadRange
Obtiene una matriz de datos del un archivo excel, ideal para leer archivos Excel con formato personalizado.

Crear una nueva instancia de FileExcelApp

```csharp
FileExcelApp excel = new FileExcelApp(@"C:\repos\MsExcelUtility\PlantillaInformatica.xls");
```

Abrir el archivo excel y velidar que se haya creado la nueva instancia

```csharp
excel.Open();
if(excel.getLevel() != "OK")
{
  Console.WriteLine(excel.getMessage());
  Console.ReadKey();
  return;
}
```

Validar si existe la hoja Excel.

```csharp
string vWorksheetname = "Hoja1";

if (!excel.existsWorkSheetName(vWorksheetname))
{
    Console.WriteLine("La hoja no existe");
    return;
}
excel.setWorksheet(vWorksheetname);
```

Obtener la matriz de datos de la hoja excel desde la celda G6 hasta la celda J[N], donde [N] corresponde a la posición de la fila donde se encuentra el texto "3" en la columna "F"

```csharp
int rowStart = 6;
int rowFinish = excel.searchRowIndexInColRange("F", "F", "3");

if(rowFinish > -1)
{
    string begin = "G" + rowStart;
    string end = "J" + rowFinish;
    Console.WriteLine(begin);
    Console.WriteLine(end);
    string[][] values = excel.getCells(begin, end);
    foreach(string[] value in values)
    {
        Console.WriteLine(value);
    }

}
else
{
    Console.WriteLine("No se encontró un resultado de busqueda");
}
excel.Close();
Console.ReadKey();
```
