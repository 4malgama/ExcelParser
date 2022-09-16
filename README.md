# ExcelParser
Excel parser for Qt Framework.
This parser uses the ActiveX control interface.
Before using it, I recommend to familiarize yourself with [COM objects](https://docs.microsoft.com/en-us/windows/win32/com/the-component-object-model) and [ActiveX](https://docs.microsoft.com/en-us/windows/win32/com/activex-controls).

# Using ExcelParser
Create an Excel object. To create it, just write:
> Excel object_name("path_to_excel_file.xlsx").

And then you can call methods on this object:
> object_name.method(arguments);

# Examples
**Reading the cell value from B3:**
> Excel excel("C:/dir_to_file/file.xlsx");
> qDebug() << excel.cellData("B3").toString();

**Search for the first cell in which there is the text "Tamara":**
> Excel excel("C:/dir_to_file/girlfriend.xlsx");
> qDebug() << "Tamara here: " << excel.findCellByText("Tamara"); // returns the QString format with the address

**Search for all cells that contain the word "Tamara":**
> Excel excel("C:/dir_to_file/girlfriend.xlsx");
> auto addresses = excel.findCellsByText("Tamara");
> for (auto i : addresses) { qDebug() << "Tamara is here: " << i; }

**Search for the word "Nikita" in the first table and "Tamara" in the second:**
> Excel excel("C:/dir_to_file/best_story.xlsx");
> QString n = excel.findCellByText("Nikita");
> excel.setTable(2);
> QString t = excel.findCellByText("Tamara");
> qDebug() << "Nikita is here: " n << "\nTamara is here: " << t;
