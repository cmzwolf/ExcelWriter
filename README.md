## Synopsis

ExcelWriter is a simple tool for quickly creating excel file in Java. This is based on Apache POI and works as a wrapping between Java common collection objects and POI specific objects.

## Code Example

The content of the excel file to write is modelled with three java objects:

```java
List<String> columnsNames;
Map<String, String> columnsTypes;
Map<String, List<String>> columnsDataContent;
```
The _columnsNames_ List contains the names of the columns.
The _columnsTypes_ map contains the type of the colums (legal arguments are _number_,  _date_ and _String_) and the key of the map are the element contained into the _columnNames_ List. 
The _columnsDataContent_ map contains the data of each column. The key of the map are the element contained into the _columnNames_ List.

Example of usage are provided in the class _com.excel.writer.example.ExampleofUse_.

The creation of the excel content is straightforward starting from these three objects:
 ```java
// create the excel content.
ExcelCreator ec = new ExcelCreator(sheetName, columnsNames,
				columnsTypes, columnsDataContent, false);
// the boolean call argument is a flag for rounding or not the display of numeric values. True --> round numerics. False --> no rounding.		

// write the excel content into a file
ExcelFileWriter fl = new ExcelFileWriter(ec, "myExcelFile.xls");
fl.writeFile();		
```

## Motivation

I needed a quick factored way for creating excel files from common Java collection objects. This tiny library fit with my needs and I'm happy to share it. Maybe it can fit with your needs too.

## License

This software is published on GitHub with a GPL3 license. 