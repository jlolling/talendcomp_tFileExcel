# Talend User Component suite tFileExcel
These components are dedicated to work with Excel in Talend in a very comfortable and powerful way.
There are a lot of features e.g: 
* support for conditional formatting, 
* data range validation, 
* formula support and 
* building new excel workbooks based on templates

List components:
| Component                      | Purpose                                                                                                                  |
|--------------------------------|--------------------------------------------------------------------------------------------------------------------------|
| tFileExcelWorkobookOpen        | Central component to hold the excel workbook. Is always needed.                                                          |
| tFileExcelWorkbookSave         | Writes a workbook back to the local file system, only needed of the excel file is written.                               |
| tFileExcelSheetList            | List all sheets in the workbook with filter                                                                              |
| tFileExcelSheetInput           | Reads data from a sheet and allows to configure the columns according to a header line                                   |
| tFileExcelSheetInputUnpivot    | Base on the outgoing flow of the tFileExcelSheetInput the component can unpivot (or normalize) the data in a generic way |
| tFileExcelSheetOutput          | Create/copy a sheet and writes data into. Have lots of functions to reuse existing formats.                              |
| tFileExcelSheetNamedCellInput  | Iterates through the named cells and provide name and value and data type                                                |
| tFileExcelSheetNamedCellOutput | Writes into named cells                                                                                                  |
| tFileExcelReferencedCellInput  | Use an incoming flow to address cells and provide the values as output flow                                              |
| tFileExcelReferencedCellOutput | Writes cell values addressed directly with row+column information                                                        |

These components can be get from the Release section here on Github: [Download](https://github.com/jlolling/talendcomp_tFileExcel/releases)

These Componentes usually does not work together with the build-in Excel-components because the build-in Excel-components uses outdated libraries.
You have to choose in Talend job between the usage of the build-in Excel components or the these custome Excel components.

[Documentation for the input components](https://github.com/jlolling/talendcomp_tFileExcel/blob/master/doc/tFileExcelSheetInput.pdf)

[Documentation for the output components](https://github.com/jlolling/talendcomp_tFileExcel/blob/master/doc/tFileExcelSheetOutput.pdf)
