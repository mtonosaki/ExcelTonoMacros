# ExcelTonoMacros
Excel Keybord Macros / Utilities for your rapid operation.

## KeyMacros.bas

|  Function  |  Shortcut key  |  Description  |
| ---- | ---- | ---- |
|  MacroCSF  |  [CTRL]+[SHIFT]+[F]  |  ON/OFF Auto Filter at the selected row  |
|  MacroCSH  |  [CTRL]+[SHIFT]+[H]  |  Format selected cells to yyyy/mm/dd hh:mm:ss  |
|  MacroCSM  |  [CTRL]+[SHIFT]+[M]  |  Merge / Unmerge the selected cells  |
|  MacroCST  |  [CTRL]+[SHIFT]+[T]  |  Quick sheet auto format/arrage  |
|  MacroCSV  |  [CTRL]+[SHIFT]+[V]  |  Paste as value  |
|  MacroCSW  |  [CTRL]+[SHIFT]+[W]  |  ON/OFF freeze pane at the selected cell  |

## Funcs.bas

|  Function  |  F/S |  Description  |
| ---- | ---- | ---- |
|  ColorIndex(cell)  | F  | Returns color index number  |
|  MakeBoxCode  | F  | make three digits inch value (25.4,50.8,25.4,"A") -> "A001002001"  |
|  CopyTextToClipboard  | F  | Copy string to clipboard  |
|  GetTextFromClipboard  | F  | Get string from clipboard  |
|  CountStr  | F  | Count number of contains character |
|  Roundup24  | F  | Get round up value divisible by 24 |
|  Rounddown24  | F  | Get round down value divisible by 24 |
|  Roundup36  | F  | Get round up value divisible by 36 |
|  Rounddown36  | F  | Get round down value divisible by 36 |
|  Rounddown12348  | F  | Get round down value divisible by values 1,2,3,4 and 8 |
|  LineDuplication  | S  | Duplicate a row by the numerical value of the specified column |
|  GetSoloName  | S  | Make unique key |

F = Function / S = Subroutine

## How to install

1. Open excel file
2. Set visible "Developer" tab in excel ribbon
3. Click [Record Macro] button in the Developer tab
4. Select the combobox "Store macro in:" to "Personal Macro Workbook" 
5. Click [OK]
6. Click [Stop Recording] button in developer tab to make personal macro automatically.
7. Click [Visual Basic] button in the Developer tab
8. Open context menu on the "VBAProject(PERSONAL.XLSB)"
9. Import the ***.bas files



