# osgl-excel CHANGE LOG

1.10.1 - 04/Mar/2020
* ExcelWriter - when passing data of Iterable type, it renders an empty spreadsheet #21
* ExcelWriter - `IllegalArgumentException` when write a single object or empty list #20

1.10.0 - 02/Mar/2020
* update to osgl-tool 1.24.0
* ExcelWriter - set auto filter #19

1.9.0 - 02/Jan/2020
* update to osgl-tool 1.23.0
* ExcelReader - make `readCellValue` methods be static public accessible 

1.8.0 - 03/Nov/2019
* update to osgl-tool 1.21.0

1.7.2 - 30/Sep/2019
* fix logic issue in ExcelWriter.Builder.bigData

1.7.1 - 15/Sep/2019
* update to osgl-tool 1.20.1
* Error rendering xls file: The maximum number of cell styles was exceeded #18

1.7.0 - 21/Jul/2019
* update to osgl-tool 1.20.0
* ExcelWriter: Freeze top row by default #16
* Provide styling support #17

1.6.0 - 19/Apr/2019
* update to osgl-tool 1.19.2

1.5.0 - 30/Oct/2018
* Add osgl.ext.list file for IO integration
* Support write big data to Excel file #15
* Support specify cell data format by header name #14
* When writing `Map` delegate to `writeSheets` function #13
* When white list filter specified, it shall follow the order defined in the filter to output data #12
* Support writing to Excel #11
* Rename to osgl-excel #10
* `ExcelReader`: read sheets into Map indexed by sheet name #8
* Plugin Excel reading into `IO.read` framework #9
* error while reading multiple empty rows separated by data #7
* update to osgl-tool-1.18.0

1.4.0 - 14/Jun/2018
* update to osgl-tool-1.15.1

1.3.4 - 19/May/2018
* update osgl-tool to 1.13.1
* update poi to 3.17

1.3.3 - 02/Apr/2018
* update osgl-tool to 1.9.0

1.3.2 - 25/Mar/2018
* update osgl-tool to 1.8.1

1.3.1 - 25/Mar/2018
* update to osgl-tool-1.8.0
* update to osgl-storage-1.5.1

1.3.0
* Add shortcut read helper method that takes a URL as input #6
* NPE on `ExcelReader$TolerantLevel.onReadCellException` #5
* update to osgl-tool-1.6

1.2.0
* update to osgl-tool-1.5

1.1.0
* update maven build process
* apply osgl-bootstrap version mechanism

1.0.0 
* base line version
