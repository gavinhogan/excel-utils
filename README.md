# Excel Utils

## Introduction
Intended as a suite of simple composable utils for working with Excel, either from the command line or as a library. 

## Utils
**json-array-to-excel-table:** Given a well formed json array convert it to a well formed Excel table. Note, that this is not just an excel range but a properly formatted excel table.

## Instructions
./gradlew shadowJar

java -jar build/libs/excel-utils-all.jar [-hV] [-t=xlsxTemplatePath] <jsonFile> <excelFile>



## Contributions
Feel free to file a request for features or improvements. This is very much a part time effort so feel free to submit merge requests if you have the time to add the features you care about.