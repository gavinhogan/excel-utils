package com.gavinhogan.excel.utilities

import com.fasterxml.jackson.core.type.TypeReference
import com.fasterxml.jackson.databind.ObjectMapper
import org.apache.commons.lang3.StringUtils
import org.apache.commons.lang3.math.NumberUtils
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFTable
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo

class JSONArraytoExcelTableUtility {

    String xlsxTemplatePath

    int tableWidth = 0
    int rowInsertionPoint = 0
    List<String> headerLabels = []
    String topCorner = "A1"
    String bottomCorner = topCorner

    InputStream jsonInputStream
    OutputStream excelOutputStream

    List<String> templateHeaderLabels = []
    List<String> humanizedHeaderLabels = []

    void addRow(Map<String, Object> data,  XSSFSheet sheet, CTTable cttable){
        rowInsertionPoint++
        def row = sheet.createRow(rowInsertionPoint)
        def header = sheet.getRow(0)
        data.keySet().each{k->
            addColumn(k, sheet, cttable)
        }

        def headerIterator = header.cellIterator()

        while(headerIterator.hasNext()){
            def headerCell = headerIterator.next()
            String headerLabel = headerCell.getStringCellValue()

            def value = data.get(headerLabel)
            if(value){
                row.createCell(headerCell.columnIndex).setCellValue(value)
            }
        }

        CellReference newRef = new CellReference(rowInsertionPoint, tableWidth-1)
        bottomCorner = newRef.formatAsString()
    }

    void addColumn(String label,  XSSFSheet sheet, CTTable cttable){
        if(headerLabels.contains(label)){
            return
        }

        tableWidth++
        sheet.getRow(0).createCell(tableWidth -1  as Integer).setCellValue(label)
        headerLabels.add(label)
        bottomCorner = new CellReference(rowInsertionPoint, tableWidth-1).formatAsString()
    }

    void convert() {
        ObjectMapper mapper = new ObjectMapper()
        TypeReference<ArrayList<HashMap<String, String>>> typeRef = new TypeReference<ArrayList<HashMap<String, String>>>() {};
        def records = mapper.readValue(jsonInputStream, typeRef)
        convertToBook(records)
    }

    void sniffDataTypes(XSSFSheet sheet){
        tableWidth.times{ columnIndex->
            def rowIt = sheet.rowIterator()
            rowIt.next()
            boolean isNumber = true
            boolean isDate = true
            while(rowIt.hasNext()){
                def row = rowIt.next()
                def cell = row.getCell(columnIndex)
                isNumber = isNumber && (cell==null || isNumberOrCurrency(cell))
            }
            if(isNumber){
                rowIt = sheet.rowIterator()
                def header = rowIt.next()
                //println "Converting ${header.getCell(columnIndex)} to numeric data type"
                while(rowIt.hasNext()){
                    def row = rowIt.next()
                    def cell = row.getCell(columnIndex)
                    if(cell!=null){
                        cell.setCellValue(NumberUtils.createDouble(cell.stringCellValue))
                    }
                }
            }
        }
    }

    boolean isNumberOrCurrency(Cell cell) {
        String value = cell.stringCellValue
        NumberUtils.isCreatable(value) //|| (value.startsWith('$') && NumberUtils.isCreatable(value?.substring(1)))
    }

    void humanizeHeaders(XSSFSheet sheet){
        def cellIterator = sheet.getRow(0).cellIterator()
        while(cellIterator.hasNext()){
            def cell = cellIterator.next()
            def cellVal = cell.stringCellValue?.trim()
            String headerName
            if(cellVal.contains("_")){
                headerName = StringUtils.split(cellVal, "_").collect{StringUtils.capitalize(it)}.join(" ").trim()
            } else if(cellVal.contains(" ")){
                //do Nothing but capitalize
                headerName =  ( StringUtils.split(cellVal, " ").collect{StringUtils.capitalize(it)}.join(" ").trim() )
            } else {
                //assume camel case.
                headerName =  ( StringUtils.splitByCharacterTypeCamelCase(cellVal).collect{StringUtils.capitalize(it)}.join(" ").trim() )
            }
            //println "Humanized ${cellVal} to >$headerName<"
            humanizedHeaderLabels << headerName
            cell.setCellValue( headerName )
        }
    }

    void autoWiden(XSSFSheet sheet){
         tableWidth.times{
             sheet.autoSizeColumn(it)
         }
    }

    void convertToBook(Iterable<Map<String, Object>> recordProvider){

        XSSFWorkbook workbook = findOrCreateWorkbook()
        XSSFSheet sheet = workbook.createSheet("Sheet1")
        sheet.createRow(0)
        recordProvider.each {it->
            addRow(it, sheet, null)
        }
        humanizeHeaders(sheet)
        sortColumns(sheet)
        sniffDataTypes(sheet)
        autoWiden(sheet)
        decorateAsTable(sheet)
        writeWorkBookToFile(workbook)
    }

    private XSSFWorkbook findOrCreateWorkbook() {
        XSSFWorkbook workbook = new XSSFWorkbook()
        if (xlsxTemplatePath) {
            if  (new File(xlsxTemplatePath).exists()){
                def workbookStream = new FileInputStream(new File(xlsxTemplatePath))
                workbook = WorkbookFactory.create(workbookStream)
            } else if (Thread.currentThread().getContextClassLoader().getResource(xlsxTemplatePath)) {
                def workbookStream = Thread.currentThread().getContextClassLoader().getResourceAsStream("${xlsxTemplatePath}")
                workbook = WorkbookFactory.create(workbookStream)
            }
        }
        //Remove Sheet1
        if (workbook.getSheet("Sheet1")) {
           templateHeaderLabels = workbook.getSheet("Sheet1").getRow(0).cellIterator().collect {it.stringCellValue}
            workbook.removeSheetAt(workbook.getSheetIndex(workbook.getSheet("Sheet1")))

        }
        return workbook
    }

    private void writeWorkBookToFile(XSSFWorkbook workbook) {
        try {
            workbook.write(excelOutputStream);
        } catch (e) {
            throw e
        }
    }

    void decorateAsTable(XSSFSheet sheet) {

        XSSFTable table = sheet.createTable()
        CTTable cttable = table.getCTTable()
        cttable.setDisplayName("Table1")
        cttable.setId(1)
        cttable.setName("Table1")
        cttable.setRef("$topCorner:$bottomCorner")
        cttable.setTotalsRowShown(false)
        cttable.addNewAutoFilter()
        CTTableStyleInfo styleInfo = cttable.addNewTableStyleInfo()
        styleInfo.setName("TableStyleMedium2")
        styleInfo.setShowColumnStripes(false)
        styleInfo.setShowRowStripes(true)

        CTTableColumns columns = cttable.addNewTableColumns()
        columns.setCount(tableWidth)

        tableWidth.times{ i->
            CTTableColumn column = columns.addNewTableColumn()
            column.setId(i+1)
            column.setName("Column_${i+1}");
        }
    }

    static void main(String[] args) {
        new JSONArraytoExcelTableUtility(jsonInputStream: new FileInputStream("repos.json"), excelOutputStream: new FileOutputStream("new-repos.xlsx")).convert()
    }

    void sortColumns(XSSFSheet sheet) {
        def headersToMove = templateHeaderLabels.findAll {humanizedHeaderLabels.contains(it)}
        def moveCount = headersToMove.size()
        //println templateHeaderLabels
        //println humanizedHeaderLabels

        //println "We need to move $moveCount columns"
        sheet.rowIterator().each{row ->
            //Create the new cells
            moveCount.times{i->
                def c = row.createCell(headerLabels.size()+i);
                c.setCellValue("")
            }
            //Move everything over
            for(int i=headerLabels.size()-1; i>=0; i--){
                row.getCell(i+moveCount).setCellValue(
                        row.getCell(i).getStringCellValue()
                )
            }
            //Put the values at the front (left) of the range
            headersToMove.eachWithIndex{ String label, int i ->
                row.getCell(i).setCellValue(
                        row.getCell(humanizedHeaderLabels.indexOf(label)+moveCount).getStringCellValue()
                )
            }




        }
        headersToMove.eachWithIndex {headerLabel, loopIndex->
            //The index changes on each loop as we are removing columns as we go.

            List<String> currentColumnLabels = sheet.getRow(0).cellIterator().collect {it.stringCellValue}
            def candidateColumnLabels = currentColumnLabels.subList(moveCount, currentColumnLabels.size())

            int columnIndexToRemove = candidateColumnLabels.indexOf(headerLabel) + moveCount
            //print "Testing $candidateColumnLabels"
            //println ".  Removing [$columnIndexToRemove] ${sheet.getRow(0).getCell(columnIndexToRemove).stringCellValue} "
            deleteColumnAndShiftLeft(sheet, columnIndexToRemove)
        }
        def headers = (sheet.getRow(0).cellIterator().collect {it.stringCellValue} as Set)

        println humanizedHeaderLabels.findAll{!headers.contains(it)}

    }

    void deleteColumnAndShiftLeft(XSSFSheet sheet, int colIndex){
        def lastCell
        sheet.rowIterator().each { row->
            def cells = row.cellIterator().collect {it}
            cells.eachWithIndex { Cell entry, int i ->
                if(i>=colIndex){

                    Cell nextCell = row.getCell(i+1)
                    entry.setCellValue(nextCell?.stringCellValue)
                    lastCell = entry
                }

            }
            //Kill The Cell
            row.removeCell(lastCell)
        }

    }
}
