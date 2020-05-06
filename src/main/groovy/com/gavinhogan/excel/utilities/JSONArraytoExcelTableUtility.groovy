package com.gavinhogan.excel.utilities

import com.fasterxml.jackson.core.type.TypeReference
import com.fasterxml.jackson.databind.ObjectMapper
import org.apache.commons.lang3.StringUtils
import org.apache.commons.lang3.math.NumberUtils
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellReference
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
    Set<String> headerLabels = []
    String topCorner = "A1"
    String bottomCorner = topCorner

    InputStream jsonInputStream
    OutputStream excelOutputStream

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
                isNumber = isNumber && (cell==null || NumberUtils.isCreatable(cell.stringCellValue))
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

        sniffDataTypes(sheet)
        autoWiden(sheet)
        decorateAsTable(sheet)
        /*workbook.sheetIterator().each { XSSFSheet it->
            //it.setForceFormulaRecalculation(true)

            it.pivotTables.each {pivot->
                pivot.pivotCacheDefinition = new XSSFPivotCacheDefinition()
                pivot.pivotCacheDefinition.CTPivotCacheDefinition.refreshOnLoad
            }

        }*/
        //workbook.forceFormulaRecalculation = true
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
}
