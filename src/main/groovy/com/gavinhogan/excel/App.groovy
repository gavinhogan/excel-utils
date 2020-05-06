package com.gavinhogan.excel

import com.gavinhogan.excel.utilities.JSONArraytoExcelTableUtility
import picocli.CommandLine

import java.util.concurrent.Callable

@CommandLine.Command(
        name = "App",
        mixinStandardHelpOptions = true,
        version = "JSON to Excel 1.0",
        description = "Converts a well formed JSON array to an Excel Table"
)
class App implements Callable<Integer>{

    @CommandLine.Parameters(index = "0", description = "The Source JSON File")
    private File jsonFile

    @CommandLine.Parameters(index = "1", description = "The Output Excel File", defaultValue = "out.xlsx")
    private File excelFile = new File("out.xlsx")

    @CommandLine.Option(names=["-t", "--template"], paramLabel = "xlsxTemplatePath")
    private String excelTemplate

    static void main(String... args) {
        int exitCode = new CommandLine(new App()).execute(args)
        System.exit(exitCode)
    }

    Integer call() {
        new JSONArraytoExcelTableUtility(
                jsonInputStream: new FileInputStream(jsonFile),
                excelOutputStream: new FileOutputStream(excelFile),
                xlsxTemplatePath: excelTemplate
        ).convert()
        return 0
    }
}
