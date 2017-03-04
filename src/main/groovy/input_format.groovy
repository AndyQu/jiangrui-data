import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory

public class EtlScript {
    def static dataSpecification = [
            [
                sheetName:"微通道干度计算",
                from :[
                    column: "B",
                    startRow: 12,
                    endRow: 12,
                ],
                to :[
                    column: "A",
                    row: 1,
                ]
            ],
            [
                sheetName: "微通道loss压降",
                from: [
                    column: "AD",
                    startRow: 25,
                    endRow: 25,
                ],
                to: [
                    column: "A",
                    row: 2,
                ]
            ],

            [
                sheetName: "微通道干度计算",
                from: [
                    column: "AS",
                    startRow: 2,
                    endRow: 11,
                ],
                to: [
                    column: "B",
                    row: 1,
                ]
            ],
            [
                sheetName: "微通道干度计算",
                from: [
                    column: "I",
                    startRow: 3,
                    endRow: 3,
                ],
                to:  [
                    column: "C",
                    row: 1,
                ]
            ],
            [
                sheetName: "微通道轴向压降",
                from:  [
                    column: "A",
                    startRow: 19,
                    endRow: 19,
                ],
                to:  [
                    column: "D",
                    row: 1,
                ]
            ],
            [
                sheetName: "微通道干度计算",
                from:  [
                    column: "AO",
                    startRow: 2,
                    endRow: 11,
                ],
                to:  [
                    column: "E",
                    row: 1,
                ]
            ],
            [
                sheetName: "微通道干度计算",
                from:  [
                    column: "AQ",
                    startRow: 2,
                    endRow: 11,
                ],
                to:  [
                    column: "F",
                    row: 1,
                ]
            ],
            [
                sheetName: "微通道干度计算",
                from:  [
                    column: "N",
                    startRow: 7,
                    endRow: 7,
                ],
                to:  [
                    column: "G",
                    row: 1,
                ]
            ],
    ]

    def static void extracFromFolder(String folderPath){
        File f = new File(folderPath)
        extractAndAssemble(f.listFiles().collect {it->it.absolutePath},dataSpecification,f.name+".xls")
    }

    def static void extractAndAssemble( inputExcelPaths, excelSpecifications, String outputExcelPath) {
        int startRow = 0;
//        File outFile = new File(outputExcelPath)
//        outFile.createNewFile()
//        Workbook outputExcel = WorkbookFactory.create(outFile)
        HSSFWorkbook outputExcel = new HSSFWorkbook()
        Sheet outputSheet = outputExcel.createSheet("汇总数据")
        inputExcelPaths.each {
            excelPath ->
                int rowNum = extractFromExcel(excelPath, excelSpecifications, outputSheet, startRow)
                startRow += (rowNum+1)
        }

        FileOutputStream fileOut = new FileOutputStream(outputExcelPath);
        outputExcel.write(fileOut);
        fileOut.close();
    }

    def static int extractFromExcel(String inputExcelPath, excelSpecifications, Sheet outputSheet, int startRow) {
        try {
            int maxRow = 0
            Workbook wb = WorkbookFactory.create(new File(inputExcelPath))
            excelSpecifications.each {
                specification ->
                    maxRow = Math.max(
                            extractFromSheet(wb.getSheet(specification.sheetName), specification, outputSheet, startRow),
                            maxRow)
            }
            return maxRow
        } catch (Exception ioe) {
            ioe.printStackTrace()
        }
    }

    def static int extractFromSheet(Sheet sheet, specification, Sheet outputExcel, int startRow) {
        def fromCol = computeColumnNum(specification.from.column)
        def toRowNum=0
        for (int r = specification.from.startRow; r <= specification.from.endRow; r++) {
            Row row = sheet.getRow(r - 1)
            double fromValue = row.getCell(fromCol - 1).getNumericCellValue()

            Row toRow = outputExcel.getRow(startRow + toRowNum + specification.to.row)
            if(toRow==null){
                toRow=outputExcel.createRow(startRow + toRowNum + specification.to.row)
            }
            def toCol = computeColumnNum(specification.to.column)
            Cell toCell = toRow.getCell(toCol-1)
            if(toCell==null){
                toCell=toRow.createCell(toCol-1)
            }
            toCell.setCellValue(fromValue)
            toRowNum++
        }
        return specification.from.endRow-specification.from.startRow+1
    }

    def static int computeColumnNum(String alphabetCol) {
        int column = 0;
        for (int i = 0; i < alphabetCol.size(); i++) {
            column = column * 26 + (alphabetCol.toLowerCase().charAt(i) - 'a'.charAt(0) + 1)
        }
        return column
    }
}
