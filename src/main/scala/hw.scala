import java.io._
// import org.apache.poi.ss.util._
import java.util.Date;
import scala.collection.mutable.ArrayBuffer
import org.apache.poi.xssf.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel._

object Hi {
    def main(args: Array[String]) = {
        println("Hi!")
    }
}

object ReportTemplateWriter extends App {
    val wb: Workbook = new XSSFWorkbook() //or new HSSFWorkbook();
    val creationHelper: CreationHelper = wb.getCreationHelper()
    val sheet: Sheet = wb.createSheet("new sheet");

    // Create a row and put some cells in it. Rows are 0 based.
    val rows: ArrayBuffer[Row] = ArrayBuffer()
    val row: Row = sheet.createRow(0)
    rows += (row)
    rows += (sheet.createRow(1))
    rows += (sheet.createRow(2))
    rows += (sheet.createRow(3))
    rows += (sheet.createRow(4))
    rows += (sheet.createRow(5))
    rows += (sheet.createRow(6))
    rows += (sheet.createRow(7))
    rows += (sheet.createRow(8))
    rows += (sheet.createRow(9))
    // Create a cell and put a value in it.

    rows.map(row => {
        var cell: Cell = row.createCell(0)
        cell.setCellValue(1)

        //numeric value
        row.createCell(1).setCellValue(1.2)

        //plain string value
        row.createCell(2).setCellValue("This is a string cell")

        //rich text string
        val str: RichTextString = creationHelper.createRichTextString("Apache")
        val font: Font = wb.createFont()
        font.setItalic(true)
        font.setUnderline(Font.U_SINGLE)
        str.applyFont(font)
        row.createCell(3).setCellValue(str)

        //boolean value
        row.createCell(4).setCellValue(true)

        //formula
        row.createCell(5).setCellFormula("SUM(A1:B1)")

        //date
        val style: CellStyle = wb.createCellStyle()
        style.setDataFormat(creationHelper.createDataFormat().getFormat("m/d/yy h:mm"))
        cell = row.createCell(6)
        cell.setCellValue(new Date())
        cell.setCellStyle(style)

        //hyperlink
        row.createCell(7).setCellFormula("SUM(A1:B1)")
        cell.setCellFormula("HYPERLINK(\"http://google.com\",\"Google\")")
        row.createCell(8).setCellValue(1000)
        row.createCell(9).setCellValue(1000)
        row.createCell(10).setCellValue(1000)
        row.createCell(11).setCellValue(1000)
        row.createCell(12).setCellValue(1000)
        row.createCell(13).setCellValue(1000)
        row.createCell(14).setCellValue(1000)
        row.createCell(15).setCellValue(1000)
        row.createCell(16).setCellValue(1000)
        row.createCell(17).setCellValue(1000)
        row.createCell(18).setCellValue(1000)
        row.createCell(19).setCellValue(1000)
    })

    val ps = sheet.getPrintSetup
    ps.setFitHeight(1)
    ps.setFitWidth(1)
    ps.setPaperSize(PrintSetup.A4_PAPERSIZE)
    sheet.setFitToPage(true)

    // Write the output to a file
    val fileOut: FileOutputStream = new FileOutputStream("ooxml-cell.xlsx")
    wb.write(fileOut)
    fileOut.close()
}

// ReportTemplateWriter.main(Array())
