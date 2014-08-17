import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.hssf.usermodel.HSSFSheet
import java.io.FileInputStream
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFCell
import java.util.Date
import java.text.SimpleDateFormat

class PoiRead {
  def poiRead = {
    try {
      val inputFile: POIFSFileSystem = new POIFSFileSystem(new FileInputStream("D:/poiTest.xls"))
      val workbook: HSSFWorkbook = new HSSFWorkbook(inputFile)
      val sheet: HSSFSheet = workbook.getSheet("出力シート名")
      var row: HSSFRow = sheet.getRow(0)
      var cell: HSSFCell = row.getCell(0)
      val a1ColumnString: String = cell.getStringCellValue()
      row = sheet.getRow(0)
      cell = row.getCell(1)
      val a2ColumnDate: Date = cell.getDateCellValue()
      row = sheet.getRow(0)
      cell = row.getCell(2)
      val a3ColumnDouble: Double = cell.getNumericCellValue()
      println("A1:" + a1ColumnString)
      println("A2:" + new SimpleDateFormat("yyyy/MM/dd").format(a2ColumnDate))
      println("A3:" + a3ColumnDouble)
    } catch {
      case e: Exception =>
        e.printStackTrace()
        println("処理失敗")
    }
  }
}