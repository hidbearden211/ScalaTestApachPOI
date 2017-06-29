import java.io.{FileInputStream, InputStream}

import ch.qos.logback.classic.Logger
import org.apache.poi.hssf.usermodel.{HSSFSheet, HSSFWorkbook}
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import org.slf4j
import org.slf4j.LoggerFactory

import scala.collection.mutable.ListBuffer

object ReadData {
  val log: slf4j.Logger = LoggerFactory.getLogger(this.getClass.getName())

  def testRead(): Unit = {
    log.debug("testRead")
    try {
      val in: InputStream = new FileInputStream("C:\\xml\\out.xls")
      val wb: HSSFWorkbook = new HSSFWorkbook(new POIFSFileSystem(in))
      wb.setActiveSheet(0)
      val currentSheet = wb.getSheetAt(0)
      val row = currentSheet.getRow(0)
      val cell = row.getCell(0)
      print(s"We read first cell ${cell.getStringCellValue}")
      getFillingSheets(wb)
      wb.close()
      in.close()
    } catch {
      case e: Exception => println(e)
    }
  }

  def getFillingSheets(workBook: HSSFWorkbook): Seq[Int] = {
    log.debug(s"getFilingSheets with workbook sheets=${workBook.getNumberOfSheets}")
    val numbers: ListBuffer[Int] = new ListBuffer[Int]
    for (index <- 0 to workBook.getNumberOfSheets - 1) {
      log.debug(s"current index=${index}")
      if (checkFillingSheet(workBook.getSheetAt(index))) numbers += index
      println(numbers)
    }
    log.debug(s"Found number of filling sheets=${numbers.size}")
    numbers
  }

  def checkFillingSheet(sheet: HSSFSheet): Boolean = {
    log.debug(s"checkFillingSheet with sheet=${sheet.getSheetName}")
    for (i <- 0 to 20) {
      val row = sheet.getRow(i)
      if(row == null) return false
      val iter = row.cellIterator()
      while (iter.hasNext){
        val cell = iter.next()
        if (!cell.getStringCellValue.isEmpty) println("value find= " + cell.getStringCellValue)
        if (!cell.getStringCellValue.isEmpty) return true
      }
    }
    false
  }
}
