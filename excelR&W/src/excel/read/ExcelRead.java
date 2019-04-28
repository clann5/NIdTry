package excel.read;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) {
		// 참고 :
		// https://m.blog.naver.com/PostView.nhn?blogId=hyoun1202&logNo=220245067954&proxyReferer=https%3A%2F%2Fwww.google.co.kr%2F
		String excelReadPath = "C:/Users/user/git/NidQuest/excelR&W/Sample.xlsx";
//		String WritePath = "C:/Users/user/git/NidQuest/excelR&W/ExcelReadResult.xlsx";

//		.option("sheetName",getFirstSheetName(file)) // Required
//		.option("useHeader", "true") // Required
		
//		try {
//			
//			InputStream stream;
//		public static void verifyZipHeader(InputStream stream) throws NotOfficeXmlFileException, IOException {
//	        InputStream is = FileMagic.prepareToCheckMagic(stream);
//	        FileMagic fm = FileMagic.valueOf(is);
//
//	        switch (fm) {
//	        case OLE2:
//	            throw new OLE2NotOfficeXmlFileException(
//	                "The supplied data appears to be in the OLE2 Format. " +
//	                "You are calling the part of POI that deals with OOXML "+
//	                "(Office Open XML) Documents. You need to call a different " +
//	                "part of POI to process this data (eg HSSF instead of XSSF)");
//	        case XML:
//	            throw new NotOfficeXmlFileException(
//	                "The supplied data appears to be a raw XML file. " +
//	                "Formats such as Office 2003 XML are not supported");
//	        default:
//	        case OOXML:
//	        case UNKNOWN:
//	            break;
//	        }
//	    }
//		} catch(Exception e) {
//			System.out.println("지원하지 않는 Excel");
//			e.printStackTrace();
//		}
		
		try {
			
			OPCPackage opcPackage = OPCPackage.open(new File(excelReadPath));
			XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);
			
			int sheetNum = workbook.getNumberOfSheets();
			for (int i = 0; i < sheetNum; i++) {
				Sheet sheet = workbook.getSheetAt(i);
				
				System.out.println("Sheet Name : " + sheet.getSheetName() + "\r\n");
				System.out.println("Sheet :" + sheet + "\r\n" + "clannbr");

				Iterator<Row> rowIterator = sheet.iterator();
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();

					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {

						Cell cell = cellIterator.next();

						switch (cell.getCellType()) {
						case BOOLEAN:
							System.out.print(cell.getBooleanCellValue() + "\t\t");
							break;
						case NUMERIC:
							System.out.print(cell.getNumericCellValue() + "\t\t");
							break;
						case STRING:
							System.out.print(cell.getStringCellValue() + "\t\t");
							break;
						case FORMULA:
							System.out.print(cell.getCellFormula() + "\t\t");
							break;
						}
					}
					System.out.println("");
				}
			}
//			workbook.close();
			opcPackage.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch(InvalidFormatException e) {
			System.out.println("지원안함");
			e.printStackTrace();
		}
	}
}