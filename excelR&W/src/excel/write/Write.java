package excel.write;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Write {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		SimpleDateFormat format1 = new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss");

		String format_time1 = format1.format(System.currentTimeMillis());

		System.out.println(format_time1);

		String id = "ID";
		String pw = "Password";
		pw = pw.replace(",", "\n");
		
		HSSFWorkbook workbook = new HSSFWorkbook(); // 새 엑셀 생성
		HSSFSheet sheet = workbook.createSheet("시트명"); // 새 시트(Sheet) 생성
		HSSFRow row = sheet.createRow(0); // 엑셀의 행은 0번부터 시작
		HSSFCell cell = row.createCell(0); // 행의 셀은 0번부터 시작
		cell.setCellValue(id); // 생성한 셀에 데이터 삽입
		HSSFCell cell1 = row.createCell(1); // 행의 셀은 0번부터 시작
		cell1.setCellValue(pw);
		
		try {
			FileOutputStream fileoutputstream = new FileOutputStream("Test" + format_time1 + ".xlsx");
			workbook.write(fileoutputstream);
			fileoutputstream.close();
			System.out.println("엑셀파일생성성공");
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("엑셀파일생성실패");
		}

	}

}
