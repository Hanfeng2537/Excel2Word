
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadFromExcel {

	public ReadFromExcel() {
		// TODO Auto-generated constructor stub
	}

	public void read() {
		//File excelFile = new File("D:\\wangjie\\eclipse-workspace\\Excel2Word\\testpkg\\test.xlsx");
		try {
			InputStream excelFile = new FileInputStream("D:\\wangjie\\eclipse-workspace\\Excel2Word\\testpkg\\test1.xlsx");
			try (XSSFWorkbook xwb = new XSSFWorkbook(excelFile)) {
				//int count = xwb.getNumberOfSheets();
				XSSFSheet sheet = xwb.getSheetAt(0);
				for (Row row: sheet) {
					for (Cell cell: row) {
						switch (cell.getCellType()) {
							case STRING:
								System.out.print(cell.getRichStringCellValue().getString());
								System.out.print(" ");
								break;
							case NUMERIC:
								if (DateUtil.isCellDateFormatted(cell)) {
				                     System.out.print(String.valueOf(cell.getDateCellValue()));
				                 } else {
				                     System.out.print(cell.getNumericCellValue());
				                 }
								System.out.print(" ");
								break;
							case BOOLEAN:
								System.out.println(cell.getBooleanCellValue());
								System.out.print(" ");
							default:
								break;
						}
					}
					System.out.println();
				}
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
