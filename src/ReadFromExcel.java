
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadFromExcel {
	
	public ReadFromExcel() {
		// TODO Auto-generated constructor stub
	}

	public PatientInfo read(String exmNum, PatientInfo pi) {
		//File excelFile = new File("D:\\wangjie\\eclipse-workspace\\Excel2Word\\testpkg\\test.xlsx");
		String exampleFilePath = "testpkg\\test1.xlsx";
			//InputStream excelFile = new FileInputStream("D:\\Excel2Word\\testpkg\\test1.xlsx");
			try (InputStream excelFile = new FileInputStream(exampleFilePath)) {
				XSSFWorkbook xwb = new XSSFWorkbook(excelFile);
				//int count = xwb.getNumberOfSheets();
				XSSFSheet sheet = xwb.getSheetAt(0);
				for (Row row: sheet) {
					for (Cell cell: row) {
						
						switch (cell.getCellType()) {
						case STRING:
							if(cell.getRichStringCellValue().getString().equals(exmNum)) { 
								int rowNo =row.getRowNum(); 
								System.out.println("find the parient info in line: " + rowNo);
								Row exampleRow = sheet.getRow(rowNo);
								getParientInfo(exampleRow, pi);
								return pi;
							}
							//System.out.print(cell.getRichStringCellValue().getString());
							//System.out.print(" ");
							continue;
						case NUMERIC:
							/*if (DateUtil.isCellDateFormatted(cell)) {
			                     System.out.print(String.valueOf(cell.getDateCellValue()));
			                 } else {
			                     System.out.print(cell.getNumericCellValue());
			                 }
							System.out.print(" ");*/
							break;
						case BOOLEAN:
							/*System.out.println(cell.getBooleanCellValue());
							System.out.print(" ");*/
							break;
						default:
							break;
					}
					//System.out.println();
				}
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		return null;
	}
	
	private void getParientInfo(Row row, PatientInfo pi) {
		int i = 0;
		for(Cell cell: row) {
			//System.out.print("#" + i);
			switch(i) {
			case 0: //�������� and �������ڣ� ��1��
				pi.acceptDate = getCellInfo(cell);
				pi.date = getCellInfo(cell);
				break;
			case 1: //�ͼ쵥λ, ��2��
				pi.AcceptPart = getCellInfo(cell);
				break;
			case 2: //�ͼ���ң���3��
				pi.partCell = getCellInfo(cell);
				break;
			case 3: //�����ţ���4��
				pi.bedNo = getCellInfo(cell);
				break;
			case 4: //�ͼ�ҽ������5��
				pi.Doctor = getCellInfo(cell);
				break;
			case 5: //������ţ���6��
				pi.exmNum = getCellInfo(cell);
				break;
			case 6: //���������� ��7��
				pi.name = getCellInfo(cell);
				break;
			case 7: //�����Ա𣬵�8��
				pi.gender = getCellInfo(cell);
				break;
			case 8: //�������䣬 ��9��
				pi.age = getCellInfo(cell);
				break;
			case 17: //�������ͣ� ��18��
				pi.sampleTyple = getCellInfo(cell);
				break;
			case 25: //�ٴ���ϣ���26��
				pi.clinicalCheck = getCellInfo(cell);
				break;
			case 26: //�ٴ���Ϣ����27��
				pi.clinicalInfo = getCellInfo(cell);
				break;
			case 27: //Ӱ��ѧ�������28��
				pi.videographyResult = getCellInfo(cell);
				break;
			case 40: //��ȫ��D -RCBIO seqTMDNA����41��
				pi.RuiQuanFen_RCBIO = getCellInfo(cell);
				break;
			case 42: //������-��ù����ԭ���ټ�⣬��43��
				pi.RuiQuSu_Antigen = getCellInfo(cell);
				break;
			case 43: //������L-��ù��������ټ�⣬��44��
				pi.RuiQuSu_Antibody = getCellInfo(cell);
				break;
			case 44: //������-��ù�������⣬��45��
				pi.RuiQuFen_QuMeiJun = getCellInfo(cell);
				break;
			case 45: //���շ�-Ү�Ϸ����Ӿ������⣬��46��
				pi.RuiPuFen_YeShiFei = getCellInfo(cell);
				break;
			case 46: //�����-����������⣬��47��
				pi.RuiQuFen_NianZhuJun = getCellInfo(cell);
				break;
			case 48: //��ë��-ëù�������⣬��49��
				pi.RuiMaoFen_MaoMeiJun = getCellInfo(cell);
				break;
			default:
					break;
			}
			i++ ; 
		}
	}
	private String getCellInfo(Cell cell) {
		String parameter = null;
		switch (cell.getCellType()) {
		case STRING:
			//System.out.print(cell.getRichStringCellValue().getString());
			//System.out.print(" ");
			parameter = cell.getRichStringCellValue().getString();
			//System.out.print(parameter + " ");
			break;
		case NUMERIC:
			String cellInfo;
			if (DateUtil.isCellDateFormatted(cell)) {
				cellInfo = formatDate(cell.getDateCellValue());
				//cellInfo = String.valueOf(cell.getDateCellValue());
                 
             } else {
            	 cellInfo = String.valueOf(cell.getNumericCellValue());
             }
			parameter = cellInfo;
			break;
		case BOOLEAN:
			parameter = String.valueOf(cell.getBooleanCellValue());
			break;
		case BLANK:
			parameter = "\\";
			break;
		case ERROR:
			parameter = String.valueOf(cell.getErrorCellValue());
			break;
		case FORMULA:
			parameter = cell.getCellFormula();
			break;
		default:
			System.out.print("Could not distinguish the cell type!");
			break;
		}
		
		return parameter;
	}
	
	private String formatDate(java.util.Date date) {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
		return sdf.format(date);
	}
}
