
import java.awt.Button;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromExcel {
	
	public final static int ERROR_NONE = 0;
	public final static int ERROR_NOLINE = 1;
	public final static int ERROR_FIRSTLINE_WRONG = 2;
	public final static int ERROR_PATIENT_NOT_FOUND = 3;
	public int errorType = 0;

	public ReadFromExcel() {
		// TODO Auto-generated constructor stub
	}

	public PatientInfo readByExmNum(String filePath, String exmNum, PatientInfo pi) {
		/*File directory = new File("");// �趨Ϊ��ǰ�ļ���
		String exampleFilePath = "";
		try {
			exampleFilePath = directory.getAbsolutePath();
			System.out.println(directory.getCanonicalPath());
			System.out.println(directory.getAbsolutePath());// ��ȡ����·��
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		String exampleFileName = "input.xlsx";
		try (InputStream excelFile = new FileInputStream(exampleFilePath + "/" + exampleFileName)) {*/
		try (InputStream excelFile = new FileInputStream(filePath)) {
			XSSFWorkbook xwb = new XSSFWorkbook(excelFile);
			// int count = xwb.getNumberOfSheets();
			XSSFSheet sheet = xwb.getSheetAt(0); // read the first sheet
			for (Row row : sheet) {
				for (Cell cell : row) {

					switch (cell.getCellType()) {
					case STRING: // to find the patient according to exmNum.
						if (cell.getRichStringCellValue().getString().equals(exmNum)) {
							int rowNo = row.getRowNum();
							System.out.println("find the parient info in line: " + rowNo);
							Row exampleRow = sheet.getRow(rowNo);
							getParientInfo(exampleRow, pi);
							errorType = ERROR_NONE;
							return pi;
						}
						// System.out.print(cell.getRichStringCellValue().getString());
						// System.out.print(" ");
						continue;
					case NUMERIC:
						/*
						 * if (DateUtil.isCellDateFormatted(cell)) {
						 * System.out.print(String.valueOf(cell.getDateCellValue())); } else {
						 * System.out.print(cell.getNumericCellValue()); } System.out.print(" ");
						 */
						break;
					case BOOLEAN:
						/*
						 * System.out.println(cell.getBooleanCellValue()); System.out.print(" ");
						 */
						break;
					default:
						break;
					}
					// System.out.println();
				}
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		errorType = ERROR_PATIENT_NOT_FOUND;
		System.out.println("Cound not find your patient....");
		System.out.println("Failed!");
		return null;
	}

	public PatientInfo readByLine(String filePath, int lineNumber, PatientInfo pi) {
		/*File directory = new File("");// �趨Ϊ��ǰ�ļ���
		String exampleFilePath = "";
		try {
			exampleFilePath = directory.getAbsolutePath();
			System.out.println(directory.getCanonicalPath());
			System.out.println(directory.getAbsolutePath());// ��ȡ����·��
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		String exampleFileName = "input.xlsx";
		try (InputStream excelFile = new FileInputStream(exampleFilePath + "/" + exampleFileName)) {*/
		try (InputStream excelFile = new FileInputStream(filePath)) {
			XSSFWorkbook xwb = new XSSFWorkbook(excelFile);
			// int count = xwb.getNumberOfSheets();
			XSSFSheet sheet = xwb.getSheetAt(0); // read the first sheet
			int rowCounts = sheet.getLastRowNum();
			System.out.println("row counts is: " + rowCounts + "; readLineNumber: " + lineNumber);
			if(lineNumber <= 0 || lineNumber > rowCounts) {
				errorType = ERROR_NOLINE;
				System.out.println("Cound not find your line number....");
				System.out.println("Failed! errorType is " + errorType);
				return null;
			}
			if(lineNumber == 1) {
				errorType = ERROR_FIRSTLINE_WRONG;
				System.out.println("The first line Number is invalid");
				System.out.println("Failed! errorType is " + errorType);
				return null;
			}
			Row row = sheet.getRow(lineNumber - 1);
			getParientInfo(row, pi);
			excelFile.close();
			errorType = ERROR_NONE;
			System.out.println(pi);
			return pi;
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		errorType = ERROR_NOLINE;
		System.out.println("Cound not find your lineNumber....");
		System.out.println("Failed!");
		return null;
	}
	
	public List<PatientInfo> readAllLines(String filePath, String exmNum, List<PatientInfo> patients) {
		try (InputStream excelFile = new FileInputStream(filePath)) {
			XSSFWorkbook xwb = new XSSFWorkbook(excelFile);
			// int count = xwb.getNumberOfSheets();
			XSSFSheet sheet = xwb.getSheetAt(0); // read the first sheet
			int firstLine = 0;
			for (Row row : sheet) {
				if(firstLine == 0) {
					firstLine++;
					continue;//skip the first line
				}
				PatientInfo pi = new PatientInfo();
				getParientInfo(row, pi);
				patients.add(pi);
			}
			errorType = ERROR_NONE;
			return patients;
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		errorType = ERROR_PATIENT_NOT_FOUND;
		System.out.println("Cound not find your patient....");
		System.out.println("Failed!");
		return null;
	}

	private void getParientInfo(Row row, PatientInfo pi) {
		int i = 0;
		for (Cell cell : row) {
			// System.out.print("#" + i);
			switch (i) {
			case 0: // �������� and �������ڣ� ��1��
				pi.acceptDate = getCellInfo(cell);
				pi.date = getCellInfo(cell);
				break;
			case 1: // �ͼ쵥λ, ��2��
				pi.AcceptPart = getCellInfo(cell);
				break;
			case 2: // �ͼ���ң���3��
				pi.partCell = getCellInfo(cell);
				break;
			case 3: // �����ţ���4��
				pi.bedNo = getCellInfo(cell);
				break;
			case 4: // �ͼ�ҽ������5��
				pi.Doctor = getCellInfo(cell);
				break;
			case 5: // ������ţ���6��
				pi.exmNum = getCellInfo(cell);
				break;
			case 6: // ���������� ��7��
				pi.name = getCellInfo(cell);
				break;
			case 7: // �����Ա𣬵�8��
				pi.gender = getCellInfo(cell);
				break;
			case 8: // �������䣬 ��9��
				pi.age = getCellInfo(cell);
				break;
			case 17: // �������ͣ� ��18��
				pi.sampleTyple = deleteSupplixForSampleType(getCellInfo(cell));
				break;
			case 25: // �ٴ���ϣ���26��
				pi.clinicalCheck = getCellInfo(cell);
				break;
			case 26: // �ٴ���Ϣ����27��
				pi.clinicalInfo = getCellInfo(cell);
				break;
			case 27: // Ӱ��ѧ�������28��
				pi.videographyResult = getCellInfo(cell);
				break;
			case 40: // ��ȫ��D -RCBIO seqTMDNA����41��
				pi.RuiQuanFen_RCBIO = getCellInfo(cell);
				break;
			case 42: // ������-��ù����ԭ���ټ�⣬��43��
				pi.RuiQuSu_Antigen = getCellInfo(cell);
				break;
			case 43: // ������L-��ù��������ټ�⣬��44��
				pi.RuiQuSu_Antibody = getCellInfo(cell);
				break;
			case 44: // ������-��ù�������⣬��45��
				pi.RuiQuFen_QuMeiJun = getCellInfo(cell);
				break;
			case 45: // ���շ�-Ү�Ϸ����Ӿ������⣬��46��
				pi.RuiPuFen_YeShiFei = getCellInfo(cell);
				break;
			case 46: // �����-����������⣬��47��
				pi.RuiQuFen_NianZhuJun = getCellInfo(cell);
				break;
			case 48: // ��ë��-ëù�������⣬��49��
				pi.RuiMaoFen_MaoMeiJun = getCellInfo(cell);
				break;
			default:
				break;
			}
			i++;
		}
	}

	private String getCellInfo(Cell cell) {
		String parameter = null;
		switch (cell.getCellType()) {
		case STRING:
			// System.out.print(cell.getRichStringCellValue().getString());
			// System.out.print(" ");
			parameter = cell.getRichStringCellValue().getString();
			// System.out.print(parameter + " ");
			break;
		case NUMERIC:
			String cellInfo;
			if (DateUtil.isCellDateFormatted(cell)) {
				cellInfo = formatDate(cell.getDateCellValue());
				// cellInfo = String.valueOf(cell.getDateCellValue());
			} else {
				//cellInfo = String.valueOf(cell.getNumericCellValue());
				//Delete the point for BedNo
				cellInfo = String.valueOf(Double.valueOf(cell.getNumericCellValue()).longValue() + "");
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
	
	private String deleteSupplixForSampleType(String sampleType) {
		if(sampleType == null || sampleType.isEmpty()) {
			return null;
		}else if(sampleType.contains("-")){
			return sampleType.substring(0, sampleType.lastIndexOf("-"));
		}else {
			return sampleType;
		}
	}
}
