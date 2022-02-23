
public class Main {

	private final static String NOTE = " Please copy your input Excel file to this folder, and the name must be \"input.xlsx\""
			+ "\n Please copy your docx sample this folder, the name must be \"example.docx\""
			+ "\n To get the patient information, you need to know the 样本编号 that you need to search"
			+ "\n This soft will output the docx file to this folder accroding input 样本编号 by you"
			+ "\n The output file name will be output_<patient name>.docx"
			+ "\n Now, please input the 样本编号 that you need to search, and then press Enter key";

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		FileChooser fc = new FileChooser();
		
		/*if(fc.executeAvai) {
			PatientInfo patient = new PatientInfo();
			ReadFromExcel rfe = new ReadFromExcel();
			System.out.println(NOTE);
			Scanner reader = new Scanner(System.in);
			String sampleNo = reader.nextLine();
			//rfe.read("22S00001", patient);
			rfe.readByExmNum(sampleNo, patient);
			System.out.println(patient);
			
			WriteToWord wtw = new WriteToWord();
			wtw.write(patient);
		}*/
		
		/*PatientInfo patient = new PatientInfo();
		ReadFromExcel rfe = new ReadFromExcel();
		rfe.readByLine(11, patient);
		WriteToWord wtw = new WriteToWord();
		wtw.write(patient);*/
	}
	
}
