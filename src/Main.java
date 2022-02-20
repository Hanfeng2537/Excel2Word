
public class Main {


	public static void main(String[] args) {
		// TODO Auto-generated method stub
		PatientInfo patient = new PatientInfo();
		ReadFromExcel rfe = new ReadFromExcel();
		rfe.read("22S00001", patient);
		System.out.println(patient);
		WriteToWord wtw = new WriteToWord();
		wtw.write(patient);
	}
}
