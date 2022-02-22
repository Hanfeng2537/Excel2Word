import java.awt.Button;
import java.awt.FlowLayout;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.Scanner;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.filechooser.FileNameExtensionFilter;

public class Main {

	private final static String NOTE = " Please copy your input Excel file to this folder, and the name must be \"input.xlsx\""
			+ "\n Please copy your docx sample this folder, the name must be \"example.docx\""
			+ "\n To get the patient information, you need to know the ������� that you need to search"
			+ "\n This soft will output the docx file to this folder accroding input ������� by you"
			+ "\n The output file name will be output_<patient name>.docx"
			+ "\n Now, please input the ������� that you need to search, and then press Enter key";

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		FileChooser fc = new FileChooser();
		
		if(fc.executeAvai) {
			PatientInfo patient = new PatientInfo();
			ReadFromExcel rfe = new ReadFromExcel();
			System.out.println(NOTE);
			Scanner reader = new Scanner(System.in);
			String sampleNo = reader.nextLine();
			//rfe.read("22S00001", patient);
			rfe.read(sampleNo, patient);
			System.out.println(patient);
			
			WriteToWord wtw = new WriteToWord();
			wtw.write(patient);
		}
	}
}
