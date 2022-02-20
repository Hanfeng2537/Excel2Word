import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

public class WriteToWord {
	String exampleWordPath = "testpkg\\out_example.docx";
	boolean writeFlag = false;
	public void write(PatientInfo patient) {
		try(XWPFDocument doc = new XWPFDocument(new FileInputStream(exampleWordPath));){
			List<XWPFTable> tables = doc.getTables();
			XWPFTable table = tables.get(0);
			List<XWPFTableRow> rows = table.getRows();
			int i = 0;
			for(XWPFTableRow row: rows) {
				//System.out.print("i = " + i);
				int j = 0;
				List<XWPFTableCell> cells = row.getTableCells();
				
				for(XWPFTableCell cell: cells) {
					String cellText = cell.getText();
					//System.out.print("j = " + j + ":");
					System.out.print(cellText);
					System.out.print(" ");
					
					
					
					System.out.print("writeFlag = " + writeFlag);
					if(writeFlag) {
						XWPFParagraph pIO2 =cell.addParagraph();
				        XWPFRun rIO2 = pIO2.createRun();
				        //IO2.setFontFamily("����");//����
				        //rIO2.setFontSize(8);//�����С
				        //rIO2.setBold(true);//�Ƿ�Ӵ�
				        //rIO2.setColor("FF0000");//������ɫ
				        rIO2.setText("����д�������");//
				         
						//cell.setText("test");
						writeFlag = false;
					}
					if(cellText.contains("�ͼ쵥λ")) {
						writeFlag = true;
					}
					
					//System.out.println(rh.getText(0));
					j++;
				}
				i++;
				System.out.println();
			}
			
			
		    
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
