import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xwpf.usermodel.IRunElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;


public class WriteToWord {
	String exampleWordPath = "testpkg/out_example.docx";
	boolean writeFlag = false;
	public void write(PatientInfo patient) {
		try(XWPFDocument doc = new XWPFDocument(new FileInputStream(exampleWordPath))){	
			List<XWPFTable> tables = doc.getTables();
			XWPFTable table = tables.get(0);
			List<XWPFTableRow> rows = table.getRows();
			String writeString = "";
			for(XWPFTableRow row: rows) {
				List<XWPFTableCell> cells = row.getTableCells();
				
				for(XWPFTableCell cell: cells) {
					String cellText = cell.getText();
					System.out.print(cellText);
					System.out.print(" ");
					
					System.out.print("writeFlag = " + writeFlag);
					if(writeFlag) {
						List<XWPFParagraph> paras = cell.getParagraphs();
						////delete the original content
						System.out.println(" paras.size = " + paras.size());
						for(int i = 0; i < paras.size(); i++) {
							cell.removeParagraph(i);
						}
						
						//add the new text
						XWPFParagraph para =cell.addParagraph();
						
				        XWPFRun run = para.createRun();
				        //run.setFontFamily("宋体");//字体
				        //run.setFontSize(8);//字体大小
				        //run.setBold(true);//是否加粗
				        //run.setColor("FF0000");//字体颜色
				        para.setAlignment(ParagraphAlignment.CENTER);
				        run.setText(writeString);
				        //cell.setParagraph(para);
						writeFlag = false;
					}
					if(cellText.contains("送检单位")) {
						writeFlag = true;
						writeString = patient.AcceptPart;
					}
				}
				System.out.println();
			}
			
			
			//write to file
			String outFileName = "output_" + patient.name + ".docx";
			OutputStream out = new FileOutputStream(outFileName);
			doc.write(out);
		    
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
