import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;


public class WriteToWord {
	String exampleWordPath = "example.docx";
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
					//System.out.print(cellText);
					//System.out.print(" ");
					
					//System.out.print("writeFlag = " + writeFlag);
					if(writeFlag) {
						List<XWPFParagraph> paras = cell.getParagraphs();
						////delete the original content
						//System.out.println(" paras.size = " + paras.size());
						for(int i = 0; i < paras.size(); i++) {
							cell.removeParagraph(i);
						}
						
						//add the new text
						XWPFParagraph para =cell.addParagraph();
				        XWPFRun run = para.createRun();
				        //run.setFontFamily("ËÎÌå");//×ÖÌå
				        //run.setFontSize(8);//×ÖÌå´óÐ¡
				        //run.setBold(true);//ÊÇ·ñ¼Ó´Ö
				        //run.setColor("FF0000");//×ÖÌåÑÕÉ«
				        para.setAlignment(ParagraphAlignment.CENTER);
				        run.setText(writeString);
				        //cell.setParagraph(para);
						writeFlag = false;
					}
					writeString = getWriteString(cellText, patient);
				}
				//System.out.println();
			}
			
			
			//write to file
			String outFileName = "output_" + patient.name + ".docx";
			OutputStream out = new FileOutputStream(outFileName);
			doc.write(out);
		    System.out.println("Finished!");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println("Failed! Please check the failed reason.");
			e.printStackTrace();
		}
	}
	
	private String getWriteString(String cellText, PatientInfo info) {
		writeFlag = true;
		if(cellText.contains("ËÍ¼ìµ¥Î»")) {
			return info.AcceptPart;
		}else if(cellText.contains("ÐÕÃû")) {
			return info.name;
		}else if(cellText.contains("ÐÔ±ð")) {
			return info.gender;
		}else if(cellText.contains("ÄêÁä")) {
			return info.age;
		}else if(cellText.contains("±ê±¾±àºÅ")) {
			return info.exmNum;
		}else if(cellText.contains("ËÍ¼ì¿ÆÊÒ")) {
			return info.partCell;
		}else if(cellText.contains("²¡´²ºÅ")) {
			return info.bedNo;
		}else if(cellText.contains("±ê±¾ÀàÐÍ")) {
			return info.sampleTyple;
		}else if(cellText.contains("²ÉÑùÈÕÆÚ")) {
			return info.date;
		}else if(cellText.contains("½ÓÊÕÈÕÆÚ")) {
			return info.acceptDate;
		}else if(cellText.contains("±ê±¾×´Ì¬")) {
			return info.sampleStatus;
		}else if(cellText.contains("ËÍ¼ìÒ½Éú")) {
			return info.Doctor;
		}else if(cellText.contains("ÁÙ´²Õï¶Ï")) {
			return info.clinicalCheck;
		}else if(cellText.contains("ÁÙ´²ÐÅÏ¢")) {
			return info.clinicalInfo;
		}else if(cellText.contains("Ó°ÏñÑ§½á¹û")) {
			return info.videographyResult;
		}else if(cellText.contains("ÇúÃ¹¾ú¿¹Ìå¿ìËÙ¼ì²â")) {
			return info.RuiQuSu_Antibody;
		}else if(cellText.contains("ÇúÃ¹¾ú¿¹Ô­¿ìËÙ¼ì²â")) {
			return info.RuiQuSu_Antigen;
		}else if(cellText.contains("ÇúÃ¹¾úºËËá¼ì²â")) {
			return info.RuiQuFen_QuMeiJun;
		}else if(cellText.contains("ÄîÖé¾úºËËá¼ì²â")) {
			return info.RuiQuFen_NianZhuJun;
		}else if(cellText.contains("Ò®ÊÏ·Îæß×Ó¾úºËËá¼ì²â")) {
			return info.RuiPuFen_YeShiFei;
		}else if(cellText.contains("Ã«Ã¹¾úºËËá¼ì²â")) {
			return info.RuiMaoFen_MaoMeiJun;
		}else if(cellText.contains("RCBIO")) {
			return info.RuiQuanFen_RCBIO;
		}else {
			writeFlag = false;
			return null;
		}
	}
}
