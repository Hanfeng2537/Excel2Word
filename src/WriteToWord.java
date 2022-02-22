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
				        //run.setFontFamily("����");//����
				        //run.setFontSize(8);//�����С
				        //run.setBold(true);//�Ƿ�Ӵ�
				        //run.setColor("FF0000");//������ɫ
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
		if(cellText.contains("�ͼ쵥λ")) {
			return info.AcceptPart;
		}else if(cellText.contains("����")) {
			return info.name;
		}else if(cellText.contains("�Ա�")) {
			return info.gender;
		}else if(cellText.contains("����")) {
			return info.age;
		}else if(cellText.contains("�걾���")) {
			return info.exmNum;
		}else if(cellText.contains("�ͼ����")) {
			return info.partCell;
		}else if(cellText.contains("������")) {
			return info.bedNo;
		}else if(cellText.contains("�걾����")) {
			return info.sampleTyple;
		}else if(cellText.contains("��������")) {
			return info.date;
		}else if(cellText.contains("��������")) {
			return info.acceptDate;
		}else if(cellText.contains("�걾״̬")) {
			return info.sampleStatus;
		}else if(cellText.contains("�ͼ�ҽ��")) {
			return info.Doctor;
		}else if(cellText.contains("�ٴ����")) {
			return info.clinicalCheck;
		}else if(cellText.contains("�ٴ���Ϣ")) {
			return info.clinicalInfo;
		}else if(cellText.contains("Ӱ��ѧ���")) {
			return info.videographyResult;
		}else if(cellText.contains("��ù��������ټ��")) {
			return info.RuiQuSu_Antibody;
		}else if(cellText.contains("��ù����ԭ���ټ��")) {
			return info.RuiQuSu_Antigen;
		}else if(cellText.contains("��ù��������")) {
			return info.RuiQuFen_QuMeiJun;
		}else if(cellText.contains("�����������")) {
			return info.RuiQuFen_NianZhuJun;
		}else if(cellText.contains("Ү�Ϸ����Ӿ�������")) {
			return info.RuiPuFen_YeShiFei;
		}else if(cellText.contains("ëù��������")) {
			return info.RuiMaoFen_MaoMeiJun;
		}else if(cellText.contains("RCBIO")) {
			return info.RuiQuanFen_RCBIO;
		}else {
			writeFlag = false;
			return null;
		}
	}
}
