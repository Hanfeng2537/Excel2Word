import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;


public class WriteToWord {
	String exampleWordPath = "example.docx";
	boolean writeFlag = false;
	public boolean writeSingleFile(String inputByUser, String outPutPath, PatientInfo patient) {
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
				        run.setFontFamily("΢���ź�");//����
				        run.setFontSize(12);//�����С
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
			
			//Set header
			CTSectPr sectPr = doc.getDocument().getBody().isSetSectPr()
							? doc.getDocument().getBody().getSectPr()
							: doc.getDocument().getBody().addNewSectPr();
			XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
			XWPFHeader header = headerFooterPolicy.getDefaultHeader();//.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
			//System.out.println(header.getText());
			
			List<XWPFTableRow> headerTableRows = header.getTables().get(1).getRows();
			for(XWPFTableRow headerTableRow: headerTableRows) {
				List<XWPFTableCell> headerTableCells = headerTableRow.getTableCells();
				//System.out.println("headerTableCells = " + headerTableCells.size());
				for(XWPFTableCell headerTablecell: headerTableCells) {
					String headerText = headerTablecell.getText();
					//System.out.println(headerText);
					List<XWPFParagraph> paras = headerTablecell.getParagraphs();
					for(int i = 0; i < paras.size(); i++) {
						headerTablecell.removeParagraph(i);
					}
					XWPFParagraph para =headerTablecell.addParagraph();
			        XWPFRun run = para.createRun();
			        run.setFontFamily("΢���ź�");//����
			        run.setFontSize(10.5);//
					if(headerText.contains("��������")) {
						run.setText(headerText + patient.name);
					}else if(headerText.contains("�걾���")) {
						run.setText(headerText + patient.exmNum); 
					}else if(headerText.contains("��������")) {
						run.setText(headerText + patient.acceptDate);
					}
					
				}
			}
			/*List<XWPFTable> headerTables = header.getTables();
			System.out.println("headerTables.size = " + headerTables.size());
			for(XWPFTable headerTable: headerTables) {
				List<XWPFTableRow> headerTableRows = headerTable.getRows();
				System.out.println("headerTableRows = " + headerTableRows.size());
				for(XWPFTableRow headerTableRow: headerTableRows) {
					List<XWPFTableCell> headerTableCells = headerTableRow.getTableCells();
					System.out.println("headerTableCells = " + headerTableCells.size());
					for(XWPFTableCell headerTablecell: headerTableCells) {
						System.out.println(headerTablecell.getText());
					}
				}
			}*/
	
			//write to file
			String newOutPutPath = outPutPath + "/��ⱨ��_" + patient.exmNum + "_" + patient.name;
			File file = new File(newOutPutPath);
			if(!file.exists()) {
				boolean mkdir_success = file.mkdir();
				if(!mkdir_success) {
					System.out.println("make directory failed. " + outPutPath);
					return false;
				}
			}
			
			String outFileName = newOutPutPath + "/��ⱨ��_" + patient.exmNum + "_" + patient.name + ".docx";
			OutputStream out = new FileOutputStream(outFileName);
			doc.write(out);
			out.close();
		    System.out.println("Finished!");
		    return true;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println("Failed! Please check the failed reason.");
			e.printStackTrace();
			return false;
		}
	}
	
	public boolean writeAllLineToFile(int lineNum, String outPutPath, PatientInfo patient) {
		try(XWPFDocument doc = new XWPFDocument(new FileInputStream(exampleWordPath))){	
			List<XWPFTable> tables = doc.getTables();
			XWPFTable table = tables.get(0);
			List<XWPFTableRow> rows = table.getRows();
			String writeString = "";
			for(XWPFTableRow row: rows) {
				List<XWPFTableCell> cells = row.getTableCells();
				for(XWPFTableCell cell: cells) {
					String cellText = cell.getText();
					if(writeFlag) {
						List<XWPFParagraph> paras = cell.getParagraphs();
						for(int i = 0; i < paras.size(); i++) {
							cell.removeParagraph(i);
						}
						
						//add the new text
						XWPFParagraph para =cell.addParagraph();
				        XWPFRun run = para.createRun();
				        run.setFontFamily("΢���ź�");//����
				        run.setFontSize(12);//�����С
				        //run.setBold(true);//�Ƿ�Ӵ�
				        //run.setColor("FF0000");//������ɫ
				        para.setAlignment(ParagraphAlignment.CENTER);
				        run.setText(writeString);
				        //cell.setParagraph(para);
						writeFlag = false;
					}
					writeString = getWriteString(cellText, patient);
				}
			}
			
			//Set header
			CTSectPr sectPr = doc.getDocument().getBody().isSetSectPr()
							? doc.getDocument().getBody().getSectPr()
							: doc.getDocument().getBody().addNewSectPr();
			XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(doc, sectPr);
			XWPFHeader header = headerFooterPolicy.getDefaultHeader();//.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
			//System.out.println(header.getText());
			
			List<XWPFTableRow> headerTableRows = header.getTables().get(1).getRows();
			for(XWPFTableRow headerTableRow: headerTableRows) {
				List<XWPFTableCell> headerTableCells = headerTableRow.getTableCells();
				//System.out.println("headerTableCells = " + headerTableCells.size());
				for(XWPFTableCell headerTablecell: headerTableCells) {
					String headerText = headerTablecell.getText();
					//System.out.println(headerText);
					List<XWPFParagraph> paras = headerTablecell.getParagraphs();
					for(int i = 0; i < paras.size(); i++) {
						headerTablecell.removeParagraph(i);
					}
					XWPFParagraph para =headerTablecell.addParagraph();
			        XWPFRun run = para.createRun();
			        run.setFontFamily("΢���ź�");//����
			        run.setFontSize(10.5);//
					if(headerText.contains("��������")) {
						run.setText(headerText + patient.name);
					}else if(headerText.contains("�걾���")) {
						run.setText(headerText + patient.exmNum); 
					}else if(headerText.contains("��������")) {
						run.setText(headerText + patient.acceptDate);
					}
					
				}
			}
			
			//write to file
			String newOutPutPath = outPutPath + "/AllPatients";
			File file = new File(newOutPutPath);
			if(!file.exists()) {
				boolean mkdir_success = file.mkdir();
				if(!mkdir_success) {
					System.out.println("make directory failed. " + outPutPath);
					return false;
				}
			}
			String outFileName = newOutPutPath + "/" + lineNum + "-���ɭ��ⱨ��-" + patient.name + "-" + patient.sampleTyple + ".docx";
			OutputStream out = new FileOutputStream(outFileName);
			doc.write(out);
			out.close();
			System.out.println("Finished!");
			return true;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println("Failed! Please check the failed reason.");
			e.printStackTrace();
			return false;
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
