import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JTextField;

public class FileChooser extends JFrame implements ActionListener {

	public final int SINGLE_LINE_TYPE = 0;
	public final int EXM_NUM_TYPE = 1;
	public final int ALL_LINE_TYPE = 2;

	public boolean executeAvai = false;

	JButton start, open, save;;
	JTextField filePath, outPath, inputString;
	JLabel inputStringLable;

	ButtonGroup convertTypes;
	JRadioButton singleLine_RB, exmNum_RB, allLine_RB;
	int convertType = SINGLE_LINE_TYPE;

	PatientInfo patient = new PatientInfo();
	List<PatientInfo> patients = new ArrayList<PatientInfo>();
	ReadFromExcel rfe = new ReadFromExcel();
	String inputByUser = null;
	WriteToWord wtw = new WriteToWord();
	String inputPath, outPutPath;

	public FileChooser() {

		File directory = new File("");
		String currentDirectory = directory.getAbsolutePath();

		this.setSize(800, 600);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		this.setLocation(300, 200);
		this.setTitle("Excel2Word");

		JPanel panel = new JPanel();
		panel.setLayout(null);

		convertTypes = new ButtonGroup();
		singleLine_RB = new JRadioButton("����", true);
		singleLine_RB.setBounds(10, 20, 100, 25);
		exmNum_RB = new JRadioButton("�������", false);
		exmNum_RB.setBounds(130, 20, 100, 25);
		allLine_RB = new JRadioButton("������", false);
		allLine_RB.setBounds(250, 20, 100, 25);
		convertTypes.add(singleLine_RB);
		convertTypes.add(exmNum_RB);
		convertTypes.add(allLine_RB);

		panel.add(singleLine_RB);
		panel.add(exmNum_RB);
		panel.add(allLine_RB);

		singleLine_RB.addActionListener(this);
		exmNum_RB.addActionListener(this);
		allLine_RB.addActionListener(this);

		inputStringLable = new JLabel("�������к�:");
		inputStringLable.setBounds(10, 50, 65, 25);
		panel.add(inputStringLable);
		inputString = new JTextField("2");
		inputString.setBounds(85, 50, 200, 25);
		panel.add(inputString);

		filePath = new JTextField();
		filePath.setBounds(10, 100, 600, 25);
		panel.add(filePath);

		open = new JButton("Open Excel file");
		open.setBounds(620, 100, 150, 25);
		panel.add(open);

		outPath = new JTextField(currentDirectory);
		outPath.setBounds(10, 150, 600, 25);
		panel.add(outPath);

		save = new JButton("Choose save folder");
		save.setBounds(620, 150, 150, 25);
		panel.add(save);

		start = new JButton("Start");
		start.setBounds(300, 200, 165, 25);
		panel.add(start);

		open.addActionListener(this);
		save.addActionListener(this);
		start.addActionListener(this);

		this.add(panel);
		this.setVisible(true);
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		// TODO Auto-generated method stub
		if (e.getSource() == open) {
			JFileChooser jfc = new JFileChooser();
			jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
			jfc.showDialog(new JLabel(), "ѡ��");
			File file = jfc.getSelectedFile();
			if (file.isDirectory()) {
				System.out.println("�ļ���:" + file.getAbsolutePath());
				executeAvai = false;
				JOptionPane.showMessageDialog(jfc, "����ѡ��һ��Excel�ļ�", "Error!!!", JOptionPane.ERROR_MESSAGE);

			} else if (file.isFile()) {
				// chooseFileName = jfc.getSelectedFile().getName();
				String chooseFileName = file.getAbsolutePath();
				System.out.println("�ļ�:" + chooseFileName);
				String suffix = chooseFileName.substring(chooseFileName.lastIndexOf("."));
				if (!suffix.equals(".xlsx")) {
					executeAvai = false;
					JOptionPane.showMessageDialog(jfc, "����ѡ��һ��Excel�ļ�", "Error!!!", JOptionPane.ERROR_MESSAGE);
				}
				executeAvai = true;
				filePath.setText(chooseFileName);
			}
		} else if (e.getSource() == save) {
			JFileChooser jfc = new JFileChooser();
			jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
			jfc.showDialog(new JLabel(), "ѡ��");
			File file = jfc.getSelectedFile();
			String saveDirectory = save.getText();
			if (file.isDirectory()) {
				saveDirectory = file.getAbsolutePath();
				System.out.println("�ļ���:" + saveDirectory);
				outPath.setText(saveDirectory);
				executeAvai = true;
				// JOptionPane.showMessageDialog(jfc, "����ѡ��һ��Excel�ļ�", "Error!!!",
				// JOptionPane.ERROR_MESSAGE);

			} else if (file.isFile()) {

				// chooseFileName = jfc.getSelectedFile().getName();
				String saveFileName = file.getAbsolutePath();
				System.out.println("�ļ�:" + saveFileName);
				String suffix = saveFileName.substring(saveFileName.lastIndexOf("."));
				System.out.println(suffix);
				if (!suffix.equals(".docx")) {
					executeAvai = false;
					JOptionPane.showMessageDialog(jfc, "����ѡ��һ��Docx�ļ�", "Error!!!", JOptionPane.ERROR_MESSAGE);
				} else {
					executeAvai = true;
					outPath.setText(saveFileName);
				}
			}
		} else if (e.getSource() == singleLine_RB) {
			inputStringLable.setBounds(10, 50, 65, 25);
			inputString.setBounds(85, 50, 200, 25);
			inputStringLable.setText("�������к�:");
			inputString.setText("2");
			inputString.setEditable(true);
			convertType = SINGLE_LINE_TYPE;

		} else if (e.getSource() == exmNum_RB) {
			inputStringLable.setBounds(10, 50, 95, 25);
			inputString.setBounds(115, 50, 200, 25);
			inputStringLable.setText("�������������:");
			inputString.setText("");
			inputString.setEditable(true);
			convertType = EXM_NUM_TYPE;
		} else if (e.getSource() == allLine_RB) {
			inputStringLable.setBackground(Color.YELLOW);
			inputString.setEditable(false);
			inputString.setText("�ù�����ʱδʵ��");
			convertType = ALL_LINE_TYPE;
		} else if (e.getSource() == start) {
			if (!startConvertion()) {
				System.out.println("startConvertion failed.");
			} else {
				switch (rfe.errorType) {
				case ReadFromExcel.ERROR_NONE:
					outPutPath = outPath.getText();
					boolean writeResult = false;
					if (convertType == ALL_LINE_TYPE) {
						for (int i = 0; i < patients.size(); i++) {
							writeResult = wtw.writeAllLineToFile(i, outPutPath, patients.get(i));
							if (!writeResult) {
								System.out.println("The " + i + " line was write failed.");
							}
						}
					} else {
						writeResult = wtw.writeSingleFile(inputByUser, outPutPath, patient);
						if (!writeResult) {
							System.out.println("The " + inputByUser + " was converted failed.");
						}
					}
					break;
				case ReadFromExcel.ERROR_NOLINE:
					showErrorNotification("��Ч�кţ������Excel��������Ч�к�!!!");
					break;
				case ReadFromExcel.ERROR_FIRSTLINE_WRONG:
					showErrorNotification("��һ��Ĭ��Ϊ�����У������������У���������Ч�к�!!!");
					break;
				case ReadFromExcel.ERROR_PATIENT_NOT_FOUND:
					showErrorNotification("û�з�����Ҫ�ҵ����ݣ�" + inputByUser);
					break;
				default:
					break;
				}
			}
		}

	}

	private boolean startConvertion() {
		if (executeAvai == false) {
			showErrorNotification("���ֲ�����������ϸ��飡����");
			return false;
		}
		inputPath = filePath.getText();
		outPutPath = outPath.getText();
		if (inputPath == null || inputPath.isEmpty()) {
			showErrorNotification("Excel�ļ�����Ϊ�գ�������ѡ��Ҫ��ȡ��Excel�ļ�������");
			return false;
		}
		if (outPutPath == null || outPutPath.isEmpty()) {
			showErrorNotification("���Ŀ¼����Ϊ�գ�������ѡ��Ҫ�����Ŀ¼������");
			return false;
		}

		System.out.println("convertType = " + convertType);

		switch (convertType) {
		case SINGLE_LINE_TYPE:
			inputByUser = inputString.getText();
			System.out.println("inputByUser = " + inputByUser);
			try {
				int inputNumber = Integer.parseInt(inputByUser);
				rfe.readByLine(inputPath, inputNumber, patient);
				return true;
			} catch (NumberFormatException e) {
				// TODO: handle exception
				showErrorNotification("�к�����������!!!");
				return false;
			}
		case EXM_NUM_TYPE:
			inputByUser = inputString.getText();
			System.out.println("inputByUser = " + inputByUser);
			rfe.readByExmNum(inputPath, inputByUser, patient);
			return true;
		case ALL_LINE_TYPE:
			if (!patients.isEmpty()) {
				patients.clear();
			}
			patients = rfe.readAllLines(inputPath, inputByUser, patients);
			return true;
		default:
			break;
		}
		showErrorNotification("δ֪����!!!");
		return false;
	}

	private void showErrorNotification(String msg) {
		JOptionPane.showMessageDialog(this, msg, "Error!!!", JOptionPane.ERROR_MESSAGE);
	}
}
