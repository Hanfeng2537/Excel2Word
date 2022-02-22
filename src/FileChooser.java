import java.awt.BorderLayout;
import java.awt.Container;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

public class FileChooser extends JFrame implements ActionListener {

	public boolean executeAvai = false;
	
	String chooseFileName = null;
	JButton start=null, open = null, choose = null;;
	JTextField filePath = null, outPath = null;
	
    public FileChooser(){
    	this.setSize(800, 600);
    	this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    	
    	JPanel panel = new JPanel();
    	panel.setLayout(null);
    	
    	filePath = new JTextField("Excel file path");
    	filePath.setBounds(10,20,200,25);
    	panel.add(filePath);
    	
    	open = new JButton("Open");
    	open.setBounds(220,20,165,25);
        panel.add(open);
        
        outPath = new JTextField("Out file path");
        outPath.setBounds(10,50,200,25);
        panel.add(outPath);
        
        choose = new JButton("Choose save folder");
        choose.setBounds(220,50,165,25);
        panel.add(choose);
        
        start = new JButton("Start");
        start.setBounds(300, 80, 165, 25);
        panel.add(start);
        
        open.addActionListener(this);
        
        this.add(panel);
        this.setVisible(true);
    }  
    @Override  
    public void actionPerformed(ActionEvent e) {  
        // TODO Auto-generated method stub 
    	switch(e.getID()) {
    		
    	}
        JFileChooser jfc=new JFileChooser();  
        jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );  
        jfc.showDialog(new JLabel(), "选择");  
        File file=jfc.getSelectedFile();  
        if(file.isDirectory()){  
            System.out.println("文件夹:"+file.getAbsolutePath());
            executeAvai = false;
            JOptionPane.showMessageDialog(jfc, "必须选择一个Excel文件", "Error!!!", JOptionPane.ERROR_MESSAGE);
            
        }else if(file.isFile()){
            System.out.println("文件:"+file.getAbsolutePath());
            chooseFileName = jfc.getSelectedFile().getName();
            String suffix = chooseFileName.substring(chooseFileName.lastIndexOf("."));
            System.out.println(suffix);
            if(!suffix.equals(".xlsx")) {
            	executeAvai = false;
            	JOptionPane.showMessageDialog(jfc, "必须选择一个Excel文件", "Error!!!", JOptionPane.ERROR_MESSAGE);
            }
            executeAvai = true;
            System.out.println(chooseFileName);
            filePath.setText(chooseFileName);
        }  
    } 
	
}
