
public class PatientInfo {
	String AcceptPart;//送检单位, 第2列
	String name;//患者姓名， 第7列
	String gender; //患者性别，第8列
	String age; //患者年龄， 第9列
	String exmNum; //样本编号，第6列
	String partCell; //送检科室，第3列
	String bedNo; //病床号，第4列
	String sampleTyple; //样本类型， 第18列
	String date; //采样日期， 第1列
	String acceptDate; //接受日期， 第1列
	String sampleStatus = ""; //样本状态， 不用管
	String Doctor; //送检医生，第5列
	String clinicalCheck; //临床诊断，第26列
	String clinicalInfo; //临床信息，第27列
	String videographyResult; //影像学结果，第28列
	String RuiQuSu_Antibody; //瑞曲速L-曲霉菌抗体快速检测，第44列
	String RuiQuSu_Antigen; //瑞曲速-曲霉菌抗原快速检测，第43列
	String RuiQuFen_QuMeiJun; //瑞曲芬-曲霉菌核酸检测，第45列
	String RuiQuFen_NianZhuJun; //瑞念芬-念珠菌核酸检测，第47列
	String RuiPuFen_YeShiFei; //瑞普芬-耶氏肺孢子菌核酸检测，第46列
	String RuiMaoFen_MaoMeiJun; //瑞毛芬-毛霉菌核酸检测，第49列
	String RuiQuanFen_RCBIO; //瑞全芬D -RCBIO seqTMDNA，第41列
	
	public void setName(String mName) {
		name = mName;
	}
	
	public String toString() {
		return "name:" + name + " "
				+ "AcceptPart:" + AcceptPart + " "
				+ "gender:" + gender + " "
				+ "age:" + age + " "
				+ "exmNum:" + exmNum + " "
				+ "partCell:" + partCell + " "
				+ "bedNo:" + bedNo + " "
				+ "sampleTyple:" + sampleTyple + " "
				+ "date:" + date + " "
				+ "acceptDate:" + acceptDate + " "
				+ "sampleStatus:" + sampleStatus + " "
				+ "Doctor:" + Doctor + " "
				+ "clinicalCheck:" + clinicalCheck + " "
				+ "clinicalInfo:" + clinicalInfo + " "
				+ "videographyResult:" + videographyResult + " "
				+ "RuiQuSu_Antibody:" + RuiQuSu_Antibody + " "
				+ "RuiQuSu_Antigen:" + RuiQuSu_Antigen + " "
				+ "RuiQuFen_QuMeiJun:" + RuiQuFen_QuMeiJun + " "
				+ "RuiQuFen_NianZhuJun:" + RuiQuFen_NianZhuJun + " "
				+ "RuiPuFen_YeShiFei:" + RuiPuFen_YeShiFei + " "
				+ "RuiMaoFen_MaoMeiJun:" + RuiMaoFen_MaoMeiJun + " "
				+ "RuiQuanFen_RCBIO:" + RuiQuanFen_RCBIO
				;
	}

}
