
public class PatientInfo {
	String AcceptPart;//�ͼ쵥λ, ��2��
	String name;//���������� ��7��
	String gender; //�����Ա𣬵�8��
	String age; //�������䣬 ��9��
	String exmNum; //������ţ���6��
	String partCell; //�ͼ���ң���3��
	String bedNo; //�����ţ���4��
	String sampleTyple; //�������ͣ� ��18��
	String date; //�������ڣ� ��1��
	String acceptDate; //�������ڣ� ��1��
	String sampleStatus = ""; //����״̬�� ���ù�
	String Doctor; //�ͼ�ҽ������5��
	String clinicalCheck; //�ٴ���ϣ���26��
	String clinicalInfo; //�ٴ���Ϣ����27��
	String videographyResult; //Ӱ��ѧ�������28��
	String RuiQuSu_Antibody; //������L-��ù��������ټ�⣬��44��
	String RuiQuSu_Antigen; //������-��ù����ԭ���ټ�⣬��43��
	String RuiQuFen_QuMeiJun; //������-��ù�������⣬��45��
	String RuiQuFen_NianZhuJun; //�����-����������⣬��47��
	String RuiPuFen_YeShiFei; //���շ�-Ү�Ϸ����Ӿ������⣬��46��
	String RuiMaoFen_MaoMeiJun; //��ë��-ëù�������⣬��49��
	String RuiQuanFen_RCBIO; //��ȫ��D -RCBIO seqTMDNA����41��
	
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
