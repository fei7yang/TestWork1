package com.dfl.report.straightmaterial;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;

public class StraightMaterialCQOp {

	private AbstractAIFUIApplication app;
	SimpleDateFormat df = new SimpleDateFormat("yyyy��MM��");// �������ڸ�ʽ
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// �������ڸ�ʽ
	private ArrayList list = new ArrayList();//���ݼ���
	private TCSession session ;
	public StraightMaterialCQOp(AbstractAIFUIApplication app) {
		// TODO Auto-generated constructor stub
		this.app = app;
		session = (TCSession) app.getSession();
		initUI();
	}

	private void initUI() {
		// TODO Auto-generated method stub
		try {
			// ��ȡѡ�ж���
			InterfaceAIFComponent ifc = app.getTargetComponent();

			TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc;

			// ������ʾ���Ȳ����ִ�в���
			ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);

			viewPanel.addInfomation("��ʼ�������...\n", 5, 100);

			// ���ݶ���BOP��ѯ���еĺ�װ����
			String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
			String[] values = new String[] { "ֱ��", "BIW Direct Material" };

			// ֱ�ļ���
			ArrayList partList = Util.searchBOMLine(topbomline, "OR", propertys, "==", values);
			
			if(partList==null || partList.size()<1) {
				viewPanel.addInfomation("��ʾ����ǰ������û��ֱ������\n", 100, 100);
				return;
			}
			viewPanel.addInfomation("�����������...\n", 20, 100);
			// ��ȡ��ǰ�����˼����ڲ���
			session = (TCSession) app.getSession();
			TCComponentUser user = session.getUser();
			TCComponentGroup group = user.getLoginGroup();
			String Manufacturing = group.getGroupName();
			String[] cover = new String[4];
			cover[0] = "      ���Ʋ��ţ�" + Manufacturing;
			cover[1] = "      �������ڣ�" + df.format(new Date());
			cover[2] = user.getUserName();
			cover[3] = df2.format(new Date());

			// ��ѯ���浼��ģ��
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_StraightMaterialCQ");

			if (inputStream == null) {
				viewPanel.addInfomation("����û���ҵ�ֱ�����Ķ����嵥ģ�壬�������ģ��(����Ϊ��DFL_Template_StraightMaterialCQ)\n", 100, 100);
				return;
			}
			viewPanel.addInfomation("�����������...\n", 40, 100);
			
			XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

			NewOutputDataToExcel.writeDataToSheetByGeneral(book, cover);
			
			viewPanel.addInfomation("�����������...\n", 60, 100);
			
			//����ֱ��������ҳ��ÿ18�����ݷ�һ��sheetҳ
			int sheetnum = (partList.size())/18 + 1;
			
			for(int i=0;i<partList.size();i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) partList.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				String str[] = new String[15];
				str[0] = Integer.toString(i+1);//���
				str[1] = Util.getProperty(bl, "bl_rev_object_name");//��������
				str[2] = "";//DFL���
				String value[] = getNESAndFactoryByDFL(str[2]);
				if(value!=null && value.length>0) {
					str[3] = value[0];//NES���
					str[13] = value[1];//����
				}
				else {
					str[3] = "";//NES���
					str[13] = "";//����
				}
				str[4] = "";//������;
				str[5] = "g";//��λ
				str[6] = null;//����1�ĵ�̨����
				str[7] = null;//����2�ĵ�̨����
				str[8] = null;//����3�ĵ�̨����
				str[9] = null;//����4�ĵ�̨����
				str[10] = null;//����5�ĵ�̨����
				str[11] = null;//����6�ĵ�̨����
				str[12] = "";//��װ
				
				str[14] = "";//ʹ�ð���
							
				list.add(str);
			}
			
			NewOutputDataToExcel.creatXSSFWorkbookByData(book,sheetnum);
			
			viewPanel.addInfomation("�����������...\n", 80, 100);
			
			NewOutputDataToExcel.writeStraightDataToSheet(book,sheetnum,cover,list);
			
			NewOutputDataToExcel.exportFile(book, "ֱ�����Ķ����嵥");
									
			NewOutputDataToExcel.openFile(FileUtil.getReportFileName("ֱ�����Ķ����嵥"));
			
			viewPanel.addInfomation("�����������...\n", 100, 100);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
    //����DFL���ͨ��ֱ�Ķ�Ӧ���ȡNES��źͳ���
	private String[] getNESAndFactoryByDFL(String code) {
		try {
			String value[]=new String[2];
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("B8_QueryStraightMaterialCorrespondTable");
			if (str == null) {
				return value;
			}
			String[] types = preferenceService.getStringValues("B8_QueryStraightMaterialCorrespondTable");
			
			for(int i=0;i<types.length;i++) {
				String temple[] = types[i].split("=");
				if(temple[0].equals(code)) {
					value[0] = temple[1];
					value[1] = temple[2];
					break;
				}
			}
			return value;
		}
		catch(Exception ex) {
			ex.printStackTrace();
		}
		return null;
	}
}
