package com.dfl.report.handlers;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.OutputDataToExcel2;
import com.dfl.report.util.OutputDataToExcel3;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class NewCarryOverListOp {
	private AbstractAIFUIApplication app;
	private TCSession session;
	private TCComponentBOMLine parent;
	ArrayList childList = new ArrayList();
	ArrayList childList1 = new ArrayList();
	private String[] value;// ��������
	private ArrayList valuelist = new ArrayList();
	int number = 0;// ��������
	String[][] data;
	private TCComponent folder;
	private Map parentMap;
	private InterfaceAIFComponent[] aifComponents;

	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMdd  HH");// �������ڸ�ʽ
	
	public NewCarryOverListOp(AbstractAIFUIApplication app, TCComponent folder, InterfaceAIFComponent[] aifComponents, TCSession session) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.folder=folder;
		this.session = session;
		this.aifComponents =aifComponents;
		parentMap = new HashMap();
		initUI();
	}

	private void initUI() {
		// TODO Auto-generated method stub
		try {
			// ������ʾ���Ȳ����ִ�в���
			ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);

			// ��ȡѡ��Ķ���
			//InterfaceAIFComponent[] aifComponents = app.getTargetComponents();
			if(aifComponents==null) {
				return;
			}
			TCComponentBOMLine topbomline = (TCComponentBOMLine) aifComponents[0];
			TCComponentBOMLine topbl=topbomline.getCachedWindow().getTopBOMLine();

			viewPanel.addInfomation("���ڻ�ȡģ��...\n", 20, 100);
			// ��ѯ������ģ��
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_NewCarryOverList");
			System.out.println("inputStream=" + inputStream);

			if (inputStream == null) {
				viewPanel.addInfomation("����û���ҵ��¹����ò�Ʒ�嵥ģ�壬�������ģ��(����Ϊ��DFL_Template_NewCarryOverList)\n", 100, 100);
				return;
			}
			
			//�޸��ж��Ƿ��ƶ���Ԫ������ 2019-08-20
			String[] propertys = new String[] { "bl_usage_address"};// �Ƿ��ƶ���Ԫ
			String[] values = new String[] { "MU"};
			ArrayList partlist = new ArrayList();
			for(InterfaceAIFComponent aif: aifComponents) {
				TCComponentBOMLine bl = (TCComponentBOMLine) aif;
				ArrayList chlist =  Util.searchBOMLine(bl, "OR", propertys, "=", values);
				for(int i=0;i<chlist.size();i++) {
					if(!partlist.contains(chlist.get(i))) {
						partlist.add(chlist.get(i));
					}
				}
			}
//			ArrayList partlist =  Util.searchBOMLine(topbomline, "OR", propertys, "=", values);
//			String[] propertys = new String[] { "B8_NoteIsBiwTrUnit","B8_NoteIsBiwTrUnit" };// �Ƿ��ƶ���Ԫ
//			String[] values = new String[] { "��","1" };
//			ArrayList partlist =  Util.searchBOMLine(topbomline, "OR", propertys, "==", values);
			String tempvehicle= topbl.getItemRevision().getProperty("project_ids");// ����
			String vehicle = Util.getDFLProjectIdVehicle(tempvehicle);
			if(vehicle==null || vehicle.isEmpty()) {
				vehicle = tempvehicle;
			}
//			System.out.println(vehicle);
			getBomLine(partlist);//��ȡ���и�������



			int index = 1;
			Iterator iterator = parentMap.entrySet().iterator();
			while (iterator.hasNext()) 
			{
				Map.Entry entry = (Entry) iterator.next();
				TCComponentBOMLine parent = (TCComponentBOMLine) entry.getKey();
				List childList = (List) entry.getValue();

				
				String parentName = parent.getProperty("bl_rev_object_name");// ����������������
				
				if(childList.isEmpty())
				{
					continue;
				}
				
				String NO = Integer.toString(index);// ��������������ţ�
				
				index++;
				for (int j = 0; j < childList.size(); j++) {
					//�ж��ƶ���Ԫ�Ƿ���ڴ��
					TCComponentBOMLine packLines[] = ((TCComponentBOMLine) childList.get(j)).getPackedLines();
					if(packLines!=null&&packLines.length>0){
						String counts=((TCComponentBOMLine)childList.get(j)).getProperty("bl_pack_count");
						int count= Integer.parseInt(counts);
						for(int p=count-1;p<count;p++) {
							//String Name =((TCComponentBOMLine)childList.get(j)).getProperty("bl_rev_object_name");// �������
							String Name =Util.getProperty(((TCComponentBOMLine)childList.get(j)).getItemRevision(), "dfl9_CADObjectName");// �������
							String ID = ((TCComponentBOMLine) childList.get(j)).getProperty("bl_DFL9SolItmPartRevision_dfl9_part_no");// �����
							String NewCaOver = ((TCComponentBOMLine) childList.get(j)).getProperty("DFL9_new_caover_mark");// ����/����

							value = new String[5];

							value[0] = NO;
							value[1] = ID;
							value[2] = Name;
							value[3] = NewCaOver;
							value[4] = parentName;

							valuelist.add(value);

							number++;
						}
						
					}else {
						//String Name =((TCComponentBOMLine)childList.get(j)).getProperty("bl_rev_object_name");// �������
						String Name =Util.getProperty(((TCComponentBOMLine)childList.get(j)).getItemRevision(), "dfl9_CADObjectName");// �������
						String ID = ((TCComponentBOMLine) childList.get(j)).getProperty("bl_DFL9SolItmPartRevision_dfl9_part_no");// �����
						String NewCaOver = ((TCComponentBOMLine) childList.get(j)).getProperty("DFL9_new_caover_mark");// ����/����

						value = new String[5];

						value[0] = NO;
						value[1] = ID;
						value[2] = Name;
						value[3] = NewCaOver;
						value[4] = parentName;

						valuelist.add(value);

						number++;
					}
					
				}

			}

			viewPanel.addInfomation("��ʼ�������...\n", 40, 100);

			// ���屨��������
			data = new String[number][13];
			// System.out.println("data=" + data.length);
			for (int i = 0; i < valuelist.size(); i++) {

				String[] str = (String[]) valuelist.get(i);

				data[i][0] = "";// ģ���ͷǰ��һ��
				data[i][1] = "";// ����
				data[i][2] = str[0];// ��һ��NO
				data[i][3] = str[4];// ���BLOCK ���������������
				data[i][4] = Integer.toString(i + 1);// �ڶ���NO
				data[i][5] = str[1];// ����� Parts No
				data[i][6] = str[2];// ������� Parts Name
				data[i][7] = str[3];// ����/���� New/Carry Over
				data[i][8] = "";// ACL
				data[i][9] = str[3];// Datum ��New/Carry Overһ�� B8_NoteNewCaOverMark
				data[i][10] = "";// PosiMark
				data[i][11] = "";// ID Mark
				data[i][12] = "";// Remark
			}
			viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 70, 100);
			//����ļ�����
			//String name=Util.getProperty(topbomline.getItemRevision(), "object_name");
			// �ļ�����
			String functionname = "";
			for (InterfaceAIFComponent aif : aifComponents) {
				TCComponentBOMLine aifbl = (TCComponentBOMLine) aif;
				if (functionname.isEmpty()) {
					functionname = Util.getProperty(aifbl, "bl_rev_object_name");
				} else {
					functionname = functionname + "&" + Util.getProperty(aifbl, "bl_rev_object_name");
				}
			}
			String date = dateformat.format(new Date());
			String datasetName = vehicle+"_"+functionname+"_"+"���輰���ò�Ʒ�嵥"+"_"+date+"ʱ";
			String fileName = Util.formatString(datasetName);
			
			OutputDataToExcel2 outdata = new OutputDataToExcel2(data, inputStream, fileName.trim(), vehicle);
			//�洢����
			Util.saveFiles(fileName.trim(),datasetName, folder, session,"AL");
			
			viewPanel.addInfomation("���������ɣ�����ѡ�񱣴���ļ����²鿴��\n", 100, 100);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}


	// ��ȡ�������ƶ���Ԫ��bomline�ĸ�������
	public void getBomLine(ArrayList partlist) throws TCException {


		for (int i = 0; i < partlist.size(); i++) {
			try {
				TCComponentBOMLine childLine =	(TCComponentBOMLine) partlist.get(i);
				// ��ȡ����ĸ�������
				parent = ((TCComponentBOMLine) partlist.get(i)).parent();

				if(parentMap.containsKey(parent))
				{
					List list = (List) parentMap.get(parent);
					list.add(childLine);
					parentMap.put(parent, list);
				}else
				{
					List list = new ArrayList();
					list.add(childLine);
					parentMap.put(parent, list);
				}

			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}


		System.out.println("parent size="+parentMap.size());

	}

}
