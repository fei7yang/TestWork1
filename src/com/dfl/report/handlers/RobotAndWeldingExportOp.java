package com.dfl.report.handlers;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.Dictionary;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.swing.SwingUtilities;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.OutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.*;
import com.teamcenter.rac.pse.AbstractPSEApplication;
import com.teamcenter.rac.util.MessageBox;

public class RobotAndWeldingExportOp {

	public RobotAndWeldingExportOp(AbstractAIFUIApplication app, InputStream inputStream) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.session = (TCSession) this.app.getSession();
		this.inputStream = inputStream;
		initUI();
	}

	private AbstractAIFUIApplication app;
	private TCSession session;
	ArrayList<TCComponentBOMLine> meresource = new ArrayList<TCComponentBOMLine>();// ��һ����Դ������
	private String[] medata;// ���ݼ���
	private ArrayList data = new ArrayList();
	private InputStream inputStream;
    private String error = "";

	private void initUI() {
		// TODO Auto-generated method stub
		try {
			// ������ʾ���Ȳ����ִ�в���
			ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
			viewPanel.setVisible(true);

			// ��ȡѡ��Ķ���
			InterfaceAIFComponent[] ifc = app.getTargetComponents();

			TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];

			String linename = Util.getProperty(topbomline, "bl_rev_object_name");

			// viewPanel.addInfomation("��ʼ�������...\n", 5, 100);

			// �ж��û�����ѡ�����Ƿ���дȨ��
//			boolean flag = Util.hasWritePrivilege(session, topbomline.getItemRevision());
//			if (!flag) {
//				viewPanel.addInfomation("�Ե�ǰ��װ����û��дȨ�ޣ�...\n", 100, 100);
//				return;
//			}
			viewPanel.addInfomation("���ڻ�ȡģ��...\n", 10, 100);
			// ��ѯ����ģ��
//			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_RobotAndWeldingList");
//
//			if (inputStream == null) {
//				viewPanel.addInfomation("����û���ҵ�������&��ǹ�嵥ģ�壬�������ģ��(����Ϊ��DFL_Template_RobotAndWeldingList)\n", 100, 100);
//				return;
//			}

			// �����˺ͺ�ǹB8_BIWRobotRevision��B8_BIWGunRevision
			long startTime = System.currentTimeMillis(); // ��ȡ��ʼʱ��

			viewPanel.addInfomation("��ʼ�������...\n", 20, 100);

			// ��ȡ��һ��Դ�㼯��
			getStationBomLine(topbomline, viewPanel);

			long endTime = System.currentTimeMillis(); // ��ȡ����ʱ��
			System.out.println("��ȡ�����˺ͺ�ǹ��������ʱ�䣺 " + (endTime - startTime) + "ms");

			// ѭ����װ��λ��ȡ�²�Ļ����˻�ǹһ���ж��ٲ�
			startTime = System.currentTimeMillis(); // ��ȡ��ʼʱ��

			viewPanel.addInfomation("", 40, 100);

			if (meresource.size() > 0) {
				data = getRobotGunPropertys(meresource, viewPanel, 40);// ���ݼ���
			}
			if(!error.isEmpty())
			{
				viewPanel.dispose();
				inputStream.close();
				MessageBox.post("���¹�λ����Դ���д��ڶ�������ˣ��������ݣ�" + error, "��ܰ��ʾ", MessageBox.INFORMATION);
				return;
			}
			
			endTime = System.currentTimeMillis(); // ��ȡ����ʱ��
			System.out.println("��ȡ���Գ�������ʱ�䣺 " + (endTime - startTime) + "ms");

			viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 60, 100);

//			Comparator comparator2 = getComParator();
//			Collections.sort(data, comparator2);

			viewPanel.addInfomation("", 80, 100);
			String datasetname = linename + "������&��ǹ�嵥";
			String filename = Util.formatString(datasetname);

			OutputDataToExcel odte = new OutputDataToExcel(data, inputStream, filename, viewPanel);

			saveFiles(datasetname, filename, topbomline);

			viewPanel.addInfomation("���������ɣ����ں�װ���߰汾�����²鿴����\n", 100, 100);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// �õݹ��㷨��ȡ��һ��Դ��MECompResourceRevision
	public void getStationBomLine(TCComponentBOMLine parent, ReportViwePanel viewPanel) {

		viewPanel.addInfomation("", 10, 100);
		try {
			AIFComponentContext[] children = parent.getChildren();
			// String parentName = parent.getProperty("bl_rev_object_name");
			for (AIFComponentContext chld : children) {
				TCComponentItemRevision rev = ((TCComponentBOMLine) chld.getComponent()).getItemRevision();
				if (rev.isTypeOf("MECompResourceRevision")) {
					meresource.add((TCComponentBOMLine) chld.getComponent());
					continue;
				} else {
					getStationBomLine((TCComponentBOMLine) chld.getComponent(), viewPanel);
				}
			}
		} catch (TCException e) {
			// TODO �Զ����ɵ� catch ��
			e.printStackTrace();
		}
	}

	// ѭ�������ˣ���ȡ�����˻�ǹ�������ֵ
	public ArrayList getRobotGunPropertys(ArrayList<TCComponentBOMLine> parent, ReportViwePanel viewPanel, int n) {
		ArrayList list = new ArrayList();

		// ���ݶ���BOP��ѯ���еĺ�װ����
		String typename = Util.getObjectDisplayName(session, "B8_BIWRobot");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { typename, typename };

		// ���ݺ�װ���߲�ѯ���еĵ㺸����
		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values2 = new String[] { guntypename, guntypename };
		int rownum = 1;
		for (int i = 0; i < parent.size(); i++) {
			viewPanel.addInfomation("", n, 100);
			TCComponentBOMLine bl = parent.get(i);
			ArrayList<TCComponentBOMLine> Robotbl = Util.searchBOMLine(bl, "OR", propertys, "==", values);// �����˼���
			ArrayList<TCComponentBOMLine> Gunbl = Util.searchBOMLine(bl, "OR", propertys2, "==", values2);// ��ǹ����

			if(Robotbl!=null && Robotbl.size()>1)
			{
				String stateresource = "��λ���ƣ�" + getStationProperty(parent.get(i)) + " ��Դ����" + Util.getProperty(parent.get(i), "bl_rev_object_name");
				if(error.isEmpty())
				{
					error = stateresource;
				}
				else
				{
					
					error = error + "," + stateresource;
									
				}			
			}
			
			// ���ݺ�ǹ����ȷ�������������
			if (Gunbl.size() > 0) {
				for (int j = 0; j < Gunbl.size(); j++) {
					medata = new String[14];
					medata[0] = Integer.toString(rownum); // ���
					medata[1] = getStationProperty(parent.get(i));// ��ȡ�����˶�Ӧ�ĺ�װ��λ
					medata[2] = Util.getProperty(parent.get(i), "bl_rev_object_name");// �����˱��

					if (Robotbl.size() > 0) {
						TCComponentItemRevision rev;
						try {
							rev = Robotbl.get(0).getItemRevision();
							medata[3] = Util.getProperty(rev, "b8_Model");// �������ͺ�
							medata[4] = Util.getProperty(rev, "b8_Features");// �����˹���
						} catch (TCException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}

					} else {
						medata[3] = "";// �������ͺ�
						medata[4] = "";// �����˹���
					}
					medata[5] = getRobotHight(parent.get(i));// �����˵����߶�

					TCComponentItemRevision gunrev;
					try {
						gunrev = Gunbl.get(j).getItemRevision();
						medata[6] = Util.getProperty(gunrev, "b8_Model");// ��ǹǹ��
						medata[7] = Util.getProperty(gunrev, "b8_MotorBrand");// ��ǹ���
						medata[8] = Util.getProperty(gunrev, "b8_Deep");// ��ǹ����
						medata[9] = Util.getProperty(gunrev, "b8_Width");// ��ǹ���
						medata[10] = Util.getProperty(gunrev, "b8_Stroke");// ��ǹ�����г�
						medata[11] = Util.getProperty(gunrev, "b8_MaxAmp");// ��ǹ����
						medata[12] = Util.getProperty(gunrev, "b8_Voltage1");// ��ǹ��ѹ
					} catch (TCException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					medata[13] = "";// ��ע
					rownum++;
					list.add(medata);
				}

			} else {
				if (Robotbl.size() > 0) {
					medata = new String[14];
					medata[0] = Integer.toString(rownum); // ���
					medata[1] = getStationProperty(parent.get(i));// ��ȡ�����˶�Ӧ�ĺ�װ��λ
					medata[2] = Util.getProperty(parent.get(i), "bl_rev_object_name");// �����˱��

					if (Robotbl.size() > 0) {
						TCComponentItemRevision rev;
						try {
							rev = Robotbl.get(0).getItemRevision();
							medata[3] = Util.getProperty(rev, "b8_Model");// �������ͺ�
							medata[4] = Util.getProperty(rev, "b8_Features");// �����˹���
						} catch (TCException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}

					} else {
						medata[3] = "";// �������ͺ�
						medata[4] = "";// �����˹���
					}
					medata[5] = getRobotHight(parent.get(i));// �����˵����߶�

					medata[6] = "";// ��ǹǹ��
					medata[7] = "";// ��ǹ���
					medata[8] = "";// ��ǹ����
					medata[9] = "";// ��ǹ���
					medata[10] = "";// ��ǹ�����г�
					medata[11] = "";// ��ǹ����
					medata[12] = "";// ��ǹ��ѹ
					medata[13] = "";// ��ע

					rownum++;
					list.add(medata);
				}

			}
		}
		return list;
	}

	// ���ݻ����˻�ȡ��Ӧ�ĺ�װ��λ��Ϣ
	public String getStationProperty(TCComponentBOMLine chilren) {
		String station_name = "";
		try {
			TCComponentBOMLine station_bl = chilren.parent();
			TCComponentItemRevision rev = station_bl.getItemRevision();
			if (rev.isTypeOf("B8_StationRevision")) {
				station_name = Util.getProperty(station_bl, "bl_rev_object_name");
//				System.out.println(station_name);
				return station_name;
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return station_name;
	}

	// ���ݻ����˻�ȡ��Ӧ�Ļ����˵����߶�
	public String getRobotHight(TCComponentBOMLine chilren) {
		String station_name = "";
		try {
			// ��һ�㸴����Դ��
			AIFComponentContext[] chilrens = chilren.getChildren();
			for (AIFComponentContext chid : chilrens) {
				AIFComponentContext[] minchilrens = ((TCComponentBOMLine) chid.getComponent()).getChildren();// �ڶ��㸴����Դ��
				for (AIFComponentContext ch : minchilrens) {
					TCComponentItemRevision rev = ((TCComponentBOMLine) ch.getComponent()).getItemRevision();
					if (rev.isTypeOf("B8_BIWDeviceRevision")) {
						station_name = Util.getProperty(rev, "b8_Spec");
						continue;
					}
				}

			}
			return station_name;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return station_name;
	}

	private Comparator getComParator() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				Object[] comp1 = (Object[]) obj;
				Object[] comp2 = (Object[]) obj1;

				return comp1[1].toString().compareTo(comp2[1].toString());
			}
		};

		return comparator;
	}

	// �����ɵı�����Ϊ���ݼ��ŵ�Newstaff�ļ�����
	public void saveFiles(String datasetname, String filename, TCComponentBOMLine topbomline) {
		try {
			String fullFileName = FileUtil.getReportFileName(filename);
			TCSession session = (TCSession) app.getSession();
			TCComponentUser user = session.getUser();
			//TCComponentFolder newstuff = user.getNewStuffFolder();

			TCComponentItemRevision rev = topbomline.getItemRevision();
			TCComponentItemRevision docrev = null;
			// �ж��Ƿ�֮ǰ�����ɹ��ĵ�
			TCComponent[] children = rev.getRelatedComponents("IMAN_reference");
			if (children.length > 0) {
				for (TCComponent child : children) {
					if (child instanceof TCComponentItem) {
						TCComponentItem item = (TCComponentItem) child;
						if (item.isTypeOf("B8_BIWProcDoc")) {
							docrev = item.getLatestItemRevision();
							break;
						}
					}
				}
			}
			TCComponentDataset ds = Util.createDataset(session, datasetname, fullFileName, "MSExcelX", "excel");
			if (docrev == null) {
				TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session
						.getTypeComponent("B8_BIWProcDoc");

				TCComponentItem tccomponentitem = tcccomponentitemtype.create("", "", "B8_BIWProcDoc", datasetname,
						"desc", null);
				tccomponentitem.setProperty("b8_BIWProcDocType", "AH");
				tccomponentitem.lock();
				tccomponentitem.save();
				tccomponentitem.unlock();
				docrev = tccomponentitem.getLatestItemRevision();
				
				// ��Ӳ��߰汾���ĵ��Ĺ�ϵ
				rev.add("IMAN_reference", docrev.getItem());
				rev.lock();
				rev.save();
				rev.unlock();
			} else {
				// ���Ƴ�֮ǰ���ɵı���
				TCComponent[] childrens = docrev.getRelatedComponents("IMAN_specification");
				if (children.length > 0) {
					for (TCComponent child : childrens) {
						if (child instanceof TCComponentDataset) {
							TCComponentDataset dataset = (TCComponentDataset) child;
							docrev.cutOperation("IMAN_specification", new TCComponent[] { dataset });
							dataset.delete();
						}
					}
				}
			}
			// ����ĵ������ݼ��Ĺ�ϵ
			docrev.add("IMAN_specification", ds);
			docrev.lock();
			docrev.save();
			docrev.unlock();
			

			// ɾ���м��ļ�
			File file = new File(fullFileName);
			if (file.isFile()) {
				file.delete();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
