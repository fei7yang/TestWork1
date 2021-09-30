package com.dfl.report.handlers;

import java.awt.Container;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.swt.widgets.DateTime;

import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class HitTheTableExportOp {

	private AbstractAIFUIApplication app;
	ArrayList<TCComponentBOMLine> list = new ArrayList<TCComponentBOMLine>();// �㺸���򼯺�
	private String[] Discrete;// �㺸��Ϣ���� {�㺸�������ƣ����ڣ������˱�ţ��������ͺţ���ǹǹ�ţ���װ��λ������}
	private String[] Weld;// ������Ϣ����{�㺸�������ƣ�����ID�����}
	ArrayList Discretelist = new ArrayList();// �㺸��Ϣ����
	ArrayList Weldlist = new ArrayList();// ������Ϣ����
	String vehicle = null; // ����
	private String beginname;
	private InputStream inputStream;
	private boolean Isupdateflag;
	private TCComponentItem tcc = null;

	public HitTheTableExportOp(AbstractAIFUIApplication app, InputStream inputStream, boolean isupdateflag, TCComponentItem tcc) {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.inputStream = inputStream;
		this.Isupdateflag = isupdateflag;
		this.tcc = tcc;
		initUI();
	}

	private void initUI() {
		// TODO Auto-generated method stub
		try {

			// ��ȡѡ��Ķ���
			InterfaceAIFComponent[] ifc = app.getTargetComponents();

			TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];

			TCSession session = (TCSession) app.getSession();
			TCComponentUser user = session.getUser();
			TCComponentFolder newstuff = user.getNewStuffFolder();
			TCComponent[] projects = topbomline.window().getTopBOMLine().getItemRevision().getRelatedComponents("project_list");

			// ����ѡ��ĺ�װ���߹����£��Ƿ��Ѿ����ɹ�����������ɹ���ֱ��ȡ֮ǰ�ı�����Ϊģ��
			TCComponentItemRevision blrev = topbomline.getItemRevision();

			// ������ļ�����
			String datasetname = topbomline.getProperty("bl_rev_object_name") + "��˳��";
			//String fileName = Util.formatString(datasetname);
			String fileName = datasetname;
			
				
			if (tcc == null) {
				// ��ȡtopbomline��һ��
				TCComponentBOMLine uplevel = topbomline.window().getTopBOMLine();

				TCComponentItemRevision uprev = uplevel.getItemRevision();

				// vehicle = Util.getProperty(uplevel, "b8_vehicle_id");
				vehicle = Util.getProperty(uprev, "project_ids");

				// ������ʾ���Ȳ����ִ�в���
				ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
				viewPanel.setVisible(true);

				// viewPanel.addInfomation("��ʼ�������...\n", 5, 100);
				
				viewPanel.addInfomation("��ʼ�������...\n", 10, 100);
				
				viewPanel.addInfomation("�����������...\n", 20, 100);

				// �õݹ��㷨��ȡ���еĵ㺸����B8_BIWDiscreteOP
				getBIWDiscreteOP(topbomline);

				viewPanel.addInfomation("", 40, 100);

				// �ж��Ƿ��е㺸��������
				if (list.size() < 1) {
					inputStream.close();
					viewPanel.dispose();
					MessageBox.post("���󣺵�ǰ��װ���߹���û�л����˵㺸��������", "��ܰ��ʾ", MessageBox.INFORMATION);
					//viewPanel.addInfomation("���󣺵�ǰ��װ���߹���û�е㺸��������\n", 100, 100);
					return;
				}

				// ѭ����ʼ�㺸�������Ʊ��
				beginname = ((TCComponentBOMLine) list.get(0)).getProperty("bl_rev_object_name");

				// ѭ���㺸�����ȡ�����ˡ��������Ϣ
				getStationBomLine(list);

				// �ж��Ƿ��е㺸��������
				if (Discretelist.size() < 1 && Weldlist.size() < 1) {
					inputStream.close();
					viewPanel.dispose();
					MessageBox.post("���󣺵�ǰ��װ���߹���û�л����˵㺸��������", "��ܰ��ʾ", MessageBox.INFORMATION);
					//viewPanel.addInfomation("���󣺵�ǰ��װ���߹���û�л����˺ͺ�������\n", 100, 100);
					return;
				}

				viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 60, 100);
				// дexcel����
				// ����ģ�崴��Excel��ģ��
				ArrayList errorname = new ArrayList();
				XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream, Discretelist,errorname);

				if (book == null) {
					viewPanel.dispose();
					String statename = "";
					for(int i=0;i<errorname.size();i++)
					{
						if(i == 0)
						{
							statename = (String) errorname.get(i);
						}
						else
						{
							statename = statename + "��" + errorname.get(i);
						}
					}
					MessageBox.post("���󣺵�ǰ��װ���߹��մ���:" + statename + "��sheet�������ȹ������޷����ɴ�˳��", "��ܰ��ʾ", MessageBox.INFORMATION);
					return;
				}

				// дsheet���ݣ��㺸��������д��
				NewOutputDataToExcel.writeDiscreteDataToSheet(book, Discretelist,Weldlist, true);

				viewPanel.addInfomation("", 80, 100);

				// дsheet���ݣ���������д��
				NewOutputDataToExcel.writeWeldDataToSheet(book, Weldlist,Discretelist, true);

				// ����ļ�
				NewOutputDataToExcel.exportFile(book, fileName);

				saveFiles(datasetname, fileName, topbomline,tcc,projects);
								

				viewPanel.addInfomation("���������ɣ����ں�װ���߹��ն��󸽼��»�Newstuff�ļ����²鿴����\n", 100, 100);

			} else {
				// �����ļ�������

				// ������ʾ���Ȳ����ִ�в���
				ReportViwePanel viewPanel = new ReportViwePanel("���ɱ���");
				viewPanel.setVisible(true);
	
				viewPanel.addInfomation("��ʼ�������...\n", 10, 100);
	
				// ��ȡtopbomline��һ��
				TCComponentBOMLine uplevel = topbomline.window().getTopBOMLine();
				// ��BOP�汾������ȥ������Ϣ
				TCComponentItemRevision uprev = uplevel.getItemRevision();

				String tempvehicle = Util.getProperty(uprev, "project_ids");
				
				vehicle = Util.getDFLProjectIdVehicle(tempvehicle);
				
				if(vehicle==null || vehicle.isEmpty()) {
					vehicle = tempvehicle;
				}

				// �õݹ��㷨��ȡ���еĵ㺸����B8_BIWDiscreteOP
				getBIWDiscreteOP(topbomline);

				viewPanel.addInfomation("�����������...\n", 20, 100);

				// �ж��Ƿ��е㺸��������
				if (list.size() < 1) {
					inputStream.close();
					viewPanel.dispose();
					MessageBox.post("���󣺵�ǰ��װ���߹���û�л����˵㺸��������", "��ܰ��ʾ", MessageBox.INFORMATION);
					//viewPanel.addInfomation("���󣺵�ǰ��װ���߹���û�е㺸��������\n", 100, 100);
					return;
				}

				// ѭ����ʼ�㺸�������Ʊ��
				beginname = ((TCComponentBOMLine) list.get(0)).getProperty("bl_rev_object_name");

				// ѭ���㺸�����ȡ�����ˡ��������Ϣ
				getStationBomLine(list);

				// �ж��Ƿ��е㺸��������
				if (Discretelist.size() < 1 && Weldlist.size() < 1) {
					inputStream.close();
					viewPanel.dispose();
					MessageBox.post("���󣺵�ǰ��װ���߹���û�л����˵㺸��������", "��ܰ��ʾ", MessageBox.INFORMATION);
					//viewPanel.addInfomation("���󣺵�ǰ��װ���߹���û�л����˺ͺ�������\n", 100, 100);
					return;
				}
				viewPanel.addInfomation("��ʼд���ݣ������ĵȴ�...\n", 40, 100);

				XSSFWorkbook book = null;
				if(Isupdateflag) {
					System.out.println("������·��� " + Isupdateflag);
					ArrayList errorname = new ArrayList();
					// ����ģ�崴��Excelģ��
					book = NewOutputDataToExcel.updateXSSFWorkbook(inputStream, Discretelist,errorname);
					
					if (book == null) {
						viewPanel.dispose();
						String statename = "";
						for(int i=0;i<errorname.size();i++)
						{
							if(i == 0)
							{
								statename = (String) errorname.get(i);
							}
							else
							{
								statename = statename + "��" + errorname.get(i);
							}
						}
						MessageBox.post("���󣺵�ǰ��װ���߹��մ���:" + statename + "��sheet�������ȹ������޷����ɴ�˳��", "��ܰ��ʾ", MessageBox.INFORMATION);
					}
					
					// дsheet���ݣ��㺸��������д��
					NewOutputDataToExcel.UpdateDiscreteDataToSheet(book, Discretelist, Weldlist,true);
					viewPanel.addInfomation("", 60, 100);
					// ����sheet���ݣ���������д��
					NewOutputDataToExcel.updateWeldDataToSheet(book, Weldlist, Discretelist, true);
				}else {
					System.out.println("������������ " + Isupdateflag);
					ArrayList errorname = new ArrayList();
					book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream, Discretelist,errorname);
					
					if (book == null) {
						viewPanel.dispose();
						String statename = "";
						for(int i=0;i<errorname.size();i++)
						{
							if(i == 0)
							{
								statename = (String) errorname.get(i);
							}
							else
							{
								statename = statename + "��" + errorname.get(i);
							}
							
						}
						MessageBox.post("���󣺵�ǰ��װ���߹��մ���:" + statename + "��sheet�������ȹ������޷����ɴ�˳��", "��ܰ��ʾ", MessageBox.INFORMATION);
						return;
					}
					
					// дsheet���ݣ��㺸��������д��
					NewOutputDataToExcel.writeDiscreteDataToSheet(book, Discretelist,Weldlist, true);
					viewPanel.addInfomation("", 60, 100);
					// ����sheet���ݣ���������д��
					NewOutputDataToExcel.writeWeldDataToSheet(book, Weldlist,Discretelist, true);
				}
				
				viewPanel.addInfomation("", 80, 100);
				// ����ļ�
				NewOutputDataToExcel.exportFile(book, fileName);

				saveFiles(datasetname, fileName, topbomline,tcc,projects);
				
				// �Ѹ��µı����ϴ�
//				String filePath = FileUtil.getReportFileName(fileName);
//				String nameRef = "excel";
//				String[] filePaths = { filePath };
//				String[] namedRefs = { nameRef };
//				dataset.setFiles(filePaths, namedRefs);
//				dataset.lock();
//				dataset.save();
//				dataset.unlock();

				viewPanel.addInfomation("���������ɣ����ں�װ���߹��ն��󸽼��»�Newstuff�ļ����²鿴����\n", 100, 100);
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// �����ɵı�����Ϊ���ݼ��ŵ�Newstaff�ļ�����
	public void saveFiles(String datasetname, String filename, TCComponentBOMLine topbomline, TCComponentItem tcc2, TCComponent[] projects) {
		try {
			TCComponentItemRevision toprev = topbomline.getItemRevision();
			TCSession session = (TCSession) app.getSession();
			TCComponentItem tccomponentitem;
			String fullFileName = FileUtil.getReportFileName(filename);
			TCComponentUser user = session.getUser();
			TCComponentFolder newstuff = user.getNewStuffFolder();
			TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session.getTypeComponent("B8_BIWProcDoc");

			if(tcc == null) {
				tccomponentitem = tcccomponentitemtype.create("", "", "B8_BIWProcDoc", datasetname, "desc",
						null);
				newstuff.add("contents", tccomponentitem);
				// MessageBox.post("�����ɹ���", "��Ϣ��ʾ", MessageBox.INFORMATION);
				tccomponentitem.setProperty("b8_BIWProcDocType", "AA");
				tccomponentitem.lock();
				tccomponentitem.save();
				tccomponentitem.unlock();
				// ��Ӻ�װ�������ĵ��Ĺ�ϵ
				toprev.add("IMAN_reference", tccomponentitem);
				toprev.lock();
				toprev.save();
				toprev.unlock();
			}else {
				tccomponentitem = tcc;
			}						
			TCComponentItemRevision rev = tccomponentitem.getLatestItemRevision();
			// �Ƴ���ʱ����Ҫ�����з��������Ķ����ҳ��������Ƴ�
			TCComponent[] children = TCComponentUtils.getCompsByRelation(rev, "IMAN_specification");
			for (TCComponent child : children) {
				if (child instanceof TCComponentDataset) {
					TCComponentDataset dataset = (TCComponentDataset) child;
					rev.cutOperation("IMAN_specification", new TCComponent[] { dataset });
					try {
						dataset.delete();
					} catch (Exception e2) {

					}
				}
			}					
			TCComponentDataset ds = Util.createDataset(session, datasetname, fullFileName, "MSExcelX", "excel");
			// ����ĵ������ݼ��Ĺ�ϵ
			rev.add("IMAN_specification", ds);
			rev.lock();
			rev.save();
			rev.unlock();
			
			// ���ĵ�ָ����Ŀ
			Util.assignProjectComp(rev, projects);
					
			//ɾ���м��ļ�
			File file = new File(fullFileName);
			if(file.isFile()) {
				file.delete();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// �õݹ��㷨��ȡ���еĻ����˵㺸����B8_BIWDiscreteOP
	public void getBIWDiscreteOP(TCComponentBOMLine parent) {
		try {
			AIFComponentContext[] children = parent.getChildren();
			//String parentName = parent.getProperty("bl_rev_object_name");
			for (AIFComponentContext chld : children) {
				TCComponentItemRevision rev = ((TCComponentBOMLine) chld.getComponent()).getItemRevision();
				// ����ǵ㺸����
				if (rev.isTypeOf("B8_BIWDiscreteOPRevision")) {
					String objectname = Util.getProperty(rev, "object_name");
					if(objectname.length()>0 && "R".equals(objectname.subSequence(0, 1))) {
						list.add((TCComponentBOMLine) chld.getComponent());
						continue;
					}				
				} else {
					getBIWDiscreteOP((TCComponentBOMLine) chld.getComponent());
				}
			}
		} catch (TCException e) {
			// TODO �Զ����ɵ� catch ��
			e.printStackTrace();
		}
	}

	SimpleDateFormat df = new SimpleDateFormat("yyyy/MM/dd");// �������ڸ�ʽ

	// ѭ���㺸�����ȡ�����ˡ��������Ϣ
	public void getStationBomLine(ArrayList<TCComponentBOMLine> parent) {
		try {
			for (int i = 0; i < parent.size(); i++) {
				AIFComponentContext[] children = parent.get(i).getChildren();
				String parentName = parent.get(i).getProperty("bl_rev_object_name");
				Discrete = new String[11];
				Discrete[0] = parentName;// �㺸��������
				Discrete[2] = parentName;// ��˳���С������˱�š�һ����Ϣ��sheetҳ������ͬ  20191008
				Discrete[1] = df.format(new Date());// ����
				Discrete[6] = vehicle;// ����
				int num = 0;// ���
				if (!beginname.equals(parentName)) {
					num = 0;
					beginname = parentName;
				}

				// ��ȡ��װ��λ����ͨ���㺸�����ȡ��һ����������һ����ȡ��װ��λ
				Discrete[5] = getStation(parent.get(i).parent());

				for (AIFComponentContext chld : children) {
					TCComponentItemRevision rev = ((TCComponentBOMLine) chld.getComponent()).getItemRevision();
					System.out.println("�������ͣ�" + rev.getType());
					// ������
					if (rev.isTypeOf("B8_BIWRobotRevision")) {
//						Discrete[2] = Util.getProperty(((TCComponentBOMLine) chld.getComponent()),
//								"bl_rev_object_name");// �����˱��
						Discrete[3] = Util.getProperty(rev, "b8_Model");// �������ͺ�
					}
					// ��ǹ
					if (rev.isTypeOf("B8_BIWGunRevision")) {
						Discrete[4] = Util.getProperty(rev, "b8_Model");// ��ǹǹ��
					}
					// ����
					if (rev.isTypeOf("WeldPointRevision")) {
						Weld = new String[8];
						Weld[0] = parentName;// �㺸��������
						Weld[1] = Integer.toString(num + 1);// ���
						Weld[2] = Util.getProperty(rev, "object_name");// ����ID
						Weld[5] = parent.get(i).parent().getProperty("bl_rev_object_name"); //��λ����
						Weld[6] =  parent.get(i).getProperty("bl_sequence_no"); //��ѯ���
						Weld[7] =  parent.get(i).getProperty("bl_rev_item_id"); //�㺸����ID
						Weldlist.add(Weld);
						num++;
					}

				}
				Discrete[7] = Integer.toString(num);// ���������
				Discrete[8] = parent.get(i).parent().getProperty("bl_rev_object_name"); //��λ����
				Discrete[9] = parent.get(i).getProperty("bl_sequence_no"); //��ѯ���
				Discrete[10] = parent.get(i).getProperty("bl_rev_item_id"); //�㺸����ID
				Discretelist.add(Discrete);
			}
		} catch (TCException e) {
			// TODO �Զ����ɵ� catch ��
			e.printStackTrace();
		}
	}

	// ��ȡ��װ��λ
	public String getStation(TCComponentBOMLine parent) {
		String stationname = "";
		try {

			AIFComponentContext[] children = parent.getChildren();
			String parentName = parent.getProperty("bl_rev_object_name");

			for (AIFComponentContext chld : children) {
				TCComponentItemRevision rev = ((TCComponentBOMLine) chld.getComponent()).getItemRevision();
				if (rev.isTypeOf("B8_StationRevision")) {
					stationname = Util.getProperty(((TCComponentBOMLine) chld.getComponent()), "bl_rev_object_name");// ��װ��λ
					continue;
				}
			}
			return stationname;
		} catch (TCException e) {
			// TODO �Զ����ɵ� catch ��
			e.printStackTrace();
		}
		return stationname;
	}
}
