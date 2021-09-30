package com.dfl.report.handlers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class HitTheTableHandler extends AbstractHandler {

	public HitTheTableHandler() {
		// TODO Auto-generated constructor stub
	}

	private AbstractAIFUIApplication app;
	private InputStream inputStream;

	@Override
	public Object execute(ExecutionEvent event) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		TCSession session = (TCSession) app.getSession();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("��ǰδѡ�������������ѡ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("��ѡ��һ��װ���߹��ն���", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("��ѡ��BOP�еĺ�װ���߹��ն���", "��ʾ", MessageBox.INFORMATION);
			return null;
		}
		TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];
		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")) {
				MessageBox.post("��ѡ��BOP�еĺ�װ���߹��ն���", "��ܰ��ʾ", MessageBox.INFORMATION);
				return null;
			}
			TCComponentBOMLine parentline = topbomline.parent();
			if (!parentline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")) {
				MessageBox.post("��ѡ��װʵ�ʲ��߹����������", "��ܰ��ʾ", MessageBox.INFORMATION);
				return null;
			}
			TCComponentBOMLine topline = topbomline.window().getTopBOMLine();
			TCComponent[] projects = topline.getItemRevision().getRelatedComponents("project_list");
			if(projects == null || projects.length<1)
			{
				MessageBox.post("�뽫BOP����ָ����Ŀ��", "��ܰ��ʾ", MessageBox.INFORMATION);
				return null;
			}
			
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		boolean Isupdateflag = true;
		TCComponentItem tcc = null;

		// ����ѡ��ĺ�װ���߹����£��Ƿ��Ѿ����ɹ�����������ɹ���ֱ��ȡ֮ǰ�ı�����Ϊģ��
		TCComponentItemRevision blrev;
		try {
			blrev = topbomline.getItemRevision();
			// ������ļ�����
			String datasetname = topbomline.getProperty("bl_rev_object_name") + "��˳��";
			//String fileName = Util.formatString(datasetname);
			String fileName = datasetname;
			TCComponent[] tccs = blrev.getRelatedComponents("IMAN_reference");
			
			System.out.println("��ϵ�������飺" + tccs);
			for (TCComponent item : tccs) {
				String type = Util.getRelProperty(item, "b8_BIWProcDocType");
				System.out.println("�����ĵ������ͣ�" + type);
				if (type.equals("AA") || type.equals("��˳��")) {
					tcc = (TCComponentItem) item;
					break;
				}
			}
			System.out.println("��ϵ����" + tcc);

			if (tcc == null) {
				// �ж��û�����ѡ�����Ƿ���дȨ��
				boolean flag = Util.hasWritePrivilege(session, blrev);
				if (!flag) {
					MessageBox.post("�Ե�ǰ��װ���߹���û��дȨ�ޣ�", "��ܰ��ʾ", MessageBox.INFORMATION);
					return null;
				}
				// ��ѯ����ģ��
				inputStream = FileUtil.getTemplateFile("DFL_Template_HitTheTable");

				if (inputStream == null) {
					MessageBox.post("����û���ҵ���˳��ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_HitTheTable)", "��ܰ��ʾ",
							MessageBox.INFORMATION);
					return null;
				}
			} else {
				TCComponentItemRevision rev = tcc.getLatestItemRevision();
				TCComponent[] tccdata = rev.getRelatedComponents("IMAN_specification");
				TCComponentDataset dataset = null;
				File file = null;
				
				// �ж��û�����ѡ�����Ƿ���дȨ��
				boolean flag = Util.hasWritePrivilege(session, rev);
				if (!flag) {
					MessageBox.post("�Ե�ǰ��װ���߹��չ�ϵ�µĴ�˳���ĵ��汾����û��дȨ�ޣ�", "��ܰ��ʾ", MessageBox.INFORMATION);
					return null;
				}
				
				if (tccdata != null && tccdata.length > 0) {
					dataset = (TCComponentDataset) tccdata[0];
				}
				System.out.println("��ȡ�����ݼ���" + dataset);
				if (dataset != null) {
					String filepath = System.getProperty("java.io.tmpdir");
					File tf=new File(filepath+fileName + ".xlsx");
					if(tf.exists())
						tf.delete();
					file = dataset.getFile("excel", fileName + ".xlsx", dataset.getWorkingDir());
				}
				if (file == null) {
					Isupdateflag = false; // �����ĵ����󣬲����ڱ������ݼ�
					System.out.println("�������⣺" + file);
					inputStream = FileUtil.getTemplateFile("DFL_Template_HitTheTable");
					if (inputStream == null) {
						MessageBox.post("����û���ҵ���˳��ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_HitTheTable)", "��ܰ��ʾ",
								MessageBox.INFORMATION);
						return null;
					}
				} else {
					// ���ݻ�ȡ�ı���Ϊģ��
					inputStream = new FileInputStream(file);					
					if (inputStream == null) {
						System.out.println("�������⣺" + inputStream);
						Isupdateflag = false; // �����ĵ����󣬲����ڱ������ݼ�
						inputStream = FileUtil.getTemplateFile("DFL_Template_HitTheTable");
						if (inputStream == null) {
							MessageBox.post("����û���ҵ���˳��ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_HitTheTable)", "��ܰ��ʾ",
									MessageBox.INFORMATION);
							return null;
						}
					}
					else
					{
						System.out.println("��ȡ�������ɵı���" + Isupdateflag);
					}
				}
				System.out.println("�����ļ���" + file);

			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		HitTheTableAction action = new HitTheTableAction(app, null, "",tcc,inputStream,Isupdateflag);
		Thread th = new Thread(action);
		th.start();

		return null;
	}

}
