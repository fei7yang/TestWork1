package com.dfl.report.Fixturestylebook;

import java.io.IOException;
import java.io.InputStream;
import java.rmi.AccessException;
import java.util.ArrayList;
import java.util.List;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class FixturestylebookHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponent savefolder;
	// private String page1;
	private String page2;
	private InterfaceAIFComponent[] aifComponents;
	private ArrayList list = new ArrayList();
	private List finishlist;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		try {
			app = AIFUtility.getCurrentApplication();
			session = (TCSession) app.getSession();
			aifComponents = app.getTargetComponents();
			if (aifComponents == null || aifComponents.length < 1) {
				MessageBox.post("����ѡ�����", "����", MessageBox.INFORMATION);
				return null;
			}
			if (aifComponents.length > 1) {
				MessageBox.post("��ѡ��һ�ĺ�װ���߹��ն����װ��λ���ն���", "����", MessageBox.INFORMATION);
				return null;
			}
			if (aifComponents[0] instanceof TCComponentBOMLine) {

			} else {
				MessageBox.post("��ѡ������д��ڲ���BOMLine����", "��ʾ", MessageBox.INFORMATION);
				return null;
			}
			// �ж���ѡ���������
			TCComponentBOMLine topbomline = (TCComponentBOMLine) aifComponents[0];

			// System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")&&!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
				MessageBox.post("��ѡ������д��ڲ��Ǻ�װ���߹��հ汾��װ��λ���հ汾����", "��ʾ", MessageBox.INFORMATION);
				return null;
			}
			//�ж��Ƿ�ά������ģ��
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_FixtureStyleBook");

			if (inputStream == null) {
				MessageBox.post("����û���ҵ��о�ʽ����ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_FixtureStyleBook)��", "��ʾ", MessageBox.INFORMATION);
				//viewPanel.addInfomation("����û���ҵ��о�ʽ����ģ�壬�������ģ��(����Ϊ��DFL_Template_FixtureStyleBook)\n", 100, 100);
				return null;
			}else {
				inputStream.close();
			}
			String typename = Util.getObjectDisplayName(session, "B8_BIWMEProcStat");
			String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
			String[] values = new String[] { typename, typename };
			if(topbomline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")) {
				// ѭ�����й�λ���ж��Ƿ��������ɹ�����
				list = Util.searchBOMLine(topbomline, "OR", propertys, "==", values);
			}else {
				if(list!=null&& list.size()>0) {
					list.clear();
				}
				list.add(topbomline);
			}	
			finishlist = new ArrayList();
			//��ǶԺ�װ��λ���հ汾�����Ƿ���дȨ��
			String allgwnames = "";
			//��ǶԺ�װ��λ���հ汾�����ϵ�µļо�ʽ�����ĵ��汾�Ƿ���дȨ��
			String alldocuments = "";
			//��ǶԺ�װ��λ���ն����Ƿ���дȨ��
			String allgw = "";
			
			for (int i = 0; i < list.size(); i++) {
				// ����ѡ��ĺ�װ��λ�����£��Ƿ��Ѿ����ɹ�����������ɹ���ֱ��ȡ֮ǰ�ı�����Ϊģ��
				TCComponentBOMLine gwbl = (TCComponentBOMLine) list.get(i);
				TCComponentItemRevision blrev = gwbl.getItemRevision();
				String gwname = Util.getProperty(blrev, "object_name");
				// ������ļ�����
				String datasetname = gwname + "�о�ʽ����";
				String filename = Util.formatString(datasetname);
				TCComponent[] tccs = blrev.getRelatedComponents("IMAN_reference");
				TCComponentItem tcc = null;
				TCComponentItemRevision oldrev = null;
				System.out.println("��ϵ�������飺" + tccs);
				for (TCComponent item : tccs) {
					String type = Util.getRelProperty(item, "object_name");
					if (type.equals(datasetname)) {
						tcc = (TCComponentItem) item;
						break;
					}
				}
				// �ж��û�����ѡ�����Ƿ���дȨ��
				boolean flag = Util.hasWritePrivilege(session, blrev);
				if (!flag) {									
					if(allgwnames.isEmpty()) {
						allgwnames = gwname;
					}else {
						allgwnames = allgwnames + "," + gwname;
					}
				}
				
				System.out.println("��ϵ����" + tcc);
				if (tcc != null) {
					oldrev = tcc.getLatestItemRevision();
					//�ж��Ƿ��ѷ���
					if(oldrev.getDateProperty("date_released") != null) {
						if(!tcc.okToModify()) {
							if(allgw.isEmpty()) {
								allgw = gwname;
							}else {
								allgw = allgw + "," + gwname;
							}
						}
					}else {
						// �ж��û�����ѡ�����Ƿ���дȨ��
						boolean flag1 = Util.hasWritePrivilege(session, oldrev);
						if (!flag1) {					
							if(alldocuments.isEmpty()) {
								alldocuments = gwname;
							}else {
								alldocuments = alldocuments + "," + gwname;
							}
						}	
					}
								
					finishlist.add(oldrev);
				}
			}
			if(!allgwnames.isEmpty()) {
				MessageBox.post("�Ե�ǰ��ѡ" + allgwnames + "��װ��λ���հ汾����û��дȨ�ޣ�", "��ʾ", MessageBox.INFORMATION);	
				return null;
			}
			if(!alldocuments.isEmpty()) {
				MessageBox.post("�Ե�ǰ��ѡ" + alldocuments + "��װ��λ���հ汾�����ϵ�µļо�ʽ�����ĵ��汾����û��дȨ�ޣ�", "��ʾ", MessageBox.INFORMATION);	
				return null;
			}
			if(!allgw.isEmpty()) {
				MessageBox.post("�Ե�ǰ��ѡ" + allgw + "��װ��λ���հ汾�����ϵ�µļо�ʽ�����ĵ��汾�ѷ�����û��Ȩ��ִ���޶���", "��ʾ", MessageBox.INFORMATION);	
				return null;
			}
			
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		page2 = "3"; // Ĭ��3ҳ
		Thread thread = new Thread() {
			public void run() {
				if (finishlist != null && finishlist.size() > 0) {
					boolean IsContinu = Util.isContinue("�����ɹ������ٴ�����Ḳ��֮ǰ�����ɵı����Ƿ�����������");
					if (!IsContinu) {
						return ;
					}
				}
				try {
					new FixturestylebookOp(session, aifComponents, page2, false,list,"");
				} catch (TCException | AccessException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
//		Display.getDefault().asyncExec(new Runnable() {
//			@Override
//			public void run() {
//				
//			execute();
//		
//			}
//		});

		return null;
	}

	protected void execute() {
		// TODO Auto-generated method stub

//		shell = AIFDesktop.getActiveDesktop().getShell();
//
//		SelectPagesDialog dialog2 = new SelectPagesDialog(shell, SWT.SHELL_TRIM);
//		dialog2.open();
//
//		//page1 = dialog2.page1;
//		page2 = dialog2.page2;
//		
//		if(page2 == null || page2.isEmpty()) {
//			return;
//		}
//
//		Thread thread = new Thread() {
//			public void run() {
//				try {
//					new FixturestylebookOp(session, aifComponents,page2,false);
//				} catch (TCException | AccessException e) {
//					// TODO Auto-generated catch block
//					e.printStackTrace();
//				}
//			}
//		};
//		thread.start();

	}
}
