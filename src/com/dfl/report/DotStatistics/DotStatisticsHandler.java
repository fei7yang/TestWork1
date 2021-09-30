package com.dfl.report.DotStatistics;

import java.io.IOException;
import java.io.InputStream;

import org.apache.xalan.xsltc.compiler.util.Util;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.FileUtil;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class DotStatisticsHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;
	private InterfaceAIFComponent[] aifComponents;

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
				MessageBox.post("��ѡ��һ�ĺ�װ�������ն���", "����", MessageBox.INFORMATION);
				return null;
			}

			if (aifComponents[0] instanceof TCComponentBOMLine) {

			} else {
				MessageBox.post("ѡ�������BOMLine����", "��ʾ", MessageBox.INFORMATION);
				return null;
			}
			// �ж���ѡ���������
			TCComponentBOMLine topbomline = (TCComponentBOMLine) aifComponents[0];

			// System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
				MessageBox.post("��ѡ��װ�������ն���", "��ʾ", MessageBox.INFORMATION);
				return null;
			}
			// ��ѯĿ¼����ģ��
			InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_DotStatistics");

			if (inputStream == null) {
				MessageBox.post("����û���ҵ����ͳ�Ʊ�ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_DotStatistics)", "��ʾ", MessageBox.INFORMATION);
				return null;
			}else {
				try {
					inputStream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
//			TCComponentGroup group = session.getGroup();			
//			String groupname = group.getGroupName();
//			
//			if(!groupname.equals("ͬ�ڹ��̿�")) {
//				MessageBox.post("������ͬ�ڹ��̿Ƶģ���Ȩ�����ɸñ���", "��ʾ", MessageBox.INFORMATION);
//				return null;
//			}
					
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Thread thread = new Thread() {
			public void run() {
				execute();
			}
		};
		thread.start();

		return null;
	}

	protected void execute() {
		// TODO Auto-generated method stub
		
		shell = AIFDesktop.getActiveDesktop().getShell();

		try {
			rootFolder = session.getUser().getHomeFolder();
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				openDialog();
			}
		});

	}

	protected void openDialog() {
		// TODO Auto-generated method stub
		OpenHomeDialog dialog = new OpenHomeDialog(shell, rootFolder, session);
		dialog.open();

		savefolder = dialog.folder;
		System.out.println("�ļ��У�" + dialog.folder);

		if (dialog.flag) {
			return;
		}

		if (savefolder == null) {
			return;
		}

		Thread thread = new Thread() {
			public void run() {
				try {
					new DotStatisticsOp(session, aifComponents,savefolder);
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
	}
}
