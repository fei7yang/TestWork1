package com.dfl.report.handlers;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.home.OpenHomeDialog;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class AntirustRequirementsCheckHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	// private AbstractAIFApplication application;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;
	private InterfaceAIFComponent[] ifc;
	private ArrayList rule;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		ifc = app.getTargetComponents();
		// ��ѯ������ģ��
		InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_AntirustRequirementsCheck");
		System.out.println("inputStream=" + inputStream);

		if (inputStream == null) {
			MessageBox.post("����û���ҵ�����Ҫ������ģ�壬����ϵϵͳ����Ա��TC�����ģ��(����Ϊ��DFL_Template_AntirustRequirementsCheck)", "����",
					MessageBox.INFORMATION);
			return null;
		} else {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		rule = getSelectStateRule();
		if (rule == null || rule.size() < 1) {
			MessageBox.post("������ѡ��δ����DFL9_Selection_test_phase������ϵϵͳ����Ա��", "����", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length < 1) {
			MessageBox.post("��ǰδѡ�������������ѡ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		for (InterfaceAIFComponent aif : ifc) {
			if (aif instanceof TCComponentBOMLine) {

			} else {
				MessageBox.post("��ѡ�����д��ڲ���BOMLine����", "��ʾ", MessageBox.INFORMATION);
				return null;
			}
			TCComponentBOMLine topbomline = (TCComponentBOMLine) aif;
			try {
				System.out.println(topbomline.getItemRevision().getType());
				if (!topbomline.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")
						&& !topbomline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")) {
					MessageBox.post("��ѡ�����д��ڲ��Ǻ�װ�������հ汾�������߰汾����", "��ܰ��ʾ", MessageBox.INFORMATION);
					return null;
				}
				if (topbomline.getItemRevision().isTypeOf("B8_BIWMEProcLineRevision")) {
					boolean flag = Util.getIsVirtualLine(topbomline);
					if (!flag) {
						MessageBox.post("��ѡ�����������߹��հ汾����", "��ܰ��ʾ", MessageBox.INFORMATION);
						return null;
					}
				}
			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
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
		session = (TCSession) app.getSession();
		shell = AIFDesktop.getActiveDesktop().getShell();

		// InterfaceAIFComponent aifComponent = app.getTargetComponent();

		try {
			rootFolder = session.getUser().getHomeFolder();
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// rootFolder = (TCComponent) aifComponent;

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
		AntirustRequirementsCheckAction action = new AntirustRequirementsCheckAction(app, null, savefolder, ifc,
				session,rule);
		new Thread(action).start();
	}

	// ��ѯ�׶���ѡ���ȡ�׶���Ϣ
	private ArrayList getSelectStateRule() {
		ArrayList rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_Selection_test_phase");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_Selection_test_phase");
				if (values != null) {
					for (int i = 0; i < values.length; i++) {
						String value = values[i];
						rule.add(value);
					}
				}
			}
			return rule;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rule;
	}
}
