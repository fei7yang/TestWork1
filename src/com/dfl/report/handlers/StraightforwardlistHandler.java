package com.dfl.report.handlers;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;

import javax.swing.SwingUtilities;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class StraightforwardlistHandler extends AbstractHandler {

	public StraightforwardlistHandler() {
		// TODO Auto-generated constructor stub
	}

	private AbstractAIFUIApplication app;
	private Shell shell;
	// private AbstractAIFApplication application;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;
	private InterfaceAIFComponent[] ifc;
	private InputStream inputStream;
	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("��ǰδѡ�������������ѡ��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		for (InterfaceAIFComponent aif : ifc) {
			if (aif instanceof TCComponentBOMLine) {

			} else {
				MessageBox.post("��ѡ�����д��ڲ���BOMLine����", "��ܰ��ʾ", MessageBox.INFORMATION);
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
						MessageBox.post("��ѡ�����д��ڲ��Ǻ�װ�������հ汾�������߰汾����", "��ܰ��ʾ", MessageBox.INFORMATION);
						return null;
					}
				}
			} catch (TCException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		String error = getSizeRule();
		if (!error.isEmpty()) {
			MessageBox.post(error, "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
		}
		// ��ѯ����ģ��
		inputStream = FileUtil.getTemplateFile("DFL_Template_Straightforwardlist");

		if (inputStream == null) {
			MessageBox.post("û���ҵ�ֱ���嵥��ģ�壬����ϵϵͳ����Ա���ģ��(����Ϊ��DFL_Template_Straightforwardlist)��", "��ܰ��ʾ", MessageBox.INFORMATION);
			return null;
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
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return;
		}
		if (savefolder == null) {
			try {
				inputStream.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			return;
		}
		StraightforwardlistAction action = new StraightforwardlistAction(app, null, savefolder, ifc, session,inputStream);
		new Thread(action).start();
	}

	// ��ѯ��С��������ѡ���ȡ��С��������Ϣ
	private String getSizeRule() {
		String error = "";
		try {

			File file = null;
			Workbook workbook = null;
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_straight_sheet_size_rule");
			if (str != null) {
				String value = preferenceService.getStringValue("DFL9_straight_sheet_size_rule");
				if (value != null) {
					TCComponentDatasetType datatype = (TCComponentDatasetType) session.getTypeComponent("Dataset");
					TCComponentDataset dataset = datatype.find(value);
					if (dataset != null) {
						String type = dataset.getType();

						TCComponentTcFile[] files;
						try {
							files = dataset.getTcFiles();
							if (files.length > 0) {
								file = files[0].getFmsFile();
							}
						} catch (TCException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}

						if (file != null) {

						} else {
							error = "��ѡ��DFL9_straight_sheet_size_rule����Ĵ�С���������ݼ����ƣ���TCϵͳδ�ҵ�������ά����";
						}
					} else {
						error = "��ѡ��DFL9_straight_sheet_size_rule����Ĵ�С���������ݼ����ƣ���TCϵͳδ�ҵ�������ά����";
					}
				} else {
					error = "��ѡ��DFL9_straight_sheet_size_ruleδ���壬����ϵϵͳ����Ա��";
				}
			} else {
				error = "��ѡ��DFL9_straight_sheet_size_ruleδ���壬����ϵϵͳ����Ա��";
			}
			return error;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return error;
	}
}
