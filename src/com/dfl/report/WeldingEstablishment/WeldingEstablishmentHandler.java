package com.dfl.report.WeldingEstablishment;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.dialog.SelectionSiteNameDialog;
import com.dfl.report.home.OpenHomeDialog;
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

public class WeldingEstablishmentHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private TCSession session;
	private TCComponentFolder rootFolder;
	private TCComponent savefolder;
	private String reportname;
	//private static Logger logger = LogManager.getLogger(WeldingEstablishmentHandler.class.getName());
	private static Logger logger = Logger.getLogger(WeldingEstablishmentHandler.class.getName()); // 日志打印类
	private InterfaceAIFComponent[] aifComponents;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		try {			
			app = AIFUtility.getCurrentApplication();
			aifComponents = app.getTargetComponents();
			session = (TCSession) app.getSession();
			if (aifComponents == null || aifComponents.length < 1) {
				MessageBox.post("请先选择对象！", "错误", MessageBox.INFORMATION);
				return null;
			}
			for (int i = 0; i < aifComponents.length; i++) {
				if (aifComponents[i] instanceof TCComponentBOMLine) {

				} else {
					MessageBox.post("所选择对象中存在不是BOMLine对象！", "提示", MessageBox.INFORMATION);
					return null;
				}
				// 判断所选对象的类型
				TCComponentBOMLine topbomline = (TCComponentBOMLine) aifComponents[i];

				// System.out.println(topbomline.getItemRevision().getType());
				if (!topbomline.getItemRevision().isTypeOf("B8_BBOMPartitionRevision")
						&& !topbomline.getItemRevision().isTypeOf("B8_WeldContainerRevision")
						&& !topbomline.getItemRevision().isTypeOf("B8_BBOMTopNodeRevision")) {
					MessageBox.post("所选择对象中存在不是BBOM总成、BBOM功能分区或焊点包对象！", "提示", MessageBox.INFORMATION);
					return null;
				}
			}
			// 获取首选项定义的Note属性
			TCPreferenceService ts = session.getPreferenceService();
			if (!ts.isDefinitionExistForPreference("B8_WeldFeasibilityReport")) {
				MessageBox.post("错误：首选项B8_WeldFeasibilityReport未定义,请联系系统管理员！", "提示", MessageBox.INFORMATION);
				logger.error("错误：首选项B8_WeldFeasibilityReport未定义");
				return null;
			}

			String error = getSizeRule();
			if (!error.isEmpty()) {
				MessageBox.post(error, "提示", MessageBox.INFORMATION);
				logger.error(error);
				return null;
			}

			InputStream inputStream = Util.getReportTempByprefercen(session, "B8_WeldFeasibilityReport", 1);
			if (inputStream == null) {
				MessageBox.post("焊接成立性一元表模板不存在，请联系系统管理员！", "提示", MessageBox.INFORMATION);
				logger.error("焊接成立性报表模板不存在！");
				return null;
			} else {
				try {
					inputStream.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
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
		System.out.println("文件夹：" + dialog.folder);

		if (dialog.flag) {
			return;
		}

		if (savefolder == null) {
			return;
		}

		SelectionSiteNameDialog dialog2 = new SelectionSiteNameDialog(shell, SWT.SHELL_TRIM);
		dialog2.open();

		reportname = dialog2.name;

		if (reportname == null || reportname.isEmpty()) {
			return;
		}

		Thread thread = new Thread() {
			public void run() {
				try {
					new WeldingEstablishmentOp(session, aifComponents, reportname, savefolder);
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();
	}

	// 根据材质获取对应的强度
	private String getSizeRule() {
		String error = "";
		try {

			File file = null;
			Workbook workbook = null;
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_part_strength");
			if (str != null) {
				String value = preferenceService.getStringValue("DFL9_get_part_strength");
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
							error = "材料强度规则表不存在，请联系系统管理员!";
						}
					} else {
						error = "材料强度规则表不存在，请联系系统管理员!";
					}
				} else {
					error = "错误：首选项DFL9_get_part_strength未定义，请联系系统管理员！";
				}
			} else {
				error = "错误：首选项DFL9_get_part_strength未定义，请联系系统管理员！";
			}

			return error;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return error;
	}
}
