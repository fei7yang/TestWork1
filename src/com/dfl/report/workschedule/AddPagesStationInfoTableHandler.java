package com.dfl.report.workschedule;

import java.io.FileNotFoundException;
import java.io.InputStream;
import java.rmi.AccessException;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AbstractAIFSession;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

/*
 * 工程作业表增加sheet页
 */
public class AddPagesStationInfoTableHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private Shell shell;
	private List shList;
	private TCSession session;
	private XSSFWorkbook book;
	private TCComponentBOMLine topbomline;
	private String sheetname;
	private String newsheetname;
	private String sheetpages;
	private String model;
	private String modelname;
	private GenerateReportInfo info;
	private TCComponentDataset factdatawet;
	Logger logger = LogManager.getLogger(AddPagesStationInfoTableHandler.class);

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		logger.info("测试但是但是发射点");
		logger.error("dsfdfs ");;
		app = AIFUtility.getCurrentApplication();
		session = (TCSession) app.getSession();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("当前未选择操作对象，请先选择！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("请选择单一焊装工位工艺对象！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("请选择焊装工位工艺对象！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		topbomline = (TCComponentBOMLine) ifc[0];

		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
				MessageBox.post("请选择焊装工厂工位对象！", "温馨提示", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// 将模板中选择的sheet复制到报表中
		TCComponentDataset dataset;

		dataset = FileUtil.getDatasetFile("DFL_Template_EngineeringWorkListStation");
		if (dataset == null) {
			MessageBox.post("请配置模板，名称为：DFL_Template_EngineeringWorkListStation，请联系系统管理员！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		dataset = FileUtil.getDatasetFile("DFL_Template_EngineeringWorkVINCarve");
		if (dataset == null) {
			MessageBox.post("请配置模板，名称为：DFL_Template_EngineeringWorkVINCarve，请联系系统管理员！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		dataset = FileUtil.getDatasetFile("DFL_Template_AdjustmentLine");
		if (dataset == null) {
			MessageBox.post("请配置模板，名称为：DFL_Template_AdjustmentLine，请联系系统管理员！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		TCComponentItemRevision toprev = null;
		try {
			toprev = topbomline.window().getTopBOMLine().getItemRevision();
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		
		// 获取已生成的报表
		info = new GenerateReportInfo();
		info.setExist(false);
		info.setIsgoon(true);
		info.setAction(""); //$NON-NLS-1$
		info.setMeDocument(null);
		info.setDFL9_process_type("H"); //$NON-NLS-1$
		info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
		info.setmeDocumentName("");
		info.setFlag(true);
		info.setProject_ids(toprev);
		

		try {
			info = ReportUtils.beforeGenerateReportAction(topbomline.getItemRevision(), info);
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
			return null;
		}
		System.out.println("The action is completed before the report operation is generated.");

		if (!info.isIsgoon()) {
			return null;
		}
		TCComponentItemRevision docmentRev = info.getMeDocument();
		if(docmentRev ==null) {
			MessageBox.post("请确认已经生成工程作业表！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
				
		String procName = Util.getProperty(docmentRev, "object_name");
		InputStream inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

		if (inputStream == null) {
			MessageBox.post("请确认" + procName + "版本对象下，存在" + procName + "数据集！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		TCComponent[] datasets = Util.getRelComponents(docmentRev, "IMAN_specification");
		if (datasets == null || datasets.length <= 0) {
			System.out.println(docmentRev.toDisplayString() + " 未获取到Excel数据集对象");
		}

		for (int j = 0; j < datasets.length; j++) {
			String type = datasets[j].getType();
			if ("MSExcelX".equals(type)) {
				factdatawet = (TCComponentDataset) datasets[j];
				break;
			}
		}
		if (factdatawet == null) {		
			System.out.println(docmentRev.toDisplayString() + " 未获取到Excel数据集对象");
		}
		book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);
		shList = new ArrayList();
		for (int i = 0; i < book.getNumberOfSheets(); i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			String sheetname = sheet.getSheetName();
			shList.add(sheetname);
		}
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				execute();
			}
		});

		return null;
	}

	protected void execute() {
		// TODO Auto-generated method stub
		shell = AIFDesktop.getActiveDesktop().getShell();

		ExistingSheetPagesDialog dialog = new ExistingSheetPagesDialog(shell, SWT.SHELL_TRIM, shList);
		dialog.open();

		sheetname = dialog.sheetname;
		if (sheetname == null || sheetname.isEmpty()) {
			return;
		}
		SelectSheetTypeDialog typedialog = new SelectSheetTypeDialog(shell, SWT.SHELL_TRIM, shList);
		typedialog.open();

		newsheetname = typedialog.sheetname;
		sheetpages = typedialog.sheetpages;
		model = typedialog.model;
		modelname = typedialog.modelname;
		if (newsheetname == null || newsheetname.isEmpty()) {
			return;
		}

		Thread thread = new Thread() {
			public void run() {
				try {
					new AddPagesStationInfoTableOp(session, factdatawet, topbomline, sheetname, newsheetname,
							sheetpages, model, modelname, info);
				} catch (TCException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		};
		thread.start();

	}

}
