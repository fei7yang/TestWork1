package com.dfl.report.workschedule;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import javax.swing.JOptionPane;

import org.apache.poi.ss.formula.functions.T;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.widgets.Display;

import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;

public class BasicInformationBZHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private GenerateReportInfo info;
	private InputStream inputStream = null;
	private Map<String, List<String>> MaterialMap;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		app = AIFUtility.getCurrentApplication();
		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		if (ifc.length < 1) {
			MessageBox.post("当前未选择操作对象，请先选择！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc.length > 1) {
			MessageBox.post("请选择单一的焊装工厂工艺对象！", "温馨提示", MessageBox.INFORMATION);
			return null;
		}
		if (ifc[0] instanceof TCComponentBOMLine) {

		} else {
			MessageBox.post("请选择焊装工厂工艺对象！", "提示", MessageBox.INFORMATION);
			return null;
		}
		TCComponentBOMLine topbomline = (TCComponentBOMLine) ifc[0];
		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWPlantBOPRevision")) {
				MessageBox.post("请选择焊装工厂工艺对象！", "温馨提示", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		// 获取材料对照表
		MaterialMap = baseinfoExcelReader.getMaterialComparisonTable(app, "DFL_MaterialMapping");
		if (MaterialMap == null || MaterialMap.size() < 1) {
			System.out.println("未找到材料对照表！");
			MessageBox.post("未配置对照表DFL_MaterialMapping，请联系系统管理员！", "提示信息", MessageBox.ERROR);
			return null;
		}

		System.out.println("用于测试无响应问题");

		/******************* 判断放到前面 *******************************/
		// 文件名称
		String procName = "222.基本信息";
		// 生成报表操作前的动作
		info = new GenerateReportInfo();
		info.setExist(false);
		info.setIsgoon(true);
		info.setAction(""); //$NON-NLS-1$
		info.setMeDocument(null);
		info.setDFL9_process_type("H"); //$NON-NLS-1$
		info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
		info.setmeDocumentName(procName);

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

		if (info.getAction() == "create") { // 都输出
			MessageBox.post("请先输出基本信息-焊点清单信息！", "信息提示", MessageBox.INFORMATION);
			return null;
		} else {
			TCComponentItemRevision docmentRev = info.getMeDocument();
			inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

			if (inputStream == null) {
				MessageBox.post("请确认222.基本信息文档版本对象下，存在222.基本信息数据集！", "信息提示", MessageBox.INFORMATION);
				return null;
			}
		}
		/*************************************************/

		Thread thread = new Thread() {
			public void run() {
				boolean IsContinu = Util.isContinue("会覆盖上一次输出的板组信息，请确认是否继续输出报表？");
				if (IsContinu) {
					try {
						new BasicInformationBZOp(app, info, inputStream,MaterialMap);
					} catch (TCException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				} else {
					try {
						inputStream.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		};
		thread.start();

		return null;
	}
}
