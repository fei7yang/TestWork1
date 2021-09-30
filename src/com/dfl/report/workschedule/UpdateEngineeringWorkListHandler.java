package com.dfl.report.workschedule;

import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.CurrentandVoltage;
import com.dfl.report.ExcelReader.RecommendedPressure;
import com.dfl.report.ExcelReader.SFSequenceWeldingConditionList;
import com.dfl.report.ExcelReader.SequenceComparisonTable;
import com.dfl.report.ExcelReader.SequenceWeldingConditionList;
import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.dealparameter.DealParameterHandler;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AIFDesktop;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.kernel.TCUserService;
import com.teamcenter.rac.util.MessageBox;

/* *****************************************
 * 更新工程作业表，只是用于计算用户维护新增焊点的参数值
 * @hgq
 * 20191026
 */
public class UpdateEngineeringWorkListHandler extends AbstractHandler {

	private AbstractAIFUIApplication app;
	private static Logger logger = Logger.getLogger(UpdateEngineeringWorkListHandler.class.getName()); // 日志打印类
	private List<WeldPointBoardInformation> baseinfolist = new ArrayList<WeldPointBoardInformation>();// 基本信息表的数据
	private static List<SequenceWeldingConditionList> swc = new ArrayList<SequenceWeldingConditionList>();// 24序列焊接条件设定表
																											// 序列号
	private static List<CurrentandVoltage> cv = new ArrayList<CurrentandVoltage>();// 24序列焊接条件设定表 电流电压
	private static List<SFSequenceWeldingConditionList> SFswc = new ArrayList<SFSequenceWeldingConditionList>();// 255序列焊接条件设定表
	private static List<RecommendedPressure> rp = new ArrayList<RecommendedPressure>();// 推荐加压力
	private static List<SequenceComparisonTable> sct = new ArrayList<SequenceComparisonTable>();// 序列对照表
	private TCSession session;
	private Map<String, List<String>> MaterialMap;
	private Shell shell;
	private TCComponentBOMLine topbomline;
	private String result;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
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
			MessageBox.post("请选择焊装工位工艺对象！", "提示", MessageBox.INFORMATION);
			return null;
		}
		topbomline = (TCComponentBOMLine) ifc[0];
		try {
			System.out.println(topbomline.getItemRevision().getType());
			if (!topbomline.getItemRevision().isTypeOf("B8_BIWMEProcStatRevision")) {
				MessageBox.post("请选择焊装工位工艺对象！", "温馨提示", MessageBox.INFORMATION);
				return null;
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// 获取首选项定义的Note属性
		TCPreferenceService ts = session.getPreferenceService();
		if (!ts.isDefinitionExistForPreference("B8_Calculation_Parameter_Name")) {
			MessageBox.post("错误：首选项B8_Calculation_Parameter_Name未定义，请联系系统管理员！", "提示信息", MessageBox.ERROR);
			return null;
		}
		// 获取材料对照表
		MaterialMap = baseinfoExcelReader.getMaterialComparisonTable(app, "DFL_MaterialMapping");
		if (MaterialMap == null || MaterialMap.size() < 1) {
			System.out.println("未找到材料对照表！");
			MessageBox.post("未配置对照表DFL_MaterialMapping，请联系系统管理员！", "提示信息", MessageBox.ERROR);
			return null;
		}

		shell = AIFDesktop.getActiveDesktop().getShell();
		Display.getDefault().asyncExec(new Runnable() {
			@Override
			public void run() {
				openDialog();
			}
		});

		return null;
	}

	protected void openDialog() {
		// TODO Auto-generated method stub
		DetermineDialog dialog = new DetermineDialog(shell, SWT.SHELL_TRIM);
		dialog.open();

		result = dialog.getMessage();

		Thread thread = new Thread() {
			public void run() {
				if (!result.isEmpty()) {
					try {
						UpdateEngineeringWorkList(topbomline, result);
					} catch (TCException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		};
		thread.start();

	}

	private void UpdateEngineeringWorkList(TCComponentBOMLine topbomline, String result) throws TCException {
		// TODO Auto-generated method stub

		String procName = "";
		// 生成报表操作前的动作
		GenerateReportInfo info = new GenerateReportInfo();
		info.setExist(false);
		info.setIsgoon(true);
		info.setAction(""); //$NON-NLS-1$
		info.setMeDocument(null);
		info.setDFL9_process_type("H"); //$NON-NLS-1$
		info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
		info.setmeDocumentName(procName);
		info.setFlag(true);
		info.setProject_ids(topbomline.window().getTopBOMLine().getItemRevision());

		try {
			info = ReportUtils.beforeGenerateReportAction(topbomline.getItemRevision(), info);
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
			return;
		}
		System.out.println("The action is completed before the report operation is generated.");

		if (!info.isIsgoon()) {
			return;
		}
		if (!info.isExist()) {
			MessageBox.post("请确认已经生成工程作业表！", "温馨提示", MessageBox.INFORMATION);
			return;
		}
		InputStream inputStream = null;
		TCComponentItemRevision docmentRev = info.getMeDocument();
		procName = Util.getProperty(docmentRev, "object_name");
		inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

		if (inputStream == null) {
			MessageBox.post("请确认" + procName + "版本对象下，存在" + procName + "数据集！", "温馨提示", MessageBox.INFORMATION);
			return;
		}
		// 获取基本信息
		String baseName = "222.基本信息";
		TCComponentBOMLine topbl = topbomline.window().getTopBOMLine();
		baseinfolist = getBaseinfomation(topbl, baseName);
		if (baseinfolist == null || baseinfolist.size() < 1) {
			System.out.println("请先生成工程作业表-基本信息表！");
			MessageBox.post("请先生成工程作业表-基本信息表！", "提示信息", MessageBox.ERROR);
			return;
		}
		// 获取计算参数
		Object[] obj = baseinfoExcelReader.getCalculationParameter(app, "B8_Calculation_Parameter_Name");
		if (obj != null) {
			if (obj[0] != null) {
				swc = (List<SequenceWeldingConditionList>) obj[0];
			} else {
				System.out.println("未获取到24序列焊接条件设定表 序列号信息。");
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
			if (obj[1] != null) {
				cv = (List<CurrentandVoltage>) obj[1];
			} else {
				System.out.println("未获取到24序列焊接条件设定表 电流电压信息。");
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
			if (obj[2] != null) {
				SFswc = (List<SFSequenceWeldingConditionList>) obj[2];
			} else {
				System.out.println("未获取到255序列焊接条件设定表信息。");
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
			if (obj[3] != null) {
				rp = (List<RecommendedPressure>) obj[3];
			} else {
				System.out.println("未获取到推荐加压力信息。");
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
			if (obj[4] != null) {
				sct = (List<SequenceComparisonTable>) obj[4];
			} else {
				System.out.println("未获取到序列对照表信息。");
				MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
				return;
			}
		} else {
			MessageBox.post("未找到焊接参数计算规则！", "提示信息", MessageBox.ERROR);
			return;
		}

		// 显示进度输出窗口
		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("报表更新");
		viewPanel.setVisible(true);
		viewPanel.addInfomation("正在更新...\n", 5, 100);

		XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

		viewPanel.addInfomation("", 20, 100);

		// 开启旁路
		{
			Util.callByPass(session, true);
		}
		// 处理PSWsheet 页
		DealPSWSheet(book, viewPanel, result);

		viewPanel.addInfomation("正在更新...\n", 40, 100);

		// 处理RSW气动
		DealRSWQDSheet(book, viewPanel);

		// 处理RSW伺服
		DealRSWSFSheet(book, viewPanel);

		viewPanel.addInfomation("", 60, 100);

		TCComponentItemRevision dfl9MEDocumentRev = info.getMeDocument();
		TCComponentDataset tagedataset = null;
		TCComponent[] children = TCComponentUtils.getCompsByRelation(dfl9MEDocumentRev, "IMAN_specification");
		for (TCComponent child : children) {
			if (child instanceof TCComponentDataset) {
				TCComponentDataset dataset = (TCComponentDataset) child;
				tagedataset = dataset;
				break;
			}
		}
		String fileName = Util.formatString(Util.getProperty(tagedataset, "object_name"));
		// 输出文件
		NewOutputDataToExcel.exportFile(book, fileName);

		viewPanel.addInfomation("", 80, 100);

		String fullFileName = FileUtil.getReportFileName(fileName);
		System.out.println(fullFileName);
		TCComponentDataset ds = Util.createDataset(session, fileName, fullFileName, "MSExcelX", "excel");
		List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
		List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
		if (ds != null) {
			datasetList.add(ds);
		}
		revlist.add(topbomline.getItemRevision());
		try {
			TCComponentItem docunment = ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "",
					session);
			// saveFileToFolder(docunment, topfoldername, childrenFoldername);

		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}
		// 关闭旁路
		{
			Util.callByPass(session, false);
		}
		viewPanel.addInfomation("新增焊点数据更新完成。", 100, 100);
	}

	private void DealRSWSFSheet(XSSFWorkbook book, ReportViwePanel viewPanel) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		ArrayList sheetAtIndexs = new ArrayList(); // RSW气动所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("RSW伺服")) {
				sheetAtIndexs.add(i);
			}
		}
		if (sheetAtIndexs == null && sheetAtIndexs.size() < 1) {
			return;
		}
		// 获取sheet内的数据
		ArrayList datalist = getSheetData(book, sheetAtIndexs, true);

		ArrayList hdlist = new ArrayList();

		// 根据焊点在基本信息表中获取板件信息
		List hdinfo = new ArrayList();// 焊点所有信息

		hdinfo = getBoardInformation(baseinfolist, datalist);

		// 循环焊点信息，计算并获取参数属性值
		Map<String, String[]> paramap = getDealParameter(hdinfo, true, viewPanel);

		// 根据新增焊点的位置，将获取到的信息写入到报表中
		writeRSWSFInfomation(book, hdinfo, paramap);
	}

	private void writeRSWSFInfomation(XSSFWorkbook book, List hdinfo, Map<String, String[]> paramap) {
		// TODO Auto-generated method stub
		if (hdinfo != null && hdinfo.size() > 0) {
			// 设置字体颜色
			Font font = book.createFont();
			font.setColor((short) 12);// 蓝色字体
			font.setFontHeightInPoints((short) 9);
			XSSFCellStyle style = book.createCellStyle();
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style.setFont(font);

			Font font2 = book.createFont();
			font2.setColor((short) 12);// 蓝色字体
			font2.setFontHeightInPoints((short) 18);
			font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderTop(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font2);

			// 粉色背景色
			XSSFCellStyle style3 = book.createCellStyle();
			style3.setFillForegroundColor(IndexedColors.PINK.getIndex());
			style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style3.setFont(font);
			// 紫色背景色
			Font font3 = book.createFont();
			font3.setColor((short) 1);// 白色字体
			font3.setFontHeightInPoints((short) 10);
			XSSFCellStyle style4 = book.createCellStyle();
			style4.setFillForegroundColor(IndexedColors.VIOLET.getIndex());
			style4.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style4.setFont(font3);
			// 蓝色背景色
			Font font4 = book.createFont();
			font4.setColor((short) 1);// 白色字体
			font4.setFontHeightInPoints((short) 10);
			XSSFCellStyle style5 = book.createCellStyle();
			style5.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
			style5.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style5.setFont(font4);

			XSSFCellStyle style6 = book.createCellStyle();
			style6.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			// style6.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style6.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style6.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style6.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style6.setFont(font);

			XSSFCellStyle style8 = book.createCellStyle();
			style8.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style8.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style8.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			// style8.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style8.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style8.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style8.setFont(font);

			// 白色背景色
			Font font5 = book.createFont();
			font4.setFontHeightInPoints((short) 10);
			XSSFCellStyle style7 = book.createCellStyle();
			style7.setFillForegroundColor(IndexedColors.WHITE.getIndex());
			style7.setFillPattern(CellStyle.SOLID_FOREGROUND);
			style7.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style7.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style7.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style7.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style7.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style7.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style7.setFont(font5);

			// 设置字体颜色
			Font font6 = book.createFont();
			font6.setColor((short) 2);// 红色字体
			font6.setFontHeightInPoints((short) 10);
			XSSFCellStyle style66 = book.createCellStyle();
			style66.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			// style6.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style66.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style66.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style66.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style66.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style66.setFont(font6);

			// 粉色背景色
			Font fontpink = book.createFont();
			fontpink.setColor((short) 12);// 蓝色字体
			fontpink.setFontName("MS PGothic");
			fontpink.setFontHeightInPoints((short) 9);

			XSSFCellStyle stylepink = book.createCellStyle();
			stylepink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
			stylepink.setFillPattern(CellStyle.SOLID_FOREGROUND);
			stylepink.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			stylepink.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			stylepink.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			stylepink.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			stylepink.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			stylepink.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			stylepink.setFont(fontpink);

			for (int i = 0; i < hdinfo.size(); i++) {
				String[] vals = (String[]) hdinfo.get(i);
				int sheetindex = Integer.parseInt(vals[2]); // sheet所在位置
				int rowindex = Integer.parseInt(vals[1]); // 所在行数
				XSSFSheet sheet = book.getSheetAt(sheetindex);

				String weldno = vals[3]; // 焊点编号
				String importance = vals[4]; // 重要度
				String boardnumber1 = vals[5]; // 板材1编号
				String boardname1 = vals[6]; // 板材1名称
				String partmaterial1 = vals[7]; // 板材1材质
				String partthickness1 = vals[8]; // 板材1板厚
				String boardnumber2 = vals[9]; // 板材2编号
				String boardname2 = vals[10]; // 板材2名称
				String partmaterial2 = vals[11]; // 板材2材质
				String partthickness2 = vals[12]; // 板材2板厚
				String boardnumber3 = vals[13]; // 板材3编号
				String boardname3 = vals[14]; // 板材3名称
				String partmaterial3 = vals[15]; // 板材3材质
				String partthickness3 = vals[16]; // 板材3板厚
				String layersnum = vals[17]; // 板层数
				String gagi = vals[18]; // GA /GI
				String sheetstrength440 = vals[19]; // 材料强度(Mpa)440
				String sheetstrength590 = vals[20]; // 材料强度(Mpa)590
				String sheetstrength = vals[21]; // 材料强度(Mpa)>590
				String basethickness = vals[22]; // 基准板厚
				String sheetstrength12 = vals[23]; // 材料强度(Mpa)1.2G
				String CurrentSerie = ""; // 参数 序列 (日产)
				String RecomWeldForce = "";// 推荐 加压力(N)
				String CurrentSeriedfi = ""; // 参数 序列 (对应)

				// 根据材质对照表，判断焊点是否参与计算焊接参数
				boolean partmaterialFlag1 = true;
				boolean partmaterialFlag2 = true;
				boolean partmaterialFlag3 = true;
				String gagi1 = vals[27];
				String gagi2 = vals[28];
				String gagi3 = vals[29];

				// 根据材质对照表获取GA/GI属性
				if (MaterialMap != null) {
					for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
						String MaterialNo = entry.getKey();
						List<String> infolist = entry.getValue();
						if (!"GA".equals(gagi1) && !"GI".equals(gagi1)) {
							if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag1 = false;
								}
							}
						}
						if (!"GA".equals(gagi2) && !"GI".equals(gagi2)) {
							if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag2 = false;
								}
							}
						}
						if (!"GA".equals(gagi3) && !"GI".equals(gagi3)) {
							if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag3 = false;
								}
							}
						}
					}
				}
				// 排除不参与计算的焊点
				if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
					if (paramap.containsKey(weldno)) {
						String[] curenre = paramap.get(weldno);
						CurrentSerie = curenre[1];
						RecomWeldForce = curenre[0];
						CurrentSeriedfi = curenre[12];
					}
				}
				boolean flag = false;
				// 如果是1.2g高强材，基准板厚都是取最薄板
				if (sheetstrength12.equals("1.2g")) {
					flag = true;
				}
				if (flag) {
					basethickness = getMinnum(vals[8], vals[12], vals[16]);
				}
				setStringCellAndStyle(sheet, importance, rowindex, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, weldno, rowindex, 8, style6, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, boardnumber1, rowindex, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, boardname1, rowindex, 16, style, Cell.CELL_TYPE_STRING);
				if (getIscontains1180(partmaterial1)) {
					XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, rowindex, 29, -1, new XSSFColor(new java.awt.Color(255,199,206)));
					setStringCellAndStyle2(sheet, partmaterial1, rowindex, 29, newstyle, Cell.CELL_TYPE_STRING);
				} else {
					XSSFCellStyle newstyle = getXSSFStyle(book, sheet,rowindex, 29, -1, IndexedColors.WHITE.getIndex());
					setStringCellAndStyle2(sheet, partmaterial1, rowindex, 29, newstyle, Cell.CELL_TYPE_STRING);
				}
//				if (!partmaterialFlag1) {
//					setStringCellAndStyle2(sheet, partmaterial1, rowindex, 29, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, partmaterial1, rowindex, 29, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, partthickness1, rowindex, 36, style, 11);
				setStringCellAndStyle(sheet, boardnumber2, rowindex, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, boardname2, rowindex, 42, style, Cell.CELL_TYPE_STRING);
				if (getIscontains1180(partmaterial2)) {
					XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, rowindex, 55, -1, new XSSFColor(new java.awt.Color(255,199,206)));
					setStringCellAndStyle2(sheet, partmaterial2, rowindex, 55, newstyle, Cell.CELL_TYPE_STRING);
				} else {
					XSSFCellStyle newstyle = getXSSFStyle(book, sheet,rowindex, 55, -1, IndexedColors.WHITE.getIndex());
					setStringCellAndStyle2(sheet, partmaterial2, rowindex, 55, newstyle, Cell.CELL_TYPE_STRING);
				}
//				if (!partmaterialFlag2) {
//					setStringCellAndStyle2(sheet, partmaterial2, rowindex, 55, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, partmaterial2, rowindex, 55, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, partthickness2, rowindex, 62, style, 11);
				setStringCellAndStyle(sheet, boardnumber3, rowindex, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, boardname3, rowindex, 68, style, Cell.CELL_TYPE_STRING);
				if (getIscontains1180(partmaterial3)) {
					XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, rowindex, 81, -1, new XSSFColor(new java.awt.Color(255,199,206)));
					setStringCellAndStyle2(sheet, partmaterial3, rowindex, 81, newstyle, Cell.CELL_TYPE_STRING);
				} else {
					XSSFCellStyle newstyle = getXSSFStyle(book, sheet,rowindex, 81, -1, IndexedColors.WHITE.getIndex());
					setStringCellAndStyle2(sheet, partmaterial3, rowindex, 81, newstyle, Cell.CELL_TYPE_STRING);
				}
//				if (!partmaterialFlag3) {
//					setStringCellAndStyle2(sheet, partmaterial3, rowindex, 81, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, partmaterial3, rowindex, 81, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, partthickness3, rowindex, 87, style, 11);
				setStringCellAndStyle(sheet, layersnum, rowindex, 90, style, 10);
				setStringCellAndStyle(sheet, gagi, rowindex, 92, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, sheetstrength440, rowindex, 94, style, 10);
				setStringCellAndStyle(sheet, sheetstrength590, rowindex, 96, style, 10);
				setStringCellAndStyle(sheet, sheetstrength, rowindex, 98, style, 10);
				if (flag) {
					setStringCellAndStyle(sheet, "○", rowindex, 100, style, Cell.CELL_TYPE_STRING);
					if (getColorDistinction(layersnum, partmaterial1, partmaterial2, partmaterial3, partthickness1,
							partthickness2, partthickness3)) {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, rowindex, 102, 1, IndexedColors.SKY_BLUE.getIndex());
						setStringCellAndStyle2(sheet, basethickness, rowindex, 102, newstyle, 11);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, rowindex, 102, 1, IndexedColors.VIOLET.getIndex());
						setStringCellAndStyle2(sheet, basethickness, rowindex, 102, newstyle, 11);
					}
					// 后面参数序列为空
					setStringCellAndStyle(sheet, "", rowindex, 105, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, "", rowindex, 108, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, "", rowindex, 111, style, Cell.CELL_TYPE_STRING);
				} else {
					XSSFCellStyle newstyle = getXSSFStyle(book, sheet, rowindex, 102, 12, IndexedColors.WHITE.getIndex());
					setStringCellAndStyle(sheet, basethickness, rowindex, 102, newstyle, 11);
					setStringCellAndStyle(sheet, "-", rowindex, 100, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, CurrentSerie, rowindex, 105, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, RecomWeldForce, rowindex, 108, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, CurrentSeriedfi, rowindex, 111, style, Cell.CELL_TYPE_STRING);
				}
			}
		}
	}
	private XSSFCellStyle getXSSFStyleByrgb(XSSFWorkbook book,XSSFSheet sheet,int rowindex,int cellindex,int colorindex,XSSFColor bgcolor)
	{
		XSSFRow row = sheet.getRow(rowindex);
		if(row!=null)
		{
			XSSFCell cell = row.getCell(cellindex);
			if(cell!=null)
			{
				XSSFCellStyle style = cell.getCellStyle();
				XSSFCellStyle newstyle = book.createCellStyle();
				newstyle = (XSSFCellStyle) style.clone();
				if(bgcolor != null)
				{
					newstyle.setFillForegroundColor(bgcolor);
					newstyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
				}
				if(colorindex > -1)
				{
					// 设置字体颜色
					Font font = book.createFont();
					Font sourcefont = style.getFont();
					font.setColor((short) colorindex);
					font.setFontHeightInPoints(sourcefont.getFontHeightInPoints());
					font.setFontName(sourcefont.getFontName());
					newstyle.setFont(font);
				}
			    return newstyle;
			}
		}
		return null;
	}

	/*
	 * 判断如果是1.2g高强才，需要判断菱形/五星决定背景色，菱形为蓝色true，五星为紫色false
	 */
	private boolean getColorDistinction(String layersnum, String partmaterial1, String partmaterial2,
			String partmaterial3, String partthickness1, String partthickness2, String partthickness3) {
		boolean flag = false;
		if (layersnum != null && !layersnum.isEmpty()) {
			int bznum = Integer.parseInt(layersnum);// 板层数
			if (bznum == 1) {
				flag = true;
			} else if (bznum == 2) { // 两层板情况
				// 先判断那块是1.2g高强才
				boolean flag1 = getIscontains1180(partmaterial1);
				boolean flag2 = getIscontains1180(partmaterial2);
				boolean flag3 = getIscontains1180(partmaterial3);
				// 有一块板材为空，需要分三种情况
				if (partmaterial1 == null || partmaterial1.isEmpty()) {
					// 如果都是1.2g高强才
					if (flag2 && flag3) {
						flag = false;
					} else {
						flag = getCompareresultByTwo(partmaterial2, partmaterial3, partthickness2, partthickness3,
								flag2, flag3);
					}
				} else if (partmaterial2 == null || partmaterial2.isEmpty()) {
					// 如果都是1.2g高强才
					if (flag1 && flag3) {
						flag = false;
					} else {
						flag = getCompareresultByTwo(partmaterial1, partmaterial3, partthickness1, partthickness3,
								flag1, flag3);
					}
				} else {
					// 如果都是1.2g高强才
					if (flag1 && flag2) {
						flag = false;
					} else {
						flag = getCompareresultByTwo(partmaterial1, partmaterial2, partthickness1, partthickness2,
								flag1, flag2);
					}
				}
			} else { // 三层板情况
				// 先判断那块是1.2g高强才
				boolean flag1 = getIscontains1180(partmaterial1);
				boolean flag2 = getIscontains1180(partmaterial2);
				boolean flag3 = getIscontains1180(partmaterial3);
				// 三层板强度都是1.2G ，五星
				if (flag1 && flag2 && flag3) {
					flag = false;
				} else if (!flag1 && flag2 && flag3) { // 板材1为非1.2g，存在两块1.2g板材
					// 先取1.2g中的薄板
					double h2 = getDoubleByString(partthickness2);
					double h3 = getDoubleByString(partthickness3);
					if (h2 < h3) {
						flag = getCompareresultByTwo(partmaterial1, partmaterial2, partthickness1, partthickness2,
								flag1, flag2);
					} else {
						flag = getCompareresultByTwo(partmaterial1, partmaterial3, partthickness1, partthickness3,
								flag1, flag3);
					}

				} else if (flag1 && !flag2 && flag3) { // 板材2为非1.2g，存在两块1.2g板材
					// 先取1.2g中的薄板
					double h1 = getDoubleByString(partthickness1);
					double h3 = getDoubleByString(partthickness3);
					if (h1 < h3) {
						flag = getCompareresultByTwo(partmaterial2, partmaterial1, partthickness2, partthickness1,
								flag2, flag1);
					} else {
						flag = getCompareresultByTwo(partmaterial2, partmaterial3, partthickness2, partthickness3,
								flag2, flag3);
					}
				} else if (flag1 && flag2 && !flag3) { // 板材3为非1.2g，存在两块1.2g板材
					// 先取1.2g中的薄板
					double h1 = getDoubleByString(partthickness1);
					double h2 = getDoubleByString(partthickness2);
					if (h1 < h2) {
						flag = getCompareresultByTwo(partmaterial3, partmaterial1, partthickness3, partthickness1,
								flag3, flag1);
					} else {
						flag = getCompareresultByTwo(partmaterial3, partmaterial2, partthickness3, partthickness2,
								flag3, flag2);
					}
				} else {// 只有一块为1.2g高强才
					double h1 = getDoubleByString(partthickness1);
					double h2 = getDoubleByString(partthickness2);
					double h3 = getDoubleByString(partthickness3);
					int kn1 = getSheetstrength(partmaterial1);
					int kn2 = getSheetstrength(partmaterial2);
					int kn3 = getSheetstrength(partmaterial3);

					if (h1 != h2 && h1 != h3 && h2 != h3) { // 板厚不相等
						// 1.2G 是最薄板，五星（板厚不相等）
						if (flag1) {
							if (h1 < h2 && h1 < h3) {
								flag = false;
							} else { // 1.2G是最厚板，菱形（板厚不相等） 1.2G板厚居中，菱形（板厚不相等）
								flag = true;
							}
						} else if (flag2) {
							if (h2 < h1 && h2 < h3) {
								flag = false;
							} else { // 1.2G是最厚板，菱形（板厚不相等） 1.2G板厚居中，菱形（板厚不相等）
								flag = true;
							}
						} else {
							if (h3 < h1 && h3 < h2) {
								flag = false;
							} else { // 1.2G是最厚板，菱形（板厚不相等） 1.2G板厚居中，菱形（板厚不相等）
								flag = true;
							}
						}
					} else { // 存在1.2g与其他板件厚度相同，则比较强度，如果另外两块板的强度都比1.2G高，五星；其他情况，菱形
						if (flag1) {
							if (kn1 < kn2 && kn1 < kn3) {
								flag = false;
							} else {
								flag = true;
							}
						} else if (flag2) {
							if (kn2 < kn1 && kn2 < kn3) {
								flag = false;
							} else {
								flag = true;
							}
						} else {
							if (kn3 < kn1 && kn3 < kn2) {
								flag = false;
							} else {
								flag = true;
							}
						}
					}
				}
			}
		}
		return flag;
	}

	/*
	 * 两层板的比较
	 */
	private boolean getCompareresultByTwo(String partmaterial, String partmateria2, String partthickness1,
			String partthickness2, boolean flag1, boolean flag2) {
		boolean flag = false;
		// 判断板厚是否相同
		if (partthickness1.equals(partthickness2)) { // 板厚相同，再判断强度
			int kn1 = getSheetstrength(partmaterial);
			int kn2 = getSheetstrength(partmateria2);
			// 另外的板强度都比1.2G板低，菱形
			if (flag1) {
				if (kn1 > kn2) {
					flag = true;
				} else {
					flag = false;
				}
			} else {
				if (kn1 > kn2) {
					flag = false;
				} else {
					flag = true;
				}
			}
		} else {// 板厚不相同，1.2G 是最薄板，五星
			double high = 0.0;
			double ordinary = 0.0;
			if (flag1) {
				high = getDoubleByString(partthickness1);
				ordinary = getDoubleByString(partthickness2);
			} else {
				ordinary = getDoubleByString(partthickness1);
				high = getDoubleByString(partthickness2);
			}
			if (high > ordinary) {
				flag = true;
			} else {
				flag = false;
			}
		}
		return flag;
	}

	/*
	 * 根据材料获取强度
	 */
	private int getSheetstrength(String partmaterial) {
		int tkness = 0;
		if (partmaterial != null && !partmaterial.isEmpty()) {

			String Sheetstrength = "";
			String[] str = partmaterial.split("-");
			if (str.length > 1) {
				String tempstr = str[1].trim();
				if (tempstr != null && !"".equals(tempstr)) {
					for (int K = 0; K < tempstr.length(); K++) {
						if (tempstr.charAt(K) >= 48 && tempstr.charAt(K) <= 57) {
							Sheetstrength += tempstr.charAt(K);
						}
					}
				}
				if (!Sheetstrength.isEmpty()) {
					tkness = Integer.parseInt(Sheetstrength);
				}
			}
		}

		return tkness;
	}

	/*
	 * 字符串转为double型，为空默认为0.0
	 */
	private double getDoubleByString(String str) {
		double num = 0.0;
		if (str != null && !str.isEmpty()) {
			num = Double.parseDouble(str);
		}
		return num;
	}

	/*
	 * 判断材质是否含有1180，也就是高强才
	 */
	private boolean getIscontains1180(String partmaterial1) {
		boolean flag = false;
		if (partmaterial1 != null && !partmaterial1.isEmpty()) {
			String Sheetstrength = "";
			String[] str = partmaterial1.split("-");
			if (str.length > 1) {
				String tempstr = str[1].trim();
				if (tempstr != null && !"".equals(tempstr)) {
					for (int K = 0; K < tempstr.length(); K++) {
						if (tempstr.charAt(K) >= 48 && tempstr.charAt(K) <= 57) {
							Sheetstrength += tempstr.charAt(K);
						}
					}
				}
			}
			if (Sheetstrength.equals("1180")) {
				flag = true;
			}
		}
		return flag;
	}

	private void DealRSWQDSheet(XSSFWorkbook book, ReportViwePanel viewPanel) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		ArrayList sheetAtIndexs = new ArrayList(); // RSW气动所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("RSW气动")) {
				sheetAtIndexs.add(i);
			}
		}
		if (sheetAtIndexs == null && sheetAtIndexs.size() < 1) {
			return;
		}
		// 获取sheet内的数据
		ArrayList datalist = getSheetData(book, sheetAtIndexs, false);

		ArrayList hdlist = new ArrayList();

		// 根据焊点在基本信息表中获取板件信息
		List hdinfo = new ArrayList();// 焊点所有信息

		hdinfo = getBoardInformation(baseinfolist, datalist);

		// 循环焊点信息，计算并获取参数属性值
		Map<String, String[]> paramap = getDealParameter(hdinfo, false, viewPanel);

		// 根据新增焊点的位置，将获取到的信息写入到报表中
		writeRSWQDInfomation(book, hdinfo, paramap);
	}

	private void writeRSWQDInfomation(XSSFWorkbook book, List hdinfo, Map<String, String[]> paramap) {
		// TODO Auto-generated method stub
		if (hdinfo != null && hdinfo.size() > 0) {
			// 设置字体颜色
			Font font = book.createFont();
			font.setColor((short) 12);// 蓝色字体
			font.setFontHeightInPoints((short) 9);
			XSSFCellStyle style = book.createCellStyle();
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style.setFont(font);

			Font font2 = book.createFont();
			font2.setColor((short) 12);// 蓝色字体
			font2.setFontHeightInPoints((short) 18);
			font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderTop(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font2);

			XSSFCellStyle style3 = book.createCellStyle();
			style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style3.setFont(font);

			XSSFCellStyle style4 = book.createCellStyle();
			style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			// style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style4.setFont(font);

			Font font3 = book.createFont();
			font3.setColor((short) 2);// 红色字体
			font3.setFontHeightInPoints((short) 10);
			XSSFCellStyle style33 = book.createCellStyle();
			style33.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style33.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style33.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style33.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style33.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style33.setFont(font3);

			// 粉色背景色
			Font fontpink = book.createFont();
			fontpink.setColor((short) 12);// 蓝色字体
			fontpink.setFontName("MS PGothic");
			fontpink.setFontHeightInPoints((short) 9);

			XSSFCellStyle stylepink = book.createCellStyle();
			stylepink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
			stylepink.setFillPattern(CellStyle.SOLID_FOREGROUND);
			stylepink.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			stylepink.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			stylepink.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			stylepink.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			stylepink.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			stylepink.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			stylepink.setFont(fontpink);

			for (int i = 0; i < hdinfo.size(); i++) {
				String[] vals = (String[]) hdinfo.get(i);
				int sheetindex = Integer.parseInt(vals[2]); // sheet所在位置
				int rowindex = Integer.parseInt(vals[1]); // 所在行数
				XSSFSheet sheet = book.getSheetAt(sheetindex);

				// 根据材质对照表，判断焊点是否参与计算焊接参数
				boolean partmaterialFlag1 = true;
				boolean partmaterialFlag2 = true;
				boolean partmaterialFlag3 = true;
				String partmaterial1 = vals[7];
				String partmaterial2 = vals[11];
				String partmaterial3 = vals[15];
				String gagi1 = vals[27];
				String gagi2 = vals[28];
				String gagi3 = vals[29];

				// 根据材质对照表获取GA/GI属性
				if (MaterialMap != null) {
					for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
						String MaterialNo = entry.getKey();
						List<String> infolist = entry.getValue();
						if (!"GA".equals(gagi1) && !"GI".equals(gagi1)) {
							if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag1 = false;
								}
							}
						}
						if (!"GA".equals(gagi2) && !"GI".equals(gagi2)) {
							if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag2 = false;
								}
							}
						}
						if (!"GA".equals(gagi3) && !"GI".equals(gagi3)) {
							if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag3 = false;
								}
							}
						}
					}
				}

				setStringCellAndStyle(sheet, vals[4], rowindex, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[3], rowindex, 8, style3, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[5], rowindex, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[6], rowindex, 16, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag1) {
//					setStringCellAndStyle2(sheet, vals[7], rowindex, 29, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[7], rowindex, 29, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[7], rowindex, 29, style, Cell.CELL_TYPE_STRING);
				
				setStringCellAndStyle(sheet, vals[8], rowindex, 36, style, 11);
				setStringCellAndStyle(sheet, vals[9], rowindex, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[10], rowindex, 42, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag2) {
//					setStringCellAndStyle2(sheet, vals[11], rowindex, 55, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[11], rowindex, 55, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[11], rowindex, 55, style, Cell.CELL_TYPE_STRING);
				
				setStringCellAndStyle(sheet, vals[12], rowindex, 62, style, 11);
				setStringCellAndStyle(sheet, vals[13], rowindex, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[14], rowindex, 68, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag3) {
//					setStringCellAndStyle2(sheet, vals[15], rowindex, 81, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[15], rowindex, 81, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[15], rowindex, 81, style, Cell.CELL_TYPE_STRING);
				
				setStringCellAndStyle(sheet, vals[16], rowindex, 88, style, 11);
				setStringCellAndStyle(sheet, vals[17], rowindex, 91, style, 10);
				setStringCellAndStyle(sheet, vals[18], rowindex, 93, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[19], rowindex, 95, style, 10);
				setStringCellAndStyle(sheet, vals[20], rowindex, 97, style, 10);
				setStringCellAndStyle(sheet, vals[21], rowindex, 99, style, 10);
				setStringCellAndStyle(sheet, vals[22], rowindex, 102, style, 11);
				// 如果是1.2g高强材，基准板厚都是取最薄板
				if (vals[23].equals("1.2g")) {
					setStringCellAndStyle(sheet, "", rowindex, 105, style, 10);
					setStringCellAndStyle(sheet, "", rowindex, 108, style, 10);
					setStringCellAndStyle(sheet, "", rowindex, 111, style, 10);
				} else {
					if (paramap.containsKey(vals[3])) {
						String[] paras = paramap.get(vals[3]);
						String poweroncurent2 = "";
						String CurrentSerie = "";
						String RecomWeldForce = "";
						// 排除不参与计算的焊点
						if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
							poweroncurent2 = paras[7];
							CurrentSerie = paras[1];
							RecomWeldForce = paras[0];
						}
						setStringCellAndStyle(sheet, CurrentSerie, rowindex, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, rowindex, 108, style, 10);
						setStringCellAndStyle(sheet, CurrentSerie, rowindex, 111, style, 10);
					}
				}
			}
		}
	}

	private void DealPSWSheet(XSSFWorkbook book, ReportViwePanel viewPanel, String result) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		ArrayList sheetAtIndexs = new ArrayList(); // PSW所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("PSW") && !sheetname.contains("点焊")) {
				sheetAtIndexs.add(i);
			}
		}
		if (sheetAtIndexs == null && sheetAtIndexs.size() < 1) {
			return;
		}
		// 获取sheet内的数据
		ArrayList datalist = getSheetData(book, sheetAtIndexs, false);

		ArrayList hdlist = new ArrayList();

		// 根据焊点在基本信息表中获取板件信息
		List hdinfo = new ArrayList();// 焊点所有信息

		hdinfo = getBoardInformation(baseinfolist, datalist);

		// 循环焊点信息，计算并获取参数属性值
		Map<String, String[]> paramap = getDealParameter(hdinfo, false, viewPanel);

		// 根据新增焊点的位置，将获取到的信息写入到报表中
		writePSWInfomation(book, hdinfo, paramap);

		// 重新获取焊点电流值
		ArrayList datalist2 = getSheetData(book, sheetAtIndexs, false);

		// 根据焊枪编号，重新计算焊枪的参数
		Map<String, String[]> Calculapara = getCalculapara(datalist2, paramap);

		System.out.println(Calculapara);
		// 重新写入计算好的焊接参数
		writePSWParaInfo(book, Calculapara, result);

	}

	private void writePSWParaInfo(XSSFWorkbook book, Map<String, String[]> calculapara, String result) {
		// TODO Auto-generated method stub
		if (calculapara != null && calculapara.size() > 0) {
			Font font2 = book.createFont();
			font2.setColor((short) 12);// 蓝色字体
			font2.setFontHeightInPoints((short) 11);
			font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderTop(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font2);

			XSSFCellStyle style6 = book.createCellStyle();
			style6.setBorderBottom(XSSFCellStyle.BORDER_NONE); // 下边框
			style6.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);// 左边框
			style6.setBorderTop(XSSFCellStyle.BORDER_NONE);// 上边框
			style6.setBorderRight(XSSFCellStyle.BORDER_NONE);// 右边框

			XSSFCellStyle style7 = book.createCellStyle();
			style7.setBorderBottom(XSSFCellStyle.BORDER_NONE); // 下边框
			style7.setBorderLeft(XSSFCellStyle.BORDER_NONE);// 左边框
			style7.setBorderTop(XSSFCellStyle.BORDER_NONE);// 上边框
			style7.setBorderRight(XSSFCellStyle.BORDER_NONE);// 右边框

			for (Map.Entry<String, String[]> entry : calculapara.entrySet()) {
				String shindex = entry.getKey();
				String[] values = entry.getValue();
				if (Util.isNumber(shindex)) {
					int index = Integer.parseInt(shindex);
					XSSFSheet sheet = book.getSheetAt(index);
					XSSFRow terow = sheet.getRow(48);
					XSSFCell tecell = terow.getCell(108);
					String preedtion = tecell.getStringCellValue();
					boolean teflag = getIsSOPAfter(preedtion);
					System.out.println("是否为SOP后：" + teflag);
					if (!teflag) // 写会数据
					{
						// 只有选择重新计算，才会写入重新计算的参数值
						if ("是".equals(result)) {
							setStringCellAndStyle(sheet, "加压力", 5, 36, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "预压时间", 5, 42, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "上升时间", 5, 48, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "第一          通电时间", 5, 54, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "第一          通电电流", 5, 60, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "冷却时间一", 5, 66, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "第二          通电时间", 5, 72, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "第二          通电电流", 5, 78, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "冷却时间二", 5, 84, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "第三          通电时间", 5, 90, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "第三         通电电流", 5, 96, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "保持", 5, 102, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							setStringCellAndStyle(sheet, "焊钳额定压力", 5, 108, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
																											// ElectrodeVol
							// 再把计算的参数，写入
							for (int j = 0; j < values.length; j++) {
								setStringCellAndStyle(sheet, values[j], 7, 36 + j * 6, style2, Cell.CELL_TYPE_STRING);// 加压力~保持
							}
						}
					} else {
						// 清空焊接参数
						for (int j = 0; j < 3; j++) {
							for (int k = 0; k < 77; k++) {
								if (k == 0) {
									setStringCellAndStyle2(sheet, "", 5 + j, 36 + k, style6, Cell.CELL_TYPE_STRING);// 加压力~保持
								} else {
									setStringCellAndStyle2(sheet, "", 5 + j, 36 + k, style7, Cell.CELL_TYPE_STRING);// 加压力~保持
								}
							}
						}
					}

				}
			}
		}
	}

	private Map<String, String[]> getCalculapara(ArrayList datalist, Map<String, String[]> paramap) {
		// TODO Auto-generated method stub
		Map<String, String[]> map = new HashMap<String, String[]>();
		List tempguncode = new ArrayList();
		boolean sopflag = false;
		for (int i = 0; i < datalist.size(); i++) {
			String[] strVal = (String[]) datalist.get(i);
			String shindex = strVal[2];
			System.out.println("sheet位置：" + shindex);
			if (shindex != null && !shindex.isEmpty()) { // 如果焊枪的编号为空，不计算，因为无法保证准确性
				if (!tempguncode.contains(shindex)) {
					tempguncode.add(shindex);
				}
			}
		}
		System.out.println(tempguncode);

		if (tempguncode.size() > 0) {
			for (int i = 0; i < tempguncode.size(); i++) {
				boolean isCucalPara = true; // 是否计算综合焊接参数
				int maxRepressure = 0;// 加压力最大值
				int minRepressure = 99999999;// 加压力最小值
				double sumrevalue = 0;// 总电流值
				// List pages = new ArrayList();// 枪对应的sheet页
				String guncode = (String) tempguncode.get(i);
				int nums = 0;
				for (int j = 0; j < datalist.size(); j++) {
					String[] values = (String[]) datalist.get(j);

					// 根据材质对照表，判断焊点是否参与计算焊接参数
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String partmaterial1 = "";
					String partmaterial2 = "";
					String partmaterial3 = "";
					if (guncode.equals(values[2])) {
						partmaterial1 = values[8];
						partmaterial2 = values[9];
						partmaterial3 = values[10];
					}

					// 根据材质对照表获取GA/GI属性
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag1 = false;
									isCucalPara = false;
								}
							}

							if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag2 = false;
									isCucalPara = false;
								}

							}

							if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag3 = false;
									isCucalPara = false;
								}
							}

						}
					}
					// 排除不参与计算的焊点
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (guncode.equals(values[2])) {
							// 电流单位转换
							String temppower = values[6];
							System.out.println("获取到的电流值:" + temppower);
							if (Util.isNumber(temppower)) {
								double curent = (Double.parseDouble(temppower) / 1000);
								temppower = Double.toString(curent);
							}
							String poweroncurent2 = temppower;// 第二通电电流
							String RecomWeldForce = values[5];// 推荐 加压力(N)
//							if (paramap.containsKey(values[3])) {
//								String[] Vals = paramap.get(values[3]);
//								poweroncurent2 = Vals[7];// 第二通电电流
//								RecomWeldForce = Vals[0];// 推荐 加压力(N)
//							}
							if (Util.isNumber(RecomWeldForce)) {
								int repress = Integer.parseInt(RecomWeldForce);
								if (minRepressure > repress) {
									minRepressure = repress;
								}
								if (maxRepressure < repress) {
									maxRepressure = repress;
								}
							}
							if (Util.isNumber(poweroncurent2)) {
								sumrevalue = sumrevalue + Double.parseDouble(poweroncurent2);

								System.out.println("平均值中间：" + sumrevalue + poweroncurent2);
							}
							nums++;

						}
					}

//					if (!pages.contains(values[2])) {
//						pages.add(values[2]);
//					}
				}
				System.out.println("最后平均值：" + sumrevalue);
				System.out.println("guncode：" + guncode);
				System.out.println("isCucalPara：" + isCucalPara);
				System.out.println("nums：" + nums);
				// 计算参数
				String[] tatolcurenre = new String[12];
				if (!isCucalPara || nums == 0) {
					tatolcurenre[0] = "";
					tatolcurenre[1] = "";
					tatolcurenre[2] = "";
					tatolcurenre[3] = "";
					tatolcurenre[4] = "";
					tatolcurenre[5] = "";
					tatolcurenre[6] = "";
					tatolcurenre[7] = "";
					tatolcurenre[8] = "";
					tatolcurenre[9] = "";
					tatolcurenre[10] = "";
					tatolcurenre[11] = "";
				} else {
					// 如果没计算参数，就不计算焊枪参数
					tatolcurenre = getAverageParameterValues(cv, maxRepressure, minRepressure, sumrevalue, nums);
					if (minRepressure == 99999999) {
						tatolcurenre[0] = "";
					}
					if (sumrevalue == 0) {
						tatolcurenre[1] = "";
						tatolcurenre[2] = "";
						tatolcurenre[3] = "";
						tatolcurenre[4] = "";
						tatolcurenre[5] = "";
						tatolcurenre[6] = "";
						tatolcurenre[7] = "";
						tatolcurenre[8] = "";
						tatolcurenre[9] = "";
						tatolcurenre[10] = "";
						tatolcurenre[11] = "";
					}
				}

				map.put((String) guncode, tatolcurenre);
				System.out.println("获取的MAp：" + map);
//				if (pages.size() > 0) {
//					for (int k = 0; k < pages.size(); k++) {
//						map.put((String) pages.get(k), tatolcurenre);
//					}
//				}
			}
		}
		return map;
	}

	// 更新PSW信息
	private void writePSWInfomation(XSSFWorkbook book, List hdinfo, Map<String, String[]> paramap) {
		// TODO Auto-generated method stub
		if (hdinfo != null && hdinfo.size() > 0) {
			// 设置字体颜色
			Font font = book.createFont();
			font.setColor((short) 12);// 蓝色字体
			font.setFontHeightInPoints((short) 9);
			XSSFCellStyle style = book.createCellStyle();
			style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style.setFont(font);

			Font font2 = book.createFont();
			font2.setColor((short) 12);// 蓝色字体
			font2.setFontHeightInPoints((short) 11);
			font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
			XSSFCellStyle style2 = book.createCellStyle();
			style2.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderLeft(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setBorderTop(CellStyle.BORDER_MEDIUM); // 粗线边框
			style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style2.setFont(font2);

			Font font3 = book.createFont();
			font3.setColor((short) 12);// 蓝色字体
			font3.setFontHeightInPoints((short) 14);
			font3.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
			XSSFCellStyle style3 = book.createCellStyle();
			style3.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
			style3.setBorderLeft(CellStyle.BORDER_MEDIUM); // 粗线边框
			style3.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
			style3.setBorderTop(CellStyle.BORDER_MEDIUM); // 粗线边框
			style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style3.setFont(font3);

			XSSFCellStyle style4 = book.createCellStyle();
			style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			// style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style4.setFont(font);

			XSSFCellStyle style5 = book.createCellStyle();
			style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			// style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style5.setFont(font);

			// 设置字体颜色
			Font font4 = book.createFont();
			font4.setColor((short) 2);// 红色字体
			font4.setFontHeightInPoints((short) 10);

			XSSFCellStyle style44 = book.createCellStyle();
			style44.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			// style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			style44.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			style44.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			style44.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			style44.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			style44.setFont(font4);

			// 粉色背景色
			Font fontpink = book.createFont();
			fontpink.setColor((short) 12);// 蓝色字体
			fontpink.setFontName("MS PGothic");
			fontpink.setFontHeightInPoints((short) 9);

			XSSFCellStyle stylepink = book.createCellStyle();
			stylepink.setFillForegroundColor(IndexedColors.ROSE.getIndex());
			stylepink.setFillPattern(CellStyle.SOLID_FOREGROUND);
			stylepink.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
			stylepink.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
			stylepink.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
			stylepink.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
			stylepink.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			stylepink.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			stylepink.setFont(fontpink);

			for (int i = 0; i < hdinfo.size(); i++) {
				String[] vals = (String[]) hdinfo.get(i);
				int sheetindex = Integer.parseInt(vals[2]); // sheet所在位置
				int rowindex = Integer.parseInt(vals[1]); // 所在行数
				XSSFSheet sheet = book.getSheetAt(sheetindex);

				// 根据材质对照表，判断焊点是否参与计算焊接参数
				boolean partmaterialFlag1 = true;
				boolean partmaterialFlag2 = true;
				boolean partmaterialFlag3 = true;
				String partmaterial1 = vals[7];
				String partmaterial2 = vals[11];
				String partmaterial3 = vals[15];
				String gagi1 = vals[27];
				String gagi2 = vals[28];
				String gagi3 = vals[29];

				// 根据材质对照表获取GA/GI属性
				if (MaterialMap != null) {
					for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
						String MaterialNo = entry.getKey();
						List<String> infolist = entry.getValue();
						if (!"GA".equals(gagi1) && !"GI".equals(gagi1)) {
							if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag1 = false;
								}
							}
						}
						if (!"GA".equals(gagi2) && !"GI".equals(gagi2)) {
							if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag2 = false;
								}
							}
						}
						if (!"GA".equals(gagi3) && !"GI".equals(gagi3)) {
							if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
								if ("否".equalsIgnoreCase(infolist.get(1))) {
									partmaterialFlag3 = false;
								}
							}
						}
					}
				}

				setStringCellAndStyle(sheet, vals[4], rowindex, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[3], rowindex, 8, style4, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[5], rowindex, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[6], rowindex, 16, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag1) {
//					setStringCellAndStyle2(sheet, vals[7], rowindex, 29, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[7], rowindex, 29, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[7], rowindex, 29, style, Cell.CELL_TYPE_STRING);

				setStringCellAndStyle(sheet, vals[8], rowindex, 36, style, 11);
				setStringCellAndStyle(sheet, vals[9], rowindex, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[10], rowindex, 42, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag2) {
//					setStringCellAndStyle2(sheet, vals[11], rowindex, 55, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[11], rowindex, 55, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[11], rowindex, 55, style, Cell.CELL_TYPE_STRING);

				setStringCellAndStyle(sheet, vals[12], rowindex, 62, style, 11);
				setStringCellAndStyle(sheet, vals[13], rowindex, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[14], rowindex, 68, style, Cell.CELL_TYPE_STRING);
//				if (!partmaterialFlag3) {
//					setStringCellAndStyle2(sheet, vals[15], rowindex, 81, stylepink, Cell.CELL_TYPE_STRING);
//				} else {
//					setStringCellAndStyle2(sheet, vals[15], rowindex, 81, style, Cell.CELL_TYPE_STRING);
//				}
				setStringCellAndStyle(sheet, vals[15], rowindex, 81, style, Cell.CELL_TYPE_STRING);
				
				setStringCellAndStyle(sheet, vals[16], rowindex, 88, style, 11);
				setStringCellAndStyle(sheet, vals[17], rowindex, 91, style, 10);
				setStringCellAndStyle(sheet, vals[18], rowindex, 93, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, vals[19], rowindex, 95, style, 10);
				setStringCellAndStyle(sheet, vals[20], rowindex, 97, style, 10);
				setStringCellAndStyle(sheet, vals[21], rowindex, 99, style, 10);
				setStringCellAndStyle(sheet, vals[22], rowindex, 102, style, 11);

				// 如果是1.2g高强材，基准板厚都是取最薄板
				if (vals[23].equals("1.2g")) {
					setStringCellAndStyle(sheet, "", rowindex, 105, style, 10);
					setStringCellAndStyle(sheet, "", rowindex, 108, style, 10);
					setStringCellAndStyle(sheet, "", rowindex, 111, style, 10);
				} else {
					if (paramap.containsKey(vals[3])) {
						String[] paras = paramap.get(vals[3]);
						String poweroncurent2 = "";
						String CurrentSerie = "";
						String RecomWeldForce = "";

						// 排除不参与计算的焊点
						if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
							poweroncurent2 = paras[7];
							CurrentSerie = paras[1];
							RecomWeldForce = paras[0];
						}

						// 电流单位转换
						if (Util.isNumber(poweroncurent2)) {
							int curent = 0;
							curent = (int) (Double.parseDouble(poweroncurent2) * 1000);
							poweroncurent2 = Integer.toString(curent);
						}

						setStringCellAndStyle(sheet, CurrentSerie, rowindex, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, rowindex, 108, style, 10);
						setStringCellAndStyle(sheet, poweroncurent2, rowindex, 111, style, 10);
					}

				}
			}
		}

	}

	// 对单元格赋值
	public static void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// 对于整型与字符型的区分 10为整型，11为double型

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		// cell.setCellType(celltype);
		if (value == null || value.isEmpty()) {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		} else {
			if (celltype == Cell.CELL_TYPE_STRING) {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			} else if (celltype == 10) {
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Integer.parseInt(value));
			} else if (celltype == 11) {
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(value));
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			}
		}

		// cell.setCellStyle(Style);

	}

	// 对单元格赋值
	public static void setStringCellAndStyle2(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// 对于整型与字符型的区分 10为整型，11为double型

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		// cell.setCellType(celltype);
		if (value == null || value.isEmpty()) {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		} else {
			if (celltype == Cell.CELL_TYPE_STRING) {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			} else if (celltype == 10) {
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Integer.parseInt(value));
			} else if (celltype == 11) {
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(value));
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			}
		}

		cell.setCellStyle(Style);

	}

	/*
	 * 计算参数
	 */
	private Map<String, String[]> getDealParameter(List hdinfo, boolean flag, ReportViwePanel viewPanel) {
		// TODO Auto-generated method stub
		Map<String, String[]> paramap = new HashMap<String, String[]>();
		for (int j = 0; j < hdinfo.size(); j++) {
			String[] str = new String[13];
			String[] wbinfo = (String[]) hdinfo.get(j);
			String wbNO = wbinfo[3];
			String boradnum = wbinfo[17];// 板层数
			String basethickness = wbinfo[22];// 基准板厚
			String sheetstrength1 = wbinfo[24];// 板件1强度
			String sheetstrength2 = wbinfo[25];// 板件2强度
			String sheetstrength3 = wbinfo[26];// 板件3强度
			String partthickness1 = wbinfo[8];// 板件1厚度
			String partthickness2 = wbinfo[12];// 板件2厚度
			String partthickness3 = wbinfo[16];// 板件3厚度
			String gagi1 = wbinfo[27];// 板件1GA/GI材
			String gagi2 = wbinfo[28];// 板件2GA/GI材
			String gagi3 = wbinfo[29];// 板件3GA/GI材

			// 推荐加压力属性
			String Repressure = "";
			Repressure = getRepressure(basethickness, boradnum, sheetstrength1, sheetstrength2, sheetstrength3);
			str[0] = Repressure;
			// 24序列焊接条件设定表 参数序列号
			String parameterSerialNo24 = "";
			parameterSerialNo24 = getParameterSerialNo24(basethickness, boradnum, gagi1, gagi2, gagi3, sheetstrength1,
					sheetstrength2, sheetstrength3);

			str[1] = parameterSerialNo24;

			// 255序列焊接条件设定表 参数序列号
			// 获取板厚度差
			double thicknessdifference = getThicknessDifference(partthickness1, partthickness2, partthickness3,
					boradnum);
			String parameterSerialNo255 = "";
			parameterSerialNo255 = getParameterSerialNo255(basethickness, boradnum, gagi1, gagi2, gagi3, sheetstrength1,
					sheetstrength2, sheetstrength3, thicknessdifference);
			// 电流参数序列（日产） 需要区分气动和伺服 ，逻辑待定
			if (flag) {
				// 只有RSW伺服计算推荐序列（对应）
				str[1] = parameterSerialNo255;
				// 255序列对照表
				String SequenceComparison = "";
				SequenceComparison = getSequenceComparison(parameterSerialNo255);
				str[12] = SequenceComparison;
				str[2] = "";
				str[3] = "";
				str[4] = "";
				str[5] = "";
				str[6] = "";
				str[7] = "";
				str[8] = "";
				str[9] = "";
				str[10] = "";
				str[11] = "";
			} else {
				// 只有PSW和RSW启动需要计算电流值
				// 24序列焊接条件设定表 推荐 电流值
				String[] recommendedvalue = getRecommendedvalue(parameterSerialNo24);
				str[2] = recommendedvalue[0];
				str[3] = recommendedvalue[1];
				str[4] = recommendedvalue[2];
				str[5] = recommendedvalue[3];
				str[6] = recommendedvalue[4];
				str[7] = recommendedvalue[5];
				str[8] = recommendedvalue[6];
				str[9] = recommendedvalue[7];
				str[10] = recommendedvalue[8];
				str[11] = recommendedvalue[9];
				str[12] = "";
			}
			paramap.put(wbNO, str);

		}

		return paramap;
	}

	// 24序列焊接条件设定表 推荐 电流值
	public static String[] getRecommendedvalue(String parameterSerialNo24) {
		// TODO Auto-generated method stub
		String[] recommendedvalue = new String[10];
		for (int i = 0; i < cv.size(); i++) {
			CurrentandVoltage cvotage = cv.get(i);
			String serialNo = cvotage.getSequenceNo();
			if (serialNo != null && serialNo.equals(parameterSerialNo24)) {
				recommendedvalue[0] = cvotage.getBvalue();// 上升时间
				recommendedvalue[1] = cvotage.getCvalue();// 第一 通电时间
				recommendedvalue[2] = cvotage.getEvalue();// 第一 通电电流
				recommendedvalue[3] = cvotage.getFvalue();// 冷却时间一
				recommendedvalue[4] = cvotage.getGvalue();// 第二通电时间
				recommendedvalue[5] = cvotage.getIvalue();// 第二通电电流
				recommendedvalue[6] = cvotage.getJvalue();// 冷却时间二
				recommendedvalue[7] = cvotage.getKvalue();// 第三 通电时间
				recommendedvalue[8] = cvotage.getMvalue();// 第三 通电电流
				recommendedvalue[9] = cvotage.getNvalue();// 保持
				break;
			}
		}
		return recommendedvalue;
	}

	// 255序列对照表
	public static String getSequenceComparison(String parameterSerialNo255) {
		// TODO Auto-generated method stub
		String SequenceComparison = "";
		for (int i = 0; i < sct.size(); i++) {
			SequenceComparisonTable sctable = sct.get(i);
			Map<String, String> map = sctable.getValues();
			if (map.containsKey("S" + parameterSerialNo255)) {
				String value = map.get("S" + parameterSerialNo255);
				if (value.trim().length() < 2) {
					value = "0" + value;
				}
				SequenceComparison = sctable.getParameterGroup() + "-" + value;
				break;
			}
		}

		return SequenceComparison;
	}

	// 255序列焊接条件设定表 参数序列号
	public static String getParameterSerialNo255(String basethickness, String boradnum, String gagi1, String gagi2,
			String gagi3, String sheetstrength1, String sheetstrength2, String sheetstrength3,
			double thicknessdifference) {
		// TODO Auto-generated method stub
		String parameterSerialNo255 = "";
		int lnum = 0; // 裸板数量
		int ganum = 0; // GA材数量
		int high = 0;// 高强材数量
		if (gagi1.isEmpty()) {
			lnum++;
		}
		if (gagi2.isEmpty()) {
			lnum++;
		}
		if (gagi3.isEmpty()) {
			lnum++;
		}
		if (boradnum.equals("2")) {
			lnum--;
		}
		if (gagi1.equals("GA") || gagi1.equals("GI")) {
			ganum++;
		}
		if (gagi2.equals("GA") || gagi1.equals("GI")) {
			ganum++;
		}
		if (gagi3.equals("GA") || gagi1.equals("GI")) {
			ganum++;
		}
		if (!sheetstrength1.isEmpty()) {
			high++;
		}
		if (!sheetstrength2.isEmpty()) {
			high++;
		}
		if (!sheetstrength3.isEmpty()) {
			high++;
		}
		for (int i = 0; i < SFswc.size(); i++) {
			SFSequenceWeldingConditionList sfsw = SFswc.get(i);
			String value = sfsw.getBasethickness();
			if (Util.isNumber(value) && Util.isNumber(basethickness)) {
				if (Double.parseDouble(value) == Double.parseDouble(basethickness)) {
					if (boradnum.equals("2")) {
						if (lnum == 2 && high != 0) {
							parameterSerialNo255 = sfsw.getBvalue();
						}
						if (lnum == 2 && high == 0) {
							parameterSerialNo255 = sfsw.getCvalue();
						}
						if (ganum == 1 && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getDvalue();
						}
						if (ganum == 1 && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getEvalue();
						}
						if (ganum == 2 && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getFvalue();
						}
						if (ganum == 2 && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getGvalue();
						}
					}
					if (boradnum.equals("3")) {
						if (lnum == 3 && (high == 2 || high == 3)) {
							parameterSerialNo255 = sfsw.getHvalue();
						}
						if (lnum == 3 && (high == 0 || high == 1)) {
							parameterSerialNo255 = sfsw.getIvalue();
						}
						if (ganum == 1 && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getJvalue();
						}
						if (ganum == 1 && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getKvalue();
						}
						if ((ganum == 2 || ganum == 3) && thicknessdifference <= 2.4) {
							parameterSerialNo255 = sfsw.getLvalue();
						}
						if ((ganum == 2 || ganum == 3) && thicknessdifference > 2.4) {
							parameterSerialNo255 = sfsw.getMvalue();
						}
					}
					break;
				}
			}

		}
		return parameterSerialNo255;
	}

	// 获取板厚度差
	public static double getThicknessDifference(String partthickness1, String partthickness2, String partthickness3,
			String boradnum) {
		// TODO Auto-generated method stub
		double pk1;
		double pk2;
		double pk3;
		double thicknessdifference = 0;
		if (Util.isNumber(partthickness1)) {
			pk1 = Double.parseDouble(partthickness1);
		} else {
			pk1 = -1.0;
		}
		if (Util.isNumber(partthickness2)) {
			pk2 = Double.parseDouble(partthickness2);
		} else {
			pk2 = -1.0;
		}
		if (Util.isNumber(partthickness3)) {
			pk3 = Double.parseDouble(partthickness3);
		} else {
			pk3 = -1.0;
		}
		if (boradnum.equals("2")) {
			if (pk1 == -1.0) {
				if (pk2 < pk3) {
					thicknessdifference = pk3 / pk2;
				} else {
					thicknessdifference = pk2 / pk3;
				}
			}
			if (pk2 == -1.0) {
				if (pk1 < pk3) {
					thicknessdifference = pk3 / pk1;
				} else {
					thicknessdifference = pk1 / pk3;
				}
			}
			if (pk3 == -1.0) {
				if (pk1 < pk2) {
					thicknessdifference = pk2 / pk1;
				} else {
					thicknessdifference = pk1 / pk2;
				}
			}
		}
		if (boradnum.equals("3")) {
			String minstr = getMinnum(partthickness1, partthickness2, partthickness3);
			String maxstr = getMaxnum(partthickness1, partthickness2, partthickness3);
			double min = Double.parseDouble(minstr);
			double max = Double.parseDouble(maxstr);
			thicknessdifference = max / min;
		}
		BigDecimal bd = new BigDecimal(thicknessdifference);
		BigDecimal fact = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
		thicknessdifference = fact.doubleValue();
		return thicknessdifference;
	}

	// 24序列焊接条件设定表 参数序列号
	public static String getParameterSerialNo24(String basethickness, String boradnum, String gagi1, String gagi2,
			String gagi3, String sheetstrength1, String sheetstrength2, String sheetstrength3) {
		// TODO Auto-generated method stub
		String parameterSerialNo24 = "";
		int lnum = 0; // 裸板数量
		int ginum = 0; // GI材数量
		int ganum = 0; // GA材数量
		int high = 0;// 高强材数量

		if (gagi1.isEmpty()) {
			lnum++;
		}
		if (gagi2.isEmpty()) {
			lnum++;
		}
		if (gagi3.isEmpty()) {
			lnum++;
		}
		if (boradnum.equals("2")) {
			lnum--;
		}

		if (gagi1.equals("GA")) {
			ganum++;
		}
		if (gagi2.equals("GA")) {
			ganum++;
		}
		if (gagi3.equals("GA")) {
			ganum++;
		}
		if (gagi1.equals("GI")) {
			ginum++;
		}
		if (gagi2.equals("GI")) {
			ginum++;
		}
		if (gagi3.equals("GI")) {
			ginum++;
		}
		if (!sheetstrength1.isEmpty()) {
			high++;
		}
		if (!sheetstrength2.isEmpty()) {
			high++;
		}
		if (!sheetstrength3.isEmpty()) {
			high++;
		}
		for (int i = 0; i < swc.size(); i++) {
			SequenceWeldingConditionList swcl = swc.get(i);
			String thick = swcl.getBasethickness();
			if (Util.isNumber(thick) && Util.isNumber(basethickness)) {
				if (Double.parseDouble(thick) == Double.parseDouble(basethickness)) {
					// 当板板组中既有GA材又有GI材时，将GA材当做GI材来考虑。

					if (boradnum.equals("2")) {
						if (lnum == 2 && high == 0) {
							parameterSerialNo24 = swcl.getBvalue();
						}
						if (lnum == 2 && high != 0) {
							parameterSerialNo24 = swcl.getCvalue();
						}
						if (lnum == 1 && ganum == 1 && high == 0) {
							parameterSerialNo24 = swcl.getDvalue();
						}
						if (lnum == 1 && ganum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getEvalue();
						}
						if (lnum == 1 && ganum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getEvalue();
						}
						if (ganum == 2 && high == 0) {
							parameterSerialNo24 = swcl.getFvalue();
						}
						if (ganum == 2 && high != 0) {
							parameterSerialNo24 = swcl.getGvalue();
						}
						if (ginum == 1 && ganum == 1 && high == 0) {
							parameterSerialNo24 = swcl.getHvalue();
						}
						if (ginum == 1 && ganum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getIvalue();
						}
						if (ginum == 2 && high == 0) {
							parameterSerialNo24 = swcl.getJvalue();
						}
						if (ginum == 2 && high != 0) {
							parameterSerialNo24 = swcl.getKvalue();
						}
					}
					if (boradnum.equals("3")) {
						if (lnum == 3 && high == 0) {
							parameterSerialNo24 = swcl.getLvalue();
						}
						if (lnum == 3 && high != 0) {
							parameterSerialNo24 = swcl.getMvalue();
						}
						if (ganum == 1 && ginum == 0 && high == 0) {
							parameterSerialNo24 = swcl.getNvalue();
						}
						if (ganum == 1 && ginum == 0 && high != 0) {
							parameterSerialNo24 = swcl.getOvalue();
						}
						if (ganum == 2 && ginum == 0 && high == 0) {
							parameterSerialNo24 = swcl.getPvalue();
						}
						if (ganum == 2 && ginum == 0 && high != 0) {
							parameterSerialNo24 = swcl.getQvalue();
						}
						if (ganum == 3 && ginum == 0 && high == 0) {
							parameterSerialNo24 = swcl.getRvalue();
						}
						if (ganum == 3 && ginum == 0 && high != 0 && high != 3) {
							parameterSerialNo24 = swcl.getSvalue();
						}
						if (ganum == 3 && ginum == 0 && high == 3) {
							parameterSerialNo24 = swcl.getTvalue();
						}
						if (ganum == 0 && ginum == 1 && high == 0) {
							parameterSerialNo24 = swcl.getUvalue();
						}
						if (ganum == 0 && ginum == 1 && high != 0) {
							parameterSerialNo24 = swcl.getVvalue();
						}
						if (lnum == 1 && ganum != 2 && high == 0) {
							parameterSerialNo24 = swcl.getWvalue();
						}
						if (lnum == 1 && ganum != 2 && high != 0) {
							parameterSerialNo24 = swcl.getXvalue();
						}
						if (lnum == 0 && ganum != 3 && high == 0) {
							parameterSerialNo24 = swcl.getYvalue();
						}
						if (lnum == 0 && ganum != 3 && high != 0 && high != 3) {
							parameterSerialNo24 = swcl.getZvalue();
						}
						if (lnum == 0 && ganum != 3 && high == 3) {
							parameterSerialNo24 = swcl.getAAvalue();
						}
					}
					break;
				}
			}

		}

		return parameterSerialNo24;
	}

	// 获取加压力
	public static String getRepressure(String basethickness, String boradnum, String sheetstrength1,
			String sheetstrength2, String sheetstrength3) {
		// TODO Auto-generated method stub
		String repressure = "";
		String distinguish = ""; // 区分
		int num1 = 0;// 440以下数量
		int num2 = 0;// 440
		int num3 = 0;// 590Mpa780Mpa980Mpa
		// 如果存在1180强度板材，不计算参数列，默认为空
		int shstrength1 = getInteger(sheetstrength1);
		int shstrength2 = getInteger(sheetstrength2);
		int shstrength3 = getInteger(sheetstrength3);
		if (shstrength1 == 1180 || shstrength2 == 1180 || shstrength3 == 1180) {
			repressure = "";
			return repressure;
		}
		// 如果存在1350强度板材，区分为I
		if (shstrength1 == 1350 || shstrength2 == 1350 || shstrength3 == 1350) {
			distinguish = "Ⅰ";
		} else {
			if (sheetstrength1.isEmpty()) {
				num1++;
			}
			if (shstrength1 == 440) {
				num2++;
			}
			if (shstrength1 == 590 || shstrength1 == 780 || shstrength1 == 980) {
				num3++;
			}
			if (sheetstrength2.isEmpty()) {
				num1++;
			}
			if (shstrength2 == 440) {
				num2++;
			}
			if (shstrength2 == 590 || shstrength2 == 780 || shstrength2 == 980) {
				num3++;
			}
			if (sheetstrength3.isEmpty()) {
				num1++;
			}
			if (shstrength3 == 440) {
				num2++;
			}
			if (shstrength3 == 590 || shstrength3 == 780 || shstrength3 == 980) {
				num3++;
			}

			// 先根据两层板规则，获取分区,再根据3层板
			if (boradnum.equals("2")) {
				if (num1 == 3) {
					distinguish = "Ⅰ";
				}
				if (num3 == 0 && num2 == 1) {
					distinguish = "Ⅱ";
				}
				if (num2 == 0 && num3 == 1) {
					distinguish = "Ⅲ";
				}
				if (num2 == 2) {
					distinguish = "Ⅲ";
				}
				if (num2 == 1 && num3 == 1) {
					distinguish = "Ⅳ";
				}
				if (num3 == 2) {
					distinguish = "Ⅴ";
				}
			} else if (boradnum.equals("3")) {
				if (num1 == 3) {
					distinguish = "Ⅰ";
				}
				if (num1 == 2 && num2 == 1) {
					distinguish = "Ⅱ";
				}
				if (num1 == 2 && num3 == 1) {
					distinguish = "Ⅲ";
				}
				if (num1 == 1 && num2 == 2) {
					distinguish = "Ⅲ";
				}
				if (num1 == 1 && num2 == 1 && num3 == 1) {
					distinguish = "Ⅳ";
				}
				if (num1 == 1 && num3 == 2) {
					distinguish = "Ⅴ";
				}
				if (num2 == 3) {
					distinguish = "Ⅲ";
				}
				if (num2 == 2 && num3 == 1) {
					distinguish = "Ⅴ";
				}
				if (num2 == 1 && num3 == 2) {
					distinguish = "Ⅴ";
				}
				if (num3 == 3) {
					distinguish = "Ⅴ";
				}
			} else {
				repressure = "";
				return repressure;
			}

		}
		for (int i = 0; i < rp.size(); i++) {
			RecommendedPressure repre = rp.get(i);
			String thickness = repre.getBasethickness();
			if (Util.isNumber(thickness) && Util.isNumber(basethickness)) {
				if (Double.parseDouble(thickness) == Double.parseDouble(basethickness)) {
					if (distinguish.equals("Ⅰ")) {
						repressure = repre.getBvalue();
					}
					if (distinguish.equals("Ⅱ")) {
						repressure = repre.getCvalue();
					}
					if (distinguish.equals("Ⅲ")) {
						repressure = repre.getDvalue();
					}
					if (distinguish.equals("Ⅳ")) {
						repressure = repre.getEvalue();
					}
					if (distinguish.equals("Ⅴ")) {
						repressure = repre.getFvalue();
					}
					break;
				}
			}

		}

		return repressure;
	}

	// 获取sheet内的数据
	private ArrayList getSheetData(XSSFWorkbook book, ArrayList sheetAtIndexs, boolean flag) {
		// TODO Auto-generated method stub

		ArrayList resultDataList = new ArrayList();
		for (int i = 0; i < sheetAtIndexs.size(); i++) {
			int sheetindex = (int) sheetAtIndexs.get(i);
			Sheet sheet = book.getSheetAt(sheetindex);
			// 校验sheet是否合法
			if (sheet == null) {
				return null;
			}
			// 获取第一行数据
			int firstRowNum = sheet.getFirstRowNum();
			Row firstRow = (Row) sheet.getRow(firstRowNum);
			if (null == firstRow) {
				logger.warning("解析Excel失败，在第一行没有读取到任何数据！");
			}

			// 先获取焊枪型号
			String guncode = "";
			Row rowgun = (Row) sheet.getRow(5);
			if (rowgun != null) {
				Cell cell = rowgun.getCell(19);
				if (cell != null) {
					guncode = convertCellValueToString(cell);
				}
			}
			// 获取版次
			String edition = "";
			Row rowedition = (Row) sheet.getRow(48);
			if (rowedition != null) {
				Cell cell = rowedition.getCell(108);
				if (cell != null) {
					edition = convertCellValueToString(cell);
				}
			}

			// 解析每一行的数据，构造数据对象
			int rowStart = 11;
			int rowEnd = 47;
			for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
				Row row = (Row) sheet.getRow(rowNum);
				if (null == row) {
					continue;
				}
				if (flag) {
					String[] resultData = convertRowToData2(row);
					if (null == resultData) {
						logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
						continue;
					}
					resultData[0] = guncode;
					resultData[1] = Integer.toString(rowNum);// 所在行数
					resultData[2] = Integer.toString(sheetindex);// 所在sheet页位置
					resultData[7] = edition;// 版次

					resultDataList.add(resultData);
				} else {
					String[] resultData = convertRowToData(row);
					if (null == resultData) {
						logger.warning("第 " + row.getRowNum() + "行数据不合法，已忽略！");
						continue;
					}
					resultData[0] = guncode;
					resultData[1] = Integer.toString(rowNum);// 所在行数
					resultData[2] = Integer.toString(sheetindex);// 所在sheet页位置
					resultData[7] = edition;// 版次

					resultDataList.add(resultData);
				}

			}
		}

		return resultDataList;
	}

	private String[] convertRowToData(Row row) {
		// TODO Auto-generated method stub
		String[] data = new String[11];
		Cell cell;
		// 焊点号
		cell = row.getCell(8);
		String weldno = convertCellValueToString(cell);
		if (weldno == null || weldno.isEmpty()) {
			return null;
		}
		data[3] = weldno;

		// 版层数
		cell = row.getCell(91);
		String boradnum = convertCellValueToString(cell);
		data[4] = boradnum;
		// 推荐 加压力(N)
		cell = row.getCell(108);
		data[5] = convertCellValueToString(cell);
		// 推荐 电流值(A)
		cell = row.getCell(111);
		data[6] = convertCellValueToString(cell);

		// 材质1
		cell = row.getCell(29);
		data[8] = convertCellValueToString(cell);
		// 材质2
		cell = row.getCell(55);
		data[9] = convertCellValueToString(cell);
		// 材质3
		cell = row.getCell(81);
		data[10] = convertCellValueToString(cell);

		return data;
	}

	private String[] convertRowToData2(Row row) {
		// TODO Auto-generated method stub
		String[] data = new String[11];
		Cell cell;
		// 焊点号
		cell = row.getCell(8);
		String weldno = convertCellValueToString(cell);
		if (weldno == null || weldno.isEmpty()) {
			return null;
		}
		data[3] = weldno;
		// 版层数
		cell = row.getCell(90);
		String boradnum = convertCellValueToString(cell);
		data[4] = boradnum;
		// 推荐 加压力(N)
		cell = row.getCell(108);
		data[5] = convertCellValueToString(cell);
		// 推荐 电流值(A)
		cell = row.getCell(111);
		data[6] = convertCellValueToString(cell);

		// 材质1
		cell = row.getCell(29);
		data[8] = convertCellValueToString(cell);
		// 材质2
		cell = row.getCell(55);
		data[9] = convertCellValueToString(cell);
		// 材质3
		cell = row.getCell(81);
		data[10] = convertCellValueToString(cell);

		return data;
	}

	private static String convertCellValueToString(Cell cell) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {

		} else {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC: // 数字
				Double doubleValue = cell.getNumericCellValue();
				// 格式化科学计数法，取一位整数
				DecimalFormat df = new DecimalFormat("0");
				returnValue = df.format(doubleValue);
				break;
			case Cell.CELL_TYPE_STRING: // 字符串
				// cell.setCellType(Cell.CELL_TYPE_STRING);
				returnValue = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_BOOLEAN: // 布尔
				Boolean booleanValue = cell.getBooleanCellValue();
				returnValue = booleanValue.toString();
				break;
			case Cell.CELL_TYPE_BLANK: // 空值
				break;
			case Cell.CELL_TYPE_FORMULA: // 公式
				returnValue = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_ERROR: // 故障
				break;
			default:
				break;
			}
		}
		return returnValue;
	}

	// 计算平均参数
	private String[] getAverageParameterValues(List<CurrentandVoltage> cv, int maxRepressure, int minRepressure,
			double sumrevalue, int size) {
		// TODO Auto-generated method stub
		String[] values = new String[12];
		String press = "";// 加压力
		String Preloadingtime = "15c.";// 预压时间 默认这个值
		String uptime = "";// 上升时间
		String powerontime1 = "";// 第一 通电时间
		String poweroncurent1 = "";// 第一 通电电流
		String coolingtime1 = "";// 冷却时间一
		String powerontime2 = "";// 第二通电时间
		String poweroncurent2 = "";// 第二通电电流
		String coolingtime2 = "";// 冷却时间二
		String powerontime3 = "";// 第三 通电时间
		String poweroncurent3 = "";// 第三 通电电流
		String maintain = "";// 保持

		int prepress = (maxRepressure + minRepressure) / 2;
		press = Integer.toString(prepress);
		// 电流平均值
		BigDecimal biga1 = new BigDecimal(Double.toString(sumrevalue));
		BigDecimal bigsize = new BigDecimal(Double.toString(size));
		double average = biga1.divide(bigsize, 8, BigDecimal.ROUND_HALF_UP).doubleValue();

		System.out.println("平均值：" + average);
		// 255序列焊接条件设定表 电流电压
		CurrentandVoltage currentandVoltage = getCurrentandVoltage(average, cv);

		System.out.println("测试打印：" + currentandVoltage.getSequenceNo());

		// 测试数据
		System.out.println("测试数据为7.7：" + getCurrentandVoltage(7.7, cv).getSequenceNo());
//		System.out.println("测试数据为7.2：" + getCurrentandVoltage(7.2, cv).getSequenceNo());
//		System.out.println("测试数据为8.3：" + getCurrentandVoltage(8.3, cv).getSequenceNo());
//		System.out.println("测试数据为9.25：" + getCurrentandVoltage(9.25, cv).getSequenceNo());
//		System.out.println("测试数据为16.8：" + getCurrentandVoltage(16.8, cv).getSequenceNo());
//		System.out.println("测试数据为18：" + getCurrentandVoltage(18, cv).getSequenceNo());

		if (currentandVoltage != null) {
			uptime = currentandVoltage.getBvalue() + "c.";// 上升时间
			powerontime1 = currentandVoltage.getCvalue() + "c.";// 第一 通电时间
			poweroncurent1 = currentandVoltage.getEvalue() + "KA";// 第一 通电电流
			coolingtime1 = currentandVoltage.getFvalue() + "c.";// 冷却时间一
			powerontime2 = currentandVoltage.getGvalue() + "c.";// 第二通电时间
			poweroncurent2 = currentandVoltage.getIvalue() + "KA";// 第二通电电流
			coolingtime2 = currentandVoltage.getJvalue() + "c.";// 冷却时间二
			powerontime3 = currentandVoltage.getKvalue() + "c.";// 第三 通电时间
			poweroncurent3 = currentandVoltage.getMvalue() + "KA";// 第三 通电电流
			maintain = currentandVoltage.getNvalue() + "c.";// 保持;
		}
		values[0] = press + "N";
		values[1] = Preloadingtime;
		values[2] = uptime;
		values[3] = powerontime1;
		values[4] = poweroncurent1;
		values[5] = coolingtime1;
		values[6] = powerontime2;
		values[7] = poweroncurent2;
		values[8] = coolingtime2;
		values[9] = powerontime3;
		values[10] = poweroncurent3;
		values[11] = maintain;

		return values;
	}

	// 255序列焊接条件设定表 电流电压
	private CurrentandVoltage getCurrentandVoltage(double average, List<CurrentandVoltage> cv) {
		// TODO Auto-generated method stub
		int index = 0;
		double fact = 0;
		double yushu = average % 0.5;
		if (yushu > 0) {
			fact = average + 0.5 - average % 0.5;
		} else {
			fact = average;
		}
		if (fact < 7) {
			fact = 7;
		}
		if (fact > 17) {
			fact = 17;
		}
		CurrentandVoltage voltage = cv.get(0);
		double first = Double.parseDouble(voltage.getIvalue());
		double difference = Math.abs(fact - first);
		for (int i = 0; i < cv.size(); i++) {
			CurrentandVoltage vol = cv.get(i);
			double bvalue = Double.parseDouble(vol.getIvalue());
			double diff = Math.abs(fact - bvalue);
			if (diff < difference) {
				if (!vol.getSequenceNo().equals("3") && !vol.getSequenceNo().equals("4")
						&& !vol.getSequenceNo().equals("5")) {
					index = i;
					difference = diff;
				}
			}
		}
		CurrentandVoltage factvaltage = cv.get(index);

		return factvaltage;
	}

	/*
	 * 获取基本信息表信息
	 */
	private List<WeldPointBoardInformation> getBaseinfomation(TCComponentBOMLine topbl, String procName) {
		List<WeldPointBoardInformation> baseinfolist = new ArrayList<WeldPointBoardInformation>();
		InputStream filein = null;
		try {
			filein = baseinfoExcelReader.getFileinbyreadExcel2(topbl.getItemRevision(), "IMAN_reference", procName);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		baseinfolist = baseinfoExcelReader.readHDExcel(filein, "xlsx");

		return baseinfolist;
	}

	/*
	 * 根据焊点在基本信息表中获取板件信息
	 */
	private List getBoardInformation(List<WeldPointBoardInformation> baseinfolist, ArrayList hdlist) {
		// TODO Auto-generated method stub
		List totalinfo = new ArrayList();
		if (baseinfolist != null) {
			for (int i = 0; i < hdlist.size(); i++) {
				String[] str = (String[]) hdlist.get(i);
				if (str[4] == null || str[4].isEmpty()) { // 根据版层数是否为空，确定是否需要重新获取板组数据
					String weldno = str[3];
					for (int j = 0; j < baseinfolist.size(); j++) {
						WeldPointBoardInformation wpb = baseinfolist.get(j);
						if (wpb.getWeldno() != null && weldno != null && wpb.getWeldno().equals(weldno)) {
							String[] values = new String[30];
							values[0] = str[0];
							values[1] = str[1];
							values[2] = str[2];
							values[3] = wpb.getWeldno(); // 焊点编号
							values[4] = wpb.getImportance(); // 重要度
							values[5] = wpb.getBoardnumber1(); // 板材1编号
							values[6] = wpb.getBoardname1(); // 板材1名称
							values[7] = wpb.getPartmaterial1(); // 板材1材质
							values[8] = wpb.getPartthickness1(); // 板材1板厚
							values[9] = wpb.getBoardnumber2(); // 板材2编号
							values[10] = wpb.getBoardname2(); // 板材2名称
							values[11] = wpb.getPartmaterial2(); // 板材2材质
							values[12] = wpb.getPartthickness2(); // 板材2板厚
							values[13] = wpb.getBoardnumber3(); // 板材3编号
							values[14] = wpb.getBoardname3(); // 板材3名称
							values[15] = wpb.getPartmaterial3(); // 板材3材质
							values[16] = wpb.getPartthickness3(); // 板材3板厚
							values[17] = wpb.getLayersnum(); // 板层数
							if (wpb.getGagi() != null && !wpb.getGagi().isEmpty()) {
								values[18] = wpb.getGagi(); // GA /GI
							} else {
								values[18] = "-"; // GA /GI
							}
							values[19] = wpb.getSheetstrength440(); // 材料强度(Mpa)440
							values[20] = wpb.getSheetstrength590(); // 材料强度(Mpa)590
							values[21] = wpb.getSheetstrength(); // 材料强度(Mpa)>590
							values[22] = wpb.getBasethickness(); // 基准板厚
							values[23] = wpb.getSheetstrength12(); // 材料强度(Mpa)1.2G
							values[24] = wpb.getStrength1();// 板件1强度
							values[25] = wpb.getStrength2();// 板件2强度
							values[26] = wpb.getStrength3();// 板件3强度
							values[27] = wpb.getGagi1();// 板件1GA/GI材
							values[28] = wpb.getGagi2();// 板件2GA/GI材
							values[29] = wpb.getGagi3();// 板件3GA/GI材
							totalinfo.add(values);
							break; // 找到就跳出本次循环，直接查找下一个焊点
						}
					}
				}
			}
		} else {
			System.out.println("获取基本信息失败！");
		}

		return totalinfo;
	}

	/*
	 * 字符转换成整数
	 */
	public static int getInteger(String str) {
		int num = -1;
		if (Util.isNumber(str)) {
			num = (int) Double.parseDouble(str);
		}
		return num;
	}

	/*
	 * 取最小值
	 */
	public static String getMinnum(String str1, String str2, String str3) {
		String minstr = "";
		if (str1 == null || str1.isEmpty()) {
			str1 = "9999";
		}
		if (str2 == null || str2.isEmpty()) {
			str2 = "9999";
		}
		if (str3 == null || str3.isEmpty()) {
			str3 = "9999";
		}
		if (Double.parseDouble(str1) > Double.parseDouble(str2)) {
			if (Double.parseDouble(str2) > Double.parseDouble(str3)) {
				minstr = str3;
			} else {
				minstr = str2;
			}
		} else {
			if (Double.parseDouble(str1) > Double.parseDouble(str3)) {
				minstr = str3;
			} else {
				minstr = str1;
			}
		}
//		if (minstr.equals("9999")) {
//			minstr = "";
//		}
		return minstr;
	}

	/*
	 * 取最大值
	 */
	public static String getMaxnum(String str1, String str2, String str3) {
		String maxstr = "";
		if (str1 == null || str1.isEmpty()) {
			str1 = "-1";
		}
		if (str2 == null || str2.isEmpty()) {
			str2 = "-1";
		}
		if (str3 == null || str3.isEmpty()) {
			str3 = "-1";
		}
		if (Double.parseDouble(str1) > Double.parseDouble(str2)) {
			if (Double.parseDouble(str1) > Double.parseDouble(str3)) {
				maxstr = str1;
			} else {
				maxstr = str3;
			}
		} else {
			if (Double.parseDouble(str2) > Double.parseDouble(str3)) {
				maxstr = str2;
			} else {
				maxstr = str3;
			}
		}
//		if (maxstr.equals("-1")) {
//			maxstr = "";
//		}
		return maxstr;
	}

	/*
	 * 判断版次是否为SOP后
	 */
	private boolean getIsSOPAfter(String bc) {
		boolean flag = false;
		ArrayList edition = getEditionSizeRule();
		if (edition != null && edition.size() > 0) {
			if (edition.contains(bc)) {
				return false;
			}
		}
		if (bc != null) {
			if (bc.length() == 1) {
				char c = bc.charAt(0);
				if (c >= 'A' && c <= 'Z') {
					flag = true;
				}
			}
			if (bc.length() == 2) {
				char c = bc.charAt(0);
				char cc = bc.charAt(1);
				if (c >= 'A' && c <= 'Z' && cc >= 'A' && cc <= 'Z') {
					flag = true;
				}
			}
		}
		return flag;
	}

	// 查询版次首选项，获取版次信息
	private ArrayList getEditionSizeRule() {
		ArrayList rule = new ArrayList();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_version_information");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_version_information");
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
	
	private XSSFCellStyle getXSSFStyle(XSSFWorkbook book,XSSFSheet sheet,int rowindex,int cellindex,int colorindex,int bgcolor)
	{
		XSSFRow row = sheet.getRow(rowindex);
		if(row!=null)
		{
			XSSFCell cell = row.getCell(cellindex);
			if(cell!=null)
			{
				XSSFCellStyle style = cell.getCellStyle();
				if(bgcolor > -1)
				{
					style.setFillForegroundColor((short)bgcolor);
					style.setFillPattern(CellStyle.SOLID_FOREGROUND);
				}
				if(colorindex > -1)
				{
					// 设置字体颜色
					Font font = book.createFont();
					Font sourcefont = style.getFont();
					font.setColor((short) colorindex);
					font.setFontHeightInPoints(sourcefont.getFontHeightInPoints());
					font.setFontName(sourcefont.getFontName());
					style.setFont(font);
				}
			    return style;
			}
		}
		return null;
	}
}
