package com.dfl.report.workschedule;

import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.rmi.AccessException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.CoverInfomation;
import com.dfl.report.ExcelReader.CurrentandVoltage;
import com.dfl.report.ExcelReader.SequenceWeldingConditionList;
import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.cme.kernel.bvr.FlowUtil;
import com.teamcenter.rac.cme.kernel.mfg.IMfgFlow;
import com.teamcenter.rac.cme.kernel.mfg.IMfgNode;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentFolderType;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.util.MessageBox;
import com.teamcenter.soa.exceptions.NotLoadedException;

public class StationInformationTableOp {

	private AbstractAIFUIApplication app;
	private static Logger logger = Logger.getLogger(baseinfoExcelReader.class.getName()); // 日志打印类
	private LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();// sheet名称与sheet页数
	private ArrayList list = new ArrayList();// sheet页集合
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// 设置日期格式
	private Map<String, String> projVehMap;// 获取首选项车型代号与familycode的关系
	private String Edition;// 版次
	private String topfoldername;
	private String model;// 模板类型
	private String nameNO;// 命名序号
	private boolean IsSameout;
	private String VehicleNo = "";// 车型代号
	private ArrayList partlist = new ArrayList();// 部品数据集
	private ArrayList tempPartlist = new ArrayList();
	private LinkedHashMap<String, String> fymap = new LinkedHashMap<String, String>();// 用于部品数据分页
	private ArrayList<TCComponentBOMLine> Discretelist = new ArrayList<>();// 点焊工序数据集
	private TCComponentBOMLine topbl = new TCComponentBOMLine();// 工位对应的顶层BOP
	private List<WeldPointBoardInformation> baseinfolist;// 基本信息表的数据
	List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private TCSession session;
	private ArrayList Import = new ArrayList();
	private boolean updateflag = false; // 是否更新标识
	private Map<String, String[]> notelist = new HashMap<String, String[]>();// 记录更新时，用户维护的页码和打点号
	private List pswlist = new ArrayList();// 记录更新时，之前的焊点信息
	private List rswqdlist = new ArrayList();// 记录更新时，之前的焊点信息
	private List rswsflist = new ArrayList();// 记录更新时，之前的焊点信息
	private ArrayList deletelist;
	private List<CurrentandVoltage> cv;
	private Map<String, List<String>> MaterialMap;
	private String stlr = "";// 记录所选工位是左工位还是右工位，1为左，2为右

	public StationInformationTableOp(AbstractAIFUIApplication app, ArrayList list, LinkedHashMap<String, String> map,
			String edition, String model, String nameNO, String topfoldername, boolean isSameout,
			List<CurrentandVoltage> cv, List<WeldPointBoardInformation> baseinfolist,
			Map<String, List<String>> materialMap) throws TCException, AccessException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.map = map;
		this.list = list;
		this.Edition = edition;
		this.topfoldername = topfoldername;
		this.model = model;
		this.nameNO = nameNO;
		this.IsSameout = isSameout;
		this.cv = cv;
		this.baseinfolist = baseinfolist;
		this.MaterialMap = materialMap;
		session = (TCSession) app.getSession();
		initUI();
	}

	public StationInformationTableOp(AbstractAIFUIApplication app, String edition, String topfoldername,
			List<CurrentandVoltage> cv, List<WeldPointBoardInformation> baseinfolist,
			Map<String, List<String>> materialMap) throws AccessException, TCException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.Edition = edition;
		this.topfoldername = topfoldername;
		this.cv = cv;
		this.baseinfolist = baseinfolist;
		session = (TCSession) app.getSession();
		this.Edition = edition;
		this.updateflag = true;
		this.MaterialMap = materialMap;
		initUI();
	}

	private void initUI() throws TCException, AccessException {
		// TODO Auto-generated method stub

		InterfaceAIFComponent[] ifc = app.getTargetComponents();
		TCComponentBOMLine gwbl = (TCComponentBOMLine) ifc[0];
		TCComponentItemRevision gwrev = gwbl.getItemRevision();
		// 获取工位所属虚层
		TCComponentBOMLine xubl = gwbl.parent().parent();
		String childrenFoldername = Util.getProperty(xubl, "bl_rev_object_name").replace("_", ".").replace("-", ".")
				.replace(" ", ".");
		try {
			topbl = gwbl.window().getTopBOMLine();
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		TCComponentGroup group = session.getGroup();
		// 科室
		String groupname = group.getLocalizedFullName();
		// 发行科
		String department = "";
		if (groupname != null
				&& (groupname.contains("同期工程科") || groupname.contains("simultaneous Engineering Section"))) {
			department = "H30";
		} else if (groupname != null
				&& (groupname.contains("焊装技术科") || groupname.contains("Body Assembly Engineering Section"))) {
			department = "VE2";
		} else {
			department = "VE2";
		}
		// 读取 项目-车型 首选项
		projVehMap = ReportUtils.getDFL_Project_VehicleNo();
		String FamlilyCode = "";
		FamlilyCode = Util.getProperty(topbl, "bl_rev_project_ids");// 基本车型
		if (projVehMap.size() < 1) {
			VehicleNo = FamlilyCode;
		} else {
			VehicleNo = projVehMap.get(FamlilyCode);
			if (VehicleNo == null) {
				if (FamlilyCode != null) {
					VehicleNo = FamlilyCode;
				}
			}
		}
		// 获取部品信息 ,部品名称如果工位名称以#开头，则为产线名称+工位名称，否则就是工位名称
		String linename = Util.getProperty(gwbl.parent(), "bl_rev_object_name");
		String staname = Util.getProperty(gwbl, "bl_rev_object_name");
		String assyname = "";
		if (staname.length() > 1) {
			if (staname.substring(0, 1).equals("#")) {
				if(linename.endsWith("LH") || linename.endsWith("RH"))
				{
					assyname = linename.trim() + staname;
				}
				else
				{
					assyname = linename.trim() + " " + staname;
				}
				
			} else {
				assyname = staname;
			}
		} else {
			if(linename.endsWith("LH") || linename.endsWith("RH"))
			{
				assyname = linename.trim() + staname;
			}
			else
			{
				assyname = linename.trim() + " " + staname;
			}
		}

		// getPartsinformation(gwbl);
		// 对称工位
		TCComponentBOMLine ssgwbl = getSymmetryState(gwbl.parent(), staname);
		// 生成报表操作前的动作
		GenerateReportInfo info = new GenerateReportInfo();
		// 文件名称
		String procName = "";
		boolean isupdatesame = false;
		// 如果是更改，先获取报表数据，根据名称判断是福需要左右工位同出
		if (updateflag) {
			info.setExist(false);
			info.setIsgoon(true);
			info.setAction(""); //$NON-NLS-1$
			info.setMeDocument(null);
			info.setDFL9_process_type("H"); //$NON-NLS-1$
			info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
			info.setmeDocumentName(procName);
			info.setFlag(true);
			info.setProject_ids(topbl.getItemRevision());

			try {
				info = ReportUtils.beforeGenerateReportAction(gwbl.getItemRevision(), info);
			} catch (TCException e) {
				e.printStackTrace();
				// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
				return;
			}
			TCComponentItemRevision docRev = info.getMeDocument();
			String docrevname = Util.getProperty(docRev, "object_name");
			if (docrevname.contains("左╱右") || docrevname.contains("右╱左")) {
				isupdatesame = true;
			}
		}

		// 如果选择左右工位不同出，对应工位为空
		if ((!IsSameout && !updateflag) || (updateflag && !isupdatesame)) {
			ssgwbl = null;
		}
		String LRassyname = "";
		if (ssgwbl != null) {
			String linename2 = Util.getProperty(ssgwbl.parent(), "bl_rev_object_name");
			String staname2 = Util.getProperty(ssgwbl, "bl_rev_object_name");
			if (staname2.length() > 1) {
				if (staname2.substring(0, 1).equals("#")) {
					LRassyname = linename2 + " " + staname2;
				} else {
					LRassyname = staname2;
				}
			} else {
				LRassyname = linename2 + " " + staname2;
			}
		}

		// 需要根据产线的中文名称取值，如果产线下游多个工位，名称后面还要增加工位名称
		// String stationname = Util.getProperty(gwrev, "b8_ChineseName");// 工位名称
		String stationname = "";
		if (Util.getIsMEProcStat(gwbl.parent())) {
			stationname = Util.getProperty(gwbl.parent().getItemRevision(), "b8_ChineseName")
					+ Util.getProperty(gwbl, "bl_rev_object_name");// 工位中文名称
		} else {
			stationname = Util.getProperty(gwbl.parent().getItemRevision(), "b8_ChineseName");// 工位中文名称
		}
		if (ssgwbl != null) {
			if (linename != null && linename.length() > 1
					&& linename.substring(linename.length() - 2, linename.length()).equals("LH")) {
				stationname = stationname.replace("左", "").replace("右", "") + " 左╱右";
				stlr = "1";
			} else {
				stationname = stationname.replace("左", "").replace("右", "") + " 右╱左";
				stlr = "2";
			}

		}

		if (!updateflag) {
			// procName = nameNO + "." + stationname.replace("#", "");
			procName = nameNO + "." + stationname;
		}
		if (!updateflag) {
			info.setExist(false);
			info.setIsgoon(true);
			info.setAction(""); //$NON-NLS-1$
			info.setMeDocument(null);
			info.setDFL9_process_type("H"); //$NON-NLS-1$
			info.setDFL9_process_file_type("AB"); // $NON-NLS-1$
			info.setmeDocumentName(procName);
			info.setFlag(true);
			info.setProject_ids(topbl.getItemRevision());

			try {
				info = ReportUtils.beforeGenerateReportAction(gwbl.getItemRevision(), info);
			} catch (TCException e) {
				e.printStackTrace();
				// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
				return;
			}
		}

		System.out.println("The action is completed before the report operation is generated.");

		if (!info.isIsgoon()) {
			return;
		}
		InputStream inputStream = null;
		if (updateflag) {
			TCComponentItemRevision docmentRev = info.getMeDocument();
			procName = Util.getProperty(docmentRev, "object_name");
			inputStream = baseinfoExcelReader.getFileinbyreadExcel(docmentRev, "IMAN_specification", procName);

			if (inputStream == null) {
				MessageBox.post("请确认" + procName + "版本对象下，存在" + procName + "数据集！", "提示信息", MessageBox.ERROR);
				return;
			}
		}
		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("正在获取模板...\n", 10, 100);

		// 查询工位导出模板
		if (updateflag) {

		} else {
			if (model.equals("普通工位模板")) {
				inputStream = FileUtil.getTemplateFile("DFL_Template_EngineeringWorkListStation");
				if (inputStream == null) {
					viewPanel.addInfomation("错误：没有找到工程作业表普通工位模板，请先添加模板(名称为：DFL_Template_EngineeringWorkListStation)\\n",
							100, 100);
					return;
				}
			} else if (model.equals("VIN码打刻模板")) {
				inputStream = FileUtil.getTemplateFile("DFL_Template_EngineeringWorkVINCarve");
				if (inputStream == null) {
					viewPanel.addInfomation("错误：没有找到工程作业表VIN码打刻模板，请先添加模板(名称为：DFL_Template_EngineeringWorkVINCarve)\\n",
							100, 100);
					return;
				}
			} else {
				inputStream = FileUtil.getTemplateFile("DFL_Template_AdjustmentLine");
				if (inputStream == null) {
					viewPanel.addInfomation("错误：没有找到工程作业表调整线模板，请先添加模板(名称为：DFL_Template_AdjustmentLine)\\n", 100, 100);
					return;
				}
			}

			System.out.println("获取空模版完成");
		}
		XSSFWorkbook book = null;
		if (updateflag) {
			book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);
		} else {
			// 根据业务选择的sheet页，加载初始模板
			book = creatEngineeringXSSFWorkbook(inputStream, list, map);

			System.out.println("初始化sheet页完成");
		}
		viewPanel.addInfomation("开始输出报表...\n", 20, 100);

		// 将公共信息批量赋值
		ArrayList plist = new ArrayList();// 获取的公共的输出数据集合
		// String username = app.getSession().getUserName();// 编制人
		TCComponentUser user = session.getUser();
		String username = user.getUserName();
		// 改为根据BOP名称取工厂产线信息
		String objectname = Util.getProperty(topbl.getItemRevision(), "object_name");
		String factoryline = ReportUtils.getFactoryLineByBOP(objectname);
		String factory = "";
		String linebody = "";
		if (factoryline.length() > 3) {
			factory = factoryline.substring(0, 3);
			linebody = factoryline.substring(factoryline.length() - 1);
		}
		String baseCarType = VehicleNo + "-" + factory + "-NO" + linebody;
		String stationcode = "";
		String[] assynos;
		TCProperty p = gwrev.getTCProperty("b8_ProcAssyNo2");//b8_ProcAssyNo2
		if (p != null) {
			assynos = p.getStringValueArray(); // Ａｓｓｙ 部番
		} else {
			assynos = null;
		}
		boolean rLflag = false;
		String[] assynos2 = null;
		if (ssgwbl != null) {
			TCProperty p2 = ssgwbl.getItemRevision().getTCProperty("b8_ProcAssyNo2");
			if (p2 != null) {
				assynos2 = p2.getStringValueArray(); // Ａｓｓｙ 部番
			} else {
				assynos2 = null;
			}
			rLflag = true;
		}
		String LRsunffix = "";
		if (assynos2 != null && assynos2.length > 0) {
			if (assynos2[0] != null) {
				if (assynos2[0].length() >= 5) {
					LRsunffix = "/" + assynos2[0].trim().substring(4, 5);// 工位编号
				}
			}
		}
		if (assynos != null && assynos.length > 0) {
			if (assynos[0] != null) {
				if (assynos[0].length() >= 5) {
					stationcode = "M" + assynos[0].trim().substring(0, 5) + LRsunffix;// 工位编号
				} else {
					stationcode = "M" + assynos[0].trim() + LRsunffix;
				}
			}
		} else {
			stationcode = "";// 工位编号
		}

		// 获取assy号
		List assylist = new ArrayList();
		List assynamelist = new ArrayList();
		if (assynos != null && assynos.length > 0) {
			for (int i = 0; i < assynos.length; i++) {
				if (assynos[i] != null) {
					String[] str = new String[2];
					str[0] = assynos[i];
					str[1] = assyname;
					assynamelist.add(str);
					assylist.add(assynos[i]);
				}
			}
		}
		if (assynos2 != null && assynos2.length > 0) {
			for (int i = 0; i < assynos2.length; i++) {
				if (assynos2[i] != null) {
					String[] str = new String[2];
					str[0] = assynos2[i];
					str[1] = LRassyname;
					assynamelist.add(str);
					assylist.add(assynos2[i]);
				}

			}
		}
		if (assynamelist != null) {
			for (int i = 0; i < assynamelist.size(); i++) {
				String[] str = (String[]) assynamelist.get(i);
				System.out.println("assynamelist子项1：" + str[0]);
				System.out.println("assynamelist子项2：" + str[1]);
			}
		}

		String pc = Edition;// 批次
		plist.add(username);
		plist.add(df2.format(new Date()));// 日期
		plist.add(baseCarType);
		plist.add(stationname);
		plist.add(stationcode);
		plist.add(pc);
		plist.add(department);

		// 获取部品信息
		List RHlist = getNewPartsinformation(gwbl);
		List LHlist = new ArrayList();

		if (ssgwbl != null) {
			// RLflag = true;
			LHlist = getNewPartsinformation(ssgwbl);
		}
		// 设置标号并排序
		SetLabelsAndSort(RHlist, gwbl, ssgwbl, LHlist);

		// getRLHStateData(sortList, LHlist);

		System.out.println("获取部品信息完成");

		viewPanel.addInfomation("", 30, 100);

		// 开启旁路
		{
			Util.callByPass(session, true);
		}

		// 构成表信息处理
		PartsinformationProcessing(book, assylist, assynamelist);

		System.out.println("构成表信息处理完成");

		// 获取图片信息
		Map<String, File> piclist = getAll3DPictures(gwbl.getItemRevision());

		// 构成图信息处理
		CompositionChartProcessing(book, assylist, assyname, rLflag, piclist);

		System.out.println("构成图信息处理完成");

		// 式样差信息处理
		PoorPatternProcessing(book, assylist, rLflag);

		System.out.println("式样差信息处理完成");

		// 获取所选点焊工序集合
		Discretelist = Util.getChildrenByBOMLine(gwbl, "B8_BIWDiscreteOPRevision");

		List<TCComponentBOMLine> symmetryDiscretelist = new ArrayList<>(); // 对称工位下的点焊工序集合
		if (ssgwbl != null) {
			// RLflag = true;
			symmetryDiscretelist = Util.getChildrenByBOMLine(ssgwbl, "B8_BIWDiscreteOPRevision");
		}

		System.out.println("获取点焊工序集合完成");

		viewPanel.addInfomation("", 40, 100);

		// 循环点焊工序集合，根据点焊工序名称是否为R开头，区分使用PSW还是RSWsheet页，并确认sheet页数
		Map<String, TCComponentBOMLine> blmap = new LinkedHashMap<String, TCComponentBOMLine>();
		int psw = 0;
		int rswq = 0;
		int rsws = 0;
		List<String> Discretenamelist = new ArrayList<>(); // 记录点焊工序的名称
		if (Discretelist.size() > 0) {
			for (int i = 0; i < Discretelist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) Discretelist.get(i);
				String Discretename = Util.getProperty(bl, "bl_rev_object_name");
				if (!Discretenamelist.contains(Discretename)) {
					Discretenamelist.add(Discretename);
				}
				if (Discretename.length() > 1) {
					if (Discretename.substring(0, 1).equals("R")) {

						// 由于RSW气动与RSW伺服不会同时出现，所以都处理
						// 复制一个sheet，并返回sheet名称
						String sheetname = CopySheet(book, "RSW气动", rswq);
						if (sheetname != null) {
							blmap.put(sheetname, bl);
							rswq++;
						}
						String sheetname1 = CopySheet(book, "RSW伺服", rsws);
						if (sheetname1 != null) {
							blmap.put(sheetname1, bl);
							rsws++;
						}

					} else {
						// PSW信息处理
						String sheetname = CopySheet(book, "PSW", psw);
						if (sheetname != null) {
							blmap.put(sheetname, bl);
							psw++;
						}
					}
				} else {
					System.out.println("The spot welding process name is incorrect and will not be processed.");
				}
			}
			/*************************
			 * 先清空系统输出内容
			 */
			if (updateflag) {
				// 获取焊点对应的页码和打点号
				getPageNumberManagement(book);

				RSWSFClearSheetContext(book, "RSW伺服");
				RSWQDClearSheetContext(book, "RSW气动");
				PSWClearSheetContext(book, "PSW");
			}
			/******************************/

		}
		// 对称工位下工序名称与点焊工序map
		Map<String, TCComponentBOMLine> symmetrymap = new HashMap<>();
		// 对称工位独有的点焊工序
		if (symmetryDiscretelist.size() > 0) {
			for (int i = 0; i < symmetryDiscretelist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) symmetryDiscretelist.get(i);
				String Discretename = Util.getProperty(bl, "bl_rev_object_name");

				if (Discretenamelist.contains(Discretename)) {
					symmetrymap.put(Discretename, bl);
					continue;
				}
				if (Discretename.length() > 1) {
					if (Discretename.substring(0, 1).equals("R")) {

						// 由于RSW气动与RSW伺服不会同时出现，所以都处理
						// 复制一个sheet，并返回sheet名称
						String sheetname = CopySheet(book, "RSW气动", rswq);
						if (sheetname != null) {
							blmap.put(sheetname, bl);
							rswq++;
						}
						String sheetname1 = CopySheet(book, "RSW伺服", rsws);
						if (sheetname1 != null) {
							blmap.put(sheetname1, bl);
							rsws++;
						}

					} else {
						// PSW信息处理
						String sheetname = CopySheet(book, "PSW", psw);
						if (sheetname != null) {
							blmap.put(sheetname, bl);
							psw++;
						}
					}
				} else {
					System.out.println("The spot welding process name is incorrect and will not be processed.");
				}
			}
		}
		if (blmap.size() > 0) {

			// 获取基本信息
//			String baseName = "222.基本信息";
//			baseinfolist = getBaseinfomation(topbl, baseName);

			for (Map.Entry<String, TCComponentBOMLine> entry : blmap.entrySet()) {
				String shname = entry.getKey();
				TCComponentBOMLine bl = entry.getValue();
				// 由于RSW气动与RSW伺服不会同时出现，所以都处理

				// RSW气动信息处理
				if (shname.contains("RSW气动")) {
					RSWpneumaticinformationProcessing(book, bl, gwbl, shname, symmetrymap);

					System.out.println("RSW气动信息处理完成");
				} else if (shname.contains("RSW伺服")) {

					// RSW伺服信息处理
					RSWServoinformationProcessing(book, bl, gwbl, shname, symmetrymap);

					System.out.println("RSW伺服信息处理完成");
				} else {
					// PSW信息处理
					String error = PSWinformationProcessing(book, bl, shname, symmetrymap);

					if (!error.isEmpty()) {
						viewPanel.dispose();
						MessageBox.post(error, "提示信息", MessageBox.ERROR);
						return;
					}

					System.out.println("PSW信息处理完成");
				}
			}
		}

		// 处理涂胶sheet页图标问题

		ProcessingGlueIcon(book, gwbl);

		// 处理安装sheet页图标问题

		ProcessingInstallationIcon(book, gwbl);

		System.out.println("涂胶和安装信息处理完成");

		// 处理打点统计表sheet页信息
		ProcessingStatistics(book, stationname, Discretelist, symmetryDiscretelist,ssgwbl);

		System.out.println("打点统计表信息处理完成");

		viewPanel.addInfomation("", 50, 100);

		// 先保存成文件，再取出
		// book = saveFileAndgetFile(book,filename);

		/* 先写数据，后删除未选的sheet页，避免图片写入错误，数据写完后，重新计算页码 */
		// 写入所有sheet页公共信息
		writePublicDataToSheet(book, plist);

		// 先删除未选择的sheet
		deleteSheets(book);

		// 重新计算页码
		writeRepatPublicDataToSheet(book);

		System.out.println("写入所有sheet页公共信息处理完成");

		if (!updateflag) {
			// 有效页信息处理
			ValidPageProcessing(book);

			System.out.println("有效页信息处理完成");
		}

		// 所有数据写完后，需要把sheet页的名称重命名
		SetSheetRename(book);

		// 对点焊sheet页设置公式，可以根据焊点号，自动获取板组编号
		// setCellFormula(book);

		// 获取作业内容
		int shs = book.getNumberOfSheets();
		String[] contents = new String[shs];
		for (int i = 0; i < shs; i++) {
			String sheetname = book.getSheetName(i);
			contents[i] = sheetname;
		}
		// 写入作业内容和焊点重要度
		TCProperty ppp = gwrev.getTCProperty("b8_OperationContent");
		if (ppp != null) {
			ppp.setStringValueArray(contents);
		}
		String[] weldimport = new String[1];
		if (Import.size() > 0) {
			String tempimport = "";
			for (int i = 0; i < Import.size(); i++) {
				String tempn = (String) Import.get(i);
				tempimport = tempimport + tempn;
			}
			weldimport[0] = tempimport;
		} else {
			weldimport[0] = "C";
		}
		TCProperty ppp2 = gwrev.getTCProperty("b8_WPImptLevel");
		if (ppp2 != null) {
			ppp2.setStringValueArray(weldimport);
		}
		// 写工序编号、工序中文名称和页码
		setPropertyValue(gwrev, "b8_OPNo", stationcode);
		setPropertyValue(gwrev, "b8_STName", stationname);
		setPropertyValue(gwrev, "b8_OpSheetNumber", Integer.toString(shs));

		gwrev.lock();
		gwrev.save();
		gwrev.unlock();

		System.out.println("sheet页的名称重命名处理完成");

		// String filename = Util.getProperty(gwbl, "bl_rev_object_name") + "工程作业表";
		viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);

		String filename = Util.formatString(procName);
		NewOutputDataToExcel.exportFile(book, filename);

		String fullFileName = FileUtil.getReportFileName(filename);
		System.out.println(fullFileName);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
		if (ds != null) {
			datasetList.add(ds);
		}
		revlist.add(gwrev);
		try {
			TCComponentItem docunment = ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "",
					session);
			saveFileToFolder(docunment, topfoldername, childrenFoldername, procName);

			// 文件编号和虚层名称
			TCProperty[] pdoc = docunment.getTCProperties(new String[] { "dfl9_vehiclePlant", "dfl9_processArea" });
			if (pdoc != null) {
				if (pdoc.length > 1 && pdoc[0] != null) {
					pdoc[0].setStringValue(topfoldername);
					pdoc[1].setStringValue(childrenFoldername);
					docunment.lock();
					docunment.save();
					docunment.unlock();
				}
			}
//			docunment.getLatestItemRevision().setProperty("dfl9_vehiclePlant", topfoldername);
//			docunment.getLatestItemRevision().setProperty("dfl9_processArea", childrenFoldername);

		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}

		// saveFiles(filename, gwbl);
		viewPanel.addInfomation("", 80, 100);

		// 关闭旁路
		{
			Util.callByPass(session, false);
		}

		viewPanel.addInfomation("输出报表完成，请在焊装工厂工位对象附件下查看报表...\n", 100, 100);

	}

	/**
	 * 打点统计表信息处理 20200727 hgq
	 * 
	 * @param book
	 * @param stationname
	 * @param discretelist2
	 * @param symmetryDiscretelist
	 * @param ssgwbl 
	 */
	private void ProcessingStatistics(XSSFWorkbook book, String stationname, ArrayList<TCComponentBOMLine> discretelist2,
			List<TCComponentBOMLine> symmetryDiscretelist, TCComponentBOMLine ssgwbl) {
		// TODO Auto-generated method stub
		// 获取打点统计表sheet
		if (!updateflag) 
		{
			for(int i=0;i<deletelist.size();i++)
			{
				if(deletelist.get(i).toString().contains("打点统计表"))
				{
					return;
				}
			}			
		}		
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex1 = -1; // 打点统计表所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("打点统计表")) {
				sheetAtIndex1 = i;
				break;
			}
		}
		if (sheetAtIndex1 == -1) {
			return;
		}
        //左右工位同出
		if(ssgwbl!=null)
		{
			if (!updateflag) 
			{
				//将打点统计表模板，分成左右两个模板
				XSSFSheet newsheet = book.cloneSheet(sheetAtIndex1);
				book.setSheetOrder(newsheet.getSheetName(), sheetAtIndex1+1);
				String modelname = book.getSheetName(sheetAtIndex1);
				book.setSheetName(sheetAtIndex1, modelname + "-左");
				book.setSheetName(sheetAtIndex1+1, modelname + "-右");
			}						
			//先处理左工位输出，根据名称后两位是否含有左或LH字眼区分
			List<TCComponentBOMLine> LHlist = new ArrayList<>();
			List<TCComponentBOMLine> RHlist = new ArrayList<>();
			String ssgwname = Util.getProperty(ssgwbl, "bl_rev_object_name");
//			if(ssgwname.length()>2 && (ssgwname.substring(ssgwname.length()-3).contains("LH") || ssgwname.substring(ssgwname.length()-3).contains("左")))
//			{
			if("1".equals(stlr))
			{
				LHlist = discretelist2;
				RHlist = symmetryDiscretelist;
			}
			else
			{
				LHlist = symmetryDiscretelist;
				RHlist = discretelist2;
				
			}
			
			WriteManagementStatistics(book,LHlist,"打点统计表-左",stationname);
			WriteManagementStatistics(book,RHlist,"打点统计表-右",stationname);
		}
		else //非左右工位同出
		{
			WriteManagementStatistics(book,discretelist2,"打点统计表",stationname);			
		}
	}

	private void WriteManagementStatistics(XSSFWorkbook book, List<TCComponentBOMLine> discretelist2, String sheettypename,
			String stationname) {
		// TODO Auto-generated method stub
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex1 = -1; // 打点统计表所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(sheettypename)) {
				sheetAtIndex1 = i;
				break;
			}
		}
		if (sheetAtIndex1 == -1) {
			return;
		}
		
		int datanum = discretelist2.size();
		int page = datanum / 12 + 1;           
		// 数据行刚好一页就会出现sheet页多了一页的情况
		if (datanum % 12 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}
		// 如果page大于1，则需要复制sheet页
		int index = sheetAtIndex1 + 1;			
		int shnum = 0;
		List<String> olddatalist = new ArrayList<>();
		if (updateflag) {
			
			//获取之前已写入的数据
			olddatalist =  getStatisticsData(book,sheettypename);
			
			int number = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains(sheettypename)) {
					if (sheetAtIndex1 <= i) {
						number++;
					}
				}
			}
			index = index + page - 1;
			if (number < page) {
				if (page - number > 0) {
					for (int i = 0; i < page - number; i++) {
						XSSFSheet newsheet = book.cloneSheet(sheetAtIndex1);
						book.setSheetOrder(newsheet.getSheetName(), index);
						index++;
					}
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex1);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}
		System.out.println("page: " + page);
		if (updateflag) 
		{
			//获取本次新增的点焊工序名称
			List<String> newnamelist = new ArrayList<>();
			for(int i=0;i<discretelist2.size(); i++)
			{
				TCComponentBOMLine  direbl = (TCComponentBOMLine) discretelist2.get(i);
				String name = Util.getProperty(direbl, "bl_rev_object_name").replace("\\", "-");
				if(!olddatalist.contains(name))
				{
					newnamelist.add(name);
				}
			}
			for (int i = sheetAtIndex1; i < index; i++) 
			{
				XSSFSheet sheet = book.getSheetAt(i);	
				setStringCellAndStyle(sheet, stationname, 8, 21, null, Cell.CELL_TYPE_STRING);
				if (i == index - 1) 
				{
					for (int j = 0; j + 12 * shnum < olddatalist.size()+newnamelist.size(); j++) 
					{
						if(j + 12 * shnum>olddatalist.size()-1)
						{
							String name = newnamelist.get(j + 12 * shnum - olddatalist.size());
							setStringCellAndStyle(sheet, name, 8 + j*3, 41, null, Cell.CELL_TYPE_STRING);
						}					
					}			
				}
				else 
				{
					for (int j = 0; j + 12 * shnum < 12 + 12 * shnum; j++) 
					{
						if(j + 12 * shnum>olddatalist.size()-1)
						{
							String name = newnamelist.get(j + 12 * shnum - olddatalist.size());
							setStringCellAndStyle(sheet, name, 8 + j*3, 41, null, Cell.CELL_TYPE_STRING);
						}						
					}
				}				
				shnum++;
			}	
		}
		else
		{
			for (int i = sheetAtIndex1; i < index; i++) 
			{
				XSSFSheet sheet = book.getSheetAt(i);	
				setStringCellAndStyle(sheet, stationname, 8, 21, null, Cell.CELL_TYPE_STRING);
				if (i == index - 1) 
				{
					for (int j = 0; j + 12 * shnum < discretelist2.size(); j++) 
					{
						TCComponentBOMLine  direbl = (TCComponentBOMLine) discretelist2.get(j + 12 * shnum);
						String name = Util.getProperty(direbl, "bl_rev_object_name").replace("\\", "-");
						setStringCellAndStyle(sheet, name, 8 + j*3, 41, null, Cell.CELL_TYPE_STRING);
					}			
				}
				else 
				{
					for (int j = 0; j + 12 * shnum < 12 + 12 * shnum; j++) 
					{
						TCComponentBOMLine  direbl = (TCComponentBOMLine) discretelist2.get(j + 12 * shnum);
						String name = Util.getProperty(direbl, "bl_rev_object_name").replace("\\", "-");
						setStringCellAndStyle(sheet, name, 8 + j*3, 41, null, Cell.CELL_TYPE_STRING);
					}
				}				
				shnum++;
			}		
		}
			
	}

	private List<String> getStatisticsData(XSSFWorkbook book, String sheettypename) {
		// TODO Auto-generated method stub
		List<String> valuelist = new ArrayList<>();
		int sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(sheettypename)) {
				XSSFSheet sheet = book.getSheetAt(i);
				// 数据从9行到42行
				int rowStart = 8;
				int rowEnd = 41;
				for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
					Row row = (Row) sheet.getRow(rowNum);
					if (null == row) {
						continue;
					}
					Cell cell;
					// 焊点编号
					cell = row.getCell(41);
					String name = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
					if(name!=null && !name.isEmpty())
					{
						valuelist.add(name);
					}
					rowNum = rowNum + 2;
				}
			}
		}
		return valuelist;
	}

	// 先保存为excel文件，再取出并删除
	public XSSFWorkbook saveFileAndgetFile(XSSFWorkbook book, String reportname) {
		try {
			String fullFileName = FileUtil.getReportFileName(reportname);
			File file = new File(fullFileName);
			if (file.exists()) {
				file.delete();
				file = new File(fullFileName);
			}

			FileOutputStream fOut = new FileOutputStream(file);
			try {
				book.write(fOut);
				fOut.flush();
				fOut.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			XSSFWorkbook newbook = new XSSFWorkbook(fullFileName);

			System.out.println("删除未选择的sheet页后：" + newbook.getNumberOfSheets());

			// 获取到后删除
			if (file.exists()) {
				file.delete();
			}

			return newbook;

			// 打开excel
			// openFile(fullFileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

	/*
	 * 设置属性值
	 */
	private void setPropertyValue(TCComponent tcc, String property, String value) throws TCException {
		TCProperty p = tcc.getTCProperty(property);
		if (p != null) {
			p.setStringValue(value);
		}
	}

	/*
	 * 如果为左右工位，需要把对应的工位部品添加到部品partlist中，根据部品名称进行匹配，如果匹配就添加到下一行标号和安装顺序为空
	 */
	private List getRLHStateData(List sortList, List lHlist,Map<String,String> partNoToNummap) {
		// TODO Auto-generated method stub
		if (sortList != null && sortList.size() > 0) {
			for (int i = 0; i < sortList.size(); i++) {
				if (lHlist != null && lHlist.size() > 0) {
					String[] values = (String[]) sortList.get(i);
					String partName = values[2];
					String partNo = values[1];
					tempPartlist.add(values);	
					for (Iterator<String[]> it = lHlist.iterator(); it.hasNext();) {
						String[] vals = it.next();
						String partName2 = vals[2];
						String partNo2 = vals[1];
						//零件号相同的不是左右件
						if(!partNo.equals(partNo2))
						{
							if ((partName != null && partName.length() > 2)
									&& (partName2 != null && partName2.length() > 2)) {
								if (partName.substring(0, partName.length() - 2)
										.equals(partName2.substring(0, partName2.length() - 2))) {
									// vals[0] = "";
									vals[7] = values[7];
									vals[0] = values[0];
									tempPartlist.add(vals);
									it.remove();
								}
							}
						}
						else
						{
							String qty = vals[3];
							partNoToNummap.put(partNo, qty);						
							it.remove();
						}
					}
				} else {
					String[] values = (String[]) sortList.get(i);
					tempPartlist.add(values);
				}
			}
		}
		return lHlist;

	}
	/** 
     * 使用java正则表达式去掉多余的.与0 
     * @param s 
     * @return  
     */  
    public static String subZeroAndDot(String s){  
        if(s.indexOf(".") > 0){  
            s = s.replaceAll("0+?$", "");//去掉多余的0  
            s = s.replaceAll("[.]$", "");//如最后一位是.则去掉  
        }  
        return s;  
    }  

	/*
	 * 设置标号并排序
	 */
	private void SetLabelsAndSort(List list, TCComponentBOMLine gwbl, TCComponentBOMLine ssgwbl, List lHlist)
			throws AccessException, TCException {

		// 获取完后，对数据进行排序处理
		// List oneList = new ArrayList();
		if (list == null) {
			return;
		}

		Comparator comparator = getComParatorBysequenceno();
		Collections.sort(list, comparator);

		int label = 0; // 标号标记
		int num = 1;// 标记同种标号的数据行数
		int Occupynum = 0;// 安装顺序为0的占用标号的顺序
		String prePartno = "";// 部品番号前5位标记
		String[] bh = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
				"T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK",
				"AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ" };
		// 标号处理
		Map<String, String> tempmap = new HashMap<String, String>();
		ArrayList tempPartlist1 = new ArrayList();
		for (int i = 0; i < list.size(); i++) {
			String[] str = (String[]) list.get(i);
			if (str[1].toString().length() > 5 && (str[1].contains(" ") || str[1].contains(" "))) {
				prePartno = str[1].toString().substring(0, 5);
			} else {
				prePartno = str[1].toString();
			}
			String note = tempmap.get(prePartno);
			// 部品番号前5位一样，则标号相同
			if (note != null && !note.isEmpty()) {
				str[7] = note;
				int spno = 0;
				for (int j = 0; j < bh.length; j++) {
					if (bh[j].equals(note)) {
						spno = j + 1 - Occupynum;
					}
				}
				if (!str[0].equals("0")) {
					str[0] = Integer.toString(spno); // 安装顺序重新定义
				}
				String strnum = fymap.get(note);
				int newnum = Integer.parseInt(strnum) + 1;
				fymap.put(note, Integer.toString(newnum));
			} else {
				if (label < 52) {
					str[7] = bh[label];
					if (!str[0].equals("0")) {
						str[0] = Integer.toString(label + 1 - Occupynum); // 安装顺序重新定义
					} else {
						Occupynum++;
					}
				} else {
					str[7] = "";
					System.out.println("超过了规定的标号。。。。");
				}
				fymap.put(bh[label], "1");
				tempmap.put(prePartno, bh[label]);
				label++;
			}
			tempPartlist1.add(str);

		}

		// 如果为左右工位，需要把对应的工位部品添加到部品tempPartlist中，返回对应工位有特有的部品
		Map<String,String> partNoToNummap = new HashMap<>();//零件号与数量的对应关系map
		List remainList = getRLHStateData(tempPartlist1, lHlist,partNoToNummap);

		if (remainList != null && remainList.size() > 0) {
			for (int i = 0; i < remainList.size(); i++) {
				String[] str = (String[]) remainList.get(i);
				if (str[1].toString().length() > 5 && (str[1].contains(" ") || str[1].contains(" "))) {
					prePartno = str[1].toString().substring(0, 5);
				} else {
					prePartno = str[1].toString();
				}
				String note = tempmap.get(prePartno);
				// 部品番号前5位一样，则标号相同
				if (note != null && !note.isEmpty()) {
					str[7] = note;
					int spno = 0;
					for (int j = 0; j < bh.length; j++) {
						if (bh[j].equals(note)) {
							spno = j + 1 - Occupynum;
						}
					}
					if (!str[0].equals("0")) {
						str[0] = Integer.toString(spno); // 安装顺序重新定义
					}
					String strnum = fymap.get(note);
					int newnum = Integer.parseInt(strnum) + 1;
					fymap.put(note, Integer.toString(newnum));
				} else {
					if (label < 52) {
						str[7] = bh[label];
						if (!str[0].equals("0")) {
							str[0] = Integer.toString(label + 1 - Occupynum); // 安装顺序重新定义
						} else {
							Occupynum++;
						}
					} else {
						str[7] = "";
						System.out.println("超过了规定的标号。。。。");
					}
					fymap.put(bh[label], "1");
					tempmap.put(prePartno, bh[label]);
					label++;
				}
				tempPartlist.add(str);
			}
		}

		// 把内制部品放到最后
		List LHlist = getLastStationPartList(gwbl);

		if (LHlist != null && LHlist.size() > 0) {
			for (int i = 0; i < LHlist.size(); i++) {
				String[] strVal = (String[]) LHlist.get(i);
				strVal[7] = bh[label];
				strVal[0] = Integer.toString(label + 1 - Occupynum); // 安装顺序重新定义
				tempPartlist.add(strVal);
			}
			if (fymap.containsKey(bh[label])) {
				String numstr = fymap.get(bh[label]);
				int newnum = Integer.parseInt(numstr) + 1;
				fymap.put(bh[label], Integer.toString(newnum));
			} else {
				fymap.put(bh[label], "1");
			}
		}

		if (ssgwbl != null) {
			List RHlist = getLastStationPartList(ssgwbl);
			if (RHlist != null && RHlist.size() > 0) {
				for (int i = 0; i < RHlist.size(); i++) {
					String[] strVal = (String[]) RHlist.get(i);
					strVal[7] = bh[label];
					strVal[0] = Integer.toString(label + 1 - Occupynum); // 安装顺序重新定义
					tempPartlist.add(strVal);
				}
				if (fymap.containsKey(bh[label])) {
					String numstr = fymap.get(bh[label]);
					int newnum = Integer.parseInt(numstr) + 1;
					fymap.put(bh[label], Integer.toString(newnum));
				} else {
					fymap.put(bh[label], "1");
				}
			}
		}
		// 根据标号排序
		Comparator comparator2 = getComParatorBybh();
		Collections.sort(tempPartlist, comparator2);

		String firstNo = "";
		for (int i = 0; i < tempPartlist.size(); i++) {
			String[] value = (String[]) tempPartlist.get(i);
			String partNo = value[1];
			 //如果存在零件号相同，数量显示为右数量/左数量
			if(partNoToNummap.containsKey(partNo))
			{
				if("1".equals(stlr))
				{
					value[3] = partNoToNummap.get(partNo) + "/" + value[3];
				}
				else
				{
					value[3] = value[3] + "/" + partNoToNummap.get(partNo);
				}
			}
			
			if (i == 0) {
				firstNo = value[7];
				partlist.add(value);
			} else {
				if (!firstNo.equals(value[7].toString())) {
					String[] str = new String[9];
					partlist.add(str);
					partlist.add(value);
					firstNo = value[7];
				} else {
					partlist.add(value);
				}
			}
		}
		System.out.println(partlist);
		// return oneList;
	}

	/*
	 * 获取部品信息
	 */
	private List getNewPartsinformation(TCComponentBOMLine gwbl) throws TCException, AccessException {
		// TODO Auto-generated method stub
		ArrayList install = new ArrayList();
		ArrayList templist = new ArrayList();
		// 先获取工位下的安装工序下的零件
		install = Util.getChildrenByBOMLine(gwbl, "B8_BIWOperationRevision");

		System.out.println("查找到的安装工序数量：" + install.size());

		for (int i = 0; i < install.size(); i++) {
			// 通过首选项获取部品来源
			Map<String, String> partsource = getSizeRule();
			TCComponentBOMLine bl = (TCComponentBOMLine) install.get(i);
			ArrayList bflist = new ArrayList();
			bflist = Util.getChildrenByBOMLine(bl, "DFL9SolItmPartRevision");
			System.out.println("查找到的部品数量：" + bflist.size());
			for (int j = 0; j < bflist.size(); j++) {
				String[] info = new String[9];
				TCComponentBOMLine bfbl = (TCComponentBOMLine) bflist.get(j);
				info[0] = Util.getProperty(bfbl, "bl_sequence_no");// 安装顺序
				if (info[0].isEmpty()) {
					info[0] = "0";
				}
				info[1] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9_part_no");// 部品番号
				// info[2] = Util.getProperty(bfbl, "bl_rev_object_name");// 部品名称
				info[2] = Util.getProperty(bfbl.getItemRevision(), "dfl9_CADObjectName");// 部品名称
				info[3] = Util.getProperty(bfbl, "bl_quantity");// 数量
				if (info[3] == null || info[3].isEmpty()) {
					info[3] = "1";
				}
				String partresoles = "";
				String partresValue = "";
				TCProperty p = bfbl.getTCProperty("B8_BiwManualMU");
				if (p != null) {
					String lovindex = p.getStringValue();
					if (lovindex != null && !lovindex.isEmpty()) {
						if (partsource.containsKey(lovindex)) {
							partresoles = partsource.get(lovindex);
						}
						partresValue = lovindex;

					}
				}
				// partresoles = Util.getProperty(bfbl, "B8_NoteManualMark");// 部品来源 待确认
				if (partresoles == null || partresoles.isEmpty()) {
					TCProperty p2 = bfbl.getTCProperty("B8_NoteIsBiwTrUnit");
					if (p2 != null) {
						String lovindex = p2.getStringValue();
						if (lovindex != null && !lovindex.isEmpty()) {
							if (partsource.containsKey(lovindex)) {
								partresoles = partsource.get(lovindex);
							}
							partresValue = lovindex;
						}
					}
					// partresoles = Util.getProperty(bfbl, "B8_NoteIsBiwTrUnit");// 部品来源 待确认
				}
				info[6] = partresoles;
				info[8] = partresValue;
				System.out.println(" 部品来源:" + partresoles);
				if (partresValue.equals("Stamping")) {
					String thick = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartThickness");// 板厚
					if (Util.isNumber(thick)) {
						Double th = Double.parseDouble(thick);
						info[4] = String.format("%.2f", th);
					} else {
						info[4] = thick;
					}
					info[5] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartMaterial");// 材质
					System.out.println(" 材质:" + Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartMaterial"));
				} else {
					info[4] = "";// 板厚
					info[5] = "";// 材质
				}
				templist.add(info);
			}
		}
		// 如果零件号相同，合并为一行，数量合计
		Map<String, String[]> map = new HashMap<String, String[]>();
		for (int i = 0; i < templist.size(); i++) {
			String[] value = (String[]) templist.get(i);
			String key = value[1];
			if (!map.containsKey(key)) {
				map.put(key, value);
			} else {
				String[] oldstr = map.get(key);
				int quality = 0;
				quality = Integer.parseInt(oldstr[3]) + Integer.parseInt(value[3]);
				oldstr[3] = Integer.toString(quality);
				map.put(key, oldstr);
			}
		}
		List newtemplist = new ArrayList();
		for (Map.Entry<String, String[]> entry : map.entrySet()) {
			String[] values = entry.getValue();
			newtemplist.add(values);
		}
		return newtemplist;

	}

	/*
	 * 获取工位的前驱工位的assy部番
	 */
	private List getLastStationPartList(TCComponentBOMLine bl) throws TCException, AccessException {
		List templist = new ArrayList();

		// 内部连接在获取工位的上一个工位的assy部番
		TCProperty pp = bl.getTCProperty("Mfg0predecessors");// 前趋工位
		// TCProperty pp = bl.getTCProperty("Mfg0successors");//后趋工位
		if (pp != null) {
			TCComponent[] obj = pp.getReferenceValueArray();
			for (int i = 0; i < obj.length; i++) {
				TCComponentBOMLine prebl = (TCComponentBOMLine) obj[i];
				String sequence_no = Util.getProperty(prebl, "bl_sequence_no");// 安装顺序
				if (sequence_no.isEmpty()) {
					sequence_no = "0";
				}
				String quantity = Util.getProperty(prebl, "bl_quantity");// 数量
				if (quantity == null || quantity.isEmpty()) {
					quantity = "1";
				}
				// 获取部品信息 ,部品名称为产线名称
				String linename = Util.getProperty(prebl.parent(), "bl_rev_object_name");
				String assyname = linename;

				TCProperty p = prebl.getItemRevision().getTCProperty("b8_ProcAssyNo2");
				String[] assynos;
				if (p != null) {
					assynos = p.getStringValueArray();// Ａｓｓｙ 部番
				} else {
					assynos = null;
				}
				if (assynos != null && assynos.length > 0) {
					for (int j = 0; j < assynos.length; j++) {
						String[] info = new String[9];
						info[0] = sequence_no;// 安装顺序
						info[1] = assynos[j];// 部品番号
						info[2] = assyname;// 部品名称
						info[3] = quantity;// 数量
						info[4] = "";// 板厚
						info[5] = "";// 材质
						info[6] = "内制总成";// 部品来源 待确认
						if(assynos[j] != null)
						{
							templist.add(info);
						}						
					}
				}
			}
		}
		// 外部连接的工位的上一个工位的assy部番
		// List<IMfgFlow> list = FlowUtil.getScopeOutputFlows(bl);//外部后趋工位
		List<IMfgFlow> list = FlowUtil.getScopeInputFlows(bl);// 外部前趋工位
		if (list != null && list.size() > 0) {
			for (IMfgFlow flow : list) {
				IMfgNode node = flow.getPredecessor();
				TCComponentBOMLine preComp = (TCComponentBOMLine) node.getComponent();
				String sequence_no = Util.getProperty(preComp, "bl_sequence_no");// 安装顺序
				if (sequence_no.isEmpty()) {
					sequence_no = "0";
				}
				String quantity = Util.getProperty(preComp, "bl_quantity");// 数量
				if (quantity == null || quantity.isEmpty()) {
					quantity = "1";
				}
				// 获取部品信息 ,部品名称为产线名称
				String linename = Util.getProperty(preComp.parent(), "bl_rev_object_name");
				String assyname = linename;

				TCProperty p = preComp.getItemRevision().getTCProperty("b8_ProcAssyNo2");
				String[] assynos;
				if (p != null) {
					assynos = p.getStringValueArray();// Ａｓｓｙ 部番
				} else {
					assynos = null;
				}
				if (assynos != null && assynos.length > 0) {
					for (int j = 0; j < assynos.length; j++) {
						String[] info = new String[9];
						info[0] = sequence_no;// 安装顺序
						info[1] = assynos[j];// 部品番号
						info[2] = assyname;// 部品名称
						info[3] = quantity;// 数量
						info[4] = "";// 板厚
						info[5] = "";// 材质
						info[6] = "内制总成";// 部品来源 待确认
						if(assynos[j]!=null)
						{
							templist.add(info);
						}						
					}
				}
			}
		}

		return templist;
	}

	/*
	 * 判断工位是否有对称工位
	 */
	private TCComponentBOMLine getSymmetryState(TCComponentBOMLine linebl, String gwname) throws TCException {
		TCComponentBOMLine ssgwbl = null;
		String ProcLinename = Util.getProperty(linebl, "bl_rev_object_name");
		if (ProcLinename.length() > 1) {
			String rl = ProcLinename.substring(ProcLinename.length() - 2, ProcLinename.length());
			System.out.println("左右工位标识：" + rl);
			if (rl.equals("LH") || rl.equals("RH")) {
				String preLinename = ProcLinename.substring(0, ProcLinename.length() - 2);
				System.out.println("产线名称：" + ProcLinename);
				ArrayList list = Util.getChildrenByBOMLine(linebl.parent(), "B8_BIWMEProcLineRevision");
				for (int i = 0; i < list.size(); i++) {
					TCComponentBOMLine plinebl = (TCComponentBOMLine) list.get(i);
					String plinename = Util.getProperty(plinebl, "bl_rev_object_name");
					System.out.println("虚层下的产线：" + plinename);
					if (!plinename.equals(ProcLinename)) {
						if (plinename.length() > 1
								&& plinename.substring(0, plinename.length() - 2).equals(preLinename)) {
							ArrayList gwlist = Util.getChildrenByBOMLine(plinebl, "B8_BIWMEProcStatRevision");
							for (int j = 0; j < gwlist.size(); j++) {
								TCComponentBOMLine bl = (TCComponentBOMLine) gwlist.get(j);
								String statename = Util.getProperty(bl, "bl_rev_object_name");
								// 如果工位名称中也有左右，也需要区分左右匹配，否则直接按照名称相同匹配
								if (gwname.length() > 1) {
									String r2 = gwname.substring(gwname.length() - 2, gwname.length());
									if (r2.equals("LH") || r2.equals("RH")) {
										if (statename.length() > 1) {
											if (statename.substring(0, statename.length() - 2)
													.equals(gwname.substring(0, gwname.length() - 2))) {
												ssgwbl = bl;
												break;
											}
										}
									} else {
										if (statename.equals(gwname)) {
											ssgwbl = bl;
											break;
										}
									}
								} else {
									if (statename.equals(gwname)) {
										ssgwbl = bl;
										break;
									}
								}
							}
						}
					}
				}
			}
		}
		return ssgwbl;
	}

	/*
	 * 根据文件编号创建顶层文件夹在home下,并创建00.封面文件夹，把封面文档放到00.封面文件夹下
	 */
	private void saveFileToFolder(TCComponentItem document, String topfoldername, String childrenFoldername,
			String procName) {
		// TODO Auto-generated method stub
		try {
			TCComponentUser user = session.getUser();
			TCComponentFolder homefolder = user.getHomeFolder();
			TCComponentFolder folder = null;
			TCComponentFolder childrenfolder = null;
			// 先判断是否已经创建了该文件夹
			AIFComponentContext[] icf = homefolder.getChildren();
			for (AIFComponentContext aif : icf) {
				TCComponent tcc = (TCComponent) aif.getComponent();
				String obejctname = Util.getProperty(tcc, "object_name");
				if (tcc.getType().equals("Folder") && obejctname.equals(topfoldername)) {
					folder = (TCComponentFolder) tcc;
					break;
				}
			}
			if (folder == null) {
				return;
			}
			TCComponentFolderType foldertype = (TCComponentFolderType) session.getTypeComponent("Folder");

			// 先判断是否已经创建了01.目录及附录文件夹
			AIFComponentContext[] icf1 = folder.getChildren();
			for (AIFComponentContext aif : icf1) {
				TCComponent tcc = (TCComponent) aif.getComponent();
				String obejctname = Util.getProperty(tcc, "object_name");
				if (tcc.getType().equals("Folder") && obejctname.equals(childrenFoldername)) {
					childrenfolder = (TCComponentFolder) tcc;
					break;
				}
			}
			if (childrenfolder == null) {
				childrenfolder = foldertype.create(childrenFoldername, "", "Folder");
				folder.add("contents", childrenfolder);
				childrenfolder.add("contents", document);
			} else {
				// folder.add("contents", childrenfolder);
				AIFComponentContext[] icf3 = childrenfolder.getChildren();
				// 先移除
				if (icf3 != null) {
					for (AIFComponentContext aif : icf3) {
						TCComponent tcc = (TCComponent) aif.getComponent();
						String gwname = Util.getProperty(tcc, "object_name");
						if (gwname.equals(procName)) {
							childrenfolder.remove("contents", tcc);
						}
					}
				}
				childrenfolder.add("contents", document);
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/*
	 * 获取焊点对应得页码和打点号
	 */
	private void getPageNumberManagement(XSSFWorkbook book) {
		// TODO Auto-generated method stub
		int sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("PSW") || sheetname.contains("RSW")) {
				XSSFSheet sheet = book.getSheetAt(i);

				// 数据从12行到47行
				int rowStart = 11;
				int rowEnd = 47;
				for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
					Row row = (Row) sheet.getRow(rowNum);
					if (null == row) {
						continue;
					}
					Cell cell;
					// 焊点编号
					cell = row.getCell(8);
					String weldno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
					if (weldno != null && !weldno.isEmpty()) {
						String[] value = new String[2];
						// 页码
						cell = row.getCell(2);
						String pagenumber = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						// 打点号
						cell = row.getCell(6);
						String dot = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						value[0] = pagenumber;
						value[1] = dot;
						notelist.put(weldno, value);
					}

				}
			}
		}
	}

	private static String convertCellValueToString(Cell cell, int type) {
		if (cell == null) {
			return null;
		}
		String returnValue = null;
		switch (type) {
		case Cell.CELL_TYPE_NUMERIC: // 数字
			Double doubleValue = cell.getNumericCellValue();
			// 格式化科学计数法，取一位整数
			DecimalFormat df = new DecimalFormat("0.00");
			returnValue = df.format(doubleValue);
			break;
		case Cell.CELL_TYPE_STRING: // 字符串
			if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
				cell.setCellType(Cell.CELL_TYPE_STRING);
			}
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
		return returnValue;
	}

	/*
	 * 清空RSWSFsheet页系统输出内容
	 */
	private void RSWSFClearSheetContext(XSSFWorkbook book, String name) {
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // RSW气动所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}

		// 设置字体颜色
		Font font = book.createFont();
		font.setColor((short) 12);// 蓝色字体
		font.setFontHeightInPoints((short) 10);
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

		int gcnum = 0;
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				gcnum++;
			}
		}
		int index = sheetAtIndex + 1;
		// 如果sheet页增加就增，减少不删除，保留
		index = sheetAtIndex + gcnum;

		// 循环构成表sheet页清空系统输出内容，手工维护内容保留
		XSSFRow row;
		XSSFCell cell;
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			// 清空内容
			setStringCellAndStyle(sheet, "", 6, 20, style2, Cell.CELL_TYPE_STRING);// 工位
			setStringCellAndStyle(sheet, "", 6, 31, style2, Cell.CELL_TYPE_STRING);// 机器人
			setStringCellAndStyle(sheet, "", 6, 48, style2, Cell.CELL_TYPE_STRING);// 焊枪编号

			for (int j = 0; j < 36; j++) {
				row = sheet.getRow(11 + j);
				if (row != null) {
					cell = row.getCell(8);
					if (cell != null) {
						String weldno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (weldno != null && !weldno.isEmpty()) {
							rswsflist.add(weldno);
						}
					}
				}
				setStringCellAndStyle(sheet, "", 11 + j, 2, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 6, style8, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 8, style6, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 16, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 29, style7, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 36, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 42, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 55, style7, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 62, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 68, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 81, style7, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 87, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 90, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 92, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 94, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 96, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 98, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 102, style7, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 105, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 108, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 111, style, Cell.CELL_TYPE_STRING);
			}
		}
	}

	/*
	 * 清空RSWQDsheet页系统输出内容
	 */
	private void RSWQDClearSheetContext(XSSFWorkbook book, String name) {
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // RSW气动所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}

		int gcnum = 0;
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				gcnum++;
			}
		}
		int index = sheetAtIndex + 1;
		// 如果sheet页增加就增，减少不删除，保留
		index = sheetAtIndex + gcnum;

		// 设置字体颜色
		Font font = book.createFont();
		font.setColor((short) 12);// 蓝色字体
		font.setFontHeightInPoints((short) 10);
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

		// 循环构成表sheet页清空系统输出内容，手工维护内容保留
		XSSFRow row;
		XSSFCell cell;
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			// 清空内容
			setStringCellAndStyle(sheet, "", 6, 19, style2, Cell.CELL_TYPE_STRING);// 工位
			setStringCellAndStyle(sheet, "", 6, 30, style2, Cell.CELL_TYPE_STRING);// 机器人
			setStringCellAndStyle(sheet, "", 6, 47, style2, Cell.CELL_TYPE_STRING);// 焊枪编号

			for (int j = 0; j < 36; j++) {
				row = sheet.getRow(11 + j);
				if (row != null) {
					cell = row.getCell(8);
					if (cell != null) {
						String weldno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (weldno != null && !weldno.isEmpty()) {
							rswqdlist.add(weldno);
						}
					}
				}
				setStringCellAndStyle(sheet, "", 11 + j, 2, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 6, style4, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 8, style3, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 16, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 29, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 36, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 42, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 55, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 62, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 68, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 81, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 88, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 91, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 95, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 97, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 99, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 102, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
			}
		}
	}

	/*
	 * 清空PSWsheet页系统输出内容
	 */
	private void PSWClearSheetContext(XSSFWorkbook book, String name) {
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // PSW所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name) && !sheetname.contains("点焊")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}

		int gcnum = 0;
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name) && !sheetname.contains("点焊")) {
				gcnum++;
			}
		}
		int index = sheetAtIndex + 1;
		// 如果sheet页增加就增，减少不删除，保留
		index = sheetAtIndex + gcnum;

		// 设置字体颜色
		Font font = book.createFont();
		font.setColor((short) 12);// 蓝色字体
		font.setFontHeightInPoints((short) 10);
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

		XSSFCellStyle style22 = book.createCellStyle();
		style22.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
		style22.setBorderLeft(CellStyle.BORDER_THIN); // 粗线边框
		style22.setBorderRight(CellStyle.BORDER_THIN); // 粗线边框
		style22.setBorderTop(CellStyle.BORDER_THIN); //
		style22.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style22.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style22.setFont(font2);

		XSSFCellStyle style20 = book.createCellStyle();
		style20.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
		style20.setBorderLeft(CellStyle.BORDER_MEDIUM); // 粗线边框
		style20.setBorderRight(CellStyle.BORDER_THIN); // 粗线边框
		style20.setBorderTop(CellStyle.BORDER_THIN); //
		style20.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style20.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style20.setFont(font2);

		XSSFCellStyle style21 = book.createCellStyle();
		style21.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
		style21.setBorderLeft(CellStyle.BORDER_THIN); // 粗线边框
		style21.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
		style21.setBorderTop(CellStyle.BORDER_THIN); //
		style21.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style21.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style21.setFont(font2);

		Font font3 = book.createFont();
		font3.setColor((short) 12);// 蓝色字体
		font3.setFontHeightInPoints((short) 18);
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

		// 循环构成表sheet页清空系统输出内容，手工维护内容保留
		XSSFRow row;
		XSSFCell cell;
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			// 清空内容

			setStringCellAndStyle(sheet, "", 5, 8, style2, Cell.CELL_TYPE_STRING);// 变压器编号
			setStringCellAndStyle(sheet, "", 5, 19, style3, Cell.CELL_TYPE_STRING);// 焊枪编号
			setStringCellAndStyle(sheet, "", 7, 36, style20, Cell.CELL_TYPE_STRING);// 加压力
			setStringCellAndStyle(sheet, "", 7, 42, style22, Cell.CELL_TYPE_STRING);// 预压时间
			setStringCellAndStyle(sheet, "", 7, 48, style22, Cell.CELL_TYPE_STRING);// 上升时间
			setStringCellAndStyle(sheet, "", 7, 54, style22, Cell.CELL_TYPE_STRING);// 第一 通电时间
			setStringCellAndStyle(sheet, "", 7, 60, style22, Cell.CELL_TYPE_STRING);// 第一 通电电流
			setStringCellAndStyle(sheet, "", 7, 66, style22, Cell.CELL_TYPE_STRING);// 冷却时间一
			setStringCellAndStyle(sheet, "", 7, 72, style22, Cell.CELL_TYPE_STRING);// 第二 通电时间
			setStringCellAndStyle(sheet, "", 7, 78, style22, Cell.CELL_TYPE_STRING);// 第二 通电电流
			setStringCellAndStyle(sheet, "", 7, 84, style22, Cell.CELL_TYPE_STRING);// 冷却时间二
			setStringCellAndStyle(sheet, "", 7, 90, style22, Cell.CELL_TYPE_STRING);// 第三 通电时间
			setStringCellAndStyle(sheet, "", 7, 96, style22, Cell.CELL_TYPE_STRING);// 第三 通电电流
			setStringCellAndStyle(sheet, "", 7, 102, style21, Cell.CELL_TYPE_STRING);// 保持

			for (int j = 0; j < 36; j++) {
				row = sheet.getRow(11 + j);
				if (row != null) {
					cell = row.getCell(8);
					if (cell != null) {
						String weldno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (weldno != null && !weldno.isEmpty()) {
							pswlist.add(weldno);
						}
					}
				}
				setStringCellAndStyle(sheet, "", 11 + j, 2, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 6, style5, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 4, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 8, style4, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 13, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 16, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 29, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 36, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 39, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 42, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 55, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 62, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 65, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 68, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 81, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 88, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 91, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet, "", 11 + j, 95, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 97, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 99, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 102, style, 11);
				setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
				setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
			}
		}
	}

	/*
	 * 根据点焊工序确认sheet页
	 */
	private String CopySheet(XSSFWorkbook book, String name, int num) {
		String shname = "";
//		if (!updateflag) 
//		{
//			for(int i=0;i<deletelist.size();i++)
//			{
//				if(deletelist.get(i).toString().contains(name) && !deletelist.get(i).toString().contains("点焊"))
//				{
//					return null;
//				}
//			}			
//		}	
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // PSW所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name) && !sheetname.contains("点焊")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return null;
		}
		if (updateflag) {
			XSSFSheet sh = book.getSheetAt(sheetAtIndex + num);
			if (sh == null) {
				XSSFSheet sheet = book.cloneSheet(sheetAtIndex);
				shname = sheet.getSheetName();
				book.setSheetOrder(sheet.getSheetName(), sheetAtIndex + 1);
			} else {
				if (sh.getSheetName().contains(name) && !sh.getSheetName().contains("点焊")) {
					shname = sh.getSheetName();
				} else {
					XSSFSheet sheet = book.cloneSheet(sheetAtIndex);
					shname = sheet.getSheetName();
					book.setSheetOrder(sheet.getSheetName(), sheetAtIndex + 1);
				}
			}
		} else {
			if (num == 0) {
				shname = book.getSheetName(sheetAtIndex);
			} else {
				XSSFSheet sheet = book.cloneSheet(sheetAtIndex);
				shname = sheet.getSheetName();
				book.setSheetOrder(sheet.getSheetName(), sheetAtIndex + num);
			}
		}

		return shname;
	}

	/*
	 * 处理安装图标问题
	 */
	private void ProcessingInstallationIcon(XSSFWorkbook book, TCComponentBOMLine gwbl) throws TCException {
		// TODO Auto-generated method stub
		// 获取铰链安装和装配sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex1 = -1; // 铰链安装所在位置
		int sheetAtIndex2 = -1; // 装配所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("铰链安装")) {
				sheetAtIndex1 = i;
			}
			if (sheetname.contains("装配")) {
				sheetAtIndex2 = i;
			}
		}
		if (sheetAtIndex1 == -1 && sheetAtIndex2 == -1) {
			return;
		}
		/**************************************************/
		// 如果是更新，则先把系统输出的信息清空，后面再写入
		if (updateflag) {
//			XSSFSheet sheet = book.getSheetAt(sheetAtIndex1);
//			List<String> delPicturesList = ReportUtils.removePictrues07((XSSFSheet) sheet, (XSSFWorkbook) book, 3, 10,
//					100, 115);
//			System.out.println("-----------符合条件的图片有-----------");
//			for (String name : delPicturesList) {
//				System.out.println(name);
//			}
//
//			XSSFSheet sheet2 = book.getSheetAt(sheetAtIndex2);
//			List<String> delPicturesList2 = ReportUtils.removePictrues07((XSSFSheet) sheet2, (XSSFWorkbook) book, 3, 10,
//					100, 115);
//			System.out.println("-----------符合条件的图片有-----------");
//			for (String name : delPicturesList2) {
//				System.out.println(name);
//			}

		}
		/**************************************************/

		// 获取安装工序集合
		ArrayList tjlist = Util.getChildrenByBOMLine(gwbl, "B8_BIWOperationRevision");
		if (tjlist != null && !tjlist.isEmpty()) {
			String b8_TorqueImptLevel = "";
			for (int i = 0; i < tjlist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) tjlist.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				b8_TorqueImptLevel = rev.getProperty("b8_TorqueImptLevel");
				System.out.println(rev.getProperty("b8_TorqueImptLevel"));
				if (b8_TorqueImptLevel != null && !b8_TorqueImptLevel.isEmpty()) {
					break;
				}
			}
			if (b8_TorqueImptLevel != null && !b8_TorqueImptLevel.isEmpty()) {
				if (sheetAtIndex1 != -1) {
					XSSFSheet sheet = book.getSheetAt(sheetAtIndex1);
					InputStream is = null;
					if (b8_TorqueImptLevel.trim().equals("A")) {
						is = this.getClass().getResourceAsStream("/com/dfl/report/imags/A.png");
					}
					if (b8_TorqueImptLevel.trim().equals("B")) {
						is = this.getClass().getResourceAsStream("/com/dfl/report/imags/B.png");
					}
					if (is != null) {
						writepicturetosheet(book, sheet, is, 105, 5, 111, 9);
					}
				}
				if (sheetAtIndex2 != -1) {
					XSSFSheet sheet = book.getSheetAt(sheetAtIndex2);
					InputStream is = null;
					if (b8_TorqueImptLevel.trim().equals("A")) {
						is = this.getClass().getResourceAsStream("/com/dfl/report/imags/A.png");
					}
					if (b8_TorqueImptLevel.trim().equals("B")) {
						is = this.getClass().getResourceAsStream("/com/dfl/report/imags/B.png");
					}
					if (is != null) {
						writepicturetosheet(book, sheet, is, 105, 5, 111, 9);
					}
				}
			}

		}
	}

	/*
	 * 处理涂胶图标问题
	 */
	private void ProcessingGlueIcon(XSSFWorkbook book, TCComponentBOMLine gwbl) throws TCException {
		// TODO Auto-generated method stub
		// 获取涂胶sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // 涂胶所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("涂胶")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}

		/**************************************************/
		// 如果是更新，则先把系统输出的信息清空，后面再写入
		if (updateflag) {
//			XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
//			List<String> delPicturesList = ReportUtils.removePictrues07((XSSFSheet) sheet, (XSSFWorkbook) book, 3, 10,
//					100, 115);
//			System.out.println("-----------符合条件的图片有-----------");
//			for (String name : delPicturesList) {
//				System.out.println(name);
//			}
		}
		/**************************************************/

		// 获取涂胶工序集合
		ArrayList tjlist = Util.getChildrenByBOMLine(gwbl, "B8_BIWArcWeldOPRevision");// B8_BIWPaintOPRevision
		if (tjlist != null && !tjlist.isEmpty()) {
			String b8_GlueFeature = "";
			for (int i = 0; i < tjlist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) tjlist.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				b8_GlueFeature = rev.getProperty("b8_GlueFeature");
				System.out.println(rev.getProperty("b8_GlueFeature"));
				if (b8_GlueFeature != null && !b8_GlueFeature.isEmpty()) {
					break;
				}
			}
			if (b8_GlueFeature != null && !b8_GlueFeature.isEmpty()) {
				XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
				InputStream is = null;
				if (b8_GlueFeature.trim().equals("水密")) {
					is = this.getClass().getResourceAsStream("/com/dfl/report/imags/SM.png");
				}
				if (b8_GlueFeature.trim().equals("防锈")) {
					is = this.getClass().getResourceAsStream("/com/dfl/report/imags/FX.png");
				}
				if (is != null) {
					writepicturetosheet(book, sheet, is, 105, 5, 111, 9);
				}
			}

		}
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
	 * 获取封面信息信息
	 */
	private List<CoverInfomation> getCoverinfomation(TCComponentBOMLine topbl, String procName) {
		List<CoverInfomation> coverinfolist = new ArrayList<CoverInfomation>();
		InputStream filein = null;
		try {
			filein = baseinfoExcelReader.getFileinbyreadExcel2(topbl.getItemRevision(), "IMAN_reference", procName);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		coverinfolist = baseinfoExcelReader.readCoverExcel(filein, "xlsx");

		return coverinfolist;
	}

	/*
	 * 根据焊点在基本信息表中获取板件信息
	 */
	private List<WeldPointBoardInformation> getBoardInformation(List<WeldPointBoardInformation> baseinfolist,
			ArrayList hdlist) {
		// TODO Auto-generated method stub

		List<WeldPointBoardInformation> totalinfo = new ArrayList<WeldPointBoardInformation>();
		if (baseinfolist != null) {
			for (int i = 0; i < hdlist.size(); i++) {
				TCComponentBOMLine hdbl = (TCComponentBOMLine) hdlist.get(i);
				String weldno = Util.getProperty(hdbl, "bl_rev_object_name");
				for (int j = 0; j < baseinfolist.size(); j++) {
					WeldPointBoardInformation wb = baseinfolist.get(j);
					if (wb.getWeldno() != null && weldno != null && wb.getWeldno().equals(weldno)) {
						totalinfo.add(wb);
						break; // 找到就跳出本次循环，直接查找下一个焊点
					}
				}
			}
		} else {
			System.out.println("获取基本信息失败！");
		}

		return totalinfo;
	}

	/*
	 * PSW信息处理
	 */
	private String PSWinformationProcessing(XSSFWorkbook book, TCComponentBOMLine bl, String name,
			Map<String, TCComponentBOMLine> symmetrymap) throws TCException {
		// TODO Auto-generated method stub
				
		String error = "";
		// 获取PSWsheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // PSW所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return error;
		}
		// 先获取工序下的枪和焊点
		ArrayList gunlist = new ArrayList();
		ArrayList hdlist = new ArrayList();

		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { guntypename, guntypename };

		String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
		String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values2 = new String[] { weldtypename, weldtypename };

		gunlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		hdlist = Util.searchBOMLine(bl, "OR", propertys2, "==", values2);

		// 根据版次是否为SOP前后，如果是SOP后，不输出焊接参数
		boolean sopflag = getIsSOPAfter();
		String Discretename = Util.getProperty(bl, "bl_rev_object_name");

		// 正常一个工序下只有一把枪
		String TransformerNumber = "";// 变压器编号
		String Guncode = "";// 焊枪编号
		String ElectrodeVol = "";// 焊枪电压 bl_B8_BIWGunRevision_b8_ElectrodeVol
		TCComponentItemRevision blrev = null;
		TransformerNumber = Util.getProperty(bl.getItemRevision(), "b8_AdapterModel");
		if (gunlist.size() > 0) {
			TCComponentBOMLine gunbl = (TCComponentBOMLine) gunlist.get(0);
			blrev = gunbl.getItemRevision();
//			TransformerNumber = Util.getProperty(blrev, "b8_AdapterModel");
//			Guncode = Util.getProperty(blrev, "b8_Model");
			ElectrodeVol = Util.getProperty(blrev, "b8_ElectrodeVol");
		}
		if (sopflag) {
			if (TransformerNumber == null || TransformerNumber.isEmpty()) {
				error = "请维护变压器型号信息再输出工程作业表。";
				return error;
			}
		}
		// 变压器编号和焊枪编号，改为从点焊工序名称中取值
		String[] nameArr = Discretename.split("\\\\");
		TransformerNumber = nameArr[0];
		if (nameArr.length > 1) {
			Guncode = nameArr[1];
		}
		if (sopflag) {
			if ("TR NO".equals(TransformerNumber)) {
				error = "请维护变压器型号信息再输出工程作业表。";
				return error;
			}
		}
		// 根据焊点在基本信息表中获取板件信息
		List<WeldPointBoardInformation> hdinfo = new ArrayList<WeldPointBoardInformation>();// 焊点所有信息

		hdinfo = getBoardInformation(baseinfolist, hdlist);

		Map<String, String[]> hjmap = getHanJieParater(hdlist);

		if (hdinfo == null || hdinfo.size() < 1) {
			System.out.println("没有获取基本信息，直接跳过！");
			return error;
		}

		System.out.println("hdinfo: " + hdinfo.size());

		// 先根据第一个板件排序
		Comparator comparator = getComParatorByfirstpart();
		Collections.sort(hdinfo, comparator);
		// 然后根据第二个板件排序
		Comparator comparator1 = getComParatorBySecondpart();
		Collections.sort(hdinfo, comparator1);

		// 再根据第三个板件排序
		Comparator comparator11 = getComParatorByThistypart();
		Collections.sort(hdinfo, comparator11);

		// 根据板层数排序，先输出3层板，在输出2层板
		Comparator comparator2 = getComParatorBylayersnum();
		Collections.sort(hdinfo, comparator2);

		// 如果是左右工位同出，还需要把所选工位的对称工位下，相同工序名称的焊点信息放在同一个sheet页输出

		// 根据焊点在基本信息表中获取板件信息
		List<WeldPointBoardInformation> symmetryhdinfo = new ArrayList<WeldPointBoardInformation>();// 焊点所有信息
		ArrayList symmetryhdlist = new ArrayList();
		if (symmetrymap.containsKey(Discretename)) {
			TCComponentBOMLine symmetrybl = symmetrymap.get(Discretename);
			symmetryhdlist = Util.searchBOMLine(symmetrybl, "OR", propertys2, "==", values2);
			symmetryhdinfo = getBoardInformation(baseinfolist, symmetryhdlist);

			// 先根据第一个板件排序
			Comparator comparators1 = getComParatorByfirstpart();
			Collections.sort(symmetryhdinfo, comparators1);
			// 然后根据第二个板件排序
			Comparator comparators2 = getComParatorBySecondpart();
			Collections.sort(symmetryhdinfo, comparators2);

			// 再根据第三个板件排序
			Comparator comparators3 = getComParatorByThistypart();
			Collections.sort(symmetryhdinfo, comparators3);

			// 根据板层数排序，先输出3层板，在输出2层板
			Comparator comparators4 = getComParatorBylayersnum();
			Collections.sort(symmetryhdinfo, comparators4);

			// 获取焊点的焊接参数
			getSymmetryHanJieParater(symmetryhdlist, hjmap);
		}

		// 根据数据判断是否需要分页,每36行数据分一页
		int hdsum = hdlist.size() + symmetryhdlist.size();
		int page = hdsum / 36 + 1;

		// 数据行刚好一页就会出现sheet页多了一页的情况
		if (hdsum % 36 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}

		// 如果page大于1，则需要复制sheet页
		int index = sheetAtIndex + 1;

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
		font2.setFontName("MS PGothic");
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

		XSSFCellStyle style22 = book.createCellStyle();
//		style22.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
//		style22.setBorderLeft(CellStyle.BORDER_THIN); // 粗线边框
//		style22.setBorderRight(CellStyle.BORDER_THIN); // 粗线边框
//		style22.setBorderTop(CellStyle.BORDER_THIN); //
		style22.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style22.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style22.setFont(font2);

		XSSFCellStyle style20 = book.createCellStyle();
//		style20.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
//		style20.setBorderLeft(CellStyle.BORDER_MEDIUM); // 粗线边框
//		style20.setBorderRight(CellStyle.BORDER_THIN); // 粗线边框
//		style20.setBorderTop(CellStyle.BORDER_THIN); //
		style20.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style20.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style20.setFont(font2);

		XSSFCellStyle style21 = book.createCellStyle();
//		style21.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
//		style21.setBorderLeft(CellStyle.BORDER_THIN); // 粗线边框
//		style21.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
//		style21.setBorderTop(CellStyle.BORDER_THIN); //
		style21.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style21.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style21.setFont(font2);

		Font font222 = book.createFont();
		// font2.setColor((short) 10);// 蓝色字体
		font222.setFontName("MS PGothic");
		font222.setFontHeightInPoints((short) 10);

		XSSFCellStyle style202 = book.createCellStyle();
//		style202.setBorderBottom(CellStyle.BORDER_THIN); // 粗线边框
//		style202.setBorderLeft(CellStyle.BORDER_MEDIUM); // 粗线边框
//		style202.setBorderRight(CellStyle.BORDER_THIN); // 粗线边框
//		style202.setBorderTop(CellStyle.BORDER_MEDIUM); //
		style202.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style202.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style202.setWrapText(true);
		style202.setFont(font222);

		XSSFCellStyle style212 = book.createCellStyle();
//		style212.setBorderBottom(CellStyle.BORDER_THIN); // 粗线边框
//		style212.setBorderLeft(CellStyle.BORDER_THIN); // 粗线边框
//		style212.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
//		style212.setBorderTop(CellStyle.BORDER_MEDIUM); //
		style212.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style212.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style212.setWrapText(true);
		style212.setFont(font222);

		XSSFCellStyle style222 = book.createCellStyle();
//		style222.setBorderBottom(CellStyle.BORDER_THIN); // 粗线边框
//		style222.setBorderLeft(CellStyle.BORDER_THIN); // 粗线边框
//		style222.setBorderRight(CellStyle.BORDER_THIN); // 粗线边框
//		style222.setBorderTop(CellStyle.BORDER_MEDIUM); //
		style222.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style222.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style222.setWrapText(true);
		style222.setFont(font222);

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

		if (updateflag) {
			int number = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("PSW")) {
					if (sheetAtIndex <= i) {
						number++;
					}
				}
			}
			index = index + page - 1;
			if (number < page) {
				if (page - number > 0) {
					for (int i = 0; i < page - number; i++) {
						XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
						book.setSheetOrder(newsheet.getSheetName(), index);
						index++;
					}
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}
		System.out.println("page: " + page);

		int shnum = 0;

		int maxRepressure = 0;// 加压力最大值
		int minRepressure = 99999999;// 加压力最小值
		double sumrevalue = 0;// 总电流值

		int datanum = 0; // 参与计算的焊点数量
		
		boolean isCucalPara = true; //是否计算综合焊接参数

		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			setStringCellAndStyle(sheet, TransformerNumber, 5, 8, style2, Cell.CELL_TYPE_STRING);// 变压器编号
			setStringCellAndStyle(sheet, Guncode, 5, 19, style3, Cell.CELL_TYPE_STRING);// 焊枪编号
			if (!sopflag) {
				setStringCellAndStyle(sheet, ElectrodeVol, 7, 108, style2, Cell.CELL_TYPE_STRING);// 焊枪电压
			}

			if (i == index - 1) {
				for (int j = 0; j + 36 * shnum < hdinfo.size() + symmetryhdinfo.size(); j++) {
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}

					String weldno = wpb.getWeldno(); // 焊点编号
					String pageNo = "";// 页码
					String dot = "";// 打点号
					// 根据焊点号获取页码和打点号
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // 重要度
					// 判断首页是否需要添加重要度图标
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // 板材1编号
					String boardname1 = wpb.getBoardname1(); // 板材1名称
					String partmaterial1 = wpb.getPartmaterial1(); // 板材1材质
					String partthickness1 = wpb.getPartthickness1(); // 板材1板厚
					String boardnumber2 = wpb.getBoardnumber2(); // 板材2编号
					String boardname2 = wpb.getBoardname2(); // 板材2名称
					String partmaterial2 = wpb.getPartmaterial2(); // 板材2材质
					String partthickness2 = wpb.getPartthickness2(); // 板材2板厚
					String boardnumber3 = wpb.getBoardnumber3(); // 板材3编号
					String boardname3 = wpb.getBoardname3(); // 板材3名称
					String partmaterial3 = wpb.getPartmaterial3(); // 板材3材质
					String partthickness3 = wpb.getPartthickness3(); // 板材3板厚
					String layersnum = wpb.getLayersnum(); // 板层数
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // 材料强度(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // 材料强度(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // 材料强度(Mpa)>590
					String basethickness = wpb.getBasethickness(); // 基准板厚
					String sheetstrength12 = wpb.getSheetstrength12(); // 材料强度(Mpa)1.2G

					String poweroncurent2 = "";// 第二通电电流
					String RecomWeldForce = "";// 推荐 加压力(N)
					String CurrentSerie = "";// 参数序列

					// 根据材质对照表，判断焊点是否参与计算焊接参数
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// 根据材质对照表获取GA/GI属性
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							System.out.println(MaterialNo + infolist);
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
										isCucalPara = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
										isCucalPara = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
										isCucalPara = false;
									}
								}
							}																				
						}
					}
					System.out.println(partmaterialFlag1);
					System.out.println(partmaterialFlag2);
					System.out.println(partmaterialFlag3);
					System.out.println(isCucalPara);
					// 排除不参与计算的焊点
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							RecomWeldForce = curenre[4];
							poweroncurent2 = curenre[3];
							CurrentSerie = curenre[1];
						}
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
							datanum++;
						}
					}

					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style5, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && pswlist != null) {
						if (!pswlist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness3, 11 + j, 88, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 91, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 95, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 97, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 99, style, 10);
					setStringCellAndStyle(sheet, basethickness, 11 + j, 102, style, 11);

					// 如果是1.2g高强材，基准板厚都是取最薄板
					if (sheetstrength12.equals("1.2g")) {
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
					} else {

						// 电流单位转换
						if (Util.isNumber(poweroncurent2)) {
							int curent = 0;
							curent = (int) (Double.parseDouble(poweroncurent2) * 1000);
							poweroncurent2 = Integer.toString(curent);
						}
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, poweroncurent2, 11 + j, 111, style, 10);
					}

				}
			} else {
				for (int j = 0; j + 36 * shnum < 36 + 36 * shnum; j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}

					String weldno = wpb.getWeldno(); // 焊点编号
					String pageNo = "";// 页码
					String dot = "";// 打点号
					// 根据焊点号获取页码和打点号
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // 重要度
					// 判断首页是否需要添加重要度图标
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // 板材1编号
					String boardname1 = wpb.getBoardname1(); // 板材1名称
					String partmaterial1 = wpb.getPartmaterial1(); // 板材1材质
					String partthickness1 = wpb.getPartthickness1(); // 板材1板厚
					String boardnumber2 = wpb.getBoardnumber2(); // 板材2编号
					String boardname2 = wpb.getBoardname2(); // 板材2名称
					String partmaterial2 = wpb.getPartmaterial2(); // 板材2材质
					String partthickness2 = wpb.getPartthickness2(); // 板材2板厚
					String boardnumber3 = wpb.getBoardnumber3(); // 板材3编号
					String boardname3 = wpb.getBoardname3(); // 板材3名称
					String partmaterial3 = wpb.getPartmaterial3(); // 板材3材质
					String partthickness3 = wpb.getPartthickness3(); // 板材3板厚
					String layersnum = wpb.getLayersnum(); // 板层数
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // 材料强度(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // 材料强度(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // 材料强度(Mpa)>590
					String basethickness = wpb.getBasethickness(); // 基准板厚
					String sheetstrength12 = wpb.getSheetstrength12(); // 材料强度(Mpa)1.2G

					String poweroncurent2 = "";// 第二通电电流
					String RecomWeldForce = "";// 推荐 加压力(N)
					String CurrentSerie = "";// 参数序列
					// 根据材质对照表，判断焊点是否参与计算焊接参数
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// 根据材质对照表获取GA/GI属性
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
										isCucalPara = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
										isCucalPara = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
										isCucalPara = false;
									}
								}
							}																				
						}
					}
					System.out.println(partmaterialFlag1);
					System.out.println(partmaterialFlag2);
					System.out.println(partmaterialFlag3);
					System.out.println(isCucalPara);
					// 排除不参与计算的焊点
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							RecomWeldForce = curenre[4];
							poweroncurent2 = curenre[3];
							CurrentSerie = curenre[1];
						}
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
							datanum++;
						}
					}

					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style5, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && pswlist != null) {
						if (!pswlist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness3, 11 + j, 88, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 91, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 95, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 97, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 99, style, 10);
					setStringCellAndStyle(sheet, basethickness, 11 + j, 102, style, 11);

					// 如果是1.2g高强材，基准板厚都是取最薄板
					if (sheetstrength12.equals("1.2g")) {
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
					} else {
						// 电流单位转换
						if (Util.isNumber(poweroncurent2)) {
							int curent = 0;
							curent = (int) (Double.parseDouble(poweroncurent2) * 1000);
							poweroncurent2 = Integer.toString(curent);
						}
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, poweroncurent2, 11 + j, 111, style, 10);
					}
				}
			}
//			// 工位名称 自适应大小
//			XSSFRow row = sheet.getRow(5);
//			if (row != null) {
//				XSSFCell cell = row.getCell(19);
//				if (cell != null) {
//					NewOutputDataToExcel.setFontSize(book, cell, (short) 16);
//				}
//			}
			shnum++;
		}

		// 计算参数
		String[] tatolcurenre = new String[12];
		// 如果没计算参数，就不计算焊枪参数
		if(!isCucalPara || datanum == 0)
		{			
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
		}
		else
		{
			tatolcurenre = getAverageParameterValues(maxRepressure, minRepressure, sumrevalue, datanum);
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
		

		// 再把计算的参数，写入
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			if (!sopflag) {
				for (int j = 0; j < tatolcurenre.length; j++) {
					if (j == 0) {
						setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style20, Cell.CELL_TYPE_STRING);// 加压力~保持
					} else if (j == tatolcurenre.length - 1) {
						setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style21, Cell.CELL_TYPE_STRING);// 加压力~保持
					} else {
						setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style22, Cell.CELL_TYPE_STRING);// 加压力~保持
					}

				}
			} else {
				if (!updateflag) {
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
			// 更改的时候不去删除焊接参数
			if (updateflag) {
				// 处理特殊情况，将焊接参数重写，获取版次信息
				XSSFRow terow = sheet.getRow(48);
				XSSFCell tecell = terow.getCell(108);
				String preedtion = tecell.getStringCellValue();
				boolean teflag = getIsTeSOPAfter(preedtion);
				if (!teflag) // 写会数据
				{
					setStringCellAndStyle(sheet, "加压力", 5, 36, style202, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "预压时间", 5, 42, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "上升时间", 5, 48, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "第一          通电时间", 5, 54, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "第一          通电电流", 5, 60, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "冷却时间一", 5, 66, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "第二          通电时间", 5, 72, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "第二          通电电流", 5, 78, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "冷却时间二", 5, 84, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "第三          通电时间", 5, 90, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "第三         通电电流", 5, 96, style222, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "保持", 5, 102, style212, Cell.CELL_TYPE_STRING);// 加压力~保持
					setStringCellAndStyle(sheet, "焊钳额定压力", 5, 108, style212, Cell.CELL_TYPE_STRING);// 加压力~保持
																									// ElectrodeVol
					setStringCellAndStyle(sheet, ElectrodeVol, 7, 108, style212, Cell.CELL_TYPE_STRING);//

					for (int j = 0; j < tatolcurenre.length; j++) {
						if (j == 0) {
							setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style20,
									Cell.CELL_TYPE_STRING);// 加压力~保持
						} else if (j == tatolcurenre.length - 1) {
							setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style21,
									Cell.CELL_TYPE_STRING);// 加压力~保持
						} else {
							setStringCellAndStyle(sheet, tatolcurenre[j], 7, 36 + j * 6, style22,
									Cell.CELL_TYPE_STRING);// 加压力~保持
						}
					}
				}
			}

		}

		String[] provalues = new String[11];
		int indexCount = 0;
		for (int i = 0; i < tatolcurenre.length; i++) {
			if (i != 1) {
				provalues[indexCount] = tatolcurenre[i].replace("c.", "").replace("KA", "").replace("N", "");
				indexCount++;
			}
		}
		TCComponentItemRevision rev = bl.getItemRevision();

		// 把计算的焊接参数写入到点焊工序上
		String[] properties = { "b8_WeldForce", "b8_RiseTime", "b8_CurrentTime1", "b8_Current1", "b8_Cool1",
				"b8_CurrentTime2", "b8_Current2", "b8_Cool2", "b8_CurrentTime3", "b8_Current3", "b8_KeepTime", };
		// 写属性值
		TCProperty[] pp = rev.getTCProperties(properties);
		if (pp != null && pp[0] != null) {
			rev.setProperties(properties, provalues);
			rev.lock();
			rev.save();
			rev.unlock();
		}

		return error;
	}

	private void getSymmetryHanJieParater(ArrayList hdlist, Map<String, String[]> hjmap) throws TCException {
		// TODO Auto-generated method stub
		if (hdlist != null && hdlist.size() > 0) {
			TCComponent[] tccs = new TCComponent[hdlist.size()];
			for (int i = 0; i < hdlist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) hdlist.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				tccs[i] = rev;
				String weldno = Util.getProperty(rev, "object_name");
			}
			String[] properties = { "object_name", "b8_CurrentSerie_Nissan", "b8_CurrentSerie_DFL", "b8_Current2",
					"b8_RecomWeldForce" };
			String[][] values = Util.getAllProperties(session, tccs, properties);
			for (int j = 0; j < values.length; j++) {
				String[] proper = values[j];
				hjmap.put(proper[0], proper);
			}
		}
		return;
	}

	// 计算平均参数
	private String[] getAverageParameterValues(int maxRepressure, int minRepressure, double sumrevalue, int size) {
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
		// 获取计算参数
//		Object[] obj = baseinfoExcelReader.getCalculationParameter(app, "B8_Calculation_Parameter_Name");
//		List<CurrentandVoltage> cv = new ArrayList<CurrentandVoltage>();
//		if (obj != null) {
//			if (obj[1] != null) {
//				cv = (List<CurrentandVoltage>) obj[1];
//			} else {
//				System.out.println("未获取到24序列焊接条件设定表 电流电压信息。");
//			}
//		}
		// 电流平均值
		BigDecimal biga1 = new BigDecimal(Double.toString(sumrevalue));
		BigDecimal bigsize = new BigDecimal(Double.toString(size));
		double average = biga1.divide(bigsize, 8, BigDecimal.ROUND_HALF_UP).doubleValue();
		// 255序列焊接条件设定表 电流电压
		CurrentandVoltage currentandVoltage = getCurrentandVoltage(average, cv);
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
	 * RSW伺服信息处理
	 */
	private void RSWServoinformationProcessing(XSSFWorkbook book, TCComponentBOMLine bl, TCComponentBOMLine gwbl,
			String name, Map<String, TCComponentBOMLine> symmetrymap) throws TCException {
		// TODO Auto-generated method stub
		// 获取RSW伺服sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // RSW伺服所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		// 先获取工序下的枪和焊点
		ArrayList gunlist = new ArrayList();
		ArrayList hdlist = new ArrayList();
		// ArrayList robotlist = new ArrayList();
		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { guntypename, guntypename };
		String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
		String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values2 = new String[] { weldtypename, weldtypename };

//		String[] propertys3 = new String[] { "bl_item_object_type", "bl_item_object_type" };
//		String[] values3 = new String[] { "机器人", "BIW Robot" };

		gunlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		hdlist = Util.searchBOMLine(bl, "OR", propertys2, "==", values2);
		// robotlist = Util.searchBOMLine(bl, "OR", propertys3, "==", values3);
		// 根据版次是否为SOP前后，如果是SOP后，不输出焊接参数
		boolean sopflag = getIsSOPAfter();
		String Discretename = Util.getProperty(bl, "bl_rev_object_name");
		// 正常一个工序下只有一把枪
		String stationname = Util.getProperty(gwbl, "bl_rev_object_name");// 工位
		String robotname = "";// 机器人
		String Guncode = "";// 焊枪编号
		String ElectrodeVol = "";// 焊枪电压 bl_B8_BIWGunRevision_b8_ElectrodeVol
		TCComponentItemRevision blrev = null;
		if (gunlist.size() > 0) {
			TCComponentBOMLine gunbl = (TCComponentBOMLine) gunlist.get(0);
			blrev = gunbl.getItemRevision();
			Guncode = Util.getProperty(blrev, "b8_Model");
			ElectrodeVol = Util.getProperty(blrev, "b8_ElectrodeVol");
		}
		robotname = Util.getProperty(bl, "bl_rev_object_name");
//		if (robotlist.size() > 0) {
//			TCComponentBOMLine robotbl = (TCComponentBOMLine) robotlist.get(0);
//			robotname = Util.getProperty(robotbl, "bl_rev_object_name");
//		}
		// 焊枪编号，改为从点焊工序名称中取值
		String[] nameArr = Discretename.split("\\\\");
		if (nameArr.length > 1) {
			Guncode = nameArr[1];
		} else {
			Guncode = "";
		}
		robotname = nameArr[0];
		
		// 根据焊点在基本信息表中获取板件信息
		List<WeldPointBoardInformation> hdinfo = new ArrayList<WeldPointBoardInformation>();// 焊点所有信息

		hdinfo = getBoardInformation(baseinfolist, hdlist);

		Map<String, String[]> hjmap = getHanJieParater(hdlist);

		if (hdinfo == null || hdinfo.size() < 1) {
			System.out.println("没有获取基本信息，直接跳过！");
			return;
		}
		// 先根据第一个板件排序
		Comparator comparator = getComParatorByfirstpart();
		Collections.sort(hdinfo, comparator);

		// 然后根据第二个板件排序
		Comparator comparator1 = getComParatorBySecondpart();
		Collections.sort(hdinfo, comparator1);

		// 再根据第三个板件排序
		Comparator comparator11 = getComParatorByThistypart();
		Collections.sort(hdinfo, comparator11);

		// 根据板层数排序，先输出3层板，在输出2层板
		Comparator comparator2 = getComParatorBylayersnum();
		Collections.sort(hdinfo, comparator2);

		// 如果是左右工位同出，还需要把所选工位的对称工位下，相同工序名称的焊点信息放在同一个sheet页输出
//		String Discretename = Util.getProperty(bl, "bl_rev_object_name");
		// 根据焊点在基本信息表中获取板件信息
		List<WeldPointBoardInformation> symmetryhdinfo = new ArrayList<WeldPointBoardInformation>();// 焊点所有信息
		ArrayList symmetryhdlist = new ArrayList();
		if (symmetrymap.containsKey(Discretename)) {
			TCComponentBOMLine symmetrybl = symmetrymap.get(Discretename);
			symmetryhdlist = Util.searchBOMLine(symmetrybl, "OR", propertys2, "==", values2);
			symmetryhdinfo = getBoardInformation(baseinfolist, symmetryhdlist);

			// 先根据第一个板件排序
			Comparator comparators1 = getComParatorByfirstpart();
			Collections.sort(symmetryhdinfo, comparators1);
			// 然后根据第二个板件排序
			Comparator comparators2 = getComParatorBySecondpart();
			Collections.sort(symmetryhdinfo, comparators2);

			// 再根据第三个板件排序
			Comparator comparators3 = getComParatorByThistypart();
			Collections.sort(symmetryhdinfo, comparators3);

			// 根据板层数排序，先输出3层板，在输出2层板
			Comparator comparators4 = getComParatorBylayersnum();
			Collections.sort(symmetryhdinfo, comparators4);

			// 获取焊点的焊接参数
			getSymmetryHanJieParater(symmetryhdlist, hjmap);
		}

		// 根据数据判断是否需要分页,每36行数据分一页
		int hdsum = hdlist.size() + symmetryhdlist.size();
		int page = hdsum / 36 + 1;
		// 数据行刚好一页就会出现sheet页多了一页的情况
		if (hdsum % 36 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}

		// 如果page大于1，则需要复制sheet页
		int index = sheetAtIndex + 1;

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

		if (updateflag) {
			int number = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("RSW伺服")) {
					if (sheetAtIndex <= i) {
						number++;
					}
				}
			}
			index = index + page - 1;
			if (number < page) {
				if (page - number > 0) {
					for (int i = 0; i < page - number; i++) {
						XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
						book.setSheetOrder(newsheet.getSheetName(), index);
						index++;
					}
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}

		int shnum = 0;

		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			setStringCellAndStyle(sheet, stationname, 6, 20, style2, Cell.CELL_TYPE_STRING);// 工位
			setStringCellAndStyle(sheet, robotname, 6, 31, style2, Cell.CELL_TYPE_STRING);// 机器人
			setStringCellAndStyle(sheet, Guncode, 6, 48, style2, Cell.CELL_TYPE_STRING);// 焊枪编号
			setStringCellAndStyle(sheet, ElectrodeVol, 6, 65, style2, Cell.CELL_TYPE_STRING);// 焊枪电压
																								// bl_B8_BIWGunRevision_b8_ElectrodeVol
			if (i == index - 1) {
				for (int j = 0; j + 36 * shnum < hdinfo.size() + symmetryhdinfo.size(); j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}
					String weldno = wpb.getWeldno(); // 焊点编号
					String pageNo = "";// 页码
					String dot = "";// 打点号

					// 根据焊点号获取页码和打点号
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // 重要度
					// 判断首页是否需要添加重要度图标
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // 板材1编号
					String boardname1 = wpb.getBoardname1(); // 板材1名称
					String partmaterial1 = wpb.getPartmaterial1(); // 板材1材质
					String partthickness1 = wpb.getPartthickness1(); // 板材1板厚
					String boardnumber2 = wpb.getBoardnumber2(); // 板材2编号
					String boardname2 = wpb.getBoardname2(); // 板材2名称
					String partmaterial2 = wpb.getPartmaterial2(); // 板材2材质
					String partthickness2 = wpb.getPartthickness2(); // 板材2板厚
					String boardnumber3 = wpb.getBoardnumber3(); // 板材3编号
					String boardname3 = wpb.getBoardname3(); // 板材3名称
					String partmaterial3 = wpb.getPartmaterial3(); // 板材3材质
					String partthickness3 = wpb.getPartthickness3(); // 板材3板厚
					String layersnum = wpb.getLayersnum(); // 板层数
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // 材料强度(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // 材料强度(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // 材料强度(Mpa)>590
					String sheetstrength12 = wpb.getSheetstrength12(); // 材料强度(Mpa)1.2G
					String basethickness = wpb.getBasethickness(); // 基准板厚
					String CurrentSerie = ""; // 参数 序列 (日产)
					String RecomWeldForce = "";// 推荐 加压力(N)
					String CurrentSeriedfi = ""; // 参数 序列 (对应)

					// 根据材质对照表，判断焊点是否参与计算焊接参数
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// 根据材质对照表获取GA/GI属性
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
									}
								}
							}																				
						}
					}
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							CurrentSerie = curenre[1];
							RecomWeldForce = curenre[4];
							CurrentSeriedfi = curenre[2];
						}
					}

					boolean flag = false;
					// 如果是1.2g高强材，基准板厚都是取最薄板
					if (sheetstrength12.equals("1.2g")) {
						flag = true;
					}
					if (flag) {
						basethickness = getMinnum(partthickness1, partthickness2, partthickness3);
					}
					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style8, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && rswsflist != null) {
						if (!rswsflist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial1)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 29, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 29, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial2)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 55, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 55, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial3)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 81, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 81, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness3, 11 + j, 87, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 90, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 92, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 92, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 94, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 96, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 98, style, 10);
					if (flag) {
						setStringCellAndStyle(sheet, "○", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
						if (getColorDistinction(layersnum, partmaterial1, partmaterial2, partmaterial3, partthickness1,
								partthickness2, partthickness3)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 1, IndexedColors.SKY_BLUE.getIndex());
							setStringCellAndStyle2(sheet, basethickness, 11 + j, 102, newstyle, 11);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 1, IndexedColors.VIOLET.getIndex());
							setStringCellAndStyle2(sheet, basethickness, 11 + j, 102, newstyle, 11);
						}
						// 后面参数序列为空
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 12, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle(sheet, basethickness, 11 + j, 102, newstyle, 11);
						setStringCellAndStyle(sheet, "-", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, CurrentSeriedfi, 11 + j, 111, style, Cell.CELL_TYPE_STRING);
					}
				}
			} else {
				for (int j = 0; j + 36 * shnum < 36 + 36 * shnum; j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}
					String weldno = wpb.getWeldno(); // 焊点编号
					String pageNo = "";// 页码
					String dot = "";// 打点号

					// 根据焊点号获取页码和打点号
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // 重要度
					// 判断首页是否需要添加重要度图标
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // 板材1编号
					String boardname1 = wpb.getBoardname1(); // 板材1名称
					String partmaterial1 = wpb.getPartmaterial1(); // 板材1材质
					String partthickness1 = wpb.getPartthickness1(); // 板材1板厚
					String boardnumber2 = wpb.getBoardnumber2(); // 板材2编号
					String boardname2 = wpb.getBoardname2(); // 板材2名称
					String partmaterial2 = wpb.getPartmaterial2(); // 板材2材质
					String partthickness2 = wpb.getPartthickness2(); // 板材2板厚
					String boardnumber3 = wpb.getBoardnumber3(); // 板材3编号
					String boardname3 = wpb.getBoardname3(); // 板材3名称
					String partmaterial3 = wpb.getPartmaterial3(); // 板材3材质
					String partthickness3 = wpb.getPartthickness3(); // 板材3板厚
					String layersnum = wpb.getLayersnum(); // 板层数
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // 材料强度(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // 材料强度(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // 材料强度(Mpa)>590
					String sheetstrength12 = wpb.getSheetstrength12(); // 材料强度(Mpa)1.2G
					String basethickness = wpb.getBasethickness(); // 基准板厚

					String CurrentSerie = ""; // 参数 序列 (日产)
					String RecomWeldForce = "";// 推荐 加压力(N)
					String CurrentSeriedfi = ""; // 参数 序列 (对应)

					// 根据材质对照表，判断焊点是否参与计算焊接参数
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// 根据材质对照表获取GA/GI属性
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
									}
								}
							}																				
						}
					}
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							CurrentSerie = curenre[1];
							RecomWeldForce = curenre[4];
							CurrentSeriedfi = curenre[2];
						}
					}
					boolean flag = false;
					// 如果是1.2g高强材，基准板厚都是取最薄板
					if (sheetstrength12.equals("1.2g")) {
						flag = true;
					}
					if (flag) {
						basethickness = getMinnum(partthickness1, partthickness2, partthickness3);
					}
					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style8, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && rswsflist != null) {
						if (!rswsflist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial1)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 29, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 29, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial2)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 55, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 55, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
					if (getIscontains1180(partmaterial3)) {
						XSSFCellStyle newstyle = getXSSFStyleByrgb(book, sheet, 11 + j, 81, -1, new XSSFColor(new java.awt.Color(255,199,206)));
						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, newstyle, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 81, -1, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, newstyle, Cell.CELL_TYPE_STRING);
					}
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partthickness3, 11 + j, 87, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 90, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 92, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 92, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 94, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 96, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 98, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 98, style, 10);
					if (flag) {
						setStringCellAndStyle(sheet, "○", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
						if (getColorDistinction(layersnum, partmaterial1, partmaterial2, partmaterial3, partthickness1,
								partthickness2, partthickness3)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 1, IndexedColors.SKY_BLUE.getIndex());
							setStringCellAndStyle2(sheet, basethickness, 11 + j, 102, newstyle, 11);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 1, IndexedColors.VIOLET.getIndex());
							setStringCellAndStyle2(sheet, basethickness, 11 + j, 102, newstyle, 11);
						}
						// 后面参数序列为空
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, Cell.CELL_TYPE_STRING);
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 102, 12, IndexedColors.WHITE.getIndex());
						setStringCellAndStyle(sheet, basethickness, 11 + j, 102, newstyle, 11);
						setStringCellAndStyle(sheet, "-", 11 + j, 100, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, Cell.CELL_TYPE_STRING);
						setStringCellAndStyle(sheet, CurrentSeriedfi, 11 + j, 111, style, Cell.CELL_TYPE_STRING);
					}
				}
			}

			shnum++;
		}
	}

	private Comparator getComParatorByfirstpart() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				WeldPointBoardInformation comp1 = (WeldPointBoardInformation) obj;
				WeldPointBoardInformation comp2 = (WeldPointBoardInformation) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1.getBoardnumber1() != null && !comp1.getBoardnumber1().isEmpty()) {
					d1 = comp1.getBoardnumber1().toString();
				}
				if (obj1 != null && comp2.getBoardnumber1() != null && !comp2.getBoardnumber1().isEmpty()) {
					d2 = comp2.getBoardnumber1().toString();
					;
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}

	private Comparator getComParatorBySecondpart() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				WeldPointBoardInformation comp1 = (WeldPointBoardInformation) obj;
				WeldPointBoardInformation comp2 = (WeldPointBoardInformation) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1.getBoardnumber2() != null && !comp1.getBoardnumber2().isEmpty()) {
					d1 = comp1.getBoardnumber2().toString();
				}
				if (obj1 != null && comp2.getBoardnumber2() != null && !comp2.getBoardnumber2().isEmpty()) {
					d2 = comp2.getBoardnumber2().toString();
					;
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}

	private Comparator getComParatorByThistypart() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				WeldPointBoardInformation comp1 = (WeldPointBoardInformation) obj;
				WeldPointBoardInformation comp2 = (WeldPointBoardInformation) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1.getBoardnumber3() != null && !comp1.getBoardnumber3().isEmpty()) {
					d1 = comp1.getBoardnumber3().toString();
				}
				if (obj1 != null && comp2.getBoardnumber3() != null && !comp2.getBoardnumber3().isEmpty()) {
					d2 = comp2.getBoardnumber3().toString();
					;
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}

	/*
	 * RSW气动信息处理
	 */
	private void RSWpneumaticinformationProcessing(XSSFWorkbook book, TCComponentBOMLine bl, TCComponentBOMLine gwbl,
			String name, Map<String, TCComponentBOMLine> symmetrymap) throws TCException {
		// TODO Auto-generated method stub
		// 获取RSW气动sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // RSW气动所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains(name)) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		// 先获取工序下的枪和焊点
		ArrayList gunlist = new ArrayList();
		ArrayList hdlist = new ArrayList();
		// ArrayList robotlist = new ArrayList();
		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { guntypename, guntypename };
		String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
		String[] propertys2 = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values2 = new String[] { weldtypename, weldtypename };

//		String[] propertys3 = new String[] { "bl_item_object_type", "bl_item_object_type" };
//		String[] values3 = new String[] { "机器人", "BIW Robot" };

		gunlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		hdlist = Util.searchBOMLine(bl, "OR", propertys2, "==", values2);
		// robotlist = Util.searchBOMLine(bl, "OR", propertys3, "==", values3);

		// 根据版次是否为SOP前后，如果是SOP后，不输出焊接参数
		boolean sopflag = getIsSOPAfter();
		String Discretename = Util.getProperty(bl, "bl_rev_object_name");

		// 正常一个工序下只有一把枪
		String stationname = Util.getProperty(gwbl, "bl_rev_object_name");// 工位
		String robotname = "";// 机器人
		String Guncode = "";// 焊枪编号
		String ElectrodeVol = "";// 焊枪电压 bl_B8_BIWGunRevision_b8_ElectrodeVol
		TCComponentItemRevision blrev = null;
		if (gunlist.size() > 0) {
			TCComponentBOMLine gunbl = (TCComponentBOMLine) gunlist.get(0);
			blrev = gunbl.getItemRevision();
			Guncode = Util.getProperty(blrev, "b8_Model");
			ElectrodeVol = Util.getProperty(blrev, "b8_ElectrodeVol");
		}
		robotname = Util.getProperty(bl, "bl_rev_object_name");
//		if (robotlist.size() > 0) {
//			TCComponentBOMLine robotbl = (TCComponentBOMLine) robotlist.get(0);
//			robotname = Util.getProperty(robotbl, "bl_rev_object_name");
//		}
		// 焊枪编号，改为从点焊工序名称中取值
		String[] nameArr = Discretename.split("\\\\");
		if (nameArr.length > 1) {
			Guncode = nameArr[1];
		} else {
			Guncode = "";
		}
		robotname = nameArr[0];

		// 根据焊点在基本信息表中获取板件信息
		List<WeldPointBoardInformation> hdinfo = new ArrayList<WeldPointBoardInformation>();// 焊点所有信息

		hdinfo = getBoardInformation(baseinfolist, hdlist);

		// 获取焊接参数
		Map<String, String[]> hjmap = getHanJieParater(hdlist);

		if (hdinfo == null || hdinfo.size() < 1) {
			System.out.println("没有获取基本信息，直接跳过！");
			return;
		}

		// 先根据第一个板件排序
		Comparator comparator = getComParatorByfirstpart();
		Collections.sort(hdinfo, comparator);

		// 然后根据第二个板件排序
		Comparator comparator1 = getComParatorBySecondpart();
		Collections.sort(hdinfo, comparator1);

		// 再根据第三个板件排序
		Comparator comparator11 = getComParatorByThistypart();
		Collections.sort(hdinfo, comparator11);

		// 根据板层数排序，先输出3层板，在输出2层板
		Comparator comparator2 = getComParatorBylayersnum();
		Collections.sort(hdinfo, comparator2);

		// 如果是左右工位同出，还需要把所选工位的对称工位下，相同工序名称的焊点信息放在同一个sheet页输出
//		String Discretename = Util.getProperty(bl, "bl_rev_object_name");
		// 根据焊点在基本信息表中获取板件信息
		List<WeldPointBoardInformation> symmetryhdinfo = new ArrayList<WeldPointBoardInformation>();// 焊点所有信息
		ArrayList symmetryhdlist = new ArrayList();
		if (symmetrymap.containsKey(Discretename)) {
			TCComponentBOMLine symmetrybl = symmetrymap.get(Discretename);
			symmetryhdlist = Util.searchBOMLine(symmetrybl, "OR", propertys2, "==", values2);
			symmetryhdinfo = getBoardInformation(baseinfolist, symmetryhdlist);

			// 先根据第一个板件排序
			Comparator comparators1 = getComParatorByfirstpart();
			Collections.sort(symmetryhdinfo, comparators1);
			// 然后根据第二个板件排序
			Comparator comparators2 = getComParatorBySecondpart();
			Collections.sort(symmetryhdinfo, comparators2);

			// 再根据第三个板件排序
			Comparator comparators3 = getComParatorByThistypart();
			Collections.sort(symmetryhdinfo, comparators3);

			// 根据板层数排序，先输出3层板，在输出2层板
			Comparator comparators4 = getComParatorBylayersnum();
			Collections.sort(symmetryhdinfo, comparators4);

			// 获取焊点的焊接参数
			getSymmetryHanJieParater(symmetryhdlist, hjmap);
		}

		// 根据数据判断是否需要分页,每36行数据分一页
		int hdsum = hdlist.size() + symmetryhdlist.size();
		int page = hdsum / 36 + 1;
		// 数据行刚好一页就会出现sheet页多了一页的情况
		if (hdsum % 36 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}
		// 如果page大于1，则需要复制sheet页
		int index = sheetAtIndex + 1;

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

		if (updateflag) {
			int number = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("RSW气动")) {
					if (sheetAtIndex <= i) {
						number++;
					}
				}
			}
			index = index + page - 1;
			if (number < page) {
				if (page - number > 0) {
					for (int i = 0; i < page - number; i++) {
						XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
						book.setSheetOrder(newsheet.getSheetName(), index);
						index++;
					}
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}

		int shnum = 0;

		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);

			setStringCellAndStyle(sheet, stationname, 6, 19, style2, Cell.CELL_TYPE_STRING);// 工位
			setStringCellAndStyle(sheet, robotname, 6, 30, style2, Cell.CELL_TYPE_STRING);// 机器人
			setStringCellAndStyle(sheet, Guncode, 6, 47, style2, Cell.CELL_TYPE_STRING);// 焊枪编号
			setStringCellAndStyle(sheet, ElectrodeVol, 6, 64, style2, Cell.CELL_TYPE_STRING);// 焊枪电压
																								// bl_B8_BIWGunRevision_b8_ElectrodeVol

			if (i == index - 1) {
				for (int j = 0; j + 36 * shnum < hdinfo.size() + symmetryhdinfo.size(); j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}
					String weldno = wpb.getWeldno(); // 焊点编号
					String pageNo = "";// 页码
					String dot = "";// 打点号

					// 根据焊点号获取页码和打点号
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // 重要度
					// 判断首页是否需要添加重要度图标
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // 板材1编号
					String boardname1 = wpb.getBoardname1(); // 板材1名称
					String partmaterial1 = wpb.getPartmaterial1(); // 板材1材质
					String partthickness1 = wpb.getPartthickness1(); // 板材1板厚
					String boardnumber2 = wpb.getBoardnumber2(); // 板材2编号
					String boardname2 = wpb.getBoardname2(); // 板材2名称
					String partmaterial2 = wpb.getPartmaterial2(); // 板材2材质
					String partthickness2 = wpb.getPartthickness2(); // 板材2板厚
					String boardnumber3 = wpb.getBoardnumber3(); // 板材3编号
					String boardname3 = wpb.getBoardname3(); // 板材3名称
					String partmaterial3 = wpb.getPartmaterial3(); // 板材3材质
					String partthickness3 = wpb.getPartthickness3(); // 板材3板厚
					String layersnum = wpb.getLayersnum(); // 板层数
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // 材料强度(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // 材料强度(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // 材料强度(Mpa)>590
					String basethickness = wpb.getBasethickness(); // 基准板厚
					String sheetstrength12 = wpb.getSheetstrength12(); // 材料强度(Mpa)1.2G

					String CurrentSerie = ""; // 参数 序列 (日产)
					String RecomWeldForce = "";// 推荐 加压力(N)

					// 根据材质对照表，判断焊点是否参与计算焊接参数
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// 根据材质对照表获取GA/GI属性
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
									}
								}
							}																				
						}
					}
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							CurrentSerie = curenre[1];
							RecomWeldForce = curenre[4];
						}
					}

					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style4, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && rswqdlist != null) {
						if (!rswqdlist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}

					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
					
					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness3, 11 + j, 88, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 91, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					}

					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 95, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 97, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 99, style, 10);
					setStringCellAndStyle(sheet, basethickness, 11 + j, 102, style, 11);
					// 如果是1.2g高强材，基准板厚都是取最薄板
					if (sheetstrength12.equals("1.2g")) {
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
					} else {
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 111, style, 10);
					}
				}
			} else {
				for (int j = 0; j + 36 * shnum < 36 + 36 * shnum; j++) {
//					WeldPointBoardInformation wpb = hdinfo.get(j + 36 * shnum);
					WeldPointBoardInformation wpb = new WeldPointBoardInformation();
					if (j + 36 * shnum > hdinfo.size() - 1) {
						wpb = symmetryhdinfo.get(j + 36 * shnum - hdinfo.size());
					} else {
						wpb = hdinfo.get(j + 36 * shnum);
					}
					String weldno = wpb.getWeldno(); // 焊点编号
					String pageNo = "";// 页码
					String dot = "";// 打点号

					// 根据焊点号获取页码和打点号
					if (notelist.containsKey(weldno)) {
						String[] pd = notelist.get(weldno);
						pageNo = pd[0];
						dot = pd[1];
					}
					String importance = wpb.getImportance(); // 重要度
					// 判断首页是否需要添加重要度图标
					if (importance != null && !importance.isEmpty()) {
						if (!Import.contains(importance)) {
							Import.add(importance);
						}
					}
					String boardnumber1 = wpb.getBoardnumber1(); // 板材1编号
					String boardname1 = wpb.getBoardname1(); // 板材1名称
					String partmaterial1 = wpb.getPartmaterial1(); // 板材1材质
					String partthickness1 = wpb.getPartthickness1(); // 板材1板厚
					String boardnumber2 = wpb.getBoardnumber2(); // 板材2编号
					String boardname2 = wpb.getBoardname2(); // 板材2名称
					String partmaterial2 = wpb.getPartmaterial2(); // 板材2材质
					String partthickness2 = wpb.getPartthickness2(); // 板材2板厚
					String boardnumber3 = wpb.getBoardnumber3(); // 板材3编号
					String boardname3 = wpb.getBoardname3(); // 板材3名称
					String partmaterial3 = wpb.getPartmaterial3(); // 板材3材质
					String partthickness3 = wpb.getPartthickness3(); // 板材3板厚
					String layersnum = wpb.getLayersnum(); // 板层数
					String gagi = wpb.getGagi(); // GA /GI
					String sheetstrength440 = wpb.getSheetstrength440(); // 材料强度(Mpa)440
					String sheetstrength590 = wpb.getSheetstrength590(); // 材料强度(Mpa)590
					String sheetstrength = wpb.getSheetstrength(); // 材料强度(Mpa)>590
					String basethickness = wpb.getBasethickness(); // 基准板厚
					String sheetstrength12 = wpb.getSheetstrength12(); // 材料强度(Mpa)1.2G

					String CurrentSerie = ""; // 参数 序列 (日产)
					String RecomWeldForce = "";// 推荐 加压力(N)

					// 根据材质对照表，判断焊点是否参与计算焊接参数
					boolean partmaterialFlag1 = true;
					boolean partmaterialFlag2 = true;
					boolean partmaterialFlag3 = true;
					String gagi1 = wpb.getGagi1();								
					String gagi2 = wpb.getGagi2();								
					String gagi3 = wpb.getGagi3();

					// 根据材质对照表获取GA/GI属性
					if (MaterialMap != null) {
						for (Map.Entry<String, List<String>> entry : MaterialMap.entrySet()) {
							String MaterialNo = entry.getKey();
							List<String> infolist = entry.getValue();
							if(!"GA".equals(gagi1) && !"GI".equals(gagi1))
							{
								if (Util.getIsEqueal(partmaterial1, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag1 = false;
									}
								}
							}
							if(!"GA".equals(gagi2) && !"GI".equals(gagi2))
							{
								if (Util.getIsEqueal(partmaterial2, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag2 = false;
									}
								}
							}
							if(!"GA".equals(gagi3) && !"GI".equals(gagi3))
							{
								if (Util.getIsEqueal(partmaterial3, MaterialNo)) {
									if ("否".equalsIgnoreCase(infolist.get(1))) {
										partmaterialFlag3 = false;
									}
								}
							}																				
						}
					}
					if (partmaterialFlag1 && partmaterialFlag2 && partmaterialFlag3) {
						if (hjmap.containsKey(weldno)) {
							String[] curenre = hjmap.get(weldno);
							CurrentSerie = curenre[1];
							RecomWeldForce = curenre[4];
						}
					}

					setStringCellAndStyle(sheet, pageNo, 11 + j, 2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, dot, 11 + j, 6, style4, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, importance, 11 + j, 4, style, Cell.CELL_TYPE_STRING);
					if (updateflag && rswqdlist != null) {
						if (!rswqdlist.contains(weldno)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 2, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
							setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 11 + j, 8, 12, -1);
						setStringCellAndStyle2(sheet, weldno, 11 + j, 8, newstyle, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, boardnumber1, 11 + j, 13, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname1, 11 + j, 16, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag1) {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial1, 11 + j, 29, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness1, 11 + j, 36, style, 11);
					setStringCellAndStyle(sheet, boardnumber2, 11 + j, 39, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname2, 11 + j, 42, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag2) {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial2, 11 + j, 55, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness2, 11 + j, 62, style, 11);
					setStringCellAndStyle(sheet, boardnumber3, 11 + j, 65, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet, boardname3, 11 + j, 68, style, Cell.CELL_TYPE_STRING);
//					if (!partmaterialFlag3) {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, stylepink, Cell.CELL_TYPE_STRING);
//					} else {
//						setStringCellAndStyle2(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);
//					}
					setStringCellAndStyle(sheet, partmaterial3, 11 + j, 81, style, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sheet, partthickness3, 11 + j, 88, style, 11);
					setStringCellAndStyle(sheet, layersnum, 11 + j, 91, style, 10);
					if (gagi != null && !gagi.isEmpty()) {
						setStringCellAndStyle(sheet, gagi, 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					} else {
						setStringCellAndStyle(sheet, "-", 11 + j, 93, style, Cell.CELL_TYPE_STRING);
					}
					setStringCellAndStyle(sheet, sheetstrength440, 11 + j, 95, style, 10);
					setStringCellAndStyle(sheet, sheetstrength590, 11 + j, 97, style, 10);
					setStringCellAndStyle(sheet, sheetstrength, 11 + j, 99, style, 10);
					setStringCellAndStyle(sheet, basethickness, 11 + j, 102, style, 11);
					// 如果是1.2g高强材，基准板厚都是取最薄板
					if (sheetstrength12.equals("1.2g")) {
						setStringCellAndStyle(sheet, "", 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, "", 11 + j, 111, style, 10);
					} else {
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 105, style, 10);
						setStringCellAndStyle(sheet, RecomWeldForce, 11 + j, 108, style, 10);
						setStringCellAndStyle(sheet, CurrentSerie, 11 + j, 111, style, 10);
					}
				}
			}

			shnum++;
		}

	}

	/*
	 * 获取焊接参数
	 */
	private Map<String, String[]> getHanJieParater(ArrayList hdlist) throws TCException {
		// TODO Auto-generated method stub
		Map<String, String[]> map = new HashMap<String, String[]>();
		if (hdlist != null && hdlist.size() > 0) {
			TCComponent[] tccs = new TCComponent[hdlist.size()];
			for (int i = 0; i < hdlist.size(); i++) {
				TCComponentBOMLine bl = (TCComponentBOMLine) hdlist.get(i);
				TCComponentItemRevision rev = bl.getItemRevision();
				tccs[i] = rev;
				String weldno = Util.getProperty(rev, "object_name");
			}
			String[] properties = { "object_name", "b8_CurrentSerie_Nissan", "b8_CurrentSerie_DFL", "b8_Current2",
					"b8_RecomWeldForce" };
			String[][] values = Util.getAllProperties(session, tccs, properties);
			for (int j = 0; j < values.length; j++) {
				String[] proper = values[j];
				map.put(proper[0], proper);
			}
		}
		return map;
	}

	/*
	 * 式样差信息处理
	 */
	private void PoorPatternProcessing(XSSFWorkbook book, List assylist, boolean rLflag) {
		// TODO Auto-generated method stub
		ArrayList poorlist = new ArrayList();
		// 获取式样差sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // 式样差所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("式样差")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		int poornum = 0;// 式样差件数量
		for (Map.Entry<String, String> entry : fymap.entrySet()) {
			String key = entry.getKey();
			int value = Integer.parseInt(entry.getValue());
			if (value > 1) {
				poornum++;
				List temp = new ArrayList();
				List afterName = new ArrayList(); // 部品番号后5位
				for (int i = 0; i < partlist.size(); i++) {
					String[] str = (String[]) partlist.get(i);
					if (key.equals(str[7])) {
						String[] station = new String[3];
						station[0] = str[1];
						if (str[1].length() > 5) {
							String afterno = str[1].substring(5);
							if (!afterName.contains(afterno)) {
								afterName.add(afterno);
							}
						} else {
							if (!afterName.contains(str[1])) {
								afterName.add(str[1]);
							}
						}
						station[1] = str[2];
						station[2] = Integer.toString(poornum);
						temp.add(station);
					}
				}
				// 如果是左右工位，需要左右工位一行输出
				if (rLflag) {
					for (int j = 0; j < afterName.size(); j++) {
						String ApartNO = (String) afterName.get(j);
						List ttlist = new ArrayList();
						for (int k = 0; k < temp.size(); k++) {
							String[] val = (String[]) temp.get(k);
							if (val[0].length() > 5) {
								if (ApartNO.equals(val[0].substring(5))) {
									ttlist.add(val);
								}
							} else {
								if (ApartNO.equals(val[0])) {
									ttlist.add(val);
								}
							}
						}
						if (ttlist.size() == 2) {
							String[] str1 = (String[]) ttlist.get(0);
							String[] str2 = (String[]) ttlist.get(1);
							String[] str3 = new String[4];
							if (str1[0].length() > 4 && str2[0].length() > 4) {
								str3[0] = str1[0].substring(0, 5) + "/" + str2[0].substring(4, 5) + ApartNO;
							} else {
								str3[0] = str1[0];
							}
							if (str3[1] != null && str3[1].length() > 2) {
								str3[1] = str1[1].substring(0, str1[1].length() - 2) + "L/RH";
							} else {
								str3[1] = str1[1] + "L/RH";
							}

							str3[2] = str1[2];
							str3[3] = ApartNO;
							poorlist.add(str3);
						} else {
							if (ttlist.size() > 0) {
								String[] val = (String[]) temp.get(j);
								String[] strvalue = new String[4];
								if (val[0].length() > 5) {
									strvalue[0] = val[0].substring(0, 5);
									strvalue[1] = val[1];
									strvalue[2] = val[2];
									strvalue[3] = val[0].substring(5);
								} else {
									strvalue[0] = val[0];
									strvalue[1] = val[1];
									strvalue[2] = val[2];
									strvalue[3] = val[0];
								}

								poorlist.add(strvalue);
							}
						}
					}

				} else {
					for (int j = 0; j < temp.size(); j++) {
						String[] val = (String[]) temp.get(j);
						String[] strvalue = new String[4];
						if (val[0].length() > 5) {
							strvalue[0] = val[0].substring(0, 5);
							strvalue[1] = val[1];
							strvalue[2] = val[2];
							strvalue[3] = val[0].substring(5);
						} else {
							strvalue[0] = val[0];
							strvalue[1] = val[1];
							strvalue[2] = val[2];
							strvalue[3] = val[0];
						}

						poorlist.add(strvalue);
					}
				}
			}
		}
//		for (Map.Entry<String, String> entry : fymap.entrySet()) {
//			String key = entry.getKey();
//			int value = Integer.parseInt(entry.getValue());
//			if (value > 1) {
//				String partno = "";
//				String partname = "";
//				poornum++;
//
//				for (int i = 0; i < partlist.size(); i++) {
//					String[] str = (String[]) partlist.get(i);
//					if (key.equals(str[7])) {
//						String[] station = new String[3];
//						station[0] = str[1];
//						station[1] = str[2];
//						station[2] = Integer.toString(poornum);
//						poorlist.add(station);
//					}
//				}
//			}
//		}
		// 根据数据判断是否需要分页,每3个行数据分一页
		int page = poornum / 3 + 1;

		// 数据行刚好一页就会出现sheet页多了一页的情况
		if (poornum % 3 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}

		// 如果page大于1，则需要复制sheet页
		int index = sheetAtIndex + 1;

		// 设置字体颜色
		Font font = book.createFont();
		font.setColor((short) 12);// 蓝色字体
		// font.setFontName("宋体");
		font.setFontHeightInPoints((short) 14);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		// style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style2.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);// 左边框
		style2.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style2.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font);

		XSSFCellStyle style3 = book.createCellStyle();
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style3.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);// 上边框
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);

		Font font2 = book.createFont();
		font2.setColor((short) 2);// 红色字体
		// font.setFontName("宋体");
		font2.setFontHeightInPoints((short) 14);
		XSSFCellStyle style00 = book.createCellStyle();
		style00.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		// style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style00.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style00.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style00.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style00.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style00.setFont(font2);

		XSSFCellStyle style22 = book.createCellStyle();
		style22.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style22.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);// 左边框
		style22.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style22.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style22.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style22.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style22.setFont(font2);

		XSSFCellStyle style33 = book.createCellStyle();
		style33.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style33.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);// 上边框
		style33.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style33.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style33.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style33.setFont(font2);

		/**************************************************/
		// 如果是更新，则先把系统输出的信息清空，后面再写入
		List assynameList = new ArrayList();// Ａｓｓｙ名称
		List assyList = new ArrayList();// Ａｓｓｙ部番
		if (updateflag) {
			int gcnum = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("式样差")) {
					gcnum++;
				}
			}
			// 如果sheet页增加就增，减少不删除，保留
			index = sheetAtIndex + page;

			XSSFCell cell;
			XSSFRow row;
			// 循环构成表sheet页清空系统输出内容，手工维护内容保留
			for (int i = sheetAtIndex; i < sheetAtIndex + gcnum; i++) {
				XSSFSheet sheet = book.getSheetAt(i);
				row = sheet.getRow(6);
				if (row != null) {
					cell = row.getCell(9);
					if (cell != null) {
						String assyname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (assyname != null && !assyname.isEmpty()) {
							assynameList.add(assyname);
						}
					}
					cell = row.getCell(46);
					if (cell != null) {
						String assyname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (assyname != null && !assyname.isEmpty()) {
							assynameList.add(assyname);
						}
					}
					cell = row.getCell(84);
					if (cell != null) {
						String assyname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (assyname != null && !assyname.isEmpty()) {
							assynameList.add(assyname);
						}
					}
				}
				// 清空内容
				setStringCellAndStyle(sheet, "", 6, 9, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
				setStringCellAndStyle(sheet, "", 6, 46, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
				setStringCellAndStyle(sheet, "", 6, 84, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称

				for (int j = 0; j < 6; j++) {
					row = sheet.getRow(34 + j * 2);
					if (row != null) {
						cell = row.getCell(1);
						if (cell != null) {
							String preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(7);
							String sufxxassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							if ((preassy != null && !preassy.isEmpty())
									|| (sufxxassy != null && !sufxxassy.isEmpty())) {
								String asstno = preassy.trim() + sufxxassy.trim();
								assyList.add(asstno);
							}
						}
						cell = row.getCell(38);
						if (cell != null) {
							String preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(44);
							String sufxxassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							if ((preassy != null && !preassy.isEmpty())
									|| (sufxxassy != null && !sufxxassy.isEmpty())) {
								String asstno = preassy.trim() + sufxxassy.trim();
								assyList.add(asstno);
							}
						}
						cell = row.getCell(76);
						if (cell != null) {
							String preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(82);
							String sufxxassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							if ((preassy != null && !preassy.isEmpty())
									|| (sufxxassy != null && !sufxxassy.isEmpty())) {
								String asstno = preassy.trim() + sufxxassy.trim();
								assyList.add(asstno);
							}
						}
					}
					setStringCellAndStyle(sheet, "", 34 + j * 2, 1, style2, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
					setStringCellAndStyle(sheet, "", 34 + j * 2, 7, style, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
					setStringCellAndStyle(sheet, "", 34 + j * 2, 38, style2, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
					setStringCellAndStyle(sheet, "", 34 + j * 2, 44, style, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
					setStringCellAndStyle(sheet, "", 34 + j * 2, 76, style2, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
					setStringCellAndStyle(sheet, "", 34 + j * 2, 82, style, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
				}
			}
			if (page > gcnum) {
				for (int i = 1; i < page - gcnum; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		} else {
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}
		/**************************************************/

		int shnum = 0;
		for (int i = sheetAtIndex; i < index; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			if (i == index - 1) {
				for (int j = 0; j + 3 * shnum < poornum; j++) {
					String partname = "";
					int rownum = 0;
					for (int k = 0; k < poorlist.size(); k++) {
						String[] str = (String[]) poorlist.get(k);
						if (j + 1 + 3 * shnum == Integer.parseInt(str[2])) {
							partname = str[1];
							String prename = str[0];
							String aftername = str[3];
//							if (str[0].length() > 5) {
//								prename = str[0].substring(0, 5);
//								aftername = str[0].substring(5).trim();
//							} else {
//								prename = str[0];
//								aftername = str[0];
//							}
							if ((j + 1 + 3 * shnum) % 3 == 1) {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style22,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style00,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style2,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style2,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
								}

							} else if ((j + 1 + 3 * shnum) % 3 == 2) {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style22,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style00,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style2,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style2,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
								}

							} else {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style22,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style00,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style2,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style2,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
								}

							}
							rownum++;
						}
					}
					if ((j + 1 + 3 * shnum) % 3 == 1) {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 9, style33, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 9, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 9, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
						}

					} else if ((j + 1 + 3 * shnum) % 3 == 2) {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 46, style33, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 46, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 46, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
						}
					} else {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 84, style33, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 84, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 84, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
						}
					}
				}
			} else {
				for (int j = 0; j + 3 * shnum < 3 + 3 * shnum; j++) {
					String partname = "";
					int rownum = 0;
					for (int k = 0; k < poorlist.size(); k++) {
						String[] str = (String[]) poorlist.get(k);
						if (j + 1 + 3 * shnum == Integer.parseInt(str[2])) {
							partname = str[1];
							String prename = str[0];
							String aftername = str[3];
//							if (str[0].length() > 5) {
//								prename = str[0].substring(0, 5);
//								aftername = str[0].substring(5).trim();
//							} else {
//								prename = str[0];
//								aftername = str[0];
//							}
							if ((j + 1 + 3 * shnum) % 3 == 1) {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style22,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style00,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style2,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 1, style2,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 7, style,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
								}

							} else if ((j + 1 + 3 * shnum) % 3 == 2) {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style22,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style00,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style2,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 38, style2,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 44, style,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
								}

							} else {
								if (updateflag && assyList != null) {
									String allassy = prename.trim() + aftername.trim();
									if (!assyList.contains(allassy)) {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style22,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style00,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									} else {
										setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style2,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
										setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style,
												Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
									}
								} else {
									setStringCellAndStyle2(sheet, prename, 34 + rownum * 2, 76, style2,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
									setStringCellAndStyle2(sheet, aftername, 34 + rownum * 2, 82, style,
											Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
								}

							}
							rownum++;
						}
					}
					if ((j + 1 + 3 * shnum) % 3 == 1) {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 9, style33, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 9, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 9, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
						}

					} else if ((j + 1 + 3 * shnum) % 3 == 2) {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 46, style33, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 46, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 46, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
						}
					} else {
						if (updateflag && assynameList != null) {
							if (!assynameList.contains(partname)) {
								setStringCellAndStyle2(sheet, partname, 6, 84, style33, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							} else {
								setStringCellAndStyle2(sheet, partname, 6, 84, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
							}
						} else {
							setStringCellAndStyle2(sheet, partname, 6, 84, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
						}
					}
				}
			}
			shnum++;
		}

	}

	/*
	 * 构成图信息处理
	 */
	private void CompositionChartProcessing(XSSFWorkbook book, List assylist, String assyname, boolean rLflag,
			Map<String, File> piclist) {
		// TODO Auto-generated method stub
//		if (!updateflag) 
//		{
//			for(int i=0;i<deletelist.size();i++)
//			{
//				if(deletelist.get(i).toString().contains("构成图"))
//				{
//					return;
//				}
//			}			
//		}	
		ArrayList complist = new ArrayList();
		// 获取构成图sheet
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // 构成图所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("构成图")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		//部品号与名称
//		Map<String,String> partTonum = new HashMap<String,String>();
//		for (int i = 0; i < partlist.size(); i++) 
//		{
//			String[] str = (String[]) partlist.get(i);
//			if(partlist.contains(str[1]))
//			{
//				partTonum.put(str[1], str[2]);
//			}
//		}
		// 式样差数据不放在一行显示，改为左右工位的放在一行显示
		for (Map.Entry<String, String> entry : fymap.entrySet()) {
			String key = entry.getKey();
			// 先把同一标号的数据取出
			List tempList = new ArrayList();
			List afterName = new ArrayList(); // 部品番号后5位
			for (int i = 0; i < partlist.size(); i++) {
				String[] str = (String[]) partlist.get(i);
				if (key.equals(str[7])) {
					tempList.add(str);
					if (str[1] != null && str[1].length() > 5) {
						String afterno = str[1].substring(5);
						if (!afterName.contains(afterno)) {
							afterName.add(afterno);
						}
					} else {
						if (!afterName.contains(str[1])) {
							afterName.add(str[1]);
						}
					}
				}
			}
			// 如果是左右工位把零件号的后5位相同的放在一行显示
			if (rLflag) {
				for (int j = 0; j < afterName.size(); j++) {
					String ApartNO = (String) afterName.get(j);
					if(ApartNO == null)
					{
						ApartNO = "";
					}
					List ttlist = new ArrayList();
					for (int k = 0; k < tempList.size(); k++) {
						String[] val = (String[]) tempList.get(k);
						if (val[1] != null && val[1].length() > 5) {
							if (ApartNO.equals(val[1].substring(5))) {
								ttlist.add(val);
							}
						} else {
							if (ApartNO.equals(val[1])) {
								ttlist.add(val);
							}
						}
					}
					if (ttlist.size() == 2) {
						String[] str1 = (String[]) ttlist.get(0);
						String[] str2 = (String[]) ttlist.get(1);
						String[] str3 = new String[3];
						if (str1[1] != null && str1[1].length() > 5 && str2[1] != null && str2[1].length() > 5) {
							str3[0] = str1[1].substring(0, 5) + "/" + str2[1].substring(4, 5) + ApartNO;
						} else {
							str3[0] = str1[1] + "/" + str2[1];
						}
						if (str1[2] != null && str1[2].length() > 2) {
							str3[1] = str1[2].substring(0, str1[2].length() - 2) + "L/RH";
						} else {
							str3[1] = str1[2] + "L/RH";
						}

						str3[2] = str1[2];
						//如果部品名称完全相同，则认为是标椎件，非左右件
						if(str1[2].equals(str2[2]))
						{
							String[] station = new String[3];
							station[0] = str1[7] + " " + str1[1] + " " + str1[2];
							station[1] = str1[6];
							station[2] = str1[8];
							complist.add(station);
							
							String[] station2 = new String[3];
							station2[0] = str2[7] + " " + str2[1] + " " + str2[2];
							station2[1] = str2[6];
							station2[2] = str2[8];
							complist.add(station2);
						}
						else
						{
							String[] station = new String[3];
							station[0] = str1[7] + " " + str3[0] + " " + str3[1];
							station[1] = str1[6];
							station[2] = str1[8];
							complist.add(station);
						}
						
					} else {
						if (ttlist.size() > 0) {
							for (int m = 0; m < ttlist.size(); m++) {
								String[] val = (String[]) ttlist.get(m);
								String[] station = new String[3];
								station[0] = val[7] + " " + val[1] + " " + val[2];
								station[1] = val[6];
								station[2] = val[8];
								complist.add(station);
							}
						}
					}
				}

			} else {
				for (int j = 0; j < tempList.size(); j++) {
					String[] val = (String[]) tempList.get(j);
					String[] station = new String[3];
					station[0] = val[7] + " " + val[1] + " " + val[2];
					station[1] = val[6];
					station[2] = val[8];
					complist.add(station);
				}
			}
		}

//		for (Map.Entry<String, String> entry : fymap.entrySet()) {
//			String key = entry.getKey();
//			String stationname = "";
//			String template = "";
//			String type = "";
//			String[] station = new String[2];
//			for (int i = 0; i < partlist.size(); i++) {
//				String[] str = (String[]) partlist.get(i);
//				if (key.equals(str[7])) {
//					if (template.isEmpty()) {
//						template = str[1];
//					} else {
//						if (str[1].length() > 5) {
//							template = template + "/" + str[1].substring(5).trim();
//						} else {
//							template = template + "/" + str[1];
//						}
//					}
//					stationname = template + " " + str[2];
//					type = str[6];
//				}
//			}
//			station[0] = stationname;
//			station[1] = type;
//			template = "";
//			complist.add(station);
//		}

		// 根据数据判断是否需要隔行输出,大于15行数据就不隔行输出
		// int sum = fymap.size();
		int sum = complist.size();
		int page = sum / 15 + 1;
		// 如果page大于1，则不隔行输出

		// 写构成图数据
		// 设置字体颜色
		Font font = book.createFont();
		font.setColor((short) 12);
		font.setFontName("宋体");
		XSSFCellStyle style = book.createCellStyle();
		style.setFont(font);
		style.setBorderBottom(CellStyle.BORDER_DOTTED); // 虚线边框 BORDER_HAIR  
		style.setBorderLeft(CellStyle.BORDER_THIN); // 虚线边框 BORDER_DOTTED
		style.setBorderRight(CellStyle.BORDER_DOTTED); // 虚线边框
		style.setBorderTop(CellStyle.BORDER_DOTTED); // 虚线边框
		

		XSSFCellStyle style2 = book.createCellStyle();
		style2.setFont(font);
		style2.setBorderBottom(CellStyle.BORDER_DOUBLE); // 双线边框 
		style2.setBorderLeft(CellStyle.BORDER_THIN); // 双线边框
		style2.setBorderRight(CellStyle.BORDER_DOUBLE); // 双线边框
		style2.setBorderTop(CellStyle.BORDER_DOUBLE); // 双线边框

		XSSFCellStyle style3 = book.createCellStyle();
		style3.setFont(font);
		style3.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
		style3.setBorderLeft(CellStyle.BORDER_THIN); // 粗线边框
		style3.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
		style3.setBorderTop(CellStyle.BORDER_MEDIUM); // 粗线边框

		XSSFCellStyle style4 = book.createCellStyle();
		style4.setFont(font);
		style4.setBorderBottom(CellStyle.BORDER_NONE); // 无边框
		style4.setBorderLeft(CellStyle.BORDER_NONE); // 无边框
		style4.setBorderRight(CellStyle.BORDER_NONE); // 无边框
		style4.setBorderTop(CellStyle.BORDER_NONE); // 无边框

		XSSFCellStyle style5 = book.createCellStyle();
		style5.setFont(font);
		style5.setBorderLeft(CellStyle.BORDER_THIN); // 左边框

		Font font2 = book.createFont();
		font2.setColor((short) 2);
		font2.setFontName("宋体");
		XSSFCellStyle style00 = book.createCellStyle();
		style00.setFont(font2);
		style00.setBorderBottom(CellStyle.BORDER_DOTTED); // 虚线边框
		style00.setBorderLeft(CellStyle.BORDER_THIN); // 虚线边框
		style00.setBorderRight(CellStyle.BORDER_DOTTED); // 虚线边框
		style00.setBorderTop(CellStyle.BORDER_DOTTED); // 虚线边框

		XSSFCellStyle style22 = book.createCellStyle();
		style22.setFont(font2);
		style22.setBorderBottom(CellStyle.BORDER_DOUBLE); // 双线边框
		style22.setBorderLeft(CellStyle.BORDER_THIN); // 双线边框
		style22.setBorderRight(CellStyle.BORDER_DOUBLE); // 双线边框
		style22.setBorderTop(CellStyle.BORDER_DOUBLE); // 双线边框

		XSSFCellStyle style33 = book.createCellStyle();
		style33.setFont(font2);
		style33.setBorderBottom(CellStyle.BORDER_MEDIUM); // 粗线边框
		style33.setBorderLeft(CellStyle.BORDER_THIN); // 粗线边框
		style33.setBorderRight(CellStyle.BORDER_MEDIUM); // 粗线边框
		style33.setBorderTop(CellStyle.BORDER_MEDIUM); // 粗线边框

		XSSFCellStyle style44 = book.createCellStyle();
		style44.setFont(font2);
		style44.setBorderBottom(CellStyle.BORDER_NONE); // 无边框
		style44.setBorderLeft(CellStyle.BORDER_NONE); // 无边框
		style44.setBorderRight(CellStyle.BORDER_NONE); // 无边框
		style44.setBorderTop(CellStyle.BORDER_NONE); // 无边框

		int shnum = 0;

		// 工位总成名称
		String procStation = "";
		String tempstation = "";
//		if (assynos != null) {
//			for (int i = 0; i < assynos.length; i++) {
//				if (i == 0) {
//					tempstation = assynos[i];
//				} else {
//					if (assynos[i].length() > 5) {
//						tempstation = tempstation + "/" + assynos[i].substring(5).trim();
//					} else {
//						tempstation = tempstation + "/" + assynos[i];
//					}
//				}
//			}
//		}
		if (assylist != null) {
			List aftertemp = new ArrayList();
			String fronttemp = "";
			String backtemp = "";
			for (int i = 0; i < assylist.size(); i++) {
				String strVal = (String) assylist.get(i);
				if (i == 0) {
					if (strVal != null && strVal.length() >= 5) {
						fronttemp = strVal.substring(0, 5);
						backtemp = strVal.substring(5).trim();
					} else {
						fronttemp = strVal;
						backtemp = "";
					}

				} else {
					if (strVal != null && strVal.length() >= 5) {
						if (!fronttemp.equals(strVal.substring(0, 5))) {
							fronttemp = fronttemp + "/" + strVal.substring(4, 5);
						}
						if (!backtemp.equals(strVal.substring(5))) {
							backtemp = backtemp + "/" + strVal.substring(5).trim();
						}
					} else {
						fronttemp = fronttemp + strVal;
						backtemp = "";
					}
				}
			}
			tempstation = fronttemp + " " + backtemp;
			System.out.println("tempstation:" + tempstation);
		}

		if (rLflag) {
			procStation = tempstation + " " + assyname.replace("LH", "").replace("RH", "") + " " + "L/RH";
		} else {
			procStation = tempstation + " " + assyname;
		}
		System.out.println("procStation:" + procStation);
		// 如果部品番号为空，不输出
		if (assylist == null || assylist.size() < 1) {
			procStation = "";
		}

		XSSFSheet sheet = book.getSheetAt(sheetAtIndex);
		/***********************************************/
		// 如果是更新，则先把系统输出的信息清空，后面再写入
		String oldproc = "FLAG";
		List oldcomplist = new ArrayList();// partlist 数据
		if (updateflag) {
			XSSFRow row;
			XSSFCell cell;
			row = sheet.getRow(8);
			if (row != null) {
				cell = row.getCell(3);
				if (cell != null) {
					oldproc = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
				}
			}
			setStringCellAndStyle(sheet, "", 8, 3, style4, Cell.CELL_TYPE_STRING);
			for (int j = 0; j < 30; j++) {
				row = sheet.getRow(10 + j);
				if (row != null) {
					cell = row.getCell(6);
					if (cell != null) {
						String partname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
						if (partname != null && !partname.isEmpty()) {
							oldcomplist.add(partname);
						}
					}
				}

				setStringCellAndStyle(sheet, "", 10 + j, 6, style4, Cell.CELL_TYPE_STRING);
				for (int n = 7; n < 32; n++) {
					setStringCellAndStyle(sheet, "", 10 + j, n, style4, Cell.CELL_TYPE_STRING);
				}
			}
		}
		/***********************************************/
		CellRangeAddress region = null;
		if (!oldproc.equals("FLAG") && !oldproc.equals(procStation.trim())) {
			setStringCellAndStyle2(sheet, procStation.trim(), 8, 3, style22, Cell.CELL_TYPE_STRING);
		} else {
			setStringCellAndStyle2(sheet, procStation.trim(), 8, 3, style2, Cell.CELL_TYPE_STRING);
		}
		// 为了加边框，每个单元格都写空值
		for (int n = 4; n < 32; n++) {
			setStringCellAndStyle2(sheet, "", 8, n, style2, Cell.CELL_TYPE_STRING);
		}
//		sheet.removeMergedRegion(getMergedRegionIndex(sheet, 8, 3));
//		region = new CellRangeAddress(8, 8, (short) 3, (short) 31);
//		sheet.addMergedRegion(region);
		if (page > 1) {
			setStringCellAndStyle2(sheet, "", 9, 6, style5, Cell.CELL_TYPE_STRING);
			for (int j = 0; j < complist.size(); j++) {
				String[] str = (String[]) complist.get(j);
				if (str[2] != null && str[2].equals("Stamping")) {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style33, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style3, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style3, Cell.CELL_TYPE_STRING);
					}
					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j, n, style3, Cell.CELL_TYPE_STRING);
					}
				} else if (str[2] != null && str[2].equals("Purchase")) {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style00, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style, Cell.CELL_TYPE_STRING);
					}

					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j, n, style, Cell.CELL_TYPE_STRING);
					}
				} else {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style22, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style2, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j, 6, style2, Cell.CELL_TYPE_STRING);
					}

					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j, n, style2, Cell.CELL_TYPE_STRING);
					}
				}
				// sheet.removeMergedRegion(getMergedRegionIndex(sheet, 10 + j, 6));
				region = new CellRangeAddress(10 + j, 10 + j, (short) 6, (short) 31);
				if (updateflag) {
					int mr1 = getMergedRegionIndex(sheet, 10 + j, 6);
					System.out.println("mr1: " + mr1);
					if (mr1 == -1) {
						sheet.addMergedRegion(region);
					}
				} else {
					sheet.addMergedRegion(region);
				}

			}
		} else {
			for (int j = 0; j < complist.size(); j++) {
				String[] str = (String[]) complist.get(j);
				setStringCellAndStyle2(sheet, "", 10 + j * 2 - 1, 6, style5, Cell.CELL_TYPE_STRING);
				if (str[2] != null && str[2].equals("Stamping")) {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style33, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style3, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style3, Cell.CELL_TYPE_STRING);
					}

					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j * 2, n, style3, Cell.CELL_TYPE_STRING);
					}
				} else if (str[2] != null && str[2].equals("Purchase")) {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style00, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style, Cell.CELL_TYPE_STRING);
					}

					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j * 2, n, style, Cell.CELL_TYPE_STRING);
					}
				} else {
					if (oldcomplist != null && updateflag) {
						if (!oldcomplist.contains(str[0].trim())) {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style22, Cell.CELL_TYPE_STRING);
						} else {
							setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style2, Cell.CELL_TYPE_STRING);
						}
					} else {
						setStringCellAndStyle2(sheet, str[0].trim(), 10 + j * 2, 6, style2, Cell.CELL_TYPE_STRING);
					}
					for (int n = 7; n < 32; n++) {
						setStringCellAndStyle2(sheet, "", 10 + j * 2, n, style2, Cell.CELL_TYPE_STRING);
					}
				}
				// sheet.removeMergedRegion(getMergedRegionIndex(sheet, 10 + j * 2, 6));
				region = new CellRangeAddress(10 + j * 2, 10 + j * 2, (short) 6, (short) 31);
				if (updateflag) {
					int mr1 = getMergedRegionIndex(sheet, 10 + j * 2, 6);
					System.out.println("mr1: " + mr1);
					if (mr1 == -1) {
						sheet.addMergedRegion(region);
					}
				} else {
					sheet.addMergedRegion(region);
				}
			}
		}

		// 写入图片到sheet中
		if (piclist != null && piclist.size() > 0) {

			// 画图的顶级管理器对象HSSFPatriarch, 一个sheet只能获取一个
			XSSFDrawing hssfPatriarch = (XSSFDrawing) sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor1 = null;
			XSSFRichTextString strValue = new XSSFRichTextString();
			int count = 0;
			int rowindex = 0;
			int colindex = 0;

			for (Entry<String, File> entry : piclist.entrySet()) {
				if ((count + 1) % 3 == 1) {
					rowindex = 0;
				} else if ((count + 1) % 3 == 2) {
					rowindex = 1;
				} else {
					rowindex = 2;
				}
				colindex = (count) / 3;
				count++;
				// String objectname = entry.getKey().replace("/", " ");
				File file = entry.getValue();
				BufferedImage image = null;
				try {
					image = ImageIO.read(file);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				int width = image.getWidth();
				int hight = image.getHeight();
				double diff = width * 1.0 / hight;
				int h = 13;
				int w = (int) (h * diff);

				writepicturetosheet(book, sheet, image, 5 + rowindex * 14, 37 + colindex * 25, 18 + rowindex * 14,
						(37 + w) + colindex * 25);
			}

		}

	}

	/*
	 * 构成表信息处理
	 */
	private void PartsinformationProcessing(XSSFWorkbook book, List assynos, List assynamelist) {
		// TODO Auto-generated method stub
		// 获取构成表sheet
//		if (!updateflag) 
//		{
//			for(int i=0;i<deletelist.size();i++)
//			{
//				if(deletelist.get(i).toString().contains("构成表"))
//				{
//					return;
//				}
//			}			
//		}	
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		int sheetAtIndex = -1; // 构成表所在位置
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("构成表")) {
				sheetAtIndex = i;
				break;
			}
		}
		if (sheetAtIndex == -1) {
			return;
		}
		// 根据数据判断是否需要分页
		int sum = 0;
//		for (Map.Entry<String, String> entry : fymap.entrySet()) {
//			sum = sum + Integer.parseInt(entry.getValue()) + 1;
//		}
		sum = partlist.size();
		// 每24行分一个sheet页
		int page = sum / 24 + 1;

		// 数据行刚好一页就会出现sheet页多了一页的情况
		if (sum % 24 == 0) {
			if (page > 1) {
				page = page - 1;
			}
		}

		int index = sheetAtIndex + 1;

		// 设置字体颜色
		Font font = book.createFont();
		font.setColor((short) 12);// 蓝色字体
		font.setFontName("MS PGothic");
		font.setFontHeightInPoints((short) 12);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		Font font2 = book.createFont();
		font2.setFontName("MS PGothic");
		font2.setColor((short) 12);// 蓝色字体
		font2.setFontHeightInPoints((short) 12);
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style2.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);// 左边框
		style2.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style2.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
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
		font3.setFontName("MS PGothic");
		font3.setFontHeightInPoints((short) 12);
		XSSFCellStyle style5 = book.createCellStyle();
		style5.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style5.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style5.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style5.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style5.setFont(font3);

		XSSFCellStyle style6 = book.createCellStyle();
		style6.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		// style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style6.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style6.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style6.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style6.setFont(font3);

		XSSFCellStyle style7 = book.createCellStyle();
		style7.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style7.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style7.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		// style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style7.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style7.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style7.setFont(font3);

		/***********************************************/
		// 如果是更新，则先把系统输出的信息清空，后面再写入
		// 清空之前先获取报表中的数据，用于跟当前获取的数据匹配，判断不一样的数据并标红
		List oldassynos = new ArrayList();
		List oldpartList = new ArrayList();
		List oldpartname = new ArrayList();
		if (updateflag) {
			int gcnum = 0;
			for (int i = 0; i < sheetnum; i++) {
				String sheetname = book.getSheetName(i);
				if (sheetname.contains("构成表")) {
					gcnum++;
				}
			}
			// 如果sheet页增加就增，减少不删除，保留
			index = sheetAtIndex + page;

			// 循环构成表sheet页清空系统输出内容，手工维护内容保留
			XSSFCell cell;
			XSSFRow row;
			for (int i = sheetAtIndex; i < sheetAtIndex + gcnum; i++) {
				XSSFSheet sheet = book.getSheetAt(i);
				// 清空Assy List内容
				for (int j = 0; j < 10; j++) {
					row = sheet.getRow(9 + j);
					String preassy;// Ａｓｓｙ部番前缀
					String suffixassy;// Ａｓｓｙ部番后缀
					if (row != null) {
						cell = row.getCell(7);
						if (cell != null) {
							preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(13);
							if (cell != null) {
								suffixassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
								if ((preassy != null && !preassy.isEmpty())
										|| (suffixassy != null && !suffixassy.isEmpty())) {
									String allassy = preassy.trim() + suffixassy.trim();
									oldassynos.add(allassy);
								}
							}
						}
					}
					setStringCellAndStyle(sheet, "", 9 + j, 7, style, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
					setStringCellAndStyle(sheet, "", 9 + j, 13, style, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
					setStringCellAndStyle(sheet, "", 9 + j, 19, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
				}
				// 清空Part List内容
				for (int j = 0; j < 24; j++) {
					row = sheet.getRow(23 + j);
					String preassy;// Ａｓｓｙ部番前缀
					String suffixassy;// Ａｓｓｙ部番后缀
					if (row != null) {
						cell = row.getCell(13);
						if (cell != null) {
							preassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
							cell = row.getCell(18);
							if (cell != null) {
								suffixassy = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
								if ((preassy != null && !preassy.isEmpty())
										|| (suffixassy != null && !suffixassy.isEmpty())) {
									String allassy = preassy.trim() + suffixassy.trim();
									oldpartList.add(allassy);
								}
							}
							cell = row.getCell(24);
							if (cell != null) {
								String partname = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
								if (partname != null && !partname.isEmpty()) {
									oldpartname.add(partname);
								}
							}
						}
					}
					setStringCellAndStyle(sheet, "", 23 + j, 4, style2, Cell.CELL_TYPE_STRING);// 标号
					setStringCellAndStyle(sheet, "", 23 + j, 7, style, Cell.CELL_TYPE_STRING);// 安装顺序
					setStringCellAndStyle(sheet, "", 23 + j, 13, style4, Cell.CELL_TYPE_STRING);// 部品番号前缀
					setStringCellAndStyle(sheet, "", 23 + j, 18, style3, Cell.CELL_TYPE_STRING);// 部品番号后缀
					setStringCellAndStyle(sheet, "", 23 + j, 24, style3, Cell.CELL_TYPE_STRING);// 部品名称
					setStringCellAndStyle(sheet, "", 23 + j, 50, style, Cell.CELL_TYPE_STRING);// 数量
					setStringCellAndStyle(sheet, "", 23 + j, 53, style, Cell.CELL_TYPE_STRING);// 板厚
					setStringCellAndStyle(sheet, "", 23 + j, 58, style, Cell.CELL_TYPE_STRING);// 材质
					setStringCellAndStyle(sheet, "", 23 + j, 72, style, Cell.CELL_TYPE_STRING);// 部品来源
				}
			}
			if (gcnum < page) {
				for (int i = 1; i < page - gcnum; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}

		} else {
			// 如果page大于1，则需要复制sheet页
			if (page > 1) {
				for (int i = 1; i < page; i++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAtIndex);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}
		}
		/***********************************************/

		// 写构成表数据
		int shnum = 0;
		for (int i = sheetAtIndex; i < index; i++) {
			int startrow = 23;// 合并起始行
			int endrow = 0;// 合并结束行
//			int Totalnum = 0;// 合并总数
			boolean flag = false;

			XSSFSheet sheet = book.getSheetAt(i);
			// 写partlist上部分信息
			if (assynos != null) {
				for (int k = 0; k < assynamelist.size(); k++) {
					String prename = "";
					String aftername = "";
					String[] assyVal = (String[]) assynamelist.get(k);
					String assyvalue = assyVal[0];
					System.out.println("测试个别数据出错的问题：" + assyvalue);
					if (assyvalue != null && assyvalue.length() > 5) {
						prename = assyvalue.substring(0, 5);
						aftername = assyvalue.substring(5).trim();
					} else {
						prename = assyvalue;
						aftername = "";
					}
					
					// 判断更新的数据是不是跟上次不同，把不同的数据标红
					if (oldassynos != null && updateflag) {
						String allassy = prename.trim() + aftername.trim();
						if (!oldassynos.contains(allassy)) {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 9 + k, 7, 2, -1);
							XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 9 + k, 13, 2, -1);
							setStringCellAndStyle2(sheet, prename, 9 + k, 7, newstyle, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
							setStringCellAndStyle2(sheet, aftername, 9 + k, 13, newstyle2, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 9 + k, 7, 12, -1);
							XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 9 + k, 13, 12, -1);
							setStringCellAndStyle2(sheet, prename, 9 + k, 7, newstyle, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
							setStringCellAndStyle2(sheet, aftername, 9 + k, 13, newstyle2, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
						}
					} else {
						XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 9 + k, 7, 12, -1);
						XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 9 + k, 13, 12, -1);
						setStringCellAndStyle2(sheet, prename, 9 + k, 7, newstyle, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番前缀
						setStringCellAndStyle2(sheet, aftername, 9 + k, 13, newstyle2, Cell.CELL_TYPE_STRING);// Ａｓｓｙ部番后缀
					}
					setStringCellAndStyle(sheet, assyVal[1], 9 + k, 19, style3, Cell.CELL_TYPE_STRING);// Ａｓｓｙ名称
				}
			}
			// 写partlist信息
			if (i == index - 1) {
				for (int j = 0; j + 24 * shnum < partlist.size(); j++) {
					String[] str = (String[]) partlist.get(j + 24 * shnum);
					// 判断是否为空行
					if (str[7] != null) {
						String prename = "";
						String aftername = "";
						System.out.println("构成表部品番号：" + str[1]);
						//②	零件号若为连续，不带空格者，不拆分前5位、后5位填写PartList  20201118
						if(str[1] != null && (str[1].contains(" ") || str[1].contains(" ")))
						{
							if (str[1] != null && str[1].length() > 5) {
								prename = str[1].substring(0, 5);
								aftername = str[1].substring(5).trim();
							} else {
								prename = str[1];
								aftername = "";
							}
						}
						else
						{
							prename = "";
							aftername = str[1];
						}
						// 判断更新的数据是不是跟上次不同，把不同的数据标红
						if (oldpartList != null && updateflag) {
							String allassyno = prename.trim() + aftername.trim();
							if (!oldpartList.contains(allassyno)) {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 2, -1);
								XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 2, -1);
								setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// 部品番号前缀
								setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// 部品番号后缀
							} else {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 12, -1);
								XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 12, -1);
								setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// 部品番号前缀
								setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// 部品番号后缀
							}
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 12, -1);
							XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 12, -1);
							setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// 部品番号前缀
							setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// 部品番号后缀

						}
						if (oldpartname != null && updateflag) {
							if (!oldpartname.contains(str[2])) {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 2, -1);
								setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// 部品名称
							} else {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 12, -1);
								setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// 部品名称
							}

						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 12, -1);
							setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// 部品名称
						}
						setStringCellAndStyle(sheet, str[7], 23 + j, 4, style2, Cell.CELL_TYPE_STRING);// 标号
						setStringCellAndStyle(sheet, str[0], 23 + j, 7, style, Cell.CELL_TYPE_STRING);// 安装顺序
						setStringCellAndStyle(sheet, str[3], 23 + j, 50, style, Cell.CELL_TYPE_STRING);// 数量
						setStringCellAndStyle(sheet, str[4], 23 + j, 53, style, Cell.CELL_TYPE_STRING);// 板厚
						setStringCellAndStyle(sheet, str[5], 23 + j, 58, style, Cell.CELL_TYPE_STRING);// 材质
						setStringCellAndStyle(sheet, str[6], 23 + j, 72, style, Cell.CELL_TYPE_STRING);// 部品来源

						flag = true;
						endrow = 23 + j;
//						if (str[3] != null && !str[3].isEmpty()) {
//							Totalnum = Totalnum + Integer.parseInt(str[3]);
//						} else {
//							Totalnum = Totalnum + 0;
//						}

					} else {
					}

				}
			} else {
				for (int j = 0; j + 24 * shnum < 24 + 24 * shnum; j++) {
					// 判断是否为空行
					String[] str = (String[]) partlist.get(j + 24 * shnum);
					// 判断是否为空行
					if (str[7] != null) {
						String prename = "";
						String aftername = "";
						//②	零件号若为连续，不带空格者，不拆分前5位、后5位填写PartList  20201118
						if(str[1] != null && (str[1].contains(" ") || str[1].contains(" ")))
						{
							if (str[1] != null && str[1].length() > 5) {
								prename = str[1].substring(0, 5);
								aftername = str[1].substring(5).trim();
							} else {
								prename = str[1];
								aftername = "";
							}
						}
						else
						{
							prename = "";
							aftername = str[1];
						}
						// 判断更新的数据是不是跟上次不同，把不同的数据标红
						if (oldpartList != null && updateflag) {
							String allassyno = prename.trim() + aftername.trim();
							if (!oldpartList.contains(allassyno)) {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 2, -1);
								XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 2, -1);
								setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// 部品番号前缀
								setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// 部品番号后缀
							} else {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 12, -1);
								XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 12, -1);
								setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// 部品番号前缀
								setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// 部品番号后缀
							}
						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 13, 12, -1);
							XSSFCellStyle newstyle2 = getXSSFStyle(book, sheet, 23 + j, 18, 12, -1);
							setStringCellAndStyle2(sheet, prename, 23 + j, 13, newstyle, Cell.CELL_TYPE_STRING);// 部品番号前缀
							setStringCellAndStyle2(sheet, aftername, 23 + j, 18, newstyle2, Cell.CELL_TYPE_STRING);// 部品番号后缀

						}
						if (oldpartname != null && updateflag) {
							if (!oldpartname.contains(str[2])) {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 2, -1);
								setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// 部品名称
							} else {
								XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 12, -1);
								setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// 部品名称
							}

						} else {
							XSSFCellStyle newstyle = getXSSFStyle(book, sheet, 23 + j, 24, 12, -1);
							setStringCellAndStyle2(sheet, str[2], 23 + j, 24, newstyle, Cell.CELL_TYPE_STRING);// 部品名称
						}
						setStringCellAndStyle(sheet, str[7], 23 + j, 4, style2, Cell.CELL_TYPE_STRING);// 标号
						setStringCellAndStyle(sheet, str[0], 23 + j, 7, style, Cell.CELL_TYPE_STRING);// 安装顺序
						setStringCellAndStyle(sheet, str[3], 23 + j, 50, style, Cell.CELL_TYPE_STRING);// 数量
						setStringCellAndStyle(sheet, str[4], 23 + j, 53, style, Cell.CELL_TYPE_STRING);// 板厚
						setStringCellAndStyle(sheet, str[5], 23 + j, 58, style, Cell.CELL_TYPE_STRING);// 材质
						setStringCellAndStyle(sheet, str[6], 23 + j, 72, style, Cell.CELL_TYPE_STRING);// 部品来源

						flag = true;
						endrow = 23 + j;
//						if (str[3] != null && !str[3].isEmpty()) {
//							Totalnum = Totalnum + Integer.parseInt(str[3]);
//						} else {
//							Totalnum = Totalnum + 0;
//						}

					} else {
					}
				}
			}
			shnum++;
		}

	}

	/*
	 * 获取部品信息
	 */
	private void getPartsinformation(TCComponentBOMLine gwbl) throws TCException {
		// TODO Auto-generated method stub
//		ArrayList install = new ArrayList();
//		ArrayList templist = new ArrayList();
//		// 先获取工位下的安装工序下的零件
//		install = Util.getChildrenByBOMLine(gwbl, "B8_BIWOperationRevision");
//
//		for (int i = 0; i < install.size(); i++) {
//			TCComponentBOMLine bl = (TCComponentBOMLine) install.get(i);
//			ArrayList bflist = new ArrayList();
//			bflist = Util.getChildrenByBOMLine(bl, "DFL9SolItmPartRevision");
//			for (int j = 0; j < bflist.size(); j++) {
//				String[] info = new String[8];
//				TCComponentBOMLine bfbl = (TCComponentBOMLine) bflist.get(j);
//				info[0] = Util.getProperty(bfbl, "bl_sequence_no");// 安装顺序
//				info[1] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9_part_no");// 部品番号
//				// info[2] = Util.getProperty(bfbl, "bl_rev_object_name");// 部品名称
//				info[2] = Util.getProperty(bfbl.getItemRevision(), "dfl9_CADObjectName");// 部品名称
//				info[3] = Util.getProperty(bfbl, "bl_quantity");// 数量
//				if (info[3] == null || info[3].isEmpty()) {
//					info[3] = "1";
//				}
//				String partresoles = "";
//				partresoles = Util.getProperty(bfbl, "B8_NoteManualMark");// 部品来源 待确认
//				if (partresoles == null || partresoles.isEmpty()) {
//					partresoles = Util.getProperty(bfbl, "B8_NoteIsBiwTrUnit");// 部品来源 待确认
//				}
//				if (partresoles.equals("外制件")) {
//					partresoles = "外制部品";
//				}
//				info[6] = partresoles;
//
//				if (partresoles.equals("冲压件")) {
//					String bh = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartThickness");
//					if (Util.isNumber(bh)) {
//						System.out.println("冲压件板厚" + bh);
//						info[4] = String.format("%.2f", Double.parseDouble(bh));// 板厚
//					} else {
//						info[4] = bh;// 板厚
//					}
//					info[5] = Util.getProperty(bfbl, "bl_DFL9SolItmPartRevision_dfl9PartMaterial");// 材质
//				} else {
//					info[4] = "";// 板厚
//					info[5] = "";// 材质
//				}
//				templist.add(info);
//			}
//		}
//		// 在获取工位的上一个工位的assy部番
//		TCProperty pp = gwbl.getTCProperty("Mfg0predecessors");
//		if (pp != null) {
//			TCComponent[] obj = pp.getReferenceValueArray();
//			for (int i = 0; i < obj.length; i++) {
//				TCComponentBOMLine prebl = (TCComponentBOMLine) obj[i];
//				String sequence_no = Util.getProperty(prebl, "bl_sequence_no");// 安装顺序
//				String quantity = Util.getProperty(prebl, "bl_quantity");// 数量
//				if (quantity == null || quantity.isEmpty()) {
//					quantity = "1";
//				}
//				// 获取部品信息 ,部品名称如果工位名称以#开头，则为产线名称+工位名称，否则就是工位名称
//				String linename = Util.getProperty(prebl.parent(), "bl_rev_object_name");
//				String staname = Util.getProperty(prebl, "bl_rev_object_name");
//				String assyname = "";
//				if (staname.length() > 1) {
//					if (staname.substring(0, 1).equals("#")) {
//						assyname = linename + " " + staname;
//					} else {
//						assyname = staname;
//					}
//				} else {
//					assyname = linename + " " + staname;
//				}
//
//				TCProperty p = prebl.getItemRevision().getTCProperty("b8_ProcAssyNo2");
//				String[] assynos;
//				if (p != null) {
//					assynos = p.getStringValueArray();// Ａｓｓｙ 部番
//				} else {
//					assynos = null;
//				}
//				if (assynos != null && assynos.length > 0) {
//					for (int j = 0; j < assynos.length; j++) {
//						String[] info = new String[8];
//						info[0] = sequence_no;// 安装顺序
//						info[1] = assynos[j];// 部品番号
//						info[2] = assyname;// 部品名称
//						info[3] = quantity;// 数量
//						info[4] = "";// 板厚
//						info[5] = "";// 材质
//						info[6] = "内制总成";// 部品来源 待确认
//						templist.add(info);
//					}
//				}
//			}
//		}
//		// 如果零件号相同，合并为一行，数量合计
//		Map<String, String[]> map = new HashMap<String, String[]>();
//		for (int i = 0; i < templist.size(); i++) {
//			String[] value = (String[]) templist.get(i);
//			String key = value[1];
//			if (!map.containsKey(key)) {
//				map.put(key, value);
//			} else {
//				String[] oldstr = map.get(key);
//				int quality = 0;
//				quality = Integer.parseInt(oldstr[3]) + Integer.parseInt(value[3]);
//				oldstr[3] = Integer.toString(quality);
//				map.put(key, oldstr);
//			}
//		}
//		ArrayList newtemplist = new ArrayList();
//		for (Map.Entry<String, String[]> entry : map.entrySet()) {
//			String[] values = entry.getValue();
//			newtemplist.add(values);
//		}
//
//		// 获取完后，对数据进行排序处理
//		Comparator comparator = getComParatorBysequenceno();
//		Collections.sort(newtemplist, comparator);
//
//		int label = 0; // 标号标记
//		int num = 1;// 标记同种标号的数据行数
//		String prePartno = "";// 部品番号前5位标记
//		String[] bh = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
//				"T", "U", "V", "W", "X", "Y", "Z" };
//		// 标号处理
//		Map<String, String> tempmap = new HashMap<String, String>();
//		ArrayList tempPartlist = new ArrayList();
//		for (int i = 0; i < newtemplist.size(); i++) {
//			String[] str = (String[]) newtemplist.get(i);
//			if (str[1].toString().length() > 5) {
//
//				prePartno = str[1].toString().substring(0, 5);
//			} else {
//				prePartno = str[1].toString();
//			}
//			String note = tempmap.get(prePartno);
//			// 部品番号前5位一样，则标号相同
//			if (note != null && !note.isEmpty()) {
//				str[7] = note;
//				String strnum = fymap.get(note);
//				int newnum = Integer.parseInt(strnum) + 1;
//				fymap.put(note, Integer.toString(newnum));
//			} else {
//				if (label < 26) {
//					str[7] = bh[label];
//				} else {
//					str[7] = "";
//					System.out.println("标号超过了给定的长度。。。");
//				}
//
//				fymap.put(bh[label], "1");
//				tempmap.put(prePartno, bh[label]);
//				label++;
//			}
//			tempPartlist.add(str);
//
//		}
//		// 根据标号排序
//		Comparator comparator2 = getComParatorBybh();
//		Collections.sort(tempPartlist, comparator2);
//
//		String firstNo = "";
//		for (int i = 0; i < tempPartlist.size(); i++) {
//			String[] value = (String[]) tempPartlist.get(i);
//			if (i == 0) {
//				firstNo = value[7];
//				partlist.add(value);
//			} else {
//				if (!firstNo.equals(value[7].toString())) {
//					String[] str = new String[8];
//					partlist.add(str);
//					partlist.add(value);
//					firstNo = value[7];
//				} else {
//					partlist.add(value);
//				}
//			}
//		}
//		System.out.println(partlist);

	}

	private Comparator getComParatorBybh() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				String[] comp1 = (String[]) obj;
				String[] comp2 = (String[]) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1[7] != null && !comp1[7].isEmpty()) {
					d1 = comp1[7].toString();
				}
				if (obj1 != null && comp2[7] != null && !comp2[7].isEmpty()) {
					d2 = comp2[7];
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}

	/*
	 * 对有效页处理
	 */
	private void ValidPageProcessing(XSSFWorkbook book) {
		// TODO Auto-generated method stub
		XSSFSheet sheet = null;
		int sheetnum = 0;
		sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("有效页")) {
				sheet = book.getSheetAt(i);
				break;
			}
		}
		if (sheet == null) {
			return;
		}
		// 一个工位作业表，sheet页不会超过120页，所以分数量处理
		// 设置字体颜色
		Font font = book.createFont();
		font.setColor((short) 12);// 蓝色字体
		font.setFontName("宋体");
		font.setFontHeightInPoints((short) 14);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		int page = (sheetnum - 1) / 40 + 1;

		for (int i = 0; i < page; i++) {
			if (i == page - 1) {
				for (int j = 0; j < sheetnum - 40 * i; j++) {
					setStringCellAndStyle2(sheet, "●", 7 + j, 11 + 35 * i, style, Cell.CELL_TYPE_STRING); // 编制
				}
			} else {
				for (int j = 0; j < 40; j++) {
					setStringCellAndStyle2(sheet, "●", 7 + j, 11 + 35 * i, style, Cell.CELL_TYPE_STRING); // 编制
				}
			}
		}
	}

	/*
	 * 对所有的sheet名称重命名
	 */
	private void SetSheetRename(XSSFWorkbook book) {
		// TODO Auto-generated method stub

		int sheetnum = book.getNumberOfSheets();
		// 定义比较名称,初始值为第一个sheet名称，如果名称相同，则需要在名称后面增加1,2......
		String tempname = "";
		String sheetAllname;
		int num = 1;
		Pattern p = Pattern.compile("[0-9a-fA-F]"); // 非数字字母
		Pattern p2 = Pattern.compile("[0-9]"); // 非数字字母
		// 先按照顺序命令，避免后续按照规则命名出现名称重复情况
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			String sheetname = sheet.getSheetName();
			sheetAllname = sheetname + (i + 1);
			book.setSheetName(i, sheetAllname);
		}

		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			// 截取前三位，去掉字母
			String sheetname = "";
			String oldsheetname = sheet.getSheetName();
			if (oldsheetname.length() > 2) {
				Matcher m = p.matcher(oldsheetname.substring(0, 3));
				Matcher m2 = p2.matcher(oldsheetname.substring(3));
				sheetname = m.replaceAll("").trim() + m2.replaceAll("").trim();
			} else {
				sheetname = oldsheetname;
			}
			// String sheetname = sheet.getSheetName().substring(2);
			// 第一个先不比较，从第二个开始和第一个比较
			if (i == 0) {
				tempname = sheetname;
				sheetAllname = String.format("%02d", i + 1) + sheetname;
				book.setSheetName(i, sheetAllname);

			} else {
				if (sheetname.contains(tempname)) {
					// 如果num为1，则说明sheet同名的第一个需要重命名增加数字后缀
					if (num == 1) {
						sheetAllname = String.format("%02d", i) + tempname + num;
						book.setSheetName(i - 1, sheetAllname);
					}
					sheetAllname = String.format("%02d", i + 1) + tempname + (num + 1);
					book.setSheetName(i, sheetAllname);
					num++;
				} else {
					sheetAllname = String.format("%02d", i + 1) + sheetname;
					tempname = sheetname;
					book.setSheetName(i, sheetAllname);

					num = 1;
				}
			}
			// 设置打印区域
			book.removePrintArea(i);
		}
		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			book.setPrintArea(i, 0, 114, 0, 51);
			PrintSetup printSetup = sheet.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 70);// 自定义缩放，此处100为无缩放
			printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)

		}
	}

	/*
	 * 写所有sheet页的公共信息
	 */
	private void writePublicDataToSheet(XSSFWorkbook book, ArrayList plist) {
		// TODO Auto-generated method stub
		// 设置字体
		Font font = book.createFont();
		font.setColor((short) 12);
		font.setFontName("宋体");
		font.setFontHeightInPoints((short) 16);
		// 创建一个样式
		XSSFCellStyle cellStyle1 = book.createCellStyle();
		cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle1.setFont(font);

		XSSFCellStyle cellStyle2 = book.createCellStyle();
		Font font2 = book.createFont();
		font2.setColor(IndexedColors.BLUE.getIndex());
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
		font2.setFontHeightInPoints((short) 16);
		font2.setFontName("宋体");
		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		cellStyle2.setFont(font2);

		XSSFCellStyle cellStyle3 = book.createCellStyle();
		Font font3 = book.createFont();
		font3.setColor(IndexedColors.BLUE.getIndex());
		// font3.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
		font3.setItalic(true); // 字体为斜体
		font3.setFontHeightInPoints((short) 72);
		font3.setFontName("宋体");
		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle3.setFont(font3);

		// 循环所有sheet页，把公共部分内容写入
		int sheetnum = book.getNumberOfSheets();
		for (int n = 0; n < sheetnum; n++) {
			XSSFSheet sh = book.getSheetAt(n);
			String sheetname = sh.getSheetName();
			if (sheetname.contains("首页")) {

				/**************************************************/
				// 如果是更新，则先把系统输出的信息清空，后面再写入
//				if (updateflag) {
//					List<String> delPicturesList = ReportUtils.removePictrues07((XSSFSheet) sh, (XSSFWorkbook) book, 3,
//							48, 100, 115);
//					System.out.println("-----------符合条件的图片有-----------");
//					for (String name : delPicturesList) {
//						System.out.println(name);
//					}
//				}
				/**************************************************/
				if (!updateflag && !model.equals("调整线模板")) { // 调整线的工位名称和编号不需要写入
					setStringCellAndStyle(sh, plist.get(3).toString(), 22, 5, cellStyle3, Cell.CELL_TYPE_STRING); // 首页中间的工位名称
				}
				if (Import != null && Import.size() > 0) {
					for (int m = 0; m < Import.size(); m++) {
						InputStream is = null;
						if (Import.get(m).toString().trim().equals("A")) {
							is = this.getClass().getResourceAsStream("/com/dfl/report/imags/A.png");
						}
						if (Import.get(m).toString().trim().equals("B")) {
							is = this.getClass().getResourceAsStream("/com/dfl/report/imags/B.png");
						}
						if (is != null) {
							if (!updateflag) {
								writepicturetosheet(book, sh, is, 105, 5 + m * 5, 111, 9 + m * 5);
							}
						}
					}
				}
			}
			// 发行科
			setStringCellAndStyle(sh, plist.get(6).toString(), 2, 0, cellStyle1, Cell.CELL_TYPE_STRING); // 编制
			// 如果是更新 编制”、“日期”、“版次”保持不变
			if (!updateflag) {
				setStringCellAndStyle(sh, plist.get(0).toString(), 2, 6, cellStyle1, Cell.CELL_TYPE_STRING); // 编制
				setStringCellAndStyle(sh, plist.get(1).toString(), 2, 30, cellStyle1, Cell.CELL_TYPE_STRING);// 日期
				setStringCellAndStyle(sh, plist.get(5).toString(), 48, 108, cellStyle2, Cell.CELL_TYPE_STRING);// 批次
			}
			setStringCellAndStyle(sh, plist.get(2).toString(), 2, 90, cellStyle1, Cell.CELL_TYPE_STRING);// 车型
			if (!updateflag && !model.equals("调整线模板")) { // 调整线的工位名称和编号不需要写入
				setStringCellAndStyle(sh, plist.get(3).toString(), 50, 72, cellStyle2, Cell.CELL_TYPE_STRING);// 工位名称
				setStringCellAndStyle(sh, plist.get(4).toString(), 51, 94, cellStyle2, Cell.CELL_TYPE_STRING);// 工位编码
			}
			setStringCellAndStyle(sh, Integer.toString(n + 1), 50, 107, cellStyle2, 10);// 当前页码
			setStringCellAndStyle(sh, Integer.toString(sheetnum), 50, 112, cellStyle2, 10);// 总页码

			// 工位名称 自适应大小
//			XSSFRow row = sh.getRow(50);
//			if (row != null) {
//				XSSFCell cell = row.getCell(72);
//				if (cell != null) {
//					NewOutputDataToExcel.setFontSize(book, cell, (short) 16);
//				}
//			}
		}
	}

	/*
	 * 写所有sheet页的公共信息
	 */
	private void writeRepatPublicDataToSheet(XSSFWorkbook book) {
		// TODO Auto-generated method stub
		// 设置字体
		Font font = book.createFont();
		font.setColor((short) 12);
		font.setFontName("宋体");
		font.setFontHeightInPoints((short) 16);
		// 创建一个样式
		XSSFCellStyle cellStyle1 = book.createCellStyle();
		cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle1.setFont(font);

		XSSFCellStyle cellStyle2 = book.createCellStyle();
		Font font2 = book.createFont();
		font2.setColor(IndexedColors.BLUE.getIndex());
		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
		font2.setFontHeightInPoints((short) 16);
		font2.setFontName("宋体");
		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		cellStyle2.setFont(font2);

		// 循环所有sheet页，把公共部分内容写入
		int sheetnum = book.getNumberOfSheets();
		for (int n = 0; n < sheetnum; n++) {
			XSSFSheet sh = book.getSheetAt(n);
			String sheetname = sh.getSheetName();
			setStringCellAndStyle(sh, Integer.toString(n + 1), 50, 107, cellStyle2, 10);// 当前页码
			setStringCellAndStyle(sh, Integer.toString(sheetnum), 50, 112, cellStyle2, 10);// 总页码

		}
	}

	/*
	 * 根据业务选择的sheet页，加载初始模板
	 */
	private XSSFWorkbook creatEngineeringXSSFWorkbook(InputStream inputStream, ArrayList list,
			LinkedHashMap<String, String> map) {
		// TODO Auto-generated method stub
		XSSFWorkbook book = null;
		try {
			book = new XSSFWorkbook(inputStream);

			// 循环所有sheet，如果用户未勾选，则移除
			int sheetnum = book.getNumberOfSheets();
			deletelist = new ArrayList();
			ArrayList copylist = new ArrayList();
			Map<String, Integer> pxmap = new LinkedHashMap<String, Integer>();// 用于sheet排序
			for (int i = 0; i < sheetnum; i++) {
				String allshname = book.getSheetName(i);
				String sheetname = allshname.substring(2);
				if (!list.contains(sheetname)) {
					deletelist.add(book.getSheetName(i));
				} else {
					int shat = 0;
					for (int j = 0; j < list.size(); j++) {
						String strval = (String) list.get(j);
						if (strval.equals(sheetname)) {
							shat = j;
							break;
						}
					}
					pxmap.put(allshname, shat);
					// 再根据用户输入的sheet页数，增加sheet页
					int sheetNum = Integer.parseInt(map.get(sheetname));
					// 如果页数大于1，则需要增加sheet页
					if (sheetNum > 1) {
						copylist.add(book.getSheetName(i));
					}
				}
			}
			for (Map.Entry<String, Integer> entry : pxmap.entrySet()) {
				String key = entry.getKey();
				Integer value = entry.getValue();
				book.setSheetOrder(key, value);
			}

			// 复制多个相同的sheet
			for (int k = 0; k < copylist.size(); k++) {
				String sheetAllname = (String) copylist.get(k);
				int sheetNums = Integer.parseInt(map.get(sheetAllname.substring(2)));
				int sheetAt = book.getSheetIndex(sheetAllname);
				int index = sheetAt + 1;
				for (int n = 1; n < sheetNums; n++) {
					XSSFSheet newsheet = book.cloneSheet(sheetAt);
					book.setSheetOrder(newsheet.getSheetName(), index);
					index++;
				}
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	/*
	 * 删除未选的sheet页
	 */
	private void deleteSheets(XSSFWorkbook book) {
		if (deletelist != null && deletelist.size() > 0) {
			for (int j = 0; j < deletelist.size(); j++) {
//				System.out.println("sheet名称：" + deletelist.get(j).toString() + " " + book.getSheetIndex(deletelist.get(j).toString()));
				book.removeSheetAt(book.getSheetIndex(deletelist.get(j).toString()));
			}
			List delist = new ArrayList();
			for (int i = 0; i < book.getNumberOfSheets(); i++) {
				String sheetname = book.getSheetName(i);
				boolean flag = getIsContain(sheetname);
				if (flag) {
					delist.add(sheetname);
				}
			}
			if (delist != null && delist.size() > 0) {
				for (int k = 0; k < delist.size(); k++) {
					book.removeSheetAt(book.getSheetIndex(delist.get(k).toString()));
				}
			}
		}
	}

	private boolean getIsContain(String name) {

		if (deletelist != null && deletelist.size() > 0) {
			for (int i = 0; i < deletelist.size(); i++) {
				String shname = (String) deletelist.get(i);
				if (name.contains(shname)) {
					return true;
				}
			}
		}
		return false;
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
        if(Style!=null)
        {
        	cell.setCellStyle(Style);
        }		

	}

	private Comparator getComParatorBysequenceno() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				Object[] comp1 = (Object[]) obj;
				Object[] comp2 = (Object[]) obj1;

				Double d1 = 0.0;
				Double d2 = 0.0;
				if (comp1[0] != null && !comp1[0].toString().isEmpty()) {
					d1 = Double.parseDouble(comp1[0].toString());
				}
				if (comp2[0] != null && !comp2[0].toString().isEmpty()) {
					d2 = Double.parseDouble(comp2[0].toString());
				}
				if (d2 > d1) {
					return -1;
				}
				if (d2 == d1) {
					return 0;
				}

				return 1;
			}
		};

		return comparator;
	}

	/*
	 * 根据板层数从大到小排序
	 */
	private Comparator getComParatorBylayersnum() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				WeldPointBoardInformation comp1 = (WeldPointBoardInformation) obj;
				WeldPointBoardInformation comp2 = (WeldPointBoardInformation) obj1;
//				if (comp1.equals(comp2)) {
//					//System.out.println("*************************");
//					return 0;
//				}

				int d1 = 0;
				int d2 = 0;
				if (obj != null && comp1.getLayersnum() != null && !comp1.getLayersnum().isEmpty()) {
					d1 = Integer.parseInt(comp1.getLayersnum());
				}
				if (obj1 != null && comp2.getLayersnum() != null && !comp2.getLayersnum().isEmpty()) {
					d2 = Integer.parseInt(comp2.getLayersnum());
				}
				if (d2 < d1) {
					return -1;
				}
				if (d2 == d1) {
					return 0;
				}

				return 1;
			}
		};

		return comparator;
	}

	/**
	 * 获取区域 Region
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public static int getMergedRegionIndex(XSSFSheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return i;
				}
			}
		}

		return -1;
	}

	/*
	 * 取最小值
	 */
	private String getMinnum(String str1, String str2, String str3) {
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
		if (minstr.equals("9999")) {
			minstr = "";
		}
		return minstr;
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

	// 根据单个文件写图片到excel
	private static void writepicturetosheet(XSSFWorkbook book, XSSFSheet sheet, InputStream is, int colindex,
			int rowindex, int endcolindex, int endrowindex) {
		// 先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		BufferedImage bufferImg;
		try {
			bufferImg = ImageIO.read(is);
			ImageIO.write(bufferImg, "png", byteArrayOut);
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) colindex, rowindex,
					(short) endcolindex, endrowindex);
			anchor.setAnchorType(2);
			int picindex = book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG);
			// 插入图片
			patriarch.createPicture(anchor, picindex);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/*
	 * 根据工位获取3D图片
	 */
	private Map<String, File> getAll3DPictures(TCComponentItemRevision blrev) throws TCException {
		Map<String, File> piclist = new HashMap<String, File>();
		TCComponent[] tccdata = blrev.getRelatedComponents("IMAN_3D_snap_shot");
		for (TCComponent tcc : tccdata) {
			String objectname = Util.getProperty(tcc, "object_name");
			// 部品构成图 名称以数字开头
			if (Util.isNumber(objectname.substring(0, 1))) {
				File file = downLoadPicture1(tcc, "ThumbnailImage");
				if (file != null) {
					piclist.put(objectname, file);
				}
			}
		}
		return piclist;
	}

	/**
	 * 下载图片数据集到本地
	 * 
	 * @param picDs1
	 * @return
	 */
	public static File downLoadPicture1(TCComponent comp, String pictype) {
		// TODO Auto-generated method stub

		// System.out.println(">>>downLoadPicture");

		TCComponentDataset dataset = null;
		if (comp instanceof TCComponentDataset) {
			dataset = (TCComponentDataset) comp;
		}
		File file = null;
		if (dataset == null) {
			// System.out.println("dataset==null");
			return null;
		}

		System.out.println("downLoadPicture:" + dataset.toString());
		String type = dataset.getType();
		// "Image","JPEG","Bitmap","TIF","GIF"
		if (!"Vis_Snapshot_2D_View_Data".equals(type) && !"SnapShotViewData".equals(type) && !"Image".equals(type)
				&& !"JPEG".equals(type) && !"Bitmap".equals(type) && !"TIF".equals(type) && !"GIF".equals(type)) {
			// System.out.println("图片类型不匹配："+type);
			return null;
		}

		TCComponentTcFile[] files;
		try {

			files = dataset.getTcFiles();
			TCComponent pic = dataset.getNamedRefComponent(pictype);
			String modelname = pic.getProperty("file_name");
			if (files == null || files.length <= 0) {
				return null;
			}
			for (int i = 0; i < files.length; i++) {
				String fileName = files[i].getProperty("file_name");
				System.out.println("fileName:" + fileName);
				if (modelname.equals(fileName)) {
					if (fileName.toLowerCase().endsWith("png") || fileName.toLowerCase().endsWith("jpeg")
							|| fileName.toLowerCase().endsWith("jpg") || fileName.toLowerCase().endsWith("bmp")
							|| fileName.toLowerCase().endsWith("tif") || fileName.toLowerCase().endsWith("gif")) {
						file = files[i].getFmsFile();
						// System.out.println("fms file:"+file.getAbsolutePath());
						return file;
					}
				}

			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return file;
	}

	// 根据单个文件写图片到excel
	private static void writepicturetosheet(XSSFWorkbook book, XSSFSheet sheet, BufferedImage bufferImg, int rowindex,
			int colindex, int rowindex2, int colindex2) {
		// 先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		try {
			ImageIO.write(bufferImg, "png", byteArrayOut);
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) colindex, rowindex,
					(short) colindex2, rowindex2);
			anchor.setAnchorType(2);
			int picindex = book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG);
			// 插入图片
			patriarch.createPicture(anchor, picindex);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	// 查询部品类型首选项，获取部品类型信息
	private Map<String, String> getSizeRule() {
		Map<String, String> rule = new HashMap<String, String>();
		try {
			AbstractAIFUIApplication app = AIFUtility.getCurrentApplication();
			TCSession session = (TCSession) app.getSession();
			TCPreferenceService preferenceService = session.getPreferenceService();

			String str = preferenceService.getPreferenceDescription("DFL9_get_parts_source");
			if (str != null) {
				String[] values = preferenceService.getStringValues("DFL9_get_parts_source");
				for (int i = 0; i < values.length; i++) {
					String value = values[i];
					if (value != null) {
						String[] val = value.split("=");
						if (val != null && val.length > 1) {
							rule.put(val[0], val[1]);
						}
					}
				}
			}
			return rule;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return rule;
	}

	// 设置点焊页根据焊点号自动获取板组编号
	private void setCellFormula(XSSFWorkbook book) {
		List shnamelist = new ArrayList();
		int sheetnum = 0;
		List sheetList = new ArrayList();
		sheetnum = book.getNumberOfSheets();
		for (int i = 0; i < sheetnum; i++) {
			String sheetname = book.getSheetName(i);
			if (sheetname.contains("点焊")) {
				shnamelist.add(sheetname);
			}
			if ((sheetname.contains("PSW") && !sheetname.contains("点焊")) || sheetname.contains("RSW伺服")
					|| sheetname.contains("RSW气动")) {
				sheetList.add(sheetname);
			}
		}
		if (shnamelist.size() < 1) {
			return;
		}
//		FormulaEvaluator evl = null;
//		evl = new XSSFFormulaEvaluator(book);
		// 循环点焊sheet页，设置公式
		for (int i = 0; i < shnamelist.size(); i++) {
			String shname = (String) shnamelist.get(i);
			XSSFSheet sheet = book.getSheet(shname);
			for (int j = 0; j < 9; j++) {
				XSSFRow row = sheet.getRow(17 + 2 * j);
				if (row == null) {
					row = sheet.createRow(17 + 2 * j);
				}
				XSSFCell cell;
				// CZ、DD、DH、FZ列
				String formula4 = "IF(ISBLANK(DM" + (18 + 2 * j) + "),\"\",DM" + (18 + 2 * j) + ")";// FZ
				cell = row.getCell(181);
				cell.setCellFormula(formula4);
				// evl.evaluateFormulaCell(cell);
				// GA、GB、GC列
				String colname = "DM" + (18 + 2 * j);
				String formula5 = getExcelFormula(sheetList, colname, 6);
				if (!formula5.isEmpty()) {
					cell = row.getCell(182);
					cell.setCellFormula(formula5);
					// evl.evaluateFormulaCell(cell);
				}
				String formula6 = getExcelFormula(sheetList, colname, 32);
				if (!formula6.isEmpty()) {
					cell = row.getCell(183);
					cell.setCellFormula(formula6);
					// evl.evaluateFormulaCell(cell);
				}
				String formula7 = getExcelFormula(sheetList, colname, 58);
				if (!formula7.isEmpty()) {
					cell = row.getCell(184);
					cell.setCellFormula(formula7);
					// evl.evaluateFormulaCell(cell);
				}
				String formula1 = "IF(ISBLANK(DM" + (18 + 2 * j) + "),\"\",GA" + (18 + 2 * j) + ")";// CZ
				cell = row.getCell(103);
				cell.setCellFormula(formula1);
				// evl.evaluateFormulaCell(cell);
				String formula2 = "IF(ISBLANK(DM" + (18 + 2 * j) + "),\"\",GB" + (18 + 2 * j) + ")";// DD
				cell = row.getCell(107);
				cell.setCellFormula(formula2);
				// evl.evaluateFormulaCell(cell);
				String formula3 = "IF(ISBLANK(DM" + (18 + 2 * j) + "),\"\",GC" + (18 + 2 * j) + ")";// DH
				cell = row.getCell(111);
				cell.setCellFormula(formula3);
				// evl.evaluateFormulaCell(cell);

			}
			sheet.setForceFormulaRecalculation(true);
		}
		book.setForceFormulaRecalculation(true);

	}

	/*
	 * 根据PSW、RSWsheet数量
	 * 
	 * @sheetname sheet名称数组
	 * 
	 * @colname 依据列明
	 * 
	 * @colnum 所取列索引
	 */
	private String getExcelFormula(List sheetname, String colname, int colnum) {
		String formula = "";
		String commonstr = "";
		if (sheetname != null && sheetname.size() > 0) {
			for (int i = 0; i < sheetname.size(); i++) {
				String name = (String) sheetname.get(i);
				if (commonstr.isEmpty()) {
					commonstr = "VLOOKUP(" + colname + ",'" + name + "'!$I$12:$BQ$47," + colnum + ",0)";
				} else {
					commonstr = "IFERROR(" + commonstr + ",VLOOKUP(" + colname + ",'" + name + "'!$I$12:$BQ$47,"
							+ colnum + ",0))";
				}
			}
			if (!commonstr.isEmpty()) {
				formula = "IFERROR(" + commonstr + ",\"\")";
				formula = "IF(" + formula + "=0,\"\"," + formula + ")";
			}
		}
		System.out.println("公式：" + formula);
		return formula;
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
				XSSFCellStyle newstyle = book.createCellStyle();
				newstyle = (XSSFCellStyle) style.clone();
				if(bgcolor > -1)
				{
					newstyle.setFillForegroundColor((short)bgcolor);
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
	 * 判断版次是否为SOP后
	 */
	private boolean getIsSOPAfter() {
		boolean flag = false;
		ArrayList edition = getEditionSizeRule();
		if (edition != null && edition.size() > 0) {
			if (edition.contains(Edition)) {
				return false;
			}
		}
		if (Edition != null) {
			if (Edition.length() == 1) {
				char c = Edition.charAt(0);
				if (c >= 'A' && c <= 'Z') {
					flag = true;
				}
			}
			if (Edition.length() == 2) {
				char c = Edition.charAt(0);
				char cc = Edition.charAt(1);
				if (c >= 'A' && c <= 'Z' && cc >= 'A' && cc <= 'Z') {
					flag = true;
				}
			}
		}
		return flag;
	}

	/*
	 * 判断版次是否为SOP后
	 */
	private boolean getIsTeSOPAfter(String bc) {
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
}
