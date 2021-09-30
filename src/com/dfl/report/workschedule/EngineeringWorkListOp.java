package com.dfl.report.workschedule;

import java.awt.Container;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.BoardInformation;
import com.dfl.report.ExcelReader.CoverInfomation;
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
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentFolderType;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentRole;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;

public class EngineeringWorkListOp {

	private AbstractAIFUIApplication app;
	SimpleDateFormat df = new SimpleDateFormat("yyyy年MM月");// 设置日期格式
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// 设置日期格式
	private String Edition;
	private String topfoldername;
	private Map<String, String> projVehMap;
	private String VehicleNo;
	private TCSession session;
	private List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	private List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private GenerateReportInfo info;
	private InputStream inputStream = null;

	public EngineeringWorkListOp(AbstractAIFUIApplication app, Object object, String edition, String topfoldername,
			GenerateReportInfo info, InputStream inputStream) throws TCException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.Edition = edition;
		this.topfoldername = topfoldername;
		session = (TCSession) app.getSession();
		this.info = info;
		this.inputStream = inputStream;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub
		InterfaceAIFComponent ift = app.getTargetComponent();
		TCComponentBOMLine topbl = (TCComponentBOMLine) ift;
		TCSession session = (TCSession) app.getSession();
		TCComponentUser user = session.getUser();

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
		// 文件名称
		String procName = "01.目录";

		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("开始输出报表...\n", 10, 100);

		viewPanel.addInfomation("", 20, 100);

		viewPanel.addInfomation("", 40, 100);

		// 新增目录
		String message2 = addNewReportContents(topbl, inputStream, info, procName, viewPanel);

		if (!message2.isEmpty()) {
			viewPanel.addInfomation(message2, 100, 100);
			return;
		}
		viewPanel.addInfomation("输出报表完成，请在焊装工厂工艺对象附件下查看\n", 100, 100);
	}

	// 新增目录
	private String addNewReportContents(TCComponentBOMLine topbl, InputStream inputStream, GenerateReportInfo info,
			String procName, ReportViwePanel viewPanel) throws TCException {

		String error = "";

		// 根据顶层BOP获取所有的焊装工位
		ArrayList dhlist = new ArrayList();
		getDiscretes(topbl, dhlist);

		ArrayList plist = new ArrayList();// 获取的公共的输出数据集合
		ArrayList list = new ArrayList();// 处理后的目录输出数据集合
		ArrayList templist = new ArrayList();// 获取的目录输出数据集合

		List<BoardInformation> bzlist = new ArrayList<BoardInformation>();// 获取的板组数据集合
		// 编制人
		// String username = app.getSession().getUserName();
		TCComponentUser user = session.getUser();
		String username = user.getUserName();
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
		plist.add(username);
		// 日期
		plist.add(df2.format(new Date()));
		// 车型
		// 改为根据BOP名称取工厂产线信息
		String objectname = Util.getProperty(topbl.getItemRevision(), "object_name");
		String factoryline = ReportUtils.getFactoryLineByBOP(objectname);
		String factory = "";
		String linebody = "";
		if (factoryline.length() > 3) {
			factory = factoryline.substring(0, 3);
			linebody = factoryline.substring(factoryline.length() - 1);
		}
		String car_type = VehicleNo + "-" + factory + "-NO" + linebody;

		plist.add(car_type);
		plist.add(Edition);
		plist.add(department);

		viewPanel.addInfomation("", 50, 100);

		ArrayList gxbh = new ArrayList();
		// 获取工位信息
		for (int i = 0; i < dhlist.size(); i++) {
			String[] str = new String[7];
			TCComponentBOMLine bomline = (TCComponentBOMLine) dhlist.get(i);
			TCComponentItemRevision blrev = bomline.getItemRevision();

			String stationcode = Util.getProperty(blrev, "b8_OPNo");
//			String[] assynos;
//			TCProperty p = blrev.getTCProperty("b8_AssyNo");
//			if (p != null) {
//				assynos = blrev.getTCProperty("b8_AssyNo").getStringValueArray(); // Ａｓｓｙ 部番
//			} else {
//				assynos = null;
//			}
//			if (assynos != null && assynos.length > 0) {
//				if (assynos[0].length() > 5) {
//					stationcode = "M" + assynos[0].trim().substring(0, 5);// 工位编号
//				} else {
//					stationcode = "M" + assynos[0].trim();
//				}
//
//			} else {
//				stationcode = "";// 工位编号
//			}

			str[0] = stationcode;// 工序编号
			str[6] = Util.getProperty(bomline, "bl_rev_object_name");// 工位名称
			// str[1] = Util.getProperty(bomline.parent().getItemRevision(),
			// "b8_ChineseName");// 工序中文名称
			String chinesename = Util.getProperty(blrev, "b8_STName");
			str[1] = chinesename;
			String englishname = Util.getProperty(bomline.parent(), "bl_rev_object_name");// 工序英文名称
			if (chinesename.contains(str[6])) {
				englishname = englishname + " " + str[6] + " ";
			}
			if (chinesename.contains("右╱左")) {
				englishname = englishname.replace("RH", "").replace("LH", "") + "RH/LH";
			}
			if (chinesename.contains("左╱右")) {
				englishname = englishname.replace("RH", "").replace("LH", "") + "LH/RH";
			}
			str[2] = englishname;
			str[3] = Edition;// 版次
			str[4] = Util.getProperty(blrev, "b8_OpSheetNumber");// 页数b8_OpSheetNumber
			// 改为读取生成的工程作业表计算sheet数
//			String baseprocName = str[1].replace("#", "");
//			String baseprocName2 = str[1].replace("#", "");
//			InputStream inputS = null;
//			inputS = baseinfoExcelReader.getFileinbyreadExcel(blrev, "IMAN_specification", baseprocName);
//			if (inputS == null) {
//				inputS = baseinfoExcelReader.getFileinbyreadExcel(blrev, "IMAN_specification", baseprocName2);
//			}
//			if (inputS != null) {
//				XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);
//				int sheetnum = book.getNumberOfSheets();
//				str[4] = Integer.toString(sheetnum);
//			} else {
//				str[4] = "0";
//			}
			str[5] = "";// 备注

			// 如果页数为空，说明该工位未输出报表，不输出
			if (str[4] != null && !str[4].isEmpty()) {
				if (!gxbh.contains(str[0])) {
					gxbh.add(str[0]);
				}
				templist.add(str);
			}
		}
		// 处理工位信息，如果工序编号一样，需要合为一条显示,页码合计
		for (int i = 0; i < gxbh.size(); i++) {
			String ProcessNumber = (String) gxbh.get(i);
			String[] value = new String[7];
			String pEnglishName = "";
			String pChineseName = "";
			int page = 0;
			String Edition = "";
			// 用于标记
			int num = 0;
			String tempName = "";
			String lastname = "";
			for (int j = 0; j < templist.size(); j++) {
				String[] str = (String[]) templist.get(j);
				if (str[0].equals(ProcessNumber)) {
					num++;
					if (num == 1) {
						pEnglishName = str[2];
						pChineseName = str[1];
						tempName = str[6];
					} else if (num == 2) {
						// tempName = tempName + "," + str[6];
						lastname = str[6];
					} else {
						// tempName = tempName + "~" + str[6];
						lastname = str[6];
					}
					if (str[4] != null && !str[4].isEmpty()) {
						page = page + Integer.parseInt(str[4]);
					}
					Edition = str[3];
				}
			}
			if (num == 1) {

			} else if (num == 2) {
				pChineseName = pChineseName.replace(tempName, "") + "(" + tempName + "," + lastname + ")";
				pEnglishName = pEnglishName.replace(tempName, "") + "(" + tempName + "," + lastname + ")";
			} else {
				pChineseName = pChineseName.replace(tempName, "") + "(" + tempName + "-" + lastname + ")";
				pEnglishName = pEnglishName.replace(tempName, "") + "(" + tempName + "-" + lastname + ")";
			}

			value[0] = Integer.toString(i + 1); // 序号
			value[1] = ProcessNumber;
			value[2] = pChineseName;
			value[3] = pEnglishName;
			value[4] = Edition;
			value[5] = Integer.toString(page);
			value[6] = "";

			list.add(value);
		}
		viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);
		// 获取板组信息
		String basename = "222.基本信息";
		bzlist = getPartData(topbl, basename);

		String filename = procName;

		// 开启旁路
		{
			Util.callByPass(session, true);
		}

		XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream, list, bzlist);
		NewOutputDataToExcel.writeDataToSheet(book, plist, list, bzlist);
		for (int i = 0; i < book.getNumberOfSheets(); i++) 
		{
			XSSFSheet sheet = book.getSheetAt(i);
			if(sheet.getSheetName().contains("附录-24参数序列") || sheet.getSheetName().contains("附录-序列对照表"))
			{
				book.setPrintArea(i, 0, 114, 0, 51);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 65);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
			if(sheet.getSheetName().contains("附录-255参数序列"))
			{
				book.setPrintArea(i, 0, 114, 0, 51);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);				
				printSetup.setScale((short) 71);// 自定义缩放，此处100为无缩放
				printSetup.setFitHeight((short) 1);
				printSetup.setFitWidth((short) 1);
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
			}
		}
	
		NewOutputDataToExcel.exportFile(book, filename);
		String fullFileName = FileUtil.getReportFileName(filename);
		viewPanel.addInfomation("", 80, 100);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
		datasetList.add(ds);
		revlist.add(topbl.getItemRevision());
		try {
			TCComponentItem docunment = ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "",
					session);
			saveFileToFolder(docunment, topfoldername);

			// 文件编号和虚层名称
			TCProperty pdoc = docunment.getTCProperty("dfl9_vehiclePlant");
			if (pdoc != null) {
				if (pdoc != null) {
					pdoc.setStringValue(topfoldername);
					docunment.lock();
					docunment.save();
					docunment.unlock();
				}
			}

		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return error;
		}
		// 关闭旁路
		{
			Util.callByPass(session, false);
		}
		return error;
	}

	// 获取板组信息
	private List<BoardInformation> getPartData(TCComponentBOMLine topbl, String procName) {
		// TODO Auto-generated method stub
		List<BoardInformation> baseinfolist = new ArrayList<BoardInformation>();

		InputStream filein = null;
		try {
			filein = baseinfoExcelReader.getFileinbyreadExcel2(topbl.getItemRevision(), "IMAN_reference", procName);
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		baseinfolist = baseinfoExcelReader.readBZExcel(filein, "xlsx");

		return baseinfolist;
	}

	// 获取焊装工位
	private ArrayList getDiscretes(TCComponentBOMLine topbl, ArrayList dhlist) {
		// TODO Auto-generated method stub
		try {
			AIFComponentContext[] children = topbl.getChildren();
			for (AIFComponentContext child : children) {
				TCComponentBOMLine bl = (TCComponentBOMLine) child.getComponent();
				TCComponentItemRevision rev = bl.getItemRevision();
				if (rev.isTypeOf("B8_BIWMEProcStatRevision")) {
					dhlist.add(bl);
					continue;
				} else {
					getDiscretes(bl, dhlist);
				}
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return dhlist;
	}

	/*
	 * 根据文件编号创建顶层文件夹在home下,并创建00.封面文件夹，把封面文档放到00.封面文件夹下
	 */
	private void saveFileToFolder(TCComponentItem document, String topfoldername) {
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
				if (tcc.getType().equals("Folder") && obejctname.equals("01.目录及附录")) {
					childrenfolder = (TCComponentFolder) tcc;
					break;
				}
			}
			if (childrenfolder == null) {
				childrenfolder = foldertype.create("01.目录及附录", "", "Folder");
				folder.add("contents", childrenfolder);
				childrenfolder.add("contents", document);
			} else {
				// folder.add("contents", childrenfolder);
				AIFComponentContext[] icf3 = childrenfolder.getChildren();
				// 先移除
				if (icf3 != null) {
					for (AIFComponentContext aif : icf3) {
						TCComponent tcc = (TCComponent) aif.getComponent();
						childrenfolder.remove("contents", tcc);
					}
				}
				childrenfolder.add("contents", document);
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
