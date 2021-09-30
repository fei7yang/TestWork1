package com.dfl.report.workschedule;

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
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.common.EclipseUtils;
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

public class EngineeringWorkListCoverOp {

	private AbstractAIFUIApplication app;
	SimpleDateFormat df = new SimpleDateFormat("yyyy年MM月");// 设置日期格式
	SimpleDateFormat df2 = new SimpleDateFormat("yyyy.MM");// 设置日期格式
	private Map<String, String> projVehMap;
	private String VehicleNo = "";// 车型代号
	private String Edition = "";// 版次
	List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private TCSession session;
	private InputStream inputStream = null;
	private GenerateReportInfo info;

	public EngineeringWorkListCoverOp(AbstractAIFUIApplication app, Object object, String edition2, GenerateReportInfo info, InputStream inputStream) throws TCException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.Edition = edition2;
		session = (TCSession) app.getSession();
		this.inputStream = inputStream;
		this.info = info ;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub
		InterfaceAIFComponent ift = app.getTargetComponent();
		TCComponentBOMLine topbl = (TCComponentBOMLine) ift;
		TCComponentItemRevision boprev = null;
		try {
			boprev = topbl.getItemRevision();
		} catch (TCException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}
		// 读取 项目-车型 首选项
		projVehMap = ReportUtils.getDFL_Project_VehicleNo();
		String FamlilyCode = "";
		FamlilyCode = Util.getProperty(boprev, "project_ids");
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
		String procName = "00.封面";


		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("开始输出报表...\n", 10, 100);

		viewPanel.addInfomation("", 20, 100);
		// 判断报表模板是否维护
		// 查询封面导出模板
		XSSFWorkbook book = null;

		viewPanel.addInfomation("", 40, 100);
		// 先生成封面报表文件
		String error = "";

		// 根据BOP顶层名称取工厂线体信息
		String objectname = Util.getProperty(boprev, "object_name");

		String factoryline = ReportUtils.getFactoryLineByBOP(objectname);
		String factory = "";
		if(factoryline!=null && factoryline.length()>2) {
			factory = factoryline.substring(0, 3);
		}

		String[] cover = new String[5];
		cover[0] = "          车    型：" + VehicleNo;
		cover[1] = "          版    次：" + Edition;
		cover[2] = "          文件编号：" + VehicleNo + "-" + factoryline + "-AB";
		cover[3] = "          编制日期：" + df.format(new Date());
		cover[4] = "          工厂工程：" + factory + "工厂焊装工程";
		
		book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

		String filename = procName;

		viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);

		// 开启旁路
		{
			Util.callByPass(session, true);
		}

		NewOutputDataToExcel.writeDataToSheet(book, cover);
		viewPanel.addInfomation("", 80, 100);
		NewOutputDataToExcel.exportFile(book, filename);

		String fullFileName = FileUtil.getReportFileName(filename);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
		datasetList.add(ds);
		revlist.add(boprev);
		try {
			TCComponentItem document = ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "",
					session);
			// 根据文件编号创建顶层文件夹在home下,并创建00.封面文件夹，把封面文档放到00.封面文件夹下
			String topfoldername = VehicleNo + "-" + factoryline + "-AB";
			saveFileToFolder(document, topfoldername);

			// 文件编号和虚层名称
			TCProperty pdoc = document.getTCProperty("dfl9_vehiclePlant");
			if (pdoc != null) {
				if (pdoc != null) {
					pdoc.setStringValue(topfoldername);
					document.lock();
					document.save();
					document.unlock();
				}
			}
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}
		// 关闭旁路
		{
			Util.callByPass(session, false);
		}
		viewPanel.addInfomation("输出报表完成，请在焊装工厂工艺版本附件下查看!", 100, 100);

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
			TCComponentFolderType foldertype = (TCComponentFolderType) session.getTypeComponent("Folder");
			if (folder == null) {
				folder = foldertype.create(topfoldername, "", "Folder");
				homefolder.add("contents", folder);
				childrenfolder = foldertype.create("00.封面", "", "Folder");
				folder.add("contents", childrenfolder);
				childrenfolder.add("contents", document);
			} else {
				// 先判断是否已经创建了00.封面文件夹
				AIFComponentContext[] icf1 = folder.getChildren();
				for (AIFComponentContext aif : icf1) {
					TCComponent tcc = (TCComponent) aif.getComponent();
					String obejctname = Util.getProperty(tcc, "object_name");
					if (tcc.getType().equals("Folder") && obejctname.equals("00.封面")) {
						childrenfolder = (TCComponentFolder) tcc;
						break;
					}
				}
				if (childrenfolder == null) {
					childrenfolder = foldertype.create("00.封面", "", "Folder");
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
			}

		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
