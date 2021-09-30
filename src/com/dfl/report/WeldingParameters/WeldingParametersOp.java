package com.dfl.report.WeldingParameters;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.swt.widgets.Shell;

import com.dfl.report.common.TCComponentUtils;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.schemas.soa._2006_03.exceptions.ServiceException;
import com.teamcenter.services.rac.core._2008_06.DataManagement.CreateResponse;

public class WeldingParametersOp {

	private String Edition;
	private TCComponent savefolder;
	private TCSession session;
	private TCComponentBOMLine topbomline;
	SimpleDateFormat dateformat = new SimpleDateFormat("yyyy.MM");// 设置日期格式
	SimpleDateFormat dateformat2 = new SimpleDateFormat("yyyy年MM月");// 设置日期格式
	private List other = new ArrayList();
	List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private InputStream inputStream;

	public WeldingParametersOp(TCComponentBOMLine topbomline, TCSession session, String edition, TCComponent savefolder, InputStream inputStream)
			throws TCException {
		// TODO Auto-generated constructor stub
		this.Edition = edition;
		this.savefolder = savefolder;
		this.session = session;
		this.topbomline = topbomline;
		this.inputStream = inputStream;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub

		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("开始输出报表...\n", 10, 100);

	
		viewPanel.addInfomation("开始获取数据...\n", 30, 100);
		String vehicle = Util.getProperty(topbomline, "bl_rev_project_ids");// 基本车型

		// 生成报表操作前的动作
//		GenerateReportInfo info = new GenerateReportInfo();
//		info.setExist(false);
//		info.setIsgoon(true);
//		info.setAction(""); //$NON-NLS-1$
//		info.setMeDocument(null);
//		info.setDFL9_process_type("H"); //$NON-NLS-1$
//		info.setDFL9_process_file_type("CS"); // $NON-NLS-1$
//		info.setmeDocumentName(procName);
//		info.setFlag(true);
//		try {
//			info = ReportUtils.beforeGenerateReportAction(topbomline.getItemRevision(), info);
//		} catch (TCException e) {
//			e.printStackTrace();
//			// EclipseUtils.info("Error : " + e.getMessage()); //$NON-NLS-1$
//			return;
//		}
//		System.out.println("The action is completed before the report operation is generated.");
//
//		if (!info.isIsgoon()) {
//			return;
//		}
//		// 只有新增，不考虑更新
//		info.setAction("create");

		String BOPname = Util.getProperty(topbomline, "bl_rev_object_name");
		String filecode = "";
		String[] BOPnames = BOPname.split("_");
		String factory = "";
		if (BOPnames != null && BOPnames.length > 2) {
			vehicle = BOPnames[1];
			filecode = BOPnames[2];
			if (filecode.length() > 3) {
				factory = filecode.substring(0, 3);
			}
		}
		// 获取封面基本信息
		String procName = vehicle + "_" + Edition + "_PSW参数汇总表"; // 车型_阶段_PSW参数汇总表;
		// 日期
		String date1 = dateformat2.format(new Date());
		TCComponentUser user = session.getUser();
		String username = user.getUserName();
		String[] cover = new String[6];
		cover[0] = "    车    型：" + vehicle;
		cover[1] = "    版    次：" + Edition;
		cover[2] = "    文件编号：" + "共通-" + filecode + "-CS";
		cover[3] = "    编制日期：" + date1;
		cover[4] = username;
		cover[5] = "    工厂工程：" + factory + "工厂焊装工程";

		// 获取BOP下的虚层产线
		List obj = new ArrayList();
		List asshilist = getAsahiLine(topbomline);
		List factasshilist = new ArrayList();
		int totalpage = 1;
		viewPanel.addInfomation("", 40, 100);
		if (asshilist != null && asshilist.size() > 0) {		
			int count = 0;
			for (int i = 0; i < asshilist.size(); i++) {
				TCComponentBOMLine xuproc = (TCComponentBOMLine) asshilist.get(i);
				List gunlist = new ArrayList();
				List humangun = getHumanGunList(xuproc);
				//boolean Isweldpoint = getWeldPointList(xuproc);
				if (humangun != null && humangun.size() > 0) {
					factasshilist.add(xuproc);
					for (int j = 0; j < humangun.size(); j++) {
						String[] strValue = new String[17];
						TCComponentBOMLine gun = (TCComponentBOMLine) humangun.get(j);
						TCComponentItemRevision gunrev = gun.getItemRevision();

//						strValue[0] = Util.getProperty(gunrev, "b8_AdapterModel");// 变压器型号
//						strValue[1] = Util.getProperty(gunrev, "b8_Model");// 焊枪型号
						// 获取对应的实际产线
						TCComponentBOMLine procbl = gun.parent().parent().parent();
						boolean flag = Util.getIsMEProcStat(procbl);
						if (flag) {
							strValue[2] = Util.getProperty(procbl.getItemRevision(), "b8_ChineseName")
									+ Util.getProperty(gun.parent().parent(), "bl_rev_object_name");// 使用工位
						} else {
							strValue[2] = Util.getProperty(procbl.getItemRevision(), "b8_ChineseName");// 使用工位
						}
						TCComponentItemRevision diearrev = gun.parent().getItemRevision();
						String diearname = Util.getProperty(diearrev, "object_name");
						// 变压器编号和焊枪编号，改为从点焊工序名称中取值
						String[] nameArr = diearname.split("\\\\");
						String TransformerNumber = "";
						String Guncode = "";
						TransformerNumber = nameArr[0];
						if (nameArr.length > 1) {
							Guncode = nameArr[1];
						}
						strValue[0] = TransformerNumber;
						strValue[1] = Guncode;
						strValue[3] = "●";// 车型
						strValue[4] = Util.getProperty(gun.getItemRevision(), "b8_ElectrodeVol");// 焊钳额定压力
						strValue[5] = Util.getProperty(diearrev, "b8_WeldForce");// 加压力
						strValue[6] = "15";// 预压时间
						strValue[7] = Util.getProperty(diearrev, "b8_RiseTime");// 上升时间
						strValue[8] = Util.getProperty(diearrev, "b8_CurrentTime1");// 第一通电时间
						strValue[9] = Util.getProperty(diearrev, "b8_Current1");// 第一通电电流
						strValue[10] = Util.getProperty(diearrev, "b8_Cool1");// 冷却时间一
						strValue[11] = Util.getProperty(diearrev, "b8_CurrentTime2");// 第二通电时间
						strValue[12] = Util.getProperty(diearrev, "b8_Current2");// 第二通电电流
						strValue[13] = Util.getProperty(diearrev, "b8_Cool2");// 冷却时间二
						strValue[14] = Util.getProperty(diearrev, "b8_CurrentTime3");// 第三通电时间
						strValue[15] = Util.getProperty(diearrev, "b8_Current3");// 第三通电电流
						strValue[16] = Util.getProperty(diearrev, "b8_KeepTime");// 保持
						gunlist.add(strValue);
					}
					
					count++;
					
					// 根据工位排序
					Comparator comparator2 = getComParatorBygwname();
					Collections.sort(gunlist, comparator2);
					
					obj.add(gunlist);
				}
				
			}
		}
		totalpage = factasshilist.size();
		// 公共信息
		String[] common = new String[5];
		common[0] = username;
		common[1] = dateformat.format(new Date());
		common[2] = Edition;
		common[3] = Integer.toString(totalpage);
		String fatory = " 一工厂 NO1";
		if (filecode != null && filecode.length() > 4) {
			String prefactory = filecode.substring(0, 2);
			String math = filecode.substring(2, 3);
			// 进行数字转换
			String upermath = getUpperMath(math);
			String after = filecode.substring(filecode.length() - 1);
			fatory = prefactory + " " + upermath + "工厂 " + "NO" + after;

		}
		common[4] = fatory;
		viewPanel.addInfomation("", 50, 100);
		// 根据虚层产线数量，加载空模板
		XSSFWorkbook book = creatEngineeringXSSFWorkbook(inputStream, factasshilist);

		// 开始写入数据
		writeDataToSheet(book, cover, obj, vehicle, common);

		viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);
		String filename = Util.formatString(procName);

		NewOutputDataToExcel.exportFile(book, filename);

		String fullFileName = FileUtil.getReportFileName(filename);
		System.out.println(fullFileName);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
//		if (ds != null) {
//			datasetList.add(ds);
//		}
//		revlist.add(topbomline.getItemRevision());
		try {
			TCComponentItem docunment = AddDocumentItem(ds, procName);

			viewPanel.addInfomation("", 80, 100);

			savefolder.add("contents", docunment);
			// 文件编号和虚层名称
			TCProperty pdoc = docunment.getTCProperty("dfl9_vehiclePlant");
			if (pdoc != null) {
				pdoc.setStringValue(fatory.replace(" ", ""));
				docunment.lock();
				docunment.save();
				docunment.unlock();
			}
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}
		viewPanel.addInfomation("输出报表完成，请在所选文件夹下查看报表...\n", 100, 100);
	}

	private boolean getWeldPointList(TCComponentBOMLine bl) throws TCException {
		// TODO Auto-generated method stub
		List list = new ArrayList();
		String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { weldtypename, weldtypename };

		ArrayList weldlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		if (weldlist != null && weldlist.size() > 0) {
			for (int i = 0; i < weldlist.size(); i++) {
				// 根据点焊工序名称判断是否为人工焊枪
				TCComponentBOMLine weld = (TCComponentBOMLine) weldlist.get(i);
				String direaname = Util.getProperty(weld.parent(), "bl_rev_object_name");
				if (weld.parent().getItemRevision().isTypeOf("B8_BIWDiscreteOPRevision")
						&& !direaname.substring(0, 1).equals("R")) {
					return true;
				}
			}
		}
		return false;
	}

	private TCComponentItem AddDocumentItem(TCComponentDataset ds, String procName) throws TCException {
		// TODO Auto-generated method stub
		// 将所有excel文件创建为MSExcelX数据集并使用规格关系挂载到一个新建MEDocument版本
		TCComponentItem docuitem = null;
		Map<String, Object> itemMap = new HashMap<String, Object>();
		Map<String, Object> itemRevisionMap = new HashMap<String, Object>();
		Map<String, Object> itemRevMasterFormMap = new HashMap<String, Object>();
		itemMap.put("item_id", ""); //$NON-NLS-1$ //$NON-NLS-2$
		itemMap.put("object_name", procName); //$NON-NLS-1$
		itemMap.put("object_desc", ""); //$NON-NLS-1$
		itemMap.put("object_type", "DFL9MEDocument"); //$NON-NLS-1$
		itemRevisionMap.put("object_type", "DFL9MEDocumentRevision"); //$NON-NLS-1$
		itemRevisionMap.put("object_name", procName); //$NON-NLS-1$
		itemRevisionMap.put("dfl9_process_type", "H"); //$NON-NLS-1$
		itemRevisionMap.put("dfl9_process_file_type", "CS"); //$NON-NLS-1$
		itemRevMasterFormMap.put("object_type", "DFL9MEDocumentRevisionMaster"); //$NON-NLS-1$

		try {
			CreateResponse respose = TCComponentUtils.create(itemMap, itemRevisionMap, itemRevMasterFormMap);
			int num = respose.serviceData.sizeOfCreatedObjects();
			if (num > 0) {
				for (int i = 0; i < num; i++) {
					TCComponent comp = respose.serviceData.getCreatedObject(i);
					if (comp instanceof TCComponentItemRevision) {
						TCComponentItemRevision docuitemrev = (TCComponentItemRevision) comp;
						docuitemrev.add("IMAN_specification", ds);
						docuitem = docuitemrev.getItem();
					}

				}
			}
		} catch (ServiceException e) {
			e.printStackTrace();
			// throw new TCException("Create " + ReportUtils.DFL9MEDocument + " Fail : "
			// +e.getMessage()); //$NON-NLS-1$ //$NON-NLS-2$
		}
		return docuitem;
	}

	private String getUpperMath(String math) {
		// TODO Auto-generated method stub
		String value = "";
		String[] Uppernumbers = { "一", "二", "三", "四", "五", "六", "七", "八", "九" };
		if (Util.isNumber(math)) {
			int num = Integer.parseInt(math);
			value = Uppernumbers[num - 1];
		}
		return value;
	}

	private void writeDataToSheet(XSSFWorkbook book, String[] cover, List obj, String vehicle, String[] common) {
		// TODO Auto-generated method stub
		// 先写入封面信息
		// 设置字体
		Font font = book.createFont();
		font.setFontName("新宋体");
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);// 加粗
		font.setFontHeightInPoints((short) 16);
		// 创建一个样式
		XSSFCellStyle cellStyle = book.createCellStyle();
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_NONE); // 下边框
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_NONE);// 左边框
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_NONE);// 上边框
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_NONE);// 右边框
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// 左对齐
		cellStyle.setFont(font);

		Font font2 = book.createFont();
		font2.setFontName("新宋体");
		// font2.setBoldweight(Font.BOLDWEIGHT_BOLD);// 加粗
		font2.setFontHeightInPoints((short) 16);
		// 创建一个样式
		XSSFCellStyle cellStyle2 = book.createCellStyle();
		cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_NONE); // 下边框
		cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_NONE);// 左边框
		cellStyle2.setBorderTop(XSSFCellStyle.BORDER_NONE);// 上边框
		cellStyle2.setBorderRight(XSSFCellStyle.BORDER_NONE);// 右边框
		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);// 内容垂直居中
		// cellStyle2.setAlignment(XSSFCellStyle.ALIGN_LEFT);// 左对齐
		cellStyle2.setFont(font2);

		Font font3 = book.createFont();
		font3.setFontName("黑体");
		// font2.setBoldweight(Font.BOLDWEIGHT_BOLD);// 加粗
		font3.setFontHeightInPoints((short) 10);
		// 创建一个样式
		XSSFCellStyle cellStyle3 = book.createCellStyle();
		cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		cellStyle3.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);// 内容垂直居中
		cellStyle3.setFont(font3);

		// 创建一个样式
		XSSFCellStyle cellStyle4 = book.createCellStyle();
		cellStyle4.setFillForegroundColor(IndexedColors.RED.getIndex());
		cellStyle4.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cellStyle4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		cellStyle4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		cellStyle4.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		cellStyle4.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		cellStyle4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle4.setAlignment(XSSFCellStyle.ALIGN_CENTER);// 内容垂直居中
		cellStyle4.setFont(font3);

		// 设置字体
		Font font4 = book.createFont();
		font4.setColor((short) 12);
		font4.setFontName("宋体");
		font4.setFontHeightInPoints((short) 16);
		// 创建一个样式
		XSSFCellStyle cellStyle5 = book.createCellStyle();
		cellStyle5.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle5.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle5.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle5.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle5.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle5.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle5.setFont(font4);

		// 设置字体
		Font font5 = book.createFont();
		font5.setColor((short) 12);
		font5.setFontName("MS PGothic");
		font5.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
		font5.setFontHeightInPoints((short) 12);
		// 创建一个样式
		XSSFCellStyle cellStyle6 = book.createCellStyle();
		cellStyle6.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle6.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle6.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle6.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle6.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle6.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle6.setFont(font5);

		// 创建一个样式
		XSSFCellStyle cellStyle61 = book.createCellStyle();
		cellStyle61.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle61.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		//cellStyle61.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle61.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle61.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle61.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle61.setFont(font5);

		// 创建一个样式
		XSSFCellStyle cellStyle62 = book.createCellStyle();
		cellStyle62.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		//cellStyle62.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle62.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle62.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle62.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle62.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle62.setFont(font5);

		// 设置字体
		Font font6 = book.createFont();
		font6.setColor((short) 12);
		font6.setFontName("宋体");
		font6.setFontHeightInPoints((short) 28);
		font6.setUnderline(Font.U_SINGLE);// 设置下划线
		// 创建一个样式
		XSSFCellStyle cellStyle7 = book.createCellStyle();
		cellStyle7.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle7.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle7.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle7.setBorderTop(XSSFCellStyle.BORDER_THIN);
		cellStyle7.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle7.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellStyle7.setFont(font6);

		Font font7 = book.createFont();
		font7.setFontName("黑体");
		// font2.setBoldweight(Font.BOLDWEIGHT_BOLD);// 加粗
		font7.setFontHeightInPoints((short) 8);
		// 创建一个样式
		XSSFCellStyle cellStyle8 = book.createCellStyle();
		cellStyle8.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		cellStyle8.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		cellStyle8.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		cellStyle8.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		cellStyle8.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle8.setAlignment(XSSFCellStyle.ALIGN_CENTER);// 内容垂直居中
		cellStyle8.setFont(font7);

		XSSFSheet sheet = book.getSheetAt(0);

		setStringCellAndStyle(sheet, cover[5], 5, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[0], 6, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[1], 7, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[2], 8, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[3], 10, 3, cellStyle, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet, cover[4], 12, 1, cellStyle2, Cell.CELL_TYPE_STRING);
        
		
		if (obj != null && obj.size() > 0) {
			for (int i = 0; i < obj.size(); i++) {							
					XSSFSheet sh = book.getSheetAt(1 + i);
					List data = (List) obj.get(i);
					setStringCellAndStyle(sh, vehicle, 6, 25, cellStyle3, Cell.CELL_TYPE_STRING);

					setStringCellAndStyle(sh, common[0], 2, 6, cellStyle5, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, common[1], 2, 30, cellStyle5, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, common[2], 50, 110, cellStyle6, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, Integer.toString(i + 1), 52, 110, cellStyle61, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, common[3], 52, 116, cellStyle62, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sh, common[4] + "    " + "PSW焊接参数表", 0, 37, cellStyle7, Cell.CELL_TYPE_STRING);
					if (data != null && data.size() > 0) {
						// 起始行
						String beginstr = "";
						int beginrow = 9;// 起始行
						int endrow = 0;// 终止行
						int n = 0;// 标记行
						// 合并单元格
						CellRangeAddress region1;
						for (int j = 0; j < data.size(); j++) {
							String[] values = (String[]) data.get(j);
							setStringCellAndStyle(sh, Integer.toString(j + 1), 9 + j, 1, cellStyle3, 10);
							setStringCellAndStyle(sh, values[0], 9 + j, 4, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[1], 9 + j, 8, cellStyle8, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[2], 9 + j, 15, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[3], 9 + j, 25, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[4], 9 + j, 46, cellStyle3, 11);
							// “加压力”大于“焊钳额定压力”，红色填充
							double hs = getDoubleByString(values[4]);
							double js = getDoubleByString(values[5]);
							if (Util.isNumber(values[4]) && Util.isNumber(values[5])) {
								if (js > hs) {
									setStringCellAndStyle(sh, values[5], 9 + j, 50, cellStyle4, 11);
								} else {
									setStringCellAndStyle(sh, values[5], 9 + j, 50, cellStyle3, 11);
								}
							} else {
								setStringCellAndStyle(sh, values[5], 9 + j, 50, cellStyle3, 11);
							}

							setStringCellAndStyle(sh, values[6], 9 + j, 56, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[7], 9 + j, 62, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[8], 9 + j, 68, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[9], 9 + j, 74, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[10], 9 + j, 80, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[11], 9 + j, 86, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[12], 9 + j, 92, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[13], 9 + j, 98, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[14], 9 + j, 104, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[15], 9 + j, 110, cellStyle3, Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sh, values[16], 9 + j, 116, cellStyle3, Cell.CELL_TYPE_STRING);

							if (j == 0) {
								beginstr = values[2];
							} else {
								/**
								 * ******************************** 工位列合并
								 */
								if (!values[2].equals(beginstr)) {
									endrow = beginrow + n;
									region1 = new CellRangeAddress(beginrow, endrow, (short) 15, (short) 24);
									sh.addMergedRegion(region1);
									beginstr = values[2].toString();
									beginrow = endrow + 1;
									n = 0;
								} else {
									n++;
								}
								if (j == data.size() - 1) {
									endrow = beginrow + n;
									region1 = new CellRangeAddress(beginrow, endrow, (short) 15, (short) 24);
									sh.addMergedRegion(region1);
								}
							}
							if(data.size() == 1) {
								region1 = new CellRangeAddress(9, 9, (short) 15, (short) 24);
								sh.addMergedRegion(region1);
							}
						}
					}
				
				
			}
		}

		for (int i = 1; i < book.getNumberOfSheets(); i++) {
			XSSFSheet sht = book.getSheetAt(i);
			book.setPrintArea(i, 0, 123, 0, 54);
			PrintSetup printSetup = sht.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 62);// 自定义缩放，此处100为无缩放
			printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
		}
	}

	private Comparator getComParatorBygwname() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				// System.setProperty("java.util.Arrays.useLegacyMergeSort", "true");
				String[] comp1 = (String[]) obj;
				String[] comp2 = (String[]) obj1;

				String d1 = "";
				String d2 = "";
				if (obj != null && comp1[2] != null && !comp1[2].isEmpty()) {
					d1 = comp1[2].toString();
				}
				if (obj1 != null && comp2[2] != null && !comp2[2].isEmpty()) {
					d2 = comp2[2];
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
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
				if(Util.isNumber(value)) {
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(Integer.parseInt(value));
				}else {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(value);
				}
				
			} else if (celltype == 11) {
				if(Util.isNumber(value)) {
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(Double.parseDouble(value));
				}else {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(value);
				}
				
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			}
		}

		//cell.setCellStyle(Style);

	}

	private XSSFWorkbook creatEngineeringXSSFWorkbook(InputStream inputStream, List asshilist) {
		// TODO Auto-generated method stub
		XSSFWorkbook book = null;
		try {
			book = new XSSFWorkbook(inputStream);
			if (asshilist != null && asshilist.size() > 0) {
				for (int i = 0; i < asshilist.size(); i++) {
					TCComponentBOMLine xuproc = (TCComponentBOMLine) asshilist.get(i);
					String sheetname = Util.getProperty(xuproc, "bl_rev_object_name");
					String[] str = sheetname.split("[-_ ]");
					if(str!=null && str.length>1) {
						sheetname = str[1];
					}
					sheetname = String.format("%02d", i + 1) + sheetname;
					if (i == 0) {
						book.setSheetName(1, sheetname);
					} else {
						XSSFSheet newsheet = book.cloneSheet(1);
						int sheetat = book.getSheetIndex(newsheet);
						book.setSheetName(sheetat, sheetname);
						// book.setSheetOrder(newsheet.getSheetName(), index);
					}
				}
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return book;
	}

	// 根据虚层产线获取其下的人工焊枪
	private List getHumanGunList(TCComponentBOMLine bl) throws TCException {
		List list = new ArrayList();
		String guntypename = Util.getObjectDisplayName(session, "B8_BIWGun");
		String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
		String[] values = new String[] { guntypename, guntypename };

		ArrayList gunlist = Util.searchBOMLine(bl, "OR", propertys, "==", values);
		if (gunlist != null && gunlist.size() > 0) {
			for (int i = 0; i < gunlist.size(); i++) {
				// 根据点焊工序名称判断是否为人工焊枪
				TCComponentBOMLine gun = (TCComponentBOMLine) gunlist.get(i);
				String direaname = Util.getProperty(gun.parent(), "bl_rev_object_name");
				if (gun.parent().getItemRevision().isTypeOf("B8_BIWDiscreteOPRevision")
						&& !direaname.substring(0, 1).equals("R")) {
					list.add(gun);
				}
			}
		}
		return list;
	}

	// 遍历BOP顶层，获取所有的虚层产线，如果产线没有放在虚层产线，归为其他
	private List getAsahiLine(TCComponentBOMLine topbl) throws TCException {
		// TODO Auto-generated method stub
		List list = new ArrayList();
		AIFComponentContext[] chilrens = topbl.getChildren();
		for (AIFComponentContext chil : chilrens) {
			TCComponentBOMLine bl = (TCComponentBOMLine) chil.getComponent();
			// 根据产线下是否有产线判断是否为虚层
			ArrayList xclist = Util.getChildrenByBOMLine(bl, "B8_BIWMEProcLineRevision");
			if (xclist != null && xclist.size() > 0) {
				list.add(bl);
			} else {
				other.add(bl);
			}
		}
		return list;
	}

	private double getDoubleByString(String str) {
		double dd = 0;
		if (str == null || !Util.isNumber(str)) {
			dd = 0;
		} else {
			dd = Double.parseDouble(str);
		}

		return dd;
	}
}
