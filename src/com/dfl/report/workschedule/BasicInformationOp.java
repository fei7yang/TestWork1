package com.dfl.report.workschedule;

import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import com.dfl.report.ExcelReader.BoardInformation;
import com.dfl.report.ExcelReader.WeldPointInfo;
import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.GenerateReportInfo;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportUtils;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFOperation;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentBOMWindow;
import com.teamcenter.rac.kernel.TCComponentBOMWindowType;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentFolderType;
import com.teamcenter.rac.kernel.TCComponentGroup;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCComponentQuery;
import com.teamcenter.rac.kernel.TCComponentQueryType;
import com.teamcenter.rac.kernel.TCComponentRole;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCComponentUser;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class BasicInformationOp  extends AbstractAIFOperation {

	private AbstractAIFUIApplication app;
	private ArrayList weld = new ArrayList();
	private List<WeldPointInfo> weldlist = new ArrayList<WeldPointInfo>();// 基本信息表的数据
	// private TCComponentItemType tcccomponentitemtype;
	private TCComponentBOMWindow bomWin;
	List<TCComponentDataset> datasetList = new ArrayList<TCComponentDataset>();
	List<TCComponentItemRevision> revlist = new ArrayList<TCComponentItemRevision>();
	private Map<String, String> projVehMap;// 获取首选项车型代号与familycode的关系
	private String VehicleNo = "";// 车型代号
	private DecimalFormat format = new DecimalFormat("0.00");
	private DecimalFormat format1 = new DecimalFormat("0.0000");
	private GenerateReportInfo info;
    private InputStream inputStream ;

	public BasicInformationOp(AbstractAIFUIApplication app, GenerateReportInfo info, InputStream inputStream) throws TCException {
		// TODO Auto-generated constructor stub
		this.app = app;
		this.info = info;
		this.inputStream = inputStream;
	}

	@Override
	public void executeOperation() throws Exception {
		InterfaceAIFComponent ifc = app.getTargetComponent();
		TCComponentBOMLine topbl = (TCComponentBOMLine) ifc;

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
		String procName = "222.基本信息";
	
		// 首次生成报表
		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
		viewPanel.setVisible(true);

		viewPanel.addInfomation("开始输出报表...\n", 10, 100);

		
		XSSFWorkbook book = null;

		if (info.getAction() == "create") { // 都输出
			
			viewPanel.addInfomation("", 20, 100);
			// 获取焊点相关信息
			getAllWeldPoint(session, topbl,viewPanel);

			book = creatXSSFWorkbook(inputStream);
			//viewPanel.addInfomation("", 40, 100);
			writeHDDataToSheet(book, weld);

		} else {			
			if (inputStream != null) {
				book = creatXSSFWorkbook(inputStream);
			} 
			viewPanel.addInfomation("", 20, 100);

			// 先清空焊点信息
			clearHDDataToSheet(book);

			// 获取焊点相关信息
			getAllWeldPoint(session, topbl,viewPanel);

			//viewPanel.addInfomation("", 40, 100);
			writeHDDataToSheet(book, weld);

		}
		// 开启旁路
		{
			Util.callByPass(session, true);
		}
		viewPanel.addInfomation("开始写数据，请耐心等待...\n", 80, 100);
		String filename = procName;
		filename = filename.replaceAll("\\s*", "");
		NewOutputDataToExcel.exportFile(book, filename);

		String fullFileName = FileUtil.getReportFileName(filename);
		System.out.println(fullFileName);
		TCComponentDataset ds = Util.createDataset(session, filename, fullFileName, "MSExcelX", "excel");
		if (ds != null) {
			datasetList.add(ds);
		}
		try {
			revlist.add(topbl.getItemRevision());
		} catch (TCException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		viewPanel.addInfomation("", 90, 100);
		try {
			ReportUtils.afterGenerateReportAction(datasetList, revlist, info, procName, "", session);
		} catch (TCException e) {
			e.printStackTrace();
			// EclipseUtils.info(Messages.FixtureOperation_15 + e.getMessage());
			return;
		}
		// 关闭旁路
		{
			Util.callByPass(session, false);
		}
		viewPanel.addInfomation("输出报表完成，请在焊装工厂工艺对象附件下查看\n", 100, 100);
		viewPanel.addInfomation("温馨提示：基本信息-焊点信息生成成功，下次再输出会把焊点信息覆盖，请谨慎操作！！\n", 100, 100);

	}

	/*
	 * 清空焊点清单数据
	 */
	private void clearHDDataToSheet(XSSFWorkbook book) {
		XSSFCellStyle style = book.createCellStyle();
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		XSSFSheet sheet = book.getSheetAt(3);
		int rownum = sheet.getPhysicalNumberOfRows();
		int row = 1;
		for (int i = 0; i < rownum; i++) {
			for (int j = 0; j < 17; j++) {
				setStringCellAndStyle(sheet, "", row, j, style, Cell.CELL_TYPE_STRING);
			}
			row++;
		}
	}

	/*
	 * 写焊点清单数据
	 */
	private void writeHDDataToSheet(XSSFWorkbook book, ArrayList hdlist) {
		XSSFCellStyle style = book.createCellStyle();
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		XSSFSheet sheet = book.getSheetAt(3);
		int row = 1;
		for (int i = 0; i < hdlist.size(); i++) {
			String[] value = (String[]) hdlist.get(i);
			for (int j = 0; j < value.length; j++) {
				setStringCellAndStyle(sheet, value[j], row, j, style, Cell.CELL_TYPE_STRING);
			}
			row++;
		}
	}

	/*
	 * 写所有板组sheet页的公共信息
	 */
	private void writeBZDataToSheet(XSSFWorkbook book, List plist) {
		// TODO Auto-generated method stub
		// 设置字体
		Font font = book.createFont();
		// font.setColor((short) 12);
		font.setFontName("宋体");
		font.setFontHeightInPoints((short) 10);
		// 创建一个样式
		XSSFCellStyle cellStyle1 = null;
//		cellStyle1.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle1.setFont(font);

		XSSFCellStyle cellStyle2 = null;
//		cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_NONE);
//		cellStyle2.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		cellStyle2.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		cellStyle2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle2.setFont(font);

		XSSFCellStyle cellStyle3 = null;
//		cellStyle3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setBorderRight(XSSFCellStyle.BORDER_NONE);
//		cellStyle3.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		cellStyle3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		cellStyle3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle3.setFont(font);

		// 焊点-板组sheet
		XSSFSheet sh = book.getSheetAt(1);
		for (int i = 0; i < plist.size(); i++) {
			String[] value = (String[]) plist.get(i);
			setStringCellAndStyle(sh, value[0], 1 + i, 0, cellStyle1, 10); // 序号
			setStringCellAndStyle(sh, value[2], 1 + i, 1, cellStyle1, Cell.CELL_TYPE_STRING);// 零件号
			setStringCellAndStyle(sh, value[1], 1 + i, 2, cellStyle1, Cell.CELL_TYPE_STRING);// 板组编号
			setStringCellAndStyle(sh, value[3], 1 + i, 4, cellStyle2, Cell.CELL_TYPE_STRING);// 零件名称
			setStringCellAndStyle(sh, value[4], 1 + i, 5, cellStyle1, Cell.CELL_TYPE_STRING);// 材质
			if (Util.isNumber(value[5])) {
				setStringCellAndStyle(sh, value[5], 1 + i, 6, cellStyle3, 11);// 板厚
			} else {
				setStringCellAndStyle(sh, value[5], 1 + i, 6, cellStyle3, Cell.CELL_TYPE_STRING);// 板厚
			}
			setStringCellAndStyle(sh, value[6], 1 + i, 7, cellStyle2, Cell.CELL_TYPE_STRING);// 板厚单位
			setStringCellAndStyle(sh, value[7], 1 + i, 8, cellStyle3, 10);// 强度
			setStringCellAndStyle(sh, value[8], 1 + i, 9, cellStyle2, Cell.CELL_TYPE_STRING);// 强度单位
			setStringCellAndStyle(sh, value[9], 1 + i, 10, cellStyle1, Cell.CELL_TYPE_STRING);// GA/GI
		}

		// 板组清单sheet
		XSSFSheet sh2 = book.getSheetAt(2);
		for (int i = 0; i < plist.size(); i++) {
			String[] value = (String[]) plist.get(i);
			setStringCellAndStyle(sh2, value[0], 1 + i, 0, cellStyle1, 10); // 序号
			setStringCellAndStyle(sh2, value[1], 1 + i, 1, cellStyle1, Cell.CELL_TYPE_STRING);// 板组编号
			setStringCellAndStyle(sh2, value[2], 1 + i, 2, cellStyle1, Cell.CELL_TYPE_STRING);// 零件号
			setStringCellAndStyle(sh2, value[3], 1 + i, 4, cellStyle2, Cell.CELL_TYPE_STRING);// 零件名称
			setStringCellAndStyle(sh2, value[4], 1 + i, 5, cellStyle1, Cell.CELL_TYPE_STRING);// 材质
			if (Util.isNumber(value[5])) {
				setStringCellAndStyle(sh2, value[5], 1 + i, 6, cellStyle3, 11);// 板厚
			} else {
				setStringCellAndStyle(sh2, value[5], 1 + i, 6, cellStyle3, Cell.CELL_TYPE_STRING);// 板厚
			}
			setStringCellAndStyle(sh2, value[6], 1 + i, 7, cellStyle2, Cell.CELL_TYPE_STRING);// 板厚单位
			setStringCellAndStyle(sh2, value[7], 1 + i, 8, cellStyle3, 10);// 强度
			setStringCellAndStyle(sh2, value[8], 1 + i, 9, cellStyle2, Cell.CELL_TYPE_STRING);// 强度单位
			setStringCellAndStyle(sh2, value[9], 1 + i, 10, cellStyle1, Cell.CELL_TYPE_STRING);// GA/GI
		}

	}

	private List getSolutionPart(List<WeldPointInfo> weldlist2, TCComponentBOMLine topbl, TCSession session)
			throws TCException {
		// TODO Auto-generated method stub
		ArrayList bzqclist = new ArrayList();
		ArrayList partlist = new ArrayList();
		int rowNum = 0;// 序号
		// 通过零件号查找对应的零件名称
		if (weldlist != null) {
			// 对板组数据根据零件号排序
			Comparator comparator = getComParatorBypartno();
			Collections.sort(weldlist, comparator);

			TCComponentItemRevision toprev = topbl.getItemRevision();
			TCComponent[] bbomlist = toprev.getRelatedComponents("IMAN_METarget");

			for (int i = 0; i < weldlist.size(); i++) {
				WeldPointInfo weldinfo = weldlist.get(i);
				String partNo = weldinfo.getPartno();

				// 根据零件号去重
				if (!bzqclist.contains(partNo)) {
					bzqclist.add(partNo);
					String[] values = new String[10];
					values[0] = Integer.toString(rowNum + 1);// 序号
					values[1] = "BZ" + String.format("%03d", rowNum + 1);// 板组编号
					values[2] = partNo; // 零件号
					String partname = "";// 零件名称
					System.out.println("关联的BBOM对象为:" + bbomlist.toString());
					if (bbomlist != null && bbomlist.length > 0) {
						for (int j = 0; j < bbomlist.length; j++) {
							TCComponentItemRevision bbomrev = (TCComponentItemRevision) bbomlist[j];
							// TCSession session = (TCSession) app.getSession();
							TCComponentBOMLine root;
							// 在制造工艺规划器中获取打开的BBOM结构
							String bbomID = Util.getProperty(bbomrev, "item_id");
							System.out.println("bbomID:" + bbomID);

							root = Util.getOpenBOMLine(bbomID);
							if (root == null) {
								TCComponentBOMWindowType bomWinType;
								bomWinType = (TCComponentBOMWindowType) session.getTypeComponent("BOMWindow");
								bomWin = bomWinType.create(null);
								root = bomWin.setWindowTopLine(null, bbomrev, null, null);
							}
							// 调用系统查询，获取相关的板件
							// 如果零件号是添加了后缀，需要截取掉，再用于查询
							String querypartno = "";
							if (partNo.length() > 2) {
								String spcilchar = partNo.substring(partNo.length() - 2, partNo.length() - 1);
								if (spcilchar.equals("-")) {
									querypartno = partNo.substring(0, partNo.length() - 2);
								} else {
									querypartno = partNo;
								}
							}
							List tcclist = Util.callStructureSearch(root, "__DFL_Find_SolutionPart",
									new String[] { "PARTNO" }, new String[] { querypartno });
							if (tcclist != null && tcclist.size() > 0) {
								TCComponentBOMLine solbl = (TCComponentBOMLine) tcclist.get(0);
								//partname = Util.getProperty(solbl, "bl_rev_object_name");
								partname = Util.getProperty(solbl.getItemRevision(), "dfl9_CADObjectName");
								break; // 找到一个对应的零件名称就中断查询
							}
							if (bomWin != null) {
								bomWin.close();
							}
						}
					}
					values[3] = partname;
					String Partmaterial = weldinfo.getPartmaterial();
					values[4] = Partmaterial; // 材质
					values[5] = weldinfo.getPartthickness(); // 板厚
					if (values[5] != null && !values[5].isEmpty()) {
						values[6] = "mm";
					} else {
						values[6] = "";
					}

					// 根据材质获取强度和GA/GI材属性
					String Sheetstrength = "";// 强度
					String Gagi = "";// GA/GI
					// 如果是厚薄板，无法获取强度和GA/GI材
					boolean flag = getJudgingThickSheet(Partmaterial);
					if (!flag) {
						if (Partmaterial != null && !Partmaterial.isEmpty()) {
							String[] str = Partmaterial.split("-");
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
							if (!Sheetstrength.isEmpty() && Integer.parseInt(Sheetstrength) >= 440) {
								values[7] = Sheetstrength;
								values[8] = "mpa";
							} else {
								values[7] = "";
								values[8] = "";
							}
						}
						if (Partmaterial != null && Partmaterial.length() > 4) {
							String gagitem = Partmaterial.trim().substring(0, 4);
							if (gagitem.equals("SP78") || gagitem.equals("SP79") || gagitem.equals("RP78")
									|| gagitem.equals("RP79")) {
								Gagi = "GA";
							} else if (gagitem.equals("SP70") || gagitem.equals("SP71") || gagitem.equals("SP72")
									|| gagitem.equals("SP73") || gagitem.equals("SP76") || gagitem.equals("RP70")
									|| gagitem.equals("RP71") || gagitem.equals("RP72") || gagitem.equals("RP73")
									|| gagitem.equals("RP76")) {
								Gagi = "GI";
							} else {
								Gagi = "";
							}
						}
					} else {
						values[7] = "";
						values[8] = "";
					}
					values[9] = Gagi;

					partlist.add(values);

					rowNum++;
				}
			}
		}

		return partlist;
	}

	// 判断是否为厚薄板
	private boolean getJudgingThickSheet(String partmaterial1) {
		// TODO Auto-generated method stub
		boolean flag = false;
		int count1 = 0;
		int count2 = 0;
		String str = "";
		if (partmaterial1 != null) {
			str = partmaterial1;
		}
		count1 = (str.length() - str.replace("SP", "").length()) / "SP".length();
		count2 = (str.length() - str.replace("RP", "").length()) / "RP".length();

		if (count1 + count2 > 1) {
			flag = true;
		}
		return flag;
	}

	private void getAllWeldPoint(TCSession session, TCComponentBOMLine topbl, ReportViwePanel viewPanel) {
		// TODO Auto-generated method stub

		System.out.println("用于测试无响应问题");
		// 获取BOP关联的BBOM对象
		ArrayList qclist = new ArrayList();
		try {
			TCComponentItemRevision toprev = topbl.getItemRevision();
			System.out.println("用于测试无响应问题");
			TCComponent[] bbomlist = toprev.getRelatedComponents("IMAN_METarget");
			System.out.println("关联的BBOM对象为:" + bbomlist.toString());
			if (bbomlist != null && bbomlist.length > 0) {
				// 定义一个Map用于判断是否为相同板件，避免重复查询
				Map<String, String[]> partMap = new HashMap<String, String[]>();
                double schedule = 60/bbomlist.length;
                int basesch = 20;
				for (int i = 0; i < bbomlist.length; i++) {
					if(i!=0) {
						basesch = basesch + (int)schedule;
					}
					TCComponentItemRevision bbomrev = (TCComponentItemRevision) bbomlist[i];
					// TCSession session = (TCSession) app.getSession();
					TCComponentBOMLine root;
					// 在制造工艺规划器中获取打开的BBOM结构
					String bbomID = Util.getProperty(bbomrev, "item_id");
					System.out.println("bbomID:" + bbomID);
					root = Util.getOpenBOMLine(bbomID);
					if (root == null) {
						TCComponentBOMWindowType bomWinType;
						bomWinType = (TCComponentBOMWindowType) session.getTypeComponent("BOMWindow");
						bomWin = bomWinType.create(null);
						root = bomWin.setWindowTopLine(null, bbomrev, null, null);
					}

					// 根据BBOM查询所有的焊点
					String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
					String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
					String[] values = new String[] { weldtypename, weldtypename };
//					List<TCComponent> lstScope = new ArrayList<TCComponent>();
//					lstScope.add(root);
//					List<TCComponent> partList = Util.callStructureSearch(lstScope, "__DFL_Find_Object_by_Name", new String[] { "NAME", "LX"},
//							new String[] { "焊点", "WeldPoint" });

					
					//modify by xiaolei 20200902
					//ArrayList partList = Util.searchBOMLine(root, "OR", propertys, "==", values);
					List<TCComponent> partList = Util.callStructureSearch(root, "__DFL_Find_Object_by_Name", new String[] { "LX"},new String[] { "WeldPoint" });
					
					
					
					System.out.println("包含的焊点：" + partList.toString());
					if (partList != null && partList.size() > 0) {

						// 根据焊点数组，一次性查询所有焊点关联的板件
//						TCComponentBOMLine[] partstr = new TCComponentBOMLine[partList.size()];
//						partList.toArray(partstr);
//						HashMap<TCComponentBOMLine,TCComponent[]> map = Util.getConnectedLines(session,partstr);

						for (int j = 0; j < partList.size(); j++) {
							double sch = (j + 1.0) / partList.size();
							int s = (int) (sch * schedule);
							if (s <=schedule) {
								viewPanel.addInfomation("", basesch+s, 100);
							}
							String[] value = new String[17];
							String[] bzvalue = new String[5];
							TCComponentBOMLine bl = (TCComponentBOMLine) partList.get(j);
							TCComponentItemRevision rev = bl.getItemRevision();

							value[0] = Util.getProperty(bl, "bl_rev_object_name");// 焊点编号
							// 获取x,y,z坐标
							String xform = Util.getProperty(bl, "bl_plmxml_abs_xform");// 绝对变换矩阵
							Double[] xyzArray = getXYZ(xform);
							Double x = xyzArray[0] * 1000;
							Double y = xyzArray[1] * 1000;
							Double z = xyzArray[2] * 1000;

							value[1] = format.format(new BigDecimal(x.toString()));// X坐标
							value[2] = format.format(new BigDecimal(y.toString()));// Y坐标
							value[3] = format.format(new BigDecimal(z.toString()));// Z坐标
							value[4] = "";// 1
							value[5] = "";// 2
							value[6] = "";// 3
							//只有A和B重要度才显示，其他的为空
							String important = Util.getProperty(rev, "b8_ImportantLevel");// 重要度b8_ImportantLevel
							if(important.equals("A") || important.equals("B")) {
								value[7] = important;// 重要度b8_ImportantLevel
							}else {
								value[7] = "";
							}	
							// 获取板件1
							String cp1 = "";
							// 获取板件2
							String cp2 = "";
							// 获取板件3
							String cp3 = "";
							// 获取板件 改为取连接到属性 bl_connected_lines
							String conlines = Util.getProperty(bl, "bl_connected_lines");
							if(conlines!=null && !conlines.isEmpty()) {
								String[] strValues = conlines.split(",");
								if(strValues.length == 1) {
									String[] strcp1 = strValues[0].split("/");
									cp1 = strcp1[0].trim();
								}else if(strValues.length == 2) {
									String[] strcp1 = strValues[0].split("/");
									cp1 = strcp1[0].trim();
									String[] strcp2 = strValues[1].split("/");
									cp2 = strcp2[0].trim();
								}else {
									String[] strcp1 = strValues[0].split("/");
									cp1 = strcp1[0].trim();
									String[] strcp2 = strValues[1].split("/");
									cp2 = strcp2[0].trim();
									String[] strcp3 = strValues[2].split("/");
									cp3 = strcp3[0].trim();
								}							
							}						
							if (cp1 != null && !cp1.equals("")) {
								// 调用系统查询，获取相关的板件
								if (partMap.containsKey(cp1)) {
									String[] strvalue = partMap.get(cp1);
									value[8] = cp1;
									value[9] = strvalue[0];
									value[10] = strvalue[1];
								} else {
									// 调用系统查询，获取相关的板件
									String[] strvalue = getPropertysBypartNo(root, cp1);
									value[8] = cp1;
									value[9] = strvalue[0];
									value[10] = strvalue[1];
									partMap.put(cp1, strvalue);
								}
								System.out.println(cp1 + " " + value[9] + " " + value[10]);
							}
							
							if (cp2 != null && !cp2.equals("")) {
								if (partMap.containsKey(cp2)) {
									String[] strvalue = partMap.get(cp2);
									value[11] = cp2;
									value[12] = strvalue[0];
									value[13] = strvalue[1];
								} else {
									// 调用系统查询，获取相关的板件
									String[] strvalue = getPropertysBypartNo(root, cp2);
									value[11] = cp2;
									value[12] = strvalue[0];
									value[13] = strvalue[1];
									partMap.put(cp2, strvalue);
								}
								System.out.println(cp2 + " " + value[12] + " " + value[13]);
							}
							
							if (cp3 != null && !cp3.equals("")) {
								// 调用系统查询，获取相关的板件
								if (partMap.containsKey(cp3)) {
									String[] strvalue = partMap.get(cp3);
									value[14] = cp3;
									value[15] = strvalue[0];
									value[16] = strvalue[1];
								} else {
									// 调用系统查询，获取相关的板件
									String[] strvalue = getPropertysBypartNo(root, cp3);
									value[14] = cp3;
									value[15] = strvalue[0];
									value[16] = strvalue[1];
									partMap.put(cp3, strvalue);
								}
								System.out.println(cp3 + " " + value[15] + " " + value[16]);
							}
							// 根据焊点编号去重
//							if (!qclist.contains(value[0])) {
//								qclist.add(value[0]);
//								weld.add(value);
//							}
							weld.add(value);
						}
					}else {
						viewPanel.addInfomation("", 40, 100);
					}					
					if (bomWin != null) {
						bomWin.close();
					}
				}
			}else {
				viewPanel.addInfomation("", 40, 100);
			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	// 调用查询获取板件属性
	private String[] getPropertysBypartNo(TCComponentBOMLine root, String partno) throws TCException {
		String[] values = new String[2];
		// 调用系统查询，获取相关的板件
		List tcclist = Util.callStructureSearch(root, "__DFL_Find_SolutionPart", new String[] { "PARTNO" },
				new String[] { partno });
		if (tcclist != null && tcclist.size() > 0) {
			TCComponentBOMLine sol = (TCComponentBOMLine) tcclist.get(0);
			
			TCComponentItemRevision solrev3 = sol.getItemRevision();
			// values[0] = Util.getProperty(solrev3, "dfl9_part_no");// 板组3
			String bh3 = Util.getProperty(solrev3, "dfl9PartThickness");// 板厚
			if (bh3 != null && !bh3.isEmpty()) {
				values[0] = format.format(new BigDecimal(bh3.toString()));
			} else {
				values[0] = bh3;
			}
			values[1] = Util.getProperty(solrev3, "dfl9PartMaterial");// 材质
		}

		return values;
	}

	// 获取焊点的坐标（x,y,z）
	private Double[] getXYZ(String xform) {
		// TODO Auto-generated method stub
		Double[] values = new Double[] { 0.0, 0.0, 0.0 };
		String[] array = xform.split(" ");
		if (array != null && array.length == 16) {
			values[0] = Double.valueOf(array[12]);
			values[1] = Double.valueOf(array[13]);
			values[2] = Double.valueOf(array[14]);
		}
		return values;
	}

	// 根据模板创建Excel空模板
	public static XSSFWorkbook creatXSSFWorkbook(InputStream input) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);
			XSSFSheet sheet1 = book.getSheetAt(3);
			//////////// 设置分组显示上方/下方
//			sheet1.setRowSumsBelow(false);
//			sheet1.setRowSumsRight(false);
//			sheet1.setRowSumsBelow(false);
//			sheet1.setRowSumsRight(false);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;

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

		if(Style!=null) {
			cell.setCellStyle(Style);
		}	
	}

	private Comparator getComParatorBypartno() {
		// TODO Auto-generated method stub
		Comparator comparator = new Comparator() {

			public int compare(Object obj, Object obj1) {
				WeldPointInfo comp1 = (WeldPointInfo) obj;
				WeldPointInfo comp2 = (WeldPointInfo) obj1;

				String d1 = "";
				String d2 = "";
				if (comp1.getPartno() != null && !comp1.getPartno().isEmpty()) {
					d1 = comp1.getPartno();
				}
				if (comp2.getPartno() != null && !comp2.getPartno().isEmpty()) {
					d2 = comp2.getPartno();
				}
				return d1.compareTo(d2);
			}
		};

		return comparator;
	}

}
