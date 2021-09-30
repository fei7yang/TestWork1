package com.dfl.report.WeldingEquipmentEst;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.WeldPointBoardInformation;
import com.dfl.report.WeldingEstablishment.WeldingEstablishmentOp;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCSession;

public class WeldingEquipmentEstOp {

	private String reportname;
	private TCComponent savefolder;
	private TCSession session;
	private InterfaceAIFComponent[] aifComponents;
	private static Logger logger = Logger.getLogger(WeldingEstablishmentOp.class);
	private ArrayList weldlist = new ArrayList(); // 焊点集合
	private ArrayList discretelist = new ArrayList(); // 点焊工序集合
	private Map<String, String> map = new HashMap<String, String>();
	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMddHH");// 设置日期格式
	private DecimalFormat format = new DecimalFormat("0.0");
	private TCComponentBOMLine root;
	private List<WeldPointBoardInformation> baseinfolist;


	public WeldingEquipmentEstOp(TCSession session, InterfaceAIFComponent[] aifComponents, String reportname,
			TCComponent savefolder, List<WeldPointBoardInformation> baseinfolist) throws TCException {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.reportname = reportname;
		this.savefolder = savefolder;
		this.aifComponents = aifComponents;
		this.baseinfolist = baseinfolist;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub
		// 显示进度输出窗口
		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
		viewPanel.setVisible(true);

		// 获取首选项定义的Note属性
		TCPreferenceService ts = session.getPreferenceService();
		if (!ts.isDefinitionExistForPreference("B8_WeldFeasibilityReport")) {
			viewPanel.addInfomation("错误：首选项B8_WeldFeasibilityReport未定义", 100, 100);
			logger.error("错误：首选项B8_WeldFeasibilityReport未定义");
			return;
		}
		viewPanel.addInfomation("正在获取模板...\n", 10, 100);
		InputStream inputStream = Util.getReportTempByprefercen(session, "B8_WeldFeasibilityReport", 2);
		if (inputStream == null) {
			viewPanel.addInfomation("焊接设备成立性报表模板不存在！", 100, 100);
			logger.error("焊接设备成立性报表模板不存在！");
			return;
		}
		// 取BBOM顶层
		TCComponentBOMLine bl = (TCComponentBOMLine) aifComponents[0];
		root = bl.window().getTopBOMLine();

		// 工厂名称
		String factoryname = "工厂：";
		String vehicle = Util.getProperty(root, "bl_rev_project_ids");// 基本车型
		String BBOMname = Util.getProperty(root, "bl_rev_object_name");
		String[] BBOMnames = BBOMname.split("_");
		String state = "";
		if (BBOMnames != null && BBOMnames.length > 2) {
			vehicle = BBOMnames[1];
			state = BBOMnames[BBOMnames.length - 1];
			String factory = BBOMnames[2];
			if(factory.length()>2) {
				factoryname = factoryname + factory.substring(0, 3);
			}
		}

		// 获取关联的BOe
//		TCComponent[] boelist = root.getItemRevision().getRelatedComponents("IMAN_MEWorkArea");
//		if (boelist != null && boelist.length > 0) {
//			TCComponentItemRevision boerev = (TCComponentItemRevision) boelist[0];
//			factoryname = factoryname + Util.getProperty(boerev, "object_name");
//		}
		System.out.println("工厂：" + factoryname);
		System.out.println("车型：" + vehicle);

		// 获取材质与强度对应关系
		map = getSizeRule();

		// 获取所有输出信息
		getAllDiiscreteWeldInfo(session, aifComponents);

		viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);

		XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

		writeDataToSheet(book, weldlist, discretelist, factoryname, vehicle);

		// 文件名称
		String linename = "";
		for (InterfaceAIFComponent aif : aifComponents) {
			TCComponentBOMLine aifbl = (TCComponentBOMLine) aif;
			if (linename.isEmpty()) {
				linename = Util.getProperty(aifbl, "bl_rev_object_name");
			} else {
				linename = linename + "&" + Util.getProperty(aifbl, "bl_rev_object_name");
			}
		}

		String date = dateformat.format(new Date());

		String procName = vehicle + "_焊接设备成立性一元表_" + reportname + "(" + linename + ")_" + state + "_" + date + "时";
		String filename = Util.formatString(procName);
		NewOutputDataToExcel.exportFile(book, filename);

		viewPanel.addInfomation("", 80, 100);

		Util.saveFilesToFolder(session, savefolder, procName, filename, "B8_BIWProcDoc", "AT");

		viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！", 100, 100);
	}

	private void writeDataToSheet(XSSFWorkbook book, ArrayList weldlist, ArrayList discretelist, String factoryname,
			String vehicle) {
		// TODO Auto-generated method stub

		// Outputsheet页
		XSSFSheet sheet = book.getSheetAt(0);
		// 设置字体颜色
		Font font = book.createFont();
		// font.setColor((short) 12);// 红色字体
		font.setFontHeightInPoints((short) 11);
		XSSFCellStyle style = book.createCellStyle();
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style.setFont(font);

		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style2.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style2.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style2.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_BOTTOM);
		style2.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		style2.setFont(font);

		for (int i = 0; i < weldlist.size(); i++) {
			String[] values = (String[]) weldlist.get(i);
			setStringCellAndStyle(sheet, Integer.toString(i + 1), 5 + i, 1, style, 10);
			for (int j = 0; j < values.length; j++) {
				setStringCellAndStyle(sheet, values[j], 5 + i, 2 + j, style, Cell.CELL_TYPE_STRING);
			}
		}
		//设置自动列宽
        for (int i = 1; i < 77; i++) {
        	sheet.autoSizeColumn(i);
        }
        // 处理中文不能自动调整列宽的问题
        this.setSizeColumn(sheet, 77);
        
		// JudgeTrans sheet页
		XSSFSheet sheet2 = book.getSheetAt(1);
		if (sheet2 == null) {
			return;
		}

		setStringCellAndStyle(sheet2, factoryname, 0, 3, style2, Cell.CELL_TYPE_STRING);// 工厂
		setStringCellAndStyle(sheet2, vehicle, 0, 8, style2, Cell.CELL_TYPE_STRING);// 车型

		for (int i = 0; i < discretelist.size(); i++) {
			String[] values = (String[]) discretelist.get(i);
			for (int j = 0; j < values.length; j++) {
				setStringCellAndStyle(sheet2, values[j], 3 + i, 1 + j, style, Cell.CELL_TYPE_STRING);
			}
		}
//		//设置自动列宽
//        for (int i = 1; i < 61; i++) {
//        	sheet.autoSizeColumn(i);
//        }
//        // 处理中文不能自动调整列宽的问题
//        this.setSizeColumn(sheet, 61);
	}

	// 自适应宽度(中文支持)
    private void setSizeColumn(XSSFSheet sheet, int size) {
        for (int columnNum = 1; columnNum < size; columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                XSSFRow currentRow;
                //当前行未被使用过
                if (sheet.getRow(rowNum) == null) {
                    currentRow = sheet.createRow(rowNum);
                } else {
                    currentRow = sheet.getRow(rowNum);
                }
 
                if (currentRow.getCell(columnNum) != null) {
                    XSSFCell currentCell = currentRow.getCell(columnNum);
                    if (currentCell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        int length = currentCell.getStringCellValue().getBytes().length;
                        if (columnWidth < length) {
                            columnWidth = length;
                        }
                    }
                }
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256);
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
			if(Util.isNumber(value) && cellIndex!=2) {
				if(value.contains(".")) {
					celltype = 11;
				}else {
					celltype = 10;
				}
			}
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

	// 获取所有输出信息
	private void getAllDiiscreteWeldInfo(TCSession session2, InterfaceAIFComponent[] aifComponents2)
			throws TCException {
		// TODO Auto-generated method stub
		// 定义一个Map用于判断是否为相同板件，避免重复查询
		Map<String, String[]> partMap = new HashMap<String, String[]>();
		for (int i = 0; i < aifComponents.length; i++) {
			TCComponentBOMLine parent = (TCComponentBOMLine) aifComponents[i];

			// 根据BBOM查询所有的点焊工序
			String typename = Util.getObjectDisplayName(session, "B8_BIWDiscreteOP");
			String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
			String[] values = new String[] { typename, typename };

			ArrayList partList = Util.searchBOMLine(parent, "OR", propertys, "==", values);
			System.out.println("包含的点焊工序：" + partList.toString());
			
			if (partList != null && partList.size() > 0) {			
				for (int j = 0; j < partList.size(); j++) {
					double maxRecomWeldForce = 0;// 该工序下的最大加压力
					double minRecomWeldForce = 99999999;// 该工序下的最小加压力
					double maxCurrent = 0;//必要熔接电流  13 16 19
					
					String[] disrevalue = new String[73];
					TCComponentBOMLine dhbl = (TCComponentBOMLine) partList.get(j);
					disrevalue[0] = Util.getProperty(dhbl.parent().parent(), "bl_rev_object_name");
					disrevalue[2] = Util.getProperty(dhbl, "bl_rev_object_name");
					disrevalue[1] = Util.getProperty(dhbl.parent(), "bl_rev_object_name") + "(" + disrevalue[2] + ")";
				
					// 焊点
					ArrayList weld = Util.getChildrenByBOMLine(dhbl, "WeldPointRevision");
					if (weld != null && weld.size() > 0) {
						disrevalue[3] = Integer.toString(weld.size());
						for (int k = 0; k < weld.size(); k++) {
							String[] value = new String[75];
							int boradnum = 0; // 用于判断板层数
							TCComponentBOMLine bl = (TCComponentBOMLine) weld.get(k);
							// 工程名称 产线名称
							value[71] = disrevalue[0];
							value[73] = disrevalue[2];
							value[72] = disrevalue[1];

							TCComponentItemRevision rev = bl.getItemRevision();
							// 获取参数值
							TCComponent[] comps = { rev };
							String[] properties = { "object_name", "b8_ImportantLevel", "object_name",
									"b8_Weld_i_value", "b8_Weld_j_value", "b8_Weld_k_value", "b8_ConnPart1",
									"b8_ConnPart2", "b8_ConnPart3", "b8_ConnPart4", "b8_RecomWeldForce", "b8_RiseTime",
									"b8_CurrentTime1", "b8_Current1", "b8_Cool1", "b8_CurrentTime2", "b8_Current2",
									"b8_Cool2", "b8_CurrentTime3", "b8_Current3", "b8_KeepTime",
									"b8_CurrentSerie_Nissan" };
							String[][] parastr = Util.getAllProperties(session2, comps, properties);
							String[] provalues = parastr[0];
							value[0] = provalues[0];// 焊点编号
							value[1] = provalues[10];// 標準条件N
							String tempRecomWeldForce = provalues[10].replace("N", "");
							if (Util.isNumber(tempRecomWeldForce)) 
							{
								double tempre = Double.parseDouble(tempRecomWeldForce);
								if (maxRecomWeldForce < tempre) {
									maxRecomWeldForce = tempre;
								}
								if (minRecomWeldForce > tempre) {
									minRecomWeldForce = tempre;
								}
							}
							String tempCurrent1 = provalues[13].replace("KA", "");
							if (Util.isNumber(tempCurrent1)) 
							{
								double temcur = Double.parseDouble(tempCurrent1);
								if (maxCurrent < temcur) {
									maxCurrent = temcur;
								}
							}
							String tempCurrent2 = provalues[16].replace("KA", "");
							if (Util.isNumber(tempCurrent2)) 
							{
								double temcur = Double.parseDouble(tempCurrent2);
								if (maxCurrent < temcur) {
									maxCurrent = temcur;
								}
							}
							String tempCurrent3 = provalues[19].replace("KA", "");
							if (Util.isNumber(tempCurrent3)) 
							{
								double temcur = Double.parseDouble(tempCurrent3);
								if (maxCurrent < temcur) {
									maxCurrent = temcur;
								}
							}
							// A表 逻辑待确定
							System.out.println("电流：" + provalues[21]);
							value[6] = provalues[21];
							value[9] = provalues[11];// 上升时间
							value[10] = provalues[12];
							value[11] = provalues[13];
							value[12] = provalues[14];
							value[13] = provalues[15];
							value[14] = provalues[16];
							value[15] = provalues[17];
							value[16] = provalues[18];
							value[17] = provalues[19];
							value[18] = provalues[20];

							value[20] = provalues[1];
							;// 重要度b8_ImportantLevel
								// 获取x,y,z坐标
							String xform = Util.getProperty(bl, "bl_plmxml_abs_xform");// 绝对变换矩阵
							Double[] xyzArray = getXYZ(xform);
							Double x = xyzArray[0];
							Double y = xyzArray[1];
							Double z = xyzArray[2];
							value[21] = format.format(new BigDecimal(x.toString()));// X坐标
							value[22] = format.format(new BigDecimal(y.toString()));// Y坐标
							value[23] = format.format(new BigDecimal(z.toString()));// Z坐标

							// 法线取4位小数
							if (Util.isNumber(provalues[3])) {
								Double cha = Double.parseDouble(provalues[3]);
								BigDecimal df = new BigDecimal(cha);
								BigDecimal bdvalue = df.setScale(3, BigDecimal.ROUND_HALF_UP);
								value[24] = bdvalue.toString();// L
								System.out.println("法线输出结果：" + bdvalue.toString());
							} else {
								value[24] = provalues[3];// L
							}
							if (Util.isNumber(provalues[4])) {
								Double cha = Double.parseDouble(provalues[4]);
								BigDecimal df = new BigDecimal(cha);
								BigDecimal bdvalue = df.setScale(3, BigDecimal.ROUND_HALF_UP);
								value[25] = bdvalue.toString();// L
							} else {
								value[25] = provalues[4];// L
							}
							if (Util.isNumber(provalues[5])) {
								Double cha = Double.parseDouble(provalues[5]);
								BigDecimal df = new BigDecimal(cha);
								BigDecimal bdvalue = df.setScale(3, BigDecimal.ROUND_HALF_UP);
								value[26] = bdvalue.toString();// L
							} else {
								value[26] = provalues[5];// L
							}
							// 强度用来判断是否为1.2g高强材
							String strength1 = "";
							String strength2 = "";
							String strength3 = "";
							
//							// 获取板件1
//							String cp1 = "";
//							// 获取板件2
//							String cp2 = "";
//							// 获取板件3
//							String cp3 = "";
//							// 获取板件4
//							String cp4 = "";
//							// 获取板件 改为取连接到属性 bl_connected_lines
//							String conlines = Util.getProperty(bl, "bl_connected_lines");
//							if(conlines!=null && !conlines.isEmpty()) {
//								String[] strValues = conlines.split(",");
//								if(strValues.length == 1) {
//									String[] strcp1 = strValues[0].split("/");
//									cp1 = strcp1[0].trim();
//								}else if(strValues.length == 2) {
//									String[] strcp1 = strValues[0].split("/");
//									cp1 = strcp1[0].trim();
//									String[] strcp2 = strValues[1].split("/");
//									cp2 = strcp2[0].trim();
//								}else if(strValues.length == 3) {
//									String[] strcp1 = strValues[0].split("/");
//									cp1 = strcp1[0].trim();
//									String[] strcp2 = strValues[1].split("/");
//									cp2 = strcp2[0].trim();
//									String[] strcp3 = strValues[2].split("/");
//									cp3 = strcp3[0].trim();
//								}else {
//									String[] strcp1 = strValues[0].split("/");
//									cp1 = strcp1[0].trim();
//									String[] strcp2 = strValues[1].split("/");
//									cp2 = strcp2[0].trim();
//									String[] strcp3 = strValues[2].split("/");
//									cp3 = strcp3[0].trim();
//									String[] strcp4 = strValues[3].split("/");
//									cp4 = strcp4[0].trim();
//								}
//							}	
//							if (cp1 != null && !cp1.equals("")) {
//								String[] strvalue;
//								// 调用系统查询，获取相关的板件
//								if (partMap.containsKey(cp1)) {
//									strvalue = partMap.get(cp1);
//								} else {
//									// 调用系统查询，获取相关的板件
//									strvalue = getPropertysBypartNo(root, cp1);
//									partMap.put(cp1, strvalue);
//								}
//								value[31] = cp1;
//								value[39] = strvalue[0];// 材质
//								strength1 = strvalue[1];// 强度
//								value[35] = strvalue[2];// 板厚
//								value[27] = strvalue[3];// 零件名称
//								boradnum++;
//							}
//							
//							if (cp2 != null && !cp2.equals("")) {
//								String[] strvalue;
//								// 调用系统查询，获取相关的板件
//								if (partMap.containsKey(cp2)) {
//									strvalue = partMap.get(cp2);
//								} else {
//									// 调用系统查询，获取相关的板件
//									strvalue = getPropertysBypartNo(root, cp2);
//									partMap.put(cp2, strvalue);
//								}
//								value[32] = cp2;
//								value[40] = strvalue[0];
//								strength2 = strvalue[1];
//								value[36] = strvalue[2];
//								value[28] = strvalue[3];// 零件名称
//								boradnum++;
//							}
//							
//							if (cp3 != null && !cp3.equals("")) {
//								String[] strvalue;
//								// 调用系统查询，获取相关的板件
//								if (partMap.containsKey(cp3)) {
//									strvalue = partMap.get(cp3);
//								} else {
//									// 调用系统查询，获取相关的板件
//									strvalue = getPropertysBypartNo(root, cp3);
//									partMap.put(cp3, strvalue);
//								}
//								value[33] = cp3;
//								value[41] = strvalue[0];
//								strength3 = strvalue[1];
//								value[37] = strvalue[2];
//								value[29] = strvalue[3];// 零件名称
//								boradnum++;
//							}
//							
//							if (cp4 != null && !cp4.equals("")) {
//								String[] strvalue;
//								// 调用系统查询，获取相关的板件
//								if (partMap.containsKey(cp4)) {
//									strvalue = partMap.get(cp4);
//								} else {
//									// 调用系统查询，获取相关的板件
//									strvalue = getPropertysBypartNo(root, cp4);
//									partMap.put(cp4, strvalue);
//								}
//								value[34] = cp4;
//								value[42] = strvalue[0];
//								// value[12] = strvalue[1];
//								value[38] = strvalue[2];
//								value[30] = strvalue[3];// 零件名称
//								boradnum++;
//							}
							//改为从基本信息表取焊点关联的板件信息
							for(WeldPointBoardInformation weldpoint: baseinfolist)
							{
								if(value[0].equals(weldpoint.getWeldno()))
								{
									String cp1 = weldpoint.getPartNo1();
									if (cp1 != null && !cp1.equals("")) {
										value[31] = cp1;
										value[39] = weldpoint.getPartmaterial1();// 材质
										strength1 = weldpoint.getStrength1();// 强度
										value[35] = weldpoint.getPartthickness1();// 板厚
										value[27] = weldpoint.getBoardname1();// 零件名称
										boradnum++;
									}
									String cp2 = weldpoint.getPartNo2();
									if (cp2 != null && !cp2.equals("")) {
										value[32] = cp2;
										value[40] = weldpoint.getPartmaterial2();// 材质
										strength1 = weldpoint.getStrength2();// 强度
										value[36] = weldpoint.getPartthickness2();// 板厚
										value[28] = weldpoint.getBoardname2();// 零件名称
										boradnum++;
									}
									
									String cp3 = weldpoint.getPartNo2();
									if (cp3 != null && !cp3.equals("")) {
										value[33] = cp3;
										value[41] = weldpoint.getPartmaterial3();// 材质
										strength1 = weldpoint.getStrength3();// 强度
										value[37] = weldpoint.getPartthickness3();// 板厚
										value[29] = weldpoint.getBoardname3();// 零件名称
										boradnum++;
									}
									break;
								}
							}
							
							
							if (boradnum == 4) {
								value[62] = "是";
							}
							if (boradnum == 2) {
								value[66] = "是";
							}
							String sumthick = getSumBoradThickness(value[35], value[36], value[37]);
							String maxthick = getMaxBoradThickness(value[35], value[36], value[37]);
							String minthick = getMinBoradThickness(value[35], value[36], value[37]);

							if (boradnum == 3) {
								if (Double.parseDouble(sumthick) != 0 && Double.parseDouble(maxthick) != 0
										&& Double.parseDouble(minthick) != -1) {
									String difference = getPlateThicknessDifference(maxthick, minthick);
									value[63] = sumthick;
									value[64] = difference;
								}
							}
							// 获取基准板厚
							boolean baseflag = false;
							if ((Util.isNumber(strength1) && Double.parseDouble(strength1) == 1180)
									|| (Util.isNumber(strength2) && Double.parseDouble(strength2) == 1180)
									|| (Util.isNumber(strength3) && Double.parseDouble(strength3) == 1180)) {
								baseflag = true;
							}

							// 基準板厚与工程作业本逻辑一致
							value[19] = getBasethickness(baseflag, boradnum, value[35], value[36], value[37]);

							System.out.println(value[0] + "的基准板厚：" + value[19]);

							weldlist.add(value);
						}
					} else {
						disrevalue[3] = "0";
					}
					// 枪
					ArrayList gun = Util.getChildrenByBOMLine(dhbl, "B8_BIWGunRevision");
					if (gun != null && gun.size() > 0) 
					{
						for (int k = 0; k < gun.size(); k++) {
							TCComponentBOMLine gunbl = (TCComponentBOMLine) gun.get(k);
							if (k == 0) {
								disrevalue[4] = Util.getProperty(gunbl, "bl_rev_object_name");
//								disrevalue[58] = Util.getProperty(gunbl, "bl_B8_BIWGunRevision_b8_ElectrodeVol");
								disrevalue[63] = Util.getProperty(gunbl, "bl_B8_BIWGunRevision_b8_ElectrodeVol");
								disrevalue[70] = Util.getProperty(gunbl, "bl_B8_BIWGunRevision_b8_MaxAmp");

//								// 許容範囲
//								if (disrevalue[2].length() > 1) {
//									if (disrevalue[2].substring(0, 1).equals("R")) {
//										if (minRecomWeldForce == 99999999) {
//											minRecomWeldForce = 0;
//										}
//										double sumre = minRecomWeldForce + maxRecomWeldForce;
//										if (sumre == 0) {
//											disrevalue[56] = "";
//										} else {
//											disrevalue[56] = Double.toString(sumre);
//										}
//									} else {
//										if (maxRecomWeldForce == 0) {
//											disrevalue[56] = "";
//										} else {
//											disrevalue[56] = Double.toString(maxRecomWeldForce);
//										}
//									}
//								} else {
//									disrevalue[56] = "";
//								}
//								// 判断符号 和 判定结果
//								double allow = 0;
//								double electrode = 0;
//								if (disrevalue[56].isEmpty() && disrevalue[58].isEmpty()) {
//									disrevalue[57] = "";
//									disrevalue[59] = "";
//								} else {
//									if (Util.isNumber(disrevalue[56])) {
//										allow = Double.parseDouble(disrevalue[56]);
//									}
//									if (Util.isNumber(disrevalue[58])) {
//										electrode = Double.parseDouble(disrevalue[58]);
//									}
//									if (allow <= electrode) {
//										disrevalue[57] = "≤";
//										disrevalue[59] = "OK";
//									} else {
//										disrevalue[57] = ">";
//										disrevalue[59] = "NG";
//									}
//								}
															
								
								//人工工位/机器人焊钳最大压力与额定压力差值(N)
								if(maxRecomWeldForce == 0)
								{
									disrevalue[62] = "";
								}
								else
								{
									disrevalue[62] = BigDecimal.valueOf(maxRecomWeldForce).toString();
								}
								//获取允许范围
								String gap = getReport_Allowed_Pres_Gap();
								if(gap == null || gap.isEmpty())
								{
									gap = "1000";
								}
								disrevalue[66] = gap;
								if(Util.isNumber(gap))
								{
									if(Util.isNumber(disrevalue[63].replace("N", "")))
									{
										disrevalue[64] = BigDecimal.valueOf(maxRecomWeldForce).subtract(BigDecimal.valueOf(Double.parseDouble(disrevalue[63].replace("N", "")))).toString();
										if(Double.parseDouble(disrevalue[64]) > Double.parseDouble(gap))
										{
											disrevalue[65] = ">";
											disrevalue[67] = "NG";
										}
										else
										{
											disrevalue[65] = "≤";
											disrevalue[67] = "OK";
										}
									}
									else
									{
										disrevalue[65] = "";
										disrevalue[67] = "缺额定压力";
									}
								}
								else
								{
									disrevalue[65] = "";
									disrevalue[67] = "";
									System.out.println("首选项配置的允许范围不是数字，无法判断");

								}
								
													
								//人工工位焊钳焊接需最大压力与最小压力差值(N)
								if(!"R".equals(disrevalue[2].subSequence(0, 1)))
								{
									//X列
									if(maxCurrent == 0)
									{
										disrevalue[22] = "";
									}
									else
									{
										disrevalue[22] = BigDecimal.valueOf(maxCurrent).toString();
									}
									if(maxRecomWeldForce == 0 && minRecomWeldForce == 99999999)
									{
										disrevalue[59] = "";
										disrevalue[60] = "";
										disrevalue[61] = "";
										disrevalue[69] = "";
										disrevalue[70] = "";
										disrevalue[71] = "";
									}
									else
									{
										if(maxRecomWeldForce == 0)
										{
											disrevalue[56] = "";
										}
										else
										{
											disrevalue[56] = BigDecimal.valueOf(maxRecomWeldForce).toString();

										}
										if (minRecomWeldForce == 99999999) {
											minRecomWeldForce = 0;
											disrevalue[57] = "";
										}
										else
										{
											disrevalue[57] = BigDecimal.valueOf(minRecomWeldForce).toString();
										}
										double chare = BigDecimal.valueOf(maxRecomWeldForce).subtract(BigDecimal.valueOf(minRecomWeldForce)).doubleValue();
										if (chare == 0) {
											disrevalue[58] = "0";
										} else {
											disrevalue[58] = BigDecimal.valueOf(chare).toString();
										}
										//获取允许范围
										String Diff = getReport_Allowed_Pres_Diff();
										if(Diff == null || Diff.isEmpty())
										{
											Diff = "2000";
										}
										disrevalue[60] = Diff;
										if(Util.isNumber(Diff))
										{
											if(chare > Double.parseDouble(Diff))
											{
												disrevalue[59] = ">";
												disrevalue[61] = "NG";
											}
											else
											{
												disrevalue[59] = "≤";
												disrevalue[61] = "OK";
											}
										}
										else
										{
											disrevalue[59] = "";
											disrevalue[61] = "";
											System.out.println("首选项配置的允许范围不是数字，无法判断");
										}
										
										//BR列
										if(maxCurrent == 0)
										{
											disrevalue[68] = "";
										}
										else
										{
											disrevalue[68] = BigDecimal.valueOf(maxCurrent).toString();
										}
										if(Util.isNumber(disrevalue[70]))
										{
											if(maxCurrent > Double.parseDouble(disrevalue[70]))
											{
												disrevalue[69] = ">";
												disrevalue[71] = "NG";
											}
											else
											{
												disrevalue[69] = "≤";
												disrevalue[71] = "OK";
											}
											if(Double.parseDouble(disrevalue[70]) == 0)
											{
												disrevalue[69] = "";
												disrevalue[71] = "缺额定电流";
											}								
										}
										else 
										{
											disrevalue[69] = "";
											disrevalue[71] = "缺额定电流";
										}
									}
																										
								}
								else
								{
									disrevalue[56] = "";
									disrevalue[57] = "";
									disrevalue[58] = "";
									disrevalue[59] = "";
									disrevalue[60] = "";
									disrevalue[61] = "";
									disrevalue[68] = "";
									disrevalue[69] = "";
									disrevalue[70] = "";
									disrevalue[71] = "";
								}
								//disrevalue[67] = "缺额定压力";
								if("缺额定压力".equals(disrevalue[67]))
								{
									disrevalue[72] = "缺额定压力";
								}
								else
								{
									if("NG".equals(disrevalue[61]) || "NG".equals(disrevalue[67]) || "NG".equals(disrevalue[71]))
									{
										disrevalue[72] = "NG";
									}
									else
									{
										disrevalue[72] = "OK";
									}	
								}
															

							} else 
							{
								String[] gunvalue = new String[73];
								gunvalue[0] = disrevalue[0];
								gunvalue[2] = disrevalue[2];
								gunvalue[1] = disrevalue[1];
								gunvalue[3] = disrevalue[3];
								gunvalue[4] = Util.getProperty(gunbl, "bl_rev_object_name");
//								gunvalue[58] = Util.getProperty(gunbl, "bl_B8_BIWGunRevision_b8_ElectrodeVol");
								gunvalue[63] = Util.getProperty(gunbl, "bl_B8_BIWGunRevision_b8_ElectrodeVol");
								gunvalue[70] = Util.getProperty(gunbl, "bl_B8_BIWGunRevision_b8_MaxAmp");
//								// 許容範囲
//								if (gunvalue[2].length() > 1) {
//									if (gunvalue[2].substring(0, 1).equals("R")) {
//										if (minRecomWeldForce == 99999999) {
//											minRecomWeldForce = 0;
//										}
//										double sumre = minRecomWeldForce + maxRecomWeldForce;
//										if (sumre == 0) {
//											gunvalue[56] = "";
//										} else {
//											gunvalue[56] = Double.toString(sumre);
//										}
//									} else {
//										if (maxRecomWeldForce == 0) {
//											gunvalue[56] = "";
//										} else {
//											gunvalue[56] = Double.toString(maxRecomWeldForce);
//										}
//									}
//								} else {
//									gunvalue[56] = "";
//								}
//								// 判断符号 和 判定结果
//								double allow = 0;
//								double electrode = 0;
//								if (gunvalue[56].isEmpty() && gunvalue[58].isEmpty()) {
//									gunvalue[57] = "";
//									gunvalue[59] = "";
//								} else {
//									if (Util.isNumber(gunvalue[56])) {
//										allow = Double.parseDouble(gunvalue[56]);
//									}
//									if (Util.isNumber(gunvalue[58])) {
//										electrode = Double.parseDouble(gunvalue[58]);
//									}
//									if (allow <= electrode) {
//										gunvalue[57] = "≤";
//										gunvalue[59] = "OK";
//									} else {
//										gunvalue[57] = ">";
//										gunvalue[59] = "NG";
//									}
//								}
								
															
								//人工工位/机器人焊钳最大压力与额定压力差值(N)
								if(maxRecomWeldForce == 0)
								{
									gunvalue[62] = "";
								}
								else
								{
									gunvalue[62] = BigDecimal.valueOf(maxRecomWeldForce).toString();
								}
								//获取允许范围
								String gap = getReport_Allowed_Pres_Gap();
								if(gap == null || gap.isEmpty())
								{
									gap = "1000";
								}
								gunvalue[66] = gap;
								if(Util.isNumber(gap))
								{
									if(Util.isNumber(gunvalue[63].replace("N", "")))
									{
										disrevalue[64] = BigDecimal.valueOf(maxRecomWeldForce).subtract(BigDecimal.valueOf(Double.parseDouble(gunvalue[63].replace("N", "")))).toString();
										if(Double.parseDouble(gunvalue[64]) > Double.parseDouble(gap))
										{
											gunvalue[65] = ">";
											gunvalue[67] = "NG";
										}
										else
										{
											gunvalue[65] = "≤";
											gunvalue[67] = "OK";
										}
									}
									else
									{
										gunvalue[65] = "";
										gunvalue[67] = "缺额定压力";
									}
								}
								else
								{
									gunvalue[65] = "";
									gunvalue[67] = "";
									System.out.println("首选项配置的允许范围不是数字，无法判断");

								}
													
								//人工工位焊钳焊接需最大压力与最小压力差值(N)
								if(!"R".equals(gunvalue[2].subSequence(0, 1)))
								{
									//X列
									if(maxCurrent == 0)
									{
										gunvalue[22] = "";
									}
									else
									{
										gunvalue[22] = BigDecimal.valueOf(maxCurrent).toString();
									}
									
									if(maxRecomWeldForce == 0 && minRecomWeldForce == 99999999)
									{
										gunvalue[59] = "";
										gunvalue[60] = "";
										gunvalue[61] = "";
										gunvalue[69] = "";
										gunvalue[70] = "";
										gunvalue[71] = "";
									}
									else
									{
										if(maxRecomWeldForce == 0)
										{
											gunvalue[56] = "";
										}
										else
										{
											gunvalue[56] = BigDecimal.valueOf(maxRecomWeldForce).toString();

										}
										if (minRecomWeldForce == 99999999) 
										{
											minRecomWeldForce = 0;
											gunvalue[57] = "";
										}
										else
										{
											gunvalue[57] = BigDecimal.valueOf(minRecomWeldForce).toString();
										}
										double chare = BigDecimal.valueOf(maxRecomWeldForce).subtract(BigDecimal.valueOf(minRecomWeldForce)).doubleValue();
										if (chare == 0) {
											gunvalue[58] = "0";
										} else {
											gunvalue[58] = BigDecimal.valueOf(chare).toString();
										}
										//获取允许范围
										String Diff = getReport_Allowed_Pres_Diff();
										if(Diff == null || Diff.isEmpty())
										{
											Diff = "2000";
										}
										gunvalue[60] = Diff;
										if(Util.isNumber(Diff))
										{
											if(chare > Double.parseDouble(Diff))
											{
												gunvalue[59] = ">";
												gunvalue[61] = "NG";
											}
											else
											{
												gunvalue[59] = "≤";
												gunvalue[61] = "OK";
											}										
										}
										else
										{
											gunvalue[59] = "";
											gunvalue[61] = "";
											System.out.println("首选项配置的允许范围不是数字，无法判断");
										}
										
										//BR列
										if(maxCurrent == 0)
										{
											gunvalue[68] = "";
										}
										else
										{
											gunvalue[68] = BigDecimal.valueOf(maxCurrent).toString();
										}
										if(Util.isNumber(gunvalue[70]))
										{
											if(maxCurrent > Double.parseDouble(gunvalue[70]))
											{
												gunvalue[69] = ">";
												gunvalue[71] = "NG";
											}
											else
											{
												gunvalue[69] = "≤";
												gunvalue[71] = "OK";
											}
											if(Double.parseDouble(gunvalue[70]) == 0)
											{
												gunvalue[69] = "";
												gunvalue[71] = "缺额定电流";
											}
										}
										else
										{
											gunvalue[69] = "";
											gunvalue[71] = "缺额定电流";
										}
									}
																										
								}
								else
								{
									gunvalue[56] = "";
									gunvalue[57] = "";
									gunvalue[58] = "";
									gunvalue[59] = "";
									gunvalue[60] = "";
									gunvalue[61] = "";
									gunvalue[68] = "";
									gunvalue[69] = "";
									gunvalue[70] = "";
									gunvalue[71] = "";
								}
								if("缺额定压力".equals(gunvalue[67]))
								{
									gunvalue[72] = "缺额定压力";
								}
								else
								{
									if("NG".equals(gunvalue[61]) || "NG".equals(gunvalue[67]) || "NG".equals(gunvalue[71]))
									{
										gunvalue[72] = "NG";
									}
									else
									{
										gunvalue[72] = "OK";
									}
								}
								
								
								discretelist.add(gunvalue);
							}
						}

					} 
					else 
					{
						disrevalue[4] = "";
						disrevalue[63] = "";
						//人工工位/机器人焊钳最大压力与额定压力差值(N)
						if(maxRecomWeldForce == 0)
						{
							disrevalue[62] = "";
						}
						else
						{
							disrevalue[62] = BigDecimal.valueOf(maxRecomWeldForce).toString();
						}
						//获取允许范围
						String gap = getReport_Allowed_Pres_Gap();
						if(gap == null || gap.isEmpty())
						{
							gap = "1000";
						}
						disrevalue[66] = gap;
						if(Util.isNumber(gap))
						{
							if(Util.isNumber(disrevalue[63].replace("N", "")))
							{
								disrevalue[64] = BigDecimal.valueOf(maxRecomWeldForce).subtract(BigDecimal.valueOf(Double.parseDouble(disrevalue[63].replace("N", "")))).toString();
								if(Double.parseDouble(disrevalue[64]) > Double.parseDouble(gap))
								{
									disrevalue[65] = ">";
									disrevalue[67] = "NG";
								}
								else
								{
									disrevalue[65] = "≤";
									disrevalue[67] = "OK";
								}
							}
							else
							{
								disrevalue[65] = "";
								disrevalue[67] = "缺额定压力";
							}
						}
						else
						{
							disrevalue[65] = "";
							disrevalue[67] = "";
							System.out.println("首选项配置的允许范围不是数字，无法判断");

						}
						//人工工位焊钳焊接需最大压力与最小压力差值(N)
						if(!"R".equals(disrevalue[2].subSequence(0, 1)))
						{
							//X列
							if(maxCurrent == 0)
							{
								disrevalue[22] = "";
							}
							else
							{
								disrevalue[22] = BigDecimal.valueOf(maxCurrent).toString();
							}
							
							if(maxRecomWeldForce == 0)
							{
								disrevalue[56] = "";
							}
							else
							{
								disrevalue[56] = BigDecimal.valueOf(maxRecomWeldForce).toString();

							}
							if (minRecomWeldForce == 99999999) {
								minRecomWeldForce = 0;
								disrevalue[57] = "";
							}
							else
							{
								disrevalue[57] = BigDecimal.valueOf(minRecomWeldForce).toString();
							}
							double chare = BigDecimal.valueOf(maxRecomWeldForce).subtract(BigDecimal.valueOf(minRecomWeldForce)).doubleValue();
							if (chare == 0) {
								disrevalue[58] = "";
							} else {
								disrevalue[58] = BigDecimal.valueOf(chare).toString();
							}
							//获取允许范围
							String Diff = getReport_Allowed_Pres_Diff();
							if(Diff == null || Diff.isEmpty())
							{
								Diff = "2000";
							}
							disrevalue[60] = Diff;
							if(Util.isNumber(Diff))
							{
								if(chare > Double.parseDouble(Diff))
								{
									disrevalue[59] = ">";
									disrevalue[61] = "NG";
								}
								else
								{
									disrevalue[59] = "≤";
									disrevalue[61] = "OK";
								}
							}
							else
							{
								disrevalue[59] = "";
								disrevalue[61] = "";
								System.out.println("首选项配置的允许范围不是数字，无法判断");
							}
							
							//BR列
							if(maxCurrent == 0)
							{
								disrevalue[68] = "";
							}
							else
							{
								disrevalue[68] = BigDecimal.valueOf(maxCurrent).toString();
							}
							if(Util.isNumber(disrevalue[70]))
							{
								if(maxCurrent > Double.parseDouble(disrevalue[70]))
								{
									disrevalue[69] = ">";
									disrevalue[71] = "NG";
								}
								else
								{
									disrevalue[69] = "≤";
									disrevalue[71] = "OK";
								}
							}
							
						}
						disrevalue[72] = "缺额定压力";
					}

					discretelist.add(disrevalue);
				}
			}
//			System.out.println("最大加压力:" + maxRecomWeldForce);
//			System.out.println("最小加压力:" + minRecomWeldForce);
		}
	}

	/*
	 * 获取基准板厚
	 */
	private String getBasethickness(boolean flag, int boradnum, String partthickness1, String partthickness2,
			String partthickness3) {
		// 基准板厚
		String basethickness = "";
		if (flag) {
			basethickness = getMinBoradThickness(partthickness1, partthickness2, partthickness3);
		} else {
			// 3层板取平均值四舍五入
			if (boradnum == 3 || boradnum == 4) {
				if ((partthickness1 != null && !partthickness1.isEmpty())
						&& (partthickness2 != null && !partthickness2.isEmpty())
						&& (partthickness3 != null && !partthickness3.isEmpty())) {
					double totalsum = Double.parseDouble(partthickness1) + Double.parseDouble(partthickness2)
							+ Double.parseDouble(partthickness3);
					double basenum = totalsum / 3;
					BigDecimal bd = new BigDecimal(basenum);
					BigDecimal bdvalue = bd.setScale(1, BigDecimal.ROUND_HALF_UP);
					basethickness = bdvalue.toString();
				}
			} else if (boradnum == 2) { // 2层板取薄板
				if (partthickness1 == null || partthickness1.isEmpty()) {
					if ((partthickness2 != null && !partthickness2.isEmpty())
							&& (partthickness3 != null && !partthickness3.isEmpty())) {

						if (Double.parseDouble(partthickness2) > Double.parseDouble(partthickness3)) {
							basethickness = partthickness3;
						} else {
							basethickness = partthickness2;
						}
					}
				} else if (partthickness2 == null || partthickness2.isEmpty()) {
					if ((partthickness1 != null && !partthickness1.isEmpty())
							&& (partthickness3 != null && !partthickness3.isEmpty())) {
						if (Double.parseDouble(partthickness1) > Double.parseDouble(partthickness3)) {
							basethickness = partthickness3;
						} else {
							basethickness = partthickness1;
						}
					}
				} else {
					if ((partthickness1 != null && !partthickness1.isEmpty())
							&& (partthickness2 != null && !partthickness2.isEmpty())) {
						if (Double.parseDouble(partthickness1) > Double.parseDouble(partthickness2)) {
							basethickness = partthickness2;
						} else {
							basethickness = partthickness1;
						}
					}
				}
			} else if (boradnum == 1) {
				if (partthickness1 != null && !partthickness1.isEmpty()) {
					basethickness = partthickness1;
				} else if (partthickness2 != null && !partthickness2.isEmpty()) {
					basethickness = partthickness2;
				} else {
					basethickness = partthickness3;
				}
			} else {
			}

		}
		return basethickness;
	}

	/*
	 * 总板厚
	 */
	private String getSumBoradThickness(String thickness1, String thickness2, String thickness3) {
		String sumth = "";
		double sum = getDouble(thickness1) + getDouble(thickness2) + getDouble(thickness3);
		sumth = String.format("%.2f", sum);
		return sumth;
	}

	/*
	 * 最大板厚
	 */
	private String getMaxBoradThickness(String thickness1, String thickness2, String thickness3) {
		String maxth = "";
		double max = 0;
		double thk1 = getDouble(thickness1);
		double thk2 = getDouble(thickness2);
		double thk3 = getDouble(thickness3);

		if (thk1 > thk2) {
			if (thk1 > thk3) {
				max = thk1;
			} else {
				max = thk3;
			}
		} else {
			if (thk2 > thk3) {
				max = thk2;
			} else {
				max = thk3;
			}
		}

		maxth = Double.toString(max);

		return maxth;
	}

	/*
	 * 最小板厚
	 */
	private String getMinBoradThickness(String thickness1, String thickness2, String thickness3) {
		String minth = "";
		double min = 0;
		double thk1 = getDouble(thickness1);
		double thk2 = getDouble(thickness2);
		double thk3 = getDouble(thickness3);

		if (thk1 == 0) {
			thk1 = 99999999999.0;
		}
		if (thk2 == 0) {
			thk2 = 99999999999.0;
		}
		if (thk3 == 0) {
			thk3 = 99999999999.0;
		}

		if (thk1 < thk2) {
			if (thk1 < thk3) {
				min = thk1;
			} else {
				min = thk3;
			}
		} else {
			if (thk2 < thk3) {
				min = thk2;
			} else {
				min = thk3;
			}
		}
		if (min == 99999999999.0) {
			min = -1;
		}

		minth = Double.toString(min);

		return minth;
	}

	/*
	 * 板厚差
	 */
	private String getPlateThicknessDifference(String maxthick, String minthick) {
		String difference = "";
		double cha = getDouble(maxthick) / getDouble(minthick);
		BigDecimal df = new BigDecimal(cha);
		BigDecimal bdvalue = df.setScale(1, BigDecimal.ROUND_HALF_UP);
		difference = bdvalue.toString();
		return difference;
	}

	/*
	 * 字符转换成整数
	 */
	private double getDouble(String str) {
		double num = 0;
		if (Util.isNumber(str)) {
			num = Double.parseDouble(str);
		}
		return num;
	}

	// 调用查询获取板件属性
	private String[] getPropertysBypartNo(TCComponentBOMLine parrent, String partno) throws TCException {
		String[] values = new String[4];
		// 调用系统查询，获取相关的板件
		List tcclist = Util.callStructureSearch(parrent, "__DFL_Find_SolutionPart", new String[] { "PARTNO" },
				new String[] { partno });
		if (tcclist != null && tcclist.size() > 0) {
			TCComponentBOMLine sol = (TCComponentBOMLine) tcclist.get(0);
			TCComponentItemRevision solrev3 = sol.getItemRevision();
			// values[0] = Util.getProperty(solrev3, "dfl9_part_no");// 板组3
			String bh3 = Util.getProperty(solrev3, "dfl9PartThickness");// 板厚
			if (bh3 != null && !bh3.isEmpty()) {
				values[2] = format.format(new BigDecimal(bh3.toString()));
			} else {
				values[2] = bh3;
			}
			values[0] = Util.getProperty(solrev3, "dfl9PartMaterial");// 材质
			if (map.containsKey(values[0])) {
				values[1] = map.get(values[0]); // 强度
			} else {
				values[1] = ""; // 强度
			}
			//values[3] = Util.getProperty(solrev3, "object_name");// 零件名称			
			String dfl9_CADObjectName = Util.getProperty(sol, "bl_DFL9SolItmPartRevision_dfl9_CADObjectName");
			System.out.println("零件名称bl_DFL9SolItmPartRevision_dfl9_CADObjectName：" + dfl9_CADObjectName);
			System.out.println("零件名称dfl9_CADObjectName：" + Util.getProperty(solrev3, "dfl9_CADObjectName"));
			if(dfl9_CADObjectName!=null && !dfl9_CADObjectName.isEmpty())
			{
				dfl9_CADObjectName = Util.getProperty(solrev3, "dfl9_CADObjectName");// 零件名称		
			}
			values[3] = dfl9_CADObjectName;// 零件名称
		}

		return values;
	}

	// 根据材质获取对应的强度
	private HashMap<String, String> getSizeRule() {
		HashMap<String, String> rule = new HashMap<String, String>();
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
							FileInputStream inputStream = new FileInputStream(file);
							if (type.equals("MSExcel")) {
								workbook = new HSSFWorkbook(inputStream);
								rule = parseCoverExcel(workbook);
							}
							if (type.equals("MSExcelX")) {
								workbook = new XSSFWorkbook(inputStream);
								rule = parseCoverExcel(workbook);
							}
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
	
	// 根据B8_WeldFeasibilityReport_Allowed_Pres_Diff的允许范围
	private String getReport_Allowed_Pres_Diff() {
		String diff = "";
		try {
			TCPreferenceService preferenceService = session.getPreferenceService();
			String str = preferenceService.getPreferenceDescription("B8_WeldFeasibilityReport_Allowed_Pres_Diff");
			if (str != null) {
				String value = preferenceService.getStringValue("B8_WeldFeasibilityReport_Allowed_Pres_Diff");
				if (value != null) {
					diff = value;
				}
			}
			return diff;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return diff;
	}
	// 根据B8_WeldFeasibilityReport_Allowed_Pres_Gap的允许范围
	private String getReport_Allowed_Pres_Gap() {
		String diff = "";
		try {
			TCPreferenceService preferenceService = session.getPreferenceService();
			String str = preferenceService.getPreferenceDescription("B8_WeldFeasibilityReport_Allowed_Pres_Gap");
			if (str != null) {
				String value = preferenceService.getStringValue("B8_WeldFeasibilityReport_Allowed_Pres_Gap");
				if (value != null) {
					diff = value;
				}
			}
			return diff;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return diff;
	}

	private static HashMap<String, String> parseCoverExcel(Workbook workbook) {
		// TODO Auto-generated method stub
		HashMap<String, String> rule = new HashMap<String, String>();
		// 解析sheet

		Sheet sheet = workbook.getSheetAt(0);
		// 校验sheet是否合法
		if (sheet == null) {
			return null;
		}
		// 获取第一行数据
		int firstRowNum = sheet.getFirstRowNum();
		Row firstRow = (Row) sheet.getRow(firstRowNum);
		if (null == firstRow) {
			logger.warn("解析Excel失败，在第一行没有读取到任何数据！");
		}

		// 解析每一行的数据，构造数据对象
		int rowStart = firstRowNum + 1;
		int rowEnd = sheet.getPhysicalNumberOfRows();
		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = (Row) sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			String[] resultData = convertRowToCoverData(row);
			if (null == resultData) {
				logger.warn("第 " + row.getRowNum() + "行数据不合法，已忽略！");
				continue;
			}
			if (resultData[0] != null && !resultData[0].isEmpty()) {
				rule.put(resultData[0], resultData[1]);
			}
		}

		return rule;
	}

	private static String[] convertRowToCoverData(Row row) {
		// TODO Auto-generated method stub
		String[] value = new String[2];
		Cell cell;
		// 材质
		cell = row.getCell(1);
		if (cell != null) {
			String partno = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
			value[0] = partno.trim();
		}
		// 强度
		cell = row.getCell(2);
		if (cell != null) {
			String parttype = convertCellValueToString(cell, Cell.CELL_TYPE_STRING);
			value[1] = parttype.trim();
		}
		return value;
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
			cell.setCellType(Cell.CELL_TYPE_STRING);
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
}
