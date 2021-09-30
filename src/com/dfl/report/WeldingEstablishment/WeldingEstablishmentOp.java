package com.dfl.report.WeldingEstablishment;

import java.awt.Color;
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
import org.apache.poi.ss.usermodel.CellStyle;
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

import com.dfl.report.ExcelReader.baseinfoExcelReader;
import com.dfl.report.dealparameter.DealParameterHandler;
import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentTcFile;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCPreferenceService;
import com.teamcenter.rac.kernel.TCProperty;
import com.teamcenter.rac.kernel.TCSession;

public class WeldingEstablishmentOp {

	private AbstractAIFUIApplication app;
	private String reportname;
	private TCComponent savefolder;
	private TCSession session;
	private InterfaceAIFComponent[] aifComponents;
	private static Logger logger = Logger.getLogger(WeldingEstablishmentOp.class);
	private ArrayList weldlist = new ArrayList();
	private DecimalFormat format = new DecimalFormat("0.0");
	private Map<String, String> map = new HashMap<String, String>();
	SimpleDateFormat dateformat = new SimpleDateFormat("yyyyMMddHH");// 设置日期格式
	private TCComponentBOMLine root;

	public WeldingEstablishmentOp(TCSession session, InterfaceAIFComponent[] aifComponents, String reportname,
			TCComponent savefolder) throws TCException {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.reportname = reportname;
		this.savefolder = savefolder;
		this.aifComponents = aifComponents;
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
		InputStream inputStream = Util.getReportTempByprefercen(session, "B8_WeldFeasibilityReport", 1);
		if (inputStream == null) {
			viewPanel.addInfomation("焊接成立性报表模板不存在！", 100, 100);
			logger.error("焊接成立性报表模板不存在！");
			return;
		}
		// 取BBOM顶层
		TCComponentBOMLine bl = (TCComponentBOMLine) aifComponents[0];
		root = bl.window().getTopBOMLine();

		// 获取材质与强度对应关系
		map = getSizeRule();

		viewPanel.addInfomation("开始输出报表...\n", 40, 100);

		// 获取所有焊点信息
		getAllWeldPoint(session, aifComponents);

		viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);

		XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);

		writeDataToSheet(book, weldlist);

		// 文件名称
		// 文件名称
		String functionname = "";
		for (InterfaceAIFComponent aif : aifComponents) {
			TCComponentBOMLine aifbl = (TCComponentBOMLine) aif;
			if (functionname.isEmpty()) {
				functionname = Util.getProperty(aifbl, "bl_rev_object_name");
			} else {
				functionname = functionname + "&" + Util.getProperty(aifbl, "bl_rev_object_name");
			}
		}

		String vehicle = Util.getProperty(root, "bl_rev_project_ids");// 基本车型
		String BBOMname = Util.getProperty(root, "bl_rev_object_name");
		String[] BBOMnames = BBOMname.split("_");
		String state = "";
		if (BBOMnames != null && BBOMnames.length > 1) {
			vehicle = BBOMnames[1];
			state = BBOMnames[BBOMnames.length - 1];
		}
		String date = dateformat.format(new Date());

		String procName = vehicle + "_焊接成立性一元表_" + reportname + "(" + functionname + ")_" + state + "_" + date + "时";
		String filename = Util.formatString(procName);
		NewOutputDataToExcel.exportFile(book, filename);

		viewPanel.addInfomation("", 80, 100);

		Util.saveFilesToFolder(session, savefolder, procName, filename, "B8_BIWProcDoc", "AS");

		viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！", 100, 100);

	}

	// 写t数据
	public static void writeDataToSheet(XSSFWorkbook book, ArrayList weldlist) {
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
		Font font2 = book.createFont();
		font2.setColor((short) 2);// 红色字体
		font2.setFontHeightInPoints((short) 11);
		XSSFCellStyle style2 = book.createCellStyle();
		style2.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style2.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style2.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style2.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style2.setFont(font2);
		// 绿色背景
		XSSFCellStyle style3 = book.createCellStyle();
		style3.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);
		// 红色背景
		XSSFCellStyle style4 = book.createCellStyle();
		style4.setFillForegroundColor(IndexedColors.RED.getIndex());
		style4.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font);

		for (int i = 0; i < weldlist.size(); i++) {
			String[] values = (String[]) weldlist.get(i);
			setStringCellAndStyle(sheet, Integer.toString(i + 1), 1 + i, 0, style, 10);
			for (int j = 0; j < values.length; j++) {
				String value = values[j];
				if (value != null && value.equals("OK")) {
					setStringCellAndStyle(sheet, values[j], 1 + i, 1 + j, style3, Cell.CELL_TYPE_STRING);
				} else if (value != null && value.equals("NG")) {
					setStringCellAndStyle(sheet, values[j], 1 + i, 1 + j, style4, Cell.CELL_TYPE_STRING);
				} else if ((j == 3 && Util.isNumber(values[4]) && Double.parseDouble(values[4]) == 1350)
						|| (j == 7 && Util.isNumber(values[8]) && Double.parseDouble(values[8]) == 1350)
						|| (j == 11 && Util.isNumber(values[12]) && Double.parseDouble(values[12]) == 1350)) {
					System.out.println("测试是否进入该判断，定位背景色不显示的问题");
					setStringCellAndStyle(sheet, values[j], 1 + i, 1 + j, style4, Cell.CELL_TYPE_STRING);
				} else if (Util.isNumber(value) && Double.parseDouble(value) == 1350) {
					setStringCellAndStyle(sheet, "热成型", 1 + i, 1 + j, style, Cell.CELL_TYPE_STRING);
				} else if (j == 20 && values[22] == "OK(中间为最薄板)") {
					setStringCellAndStyle(sheet, values[j], 1 + i, 1 + j, style2, Cell.CELL_TYPE_STRING);
				} else {
					setStringCellAndStyle(sheet, values[j], 1 + i, 1 + j, style, Cell.CELL_TYPE_STRING);
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
			if (Util.isNumber(value) && cellIndex != 1) {
				if (value.contains(".")) {
					celltype = 11;
				} else {
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

	/*
	 * 获取所有焊点信息
	 */
	private void getAllWeldPoint(TCSession session, InterfaceAIFComponent[] aifComponents) throws TCException {
		// 定义一个Map用于判断是否为相同板件，避免重复查询
		Map<String, String[]> partMap = new HashMap<String, String[]>();
		for (int i = 0; i < aifComponents.length; i++) {
			TCComponentBOMLine parent = (TCComponentBOMLine) aifComponents[i];
			// 根据BBOM查询所有的焊点
			String weldtypename = Util.getObjectDisplayName(session, "WeldPoint");
			String[] propertys = new String[] { "bl_item_object_type", "bl_item_object_type" };
			String[] values = new String[] { weldtypename, weldtypename };

			ArrayList partList = Util.searchBOMLine(parent, "OR", propertys, "==", values);
			System.out.println("包含的焊点：" + partList.toString());
			if (partList != null && partList.size() > 0) {

				for (int j = 0; j < partList.size(); j++) {
					String[] value = new String[29];
					int boradnum = 0; // 用于判断板层数
					TCComponentBOMLine bl = (TCComponentBOMLine) partList.get(j);
					TCComponentItemRevision rev = bl.getItemRevision();

					// 获取x,y,z坐标
					String xform = Util.getProperty(bl, "bl_plmxml_abs_xform");// 绝对变换矩阵
					Double[] xyzArray = getXYZ(xform);
					Double x = xyzArray[0];
					Double y = xyzArray[1];
					Double z = xyzArray[2];

					value[0] = Util.getProperty(bl, "bl_rev_object_name");// 焊点编号
					
					// 获取板件1
					String cp1 = "";
					// 获取板件2
					String cp2 = "";
					// 获取板件3
					String cp3 = "";
					// 获取板件4
					String cp4 = "";
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
						}else if(strValues.length == 3) {
							String[] strcp1 = strValues[0].split("/");
							cp1 = strcp1[0].trim();
							String[] strcp2 = strValues[1].split("/");
							cp2 = strcp2[0].trim();
							String[] strcp3 = strValues[2].split("/");
							cp3 = strcp3[0].trim();
						}else {
							String[] strcp1 = strValues[0].split("/");
							cp1 = strcp1[0].trim();
							String[] strcp2 = strValues[1].split("/");
							cp2 = strcp2[0].trim();
							String[] strcp3 = strValues[2].split("/");
							cp3 = strcp3[0].trim();
							String[] strcp4 = strValues[3].split("/");
							cp4 = strcp4[0].trim();
						}
					}
					if (cp1 != null && !cp1.isEmpty()) {
						String[] strvalue;
						// 调用系统查询，获取相关的板件
						if (partMap.containsKey(cp1)) {
							strvalue = partMap.get(cp1);
						} else {
							// 调用系统查询，获取相关的板件
							strvalue = getPropertysBypartNo(root, cp1, x, y, z);
							partMap.put(cp1, strvalue);
						}
						value[2] = cp1;
						value[3] = strvalue[0];// 材质
						value[4] = strvalue[1];// 强度
						value[5] = strvalue[2];// 板厚

						if (value[4] == null || value[4].isEmpty()) {
							value[4] = "缺少强度值";
						}

						boradnum++;
					} else {
						value[4] = "";// 强度
					}
					if (cp2 != null && !cp2.isEmpty()) {
						String[] strvalue;
						// 调用系统查询，获取相关的板件
						if (partMap.containsKey(cp2)) {
							strvalue = partMap.get(cp2);
						} else {
							// 调用系统查询，获取相关的板件
							strvalue = getPropertysBypartNo(root, cp2, x, y, z);
							partMap.put(cp2, strvalue);
						}
						value[6] = cp2;
						value[7] = strvalue[0];
						value[8] = strvalue[1];
						value[9] = strvalue[2];
						if (value[8] == null || value[8].isEmpty()) {
							value[8] = "缺少强度值";
						}
						boradnum++;
					} else {
						value[8] = "";
					}
					if (cp3 != null && !cp3.isEmpty()) {
						String[] strvalue;
						// 调用系统查询，获取相关的板件
						if (partMap.containsKey(cp3)) {
							strvalue = partMap.get(cp3);
						} else {
							// 调用系统查询，获取相关的板件
							strvalue = getPropertysBypartNo(root, cp3, x, y, z);
							partMap.put(cp3, strvalue);
						}
						value[10] = cp3;
						value[11] = strvalue[0];
						value[12] = strvalue[1];
						value[13] = strvalue[2];
						if (value[12] == null || value[12].isEmpty()) {
							value[12] = "缺少强度值";
						}
						boradnum++;
					} else {
						value[12] = "";
					}
					if (cp4 != null && !cp4.isEmpty()) {
						String[] strvalue;
						// 调用系统查询，获取相关的板件
						if (partMap.containsKey(cp4)) {
							strvalue = partMap.get(cp4);
						} else {
							// 调用系统查询，获取相关的板件
							strvalue = getPropertysBypartNo(root, cp4, x, y, z);
							partMap.put(cp4, strvalue);
						}
						boradnum++;
					}
					int maxstrength = getMaxstrength(value[4], value[8], value[12]);// 最大强度
					String sumthick = getSumBoradThickness(value[5], value[9], value[13]);
					String maxthick = getMaxBoradThickness(value[5], value[9], value[13]);
					String minthick = getMinBoradThickness(value[5], value[9], value[13]);
					value[14] = Integer.toString(boradnum);

					System.out.println("板层数量：" + boradnum);

					// 如果存在板件强度为空，不做逻辑处理
					boolean strengthflag1 = true;
					boolean strengthflag2 = true;
					boolean strengthflag3 = true;
					if (!Util.isNumber(value[4])) {
						strengthflag1 = false;
					}
					if (!Util.isNumber(value[8])) {
						strengthflag2 = false;
					}
					if (!Util.isNumber(value[12])) {
						strengthflag3 = false;
					}
					System.out.println("板件强度情况：" + strengthflag1 + strengthflag2 + strengthflag3);

					if (boradnum == 2) {

						value[15] = sumthick;// 总厚度
						value[18] = maxthick;// 最大厚度
						value[19] = minthick;// 最小厚度
						
						// 如果存在板材的板厚为空，则不判断以下逻辑
						if (Double.parseDouble(sumthick) != 0 && Double.parseDouble(maxthick) != 0
								&& Double.parseDouble(minthick) != -1) {
							
							String difference = getPlateThicknessDifference(maxthick, minthick);
							value[20] = difference;// 板厚差
							System.out.println("2层板厚差：" + difference);

							// 如果存在强度为空，则不判断以下逻辑
							if ((strengthflag1 && strengthflag2) || (strengthflag1 && strengthflag3)
									|| (strengthflag2 && strengthflag3)) {

								if (maxstrength <= 590 || maxstrength == 780 || maxstrength == 980
										|| maxstrength == 1180) {
									value[16] = "≤6";
									if (Double.parseDouble(sumthick) <= 6) {
										value[17] = "OK";
									} else {
										value[17] = "NG";
									}
								} else if (maxstrength == 1350) {
									value[16] = "≤1.8";
									if (Double.parseDouble(minthick) <= 1.6) {
										value[17] = "OK";
									} else {
										value[17] = "NG";
									}
								} else {
									value[16] = "";
									value[17] = "";
								}
							}
						}

						// 如果S列、X列、AC列均为OK
						if (value[17] == "OK") {
							value[1] = "OK";
						} else {
							value[1] = "NG";
						}

					} else if (boradnum == 3) {

						value[15] = sumthick;// 总厚度
						value[18] = maxthick;// 最大厚度
						value[19] = minthick;// 最小厚度

						// 如果存在板材的板厚为空，则不判断以下逻辑
						if (Double.parseDouble(sumthick) != 0 && Double.parseDouble(maxthick) != 0
								&& Double.parseDouble(minthick) != -1) {
							String difference = getPlateThicknessDifference(maxthick, minthick);
							value[20] = difference;// 板厚差
							// 如果存在强度为空，则不判断以下逻辑
							if (strengthflag1 && strengthflag2 && strengthflag3) {
								
								double outsidestrenth = getDouble(value[4]);
								double outsidethick = getDouble(value[5]);
								double insidestrenth = getDouble(value[12]);
								double insidethick = getDouble(value[13]);
								double intensityratio = Math.max(outsidethick * outsidethick * outsidestrenth,
										insidethick * insidethick * insidestrenth)
										/ Math.min(outsidethick * outsidethick * outsidestrenth,
												insidethick * insidethick * insidestrenth);// AB列【强度比】
								System.out.println("强度比：" + String.format("%.1f", intensityratio));
								if (maxstrength < 590) {
									if (Double.parseDouble(sumthick) <= 4.6) {
										value[16] = "≤4.6";
										value[17] = "OK";
										value[21] = "≤3.0";
										if (Double.parseDouble(difference) <= 3.0) {
											value[22] = "OK";
										} else if (Double.parseDouble(minthick) == getDouble(value[9])) {
											value[22] = "OK(中间为最薄板)"; // 字体标红
										} else {
											value[22] = "NG";
										}
									} else if (Double.parseDouble(sumthick) > 4.6
											&& Double.parseDouble(sumthick) <= 6) {
										value[16] = "4.6<t≤6";
										value[17] = "OK";
										value[21] = "≤1.4";
										if (Double.parseDouble(difference) <= 1.4) {
											value[22] = "OK";
										} else if (Double.parseDouble(minthick) == getDouble(value[9])) {
											value[22] = "OK(中间为最薄板)"; // 字体标红
										} else {
											value[22] = "NG";
										}
									} else {
										value[16] = "≤6";
										value[17] = "NG";
									}
									// 如果S列、X列均为OK
									if (value[17] == "OK" && value[22] == "OK") {
										value[1] = "OK";
									} else {
										value[1] = "NG";
									}
								} else if (maxstrength == 590 || maxstrength == 780 || maxstrength == 980) {
									value[23] = "590、780及980的高强材";
									value[21] = "≤3.0";
									if (Double.parseDouble(difference) <= 3.0) {
										value[22] = "OK";
									} else {
										value[22] = "NG";
									}
									if (Double.parseDouble(sumthick) <= 3.6) {
										value[16] = "≤3.6";
										value[17] = "OK";
										value[24] = "总板厚≤3.6mm，强度比≤16.5";
										value[25] = "≤16.5";
										value[26] = String.format("%.1f", intensityratio);
										;
										if (intensityratio <= 16.5) {
											value[27] = "OK";
										} else {
											value[27] = "NG";
										}
									} else if (Double.parseDouble(sumthick) > 3.6
											&& Double.parseDouble(sumthick) <= 5.5) {
										value[16] = "3.6<t≤5.5";
										value[17] = "OK";
										double t = Double.parseDouble(sumthick);
										double tempthick = -2.1743 * Math.pow(t, 3) + 33.93 * t * t - 178.01 * t
												+ 318.78;
										value[24] = "总板厚3.6mm<t≤5.5mm,强度比≤" + String.format("%.1f", tempthick);
										value[25] = "≤" + String.format("%.1f", tempthick);
										value[26] = String.format("%.1f", intensityratio);
										if (intensityratio <= tempthick) {
											value[27] = "OK";
										} else {
											value[27] = "NG";
										}
									} else if (Double.parseDouble(sumthick) > 5.5
											&& Double.parseDouble(sumthick) <= 6.0) {
										value[16] = "5.5<t≤6";
										value[17] = "OK";
										value[24] = "总板厚5.5mm<t≤6.0mm ，强度比≤4.5";
										value[25] = "≤4.5";
										value[26] = String.format("%.1f", intensityratio);
										if (intensityratio <= 4.5) {
											value[27] = "OK";
										} else {
											value[27] = "NG";
										}
									} else {

									}
									// 如果S列、X列、AC列均为OK
									if (value[17] == "OK" && value[22] == "OK" && value[27] == "OK") {
										value[1] = "OK";
									} else {
										value[1] = "NG";
									}
								} else if (maxstrength == 1180) {
									value[23] = "RP154-1180高强材";
									value[21] = "≤3.0";
									if (Double.parseDouble(difference) <= 3.0) {
										value[22] = "OK";
									} else {
										value[22] = "NG";
									}
									if (Double.parseDouble(sumthick) <= 3.3) {
										value[16] = "≤3.3";
										value[17] = "OK";
										value[24] = "总板厚≤3.3mm，强度比≤11.2";
										value[25] = "≤11.2";
										value[26] = String.format("%.1f", intensityratio);
										if (intensityratio <= 11.2) {
											value[27] = "OK";
										} else {
											value[27] = "NG";
										}
									} else if (Double.parseDouble(sumthick) > 3.6
											&& Double.parseDouble(sumthick) <= 5.5) {
										// 判断中间板的强度是否为1180
										if (Double.parseDouble(value[8]) == 1180) {
											if (Double.parseDouble(sumthick) > 3.6
													&& Double.parseDouble(sumthick) <= 3.8) {
												value[16] = "3.6<t≤5.5";
												value[17] = "OK";
												value[24] = "总板厚3.3mm<t≤3.8mm，强度比≤9t-18.2";
												double t = Double.parseDouble(sumthick);
												double tempthick = 9 * t - 18.2;
												value[25] = "≤" + String.format("%.1f", tempthick);
												value[26] = String.format("%.1f", intensityratio);
												if (intensityratio <= tempthick) {
													value[27] = "OK";
												} else {
													value[27] = "NG";
												}
											} else if (Double.parseDouble(sumthick) > 3.8
													&& Double.parseDouble(sumthick) <= 4.6) {
												value[16] = "3.8<t≤4.6";
												value[17] = "OK";
												value[24] = "总板厚3.8mm<t≤4.6mm，3t-3≤强度比<16";
												double t = Double.parseDouble(sumthick);
												double tempthick = 3 * t - 3;
												value[25] = Double.toString(tempthick) + "≤强度比<16";
												value[26] = String.format("%.1f", intensityratio);
												if (intensityratio <= tempthick) {
													value[27] = "OK";
												} else {
													value[27] = "NG";
												}
											} else {
												value[16] = "≤4.6";
												value[17] = "NG";
											}
										} else {
											value[16] = "3.6<t≤5.5";
											value[17] = "OK";
											value[24] = "总板厚3.3mm<t≤4.6mm，强度比≤-6t+31.3";
											double t = Double.parseDouble(sumthick);
											double tempthick = -6 * t + 31.3;
											value[25] = "≤" + String.format("%.1f", tempthick);
											value[26] = String.format("%.1f", intensityratio);
											if (intensityratio <= tempthick) {
												value[27] = "OK";
											} else {
												value[27] = "NG";
											}
										}
									}
									// 如果S列、X列、AC列均为OK
									if (value[17] == "OK" && value[22] == "OK" && value[27] == "OK") {
										value[1] = "OK";
									} else {
										value[1] = "NG";
									}
								} else if (maxstrength == 1350) {
									value[23] = "SP151-1350H(热成型)高强材";
									value[21] = "≤3.0";
									if (Double.parseDouble(difference) <= 3.0) {
										value[22] = "OK";
									} else {
										value[22] = "NG";
									}
									value[16] = "t≤4.6";
									if (Double.parseDouble(sumthick) <= 4.6) {
										value[17] = "OK";
									} else {
										value[17] = "NG";
									}
									// 如果S列、X列、AC列均为OK
									if (value[17] == "OK" && value[22] == "OK") {
										value[1] = "OK";
									} else {
										value[1] = "NG";
									}
								} else {

								}

							} else {
								value[1] = "NG";
							}
						} else {
							value[1] = "NG";
						}

					} else { // 如果不是2或3层板，直接输出NG
						value[1] = "NG";// 焊接成立性
						//value[2] = "";
						//value[3] = "";
						//value[4] = "";
						//value[5] = "";
						//value[6] = "";
						//value[7] = "";
						//value[8] = "";
						//value[9] = "";
						//value[10] = "";
						//value[11] = "";
						//value[12] = "";
						//value[13] = "";
						value[14] = Integer.toString(boradnum); // 板层数
						// 焊点如只关联一层板，在其余connected_part列下输出“未关联板件”；如均未关联板件，在所有connected_part列下输出“未关联板件”，其余情况不做判断
						if (boradnum == 0) {
							value[2] = "未关联板件";
							value[6] = "未关联板件";
							value[10] = "未关联板件";
						}
						if (boradnum == 1) {
							if (value[2] == null || value[2].isEmpty()) {
								value[2] = "未关联板件";
							}
							if (value[6] == null || value[6].isEmpty()) {
								value[6] = "未关联板件";
							}
							if (value[10] == null || value[10].isEmpty()) {
								value[10] = "未关联板件";
							}
						}
						// 其他信息都为空
					}

					weldlist.add(value);
				}
			}
		}
	}

	/*
	 * 板件为厚薄板的强度 和板厚取值逻辑
	 */
	private String[] getSpeclalBoradStrenthThick(double x, String bcX, String particl) {
		String[] value = new String[2];

		return value;
	}

	/*
	 * 最大强度
	 */
	private int getMaxstrength(String strength1, String strength2, String strength3) {
		int maxstrength = 0;
		double max = 0;
		double thk1 = getDouble(strength1);
		double thk2 = getDouble(strength2);
		double thk3 = getDouble(strength3);

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
		maxstrength = (int) max;

		return maxstrength;
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
		maxth = String.format("%.1f", max);
		// maxth = Double.toString(max);

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
	private String[] getPropertysBypartNo(TCComponentBOMLine parrent, String partno, double Xcoordinate,
			double Ycoordinate, double Zcoordinate) throws TCException {
		String[] values = new String[3];
		// 调用系统查询，获取相关的板件
		List tcclist = Util.callStructureSearch(parrent, "__DFL_Find_SolutionPart", new String[] { "PARTNO" },
				new String[] { partno });
		if (tcclist != null && tcclist.size() > 0) {
			TCComponentBOMLine sol = (TCComponentBOMLine) tcclist.get(0);
			TCComponentItemRevision solrev3 = sol.getItemRevision();
			// values[0] = Util.getProperty(solrev3, "dfl9_part_no");// 板组3
			String partcal = Util.getProperty(solrev3, "dfl9PartMaterial");// 材质
			// 如果是厚薄板，根据板件上的属性与焊点坐标比较确定，材质和板厚
			if ((partcal != null && partcal.trim().equals("590T1.7/980T2.0"))
					|| (partcal != null && partcal.trim().equals("783T1.7/980T2.0"))) {
				// 厚度差零件材质板厚属性，值样例为：X小于等于200，SP-791-440PQ，1.5；X大于等于200，SP121，2.0
				String TPMTAxis = Util.getProperty(solrev3, "B8_TPMTAxis");
				String TPMTNum = Util.getProperty(solrev3, "B8_TPMTNum");
				String TPMTBig = Util.getProperty(solrev3, "B8_TPMTBig");
				String TPMTSmall = Util.getProperty(solrev3, "B8_TPMTSmall");

				if (TPMTNum != null && Util.isNumber(TPMTNum)) {
					if (TPMTAxis != null && TPMTAxis.equals("X")) {
						if (Double.parseDouble(TPMTNum) <= Xcoordinate * 1000) {
							if (!TPMTBig.isEmpty()) {
								String[] str = TPMTBig.split("/");
								if (str.length > 1) {
									values[0] = str[0];
									if (map.containsKey(values[0])) {
										values[1] = map.get(values[0]); // 强度
									} else {
										values[1] = ""; // 强度
									}
									values[2] = str[1];
								}
							}
						} else {
							if (!TPMTSmall.isEmpty()) {
								String[] str = TPMTSmall.split("/");
								if (str.length > 1) {
									values[0] = str[0];
									if (map.containsKey(values[0])) {
										values[1] = map.get(values[0]); // 强度
									} else {
										values[1] = ""; // 强度
									}
									values[2] = str[1];
								}
							}
						}
					} else if (TPMTAxis != null && TPMTAxis.equals("Y")) {
						if (Double.parseDouble(TPMTNum) <= Ycoordinate * 1000) {
							if (!TPMTBig.isEmpty()) {
								String[] str = TPMTBig.split("/");
								if (str.length > 1) {
									values[0] = str[0];
									if (map.containsKey(values[0])) {
										values[1] = map.get(values[0]); // 强度
									} else {
										values[1] = ""; // 强度
									}
									values[2] = str[1];
								}
							}
						} else {
							if (!TPMTSmall.isEmpty()) {
								String[] str = TPMTSmall.split("/");
								if (str.length > 1) {
									values[0] = str[0];
									if (map.containsKey(values[0])) {
										values[1] = map.get(values[0]); // 强度
									} else {
										values[1] = ""; // 强度
									}
									values[2] = str[1];
								}
							}
						}

					} else if (TPMTAxis != null && TPMTAxis.equals("Z")) {
						if (Double.parseDouble(TPMTNum) <= Zcoordinate * 1000) {
							if (!TPMTBig.isEmpty()) {
								String[] str = TPMTBig.split("/");
								if (str.length > 1) {
									values[0] = str[0];
									if (map.containsKey(values[0])) {
										values[1] = map.get(values[0]); // 强度
									} else {
										values[1] = ""; // 强度
									}
									values[2] = str[1];
								}
							}
						} else {
							if (!TPMTSmall.isEmpty()) {
								String[] str = TPMTSmall.split("/");
								if (str.length > 1) {
									values[0] = str[0];
									if (map.containsKey(values[0])) {
										values[1] = map.get(values[0]); // 强度
									} else {
										values[1] = ""; // 强度
									}
									values[2] = str[1];
								}
							}
						}

					} else {

					}
				}

			} else {
				String bh3 = Util.getProperty(solrev3, "dfl9PartThickness");// 板厚
				if (bh3 != null && !bh3.isEmpty()) {
					values[2] = format.format(new BigDecimal(bh3.toString()));
				} else {
					values[2] = bh3;
				}
				values[0] = partcal;// 材质

				if (map.containsKey(values[0])) {
					values[1] = map.get(values[0]); // 强度
				} else {
					values[1] = ""; // 强度
				}
			}

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
