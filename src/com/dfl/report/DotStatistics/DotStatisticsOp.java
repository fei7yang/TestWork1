package com.dfl.report.DotStatistics;

import java.io.File;
import java.io.InputStream;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.NewOutputDataToExcel;
import com.dfl.report.util.ReportViwePanel;
import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.AIFComponentContext;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentBOMLine;
import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentFolder;
import com.teamcenter.rac.kernel.TCComponentItem;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCComponentItemType;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class DotStatisticsOp {

	private TCSession session;
	private InterfaceAIFComponent[] aifComponents;
	private TCComponent savefolder;
	private ArrayList other = new ArrayList();
	private int allsum = 0;// 总焊点数
	private int rswsum = 0;// 自动化焊点
	private Map<String, Integer[]> statistics = new HashMap<String, Integer[]>();// 人工、机器焊统计集合
	SimpleDateFormat df = new SimpleDateFormat("yyyyMMddHH");// 设置日期格式

	public DotStatisticsOp(TCSession session, InterfaceAIFComponent[] aifComponents, TCComponent savefolder)
			throws TCException {
		// TODO Auto-generated constructor stub
		this.session = session;
		this.aifComponents = aifComponents;
		this.savefolder = savefolder;
		initUI();
	}

	private void initUI() throws TCException {
		// TODO Auto-generated method stub

		// 显示进度输出窗口
		// 界面显示进度并输出执行步骤
		ReportViwePanel viewPanel = new ReportViwePanel("生成报表");
		viewPanel.setVisible(true);
		TCComponentBOMLine topbl = (TCComponentBOMLine) aifComponents[0];

		viewPanel.addInfomation("正在获取模板...\n", 10, 100);
		// 查询目录导出模板
		InputStream inputStream = FileUtil.getTemplateFile("DFL_Template_DotStatistics");

		if (inputStream == null) {
			viewPanel.addInfomation("错误：没有找到打点统计表模板，请先添加模板(名称为：DFL_Template_DotStatistics)\n", 100, 100);
			return;
		}
		String FamlilyCode = Util.getProperty(topbl, "bl_rev_project_ids");// 基本车型
		String VehicleNo = Util.getDFLProjectIdVehicle(FamlilyCode);
		if(VehicleNo== null || VehicleNo.isEmpty()) {
			VehicleNo = FamlilyCode;
		}
		// 工厂名称  改为从BOP名称中取值
		String factoryname = "";
		String bopname = Util.getProperty(topbl, "bl_rev_object_name");
		String[] bopnames = bopname.split("_");
		if(bopnames.length>2) {
			String factory = bopnames[2];
			if(factory.length()>2) {
				factoryname = factory.substring(0, 3);
			}
		}
		// 获取关联的BOe
//		TCComponent[] boelist = topbl.getItemRevision().getRelatedComponents("IMAN_MEWorkArea");
//		if (boelist != null && boelist.length > 0) {
//			TCComponentItemRevision boerev = (TCComponentItemRevision) boelist[0];
//			factoryname = factoryname + Util.getProperty(boerev, "object_name");
//		}
		viewPanel.addInfomation("", 20, 100);
		// 遍历BOP顶层，获取所有的虚层产线，如果产线没有放在虚层产线，归为其他
		ArrayList asahi = getAsahiLine(topbl);

		viewPanel.addInfomation("开始输出报表...\n", 40, 100);

		// 统计点焊工序下的焊点数
		Map<String, ArrayList> map = getAllWeldnumInfo(asahi);

		String[] str = new String[6];
		NumberFormat numberFormat = NumberFormat.getInstance();
		String result = "";
		if (allsum == 0) {
			result = "0";
		} else {
			result = numberFormat.format((float) rswsum / (float) allsum * 100);
		}
		str[0] = factoryname;
		str[1] = VehicleNo;
		str[2] = result + "%";
		str[3] = Integer.toString(allsum);
		str[4] = Integer.toString(rswsum);

		viewPanel.addInfomation("开始写数据，请耐心等待...\n", 60, 100);

		XSSFWorkbook book = NewOutputDataToExcel.creatXSSFWorkbook(inputStream);
		// 写入数据
		writeDataToSheet(book, statistics, str, map, asahi);
				
		String sheetname = 	VehicleNo + "-" + "打点统计详细页";
		book.setSheetName(1, sheetname);
			

		String date = df.format(new Date());
		String datasetname = VehicleNo  + "_打点统计表" + "_" + date + "时";
		String filename = Util.formatString(datasetname);

		NewOutputDataToExcel.exportFile(book, filename);
		
		
		viewPanel.addInfomation("", 80, 100);

		// NewOutputDataToExcel.openFile(FileUtil.getReportFileName(filename.trim()));
		saveFiles(filename, datasetname, savefolder, session);

		viewPanel.addInfomation("输出报表完成，请在选择保存的文件夹下查看！", 100, 100);
	}

	/*
	 * ****************************************************** 把生成的报表保存在指定的文件夹下
	 */
	public void saveFiles(String filename, String datasetName, TCComponent folder, TCSession session) {
		try {
			String fullFileName = FileUtil.getReportFileName(filename);
			TCComponentFolder savefolder = (TCComponentFolder) folder;
			TCComponentItemType tcccomponentitemtype = (TCComponentItemType) session.getTypeComponent("B8_BIWProcDoc");
			TCComponentItem tccomponentitem = tcccomponentitemtype.create("", "", "B8_BIWProcDoc", datasetName, "desc",
					null);
			tccomponentitem.setProperty("b8_BIWProcDocType", "AO");
			tccomponentitem.lock();
			tccomponentitem.save();
			tccomponentitem.unlock();
			TCComponentItemRevision rev = tccomponentitem.getLatestItemRevision();
			TCComponentDataset ds = Util.createDataset(session, datasetName, fullFileName, "MSExcelX", "excel");

			rev.add("IMAN_specification", ds);

			// 添加文档与数据集的关系
			savefolder.add("contents", tccomponentitem);

			// 删除中间文件
			File file = new File(fullFileName);
			if (file.isFile()) {
				file.delete();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 写入数据
	private void writeDataToSheet(XSSFWorkbook book, Map<String, Integer[]> statistics2, String[] str,
			Map<String, ArrayList> map, ArrayList asahi) {
		// TODO Auto-generated method stub

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
		// 紫色背景
		XSSFCellStyle style3 = book.createCellStyle();
		style3.setFillForegroundColor(IndexedColors.VIOLET.getIndex());
		style3.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style3.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style3.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style3.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style3.setFont(font);
		// 黄色背景
		XSSFCellStyle style4 = book.createCellStyle();
		style4.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		style4.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style4.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style4.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style4.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style4.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style4.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		style4.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		style4.setFont(font);
		// 先写入第一个sheet数据
		XSSFSheet sheet1 = book.getSheetAt(0);
		// 工厂-自动化焊点
		setStringCellAndStyle(sheet1, str[0], 1, 1, style, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet1, str[1], 1, 2, style, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet1, str[2], 1, 3, style, Cell.CELL_TYPE_STRING);
		setStringCellAndStyle(sheet1, str[3], 1, 4, style, 10);
		setStringCellAndStyle(sheet1, str[4], 1, 5, style2, 10);
		
		//设置自动列宽
        for (int i = 0; i < 10; i++) {
            sheet1.autoSizeColumn(i);
        }


		// 再写第二个sheet
		XSSFSheet sheet2 = book.getSheetAt(1);
		// 动态加载表头
		int colnum = 0;
		int colnum2 = 0;
		CellRangeAddress region;
		for (int i = 0; i < asahi.size(); i++) {
			TCComponentBOMLine bl = (TCComponentBOMLine) asahi.get(i);
			String templinename = Util.getProperty(bl, "bl_rev_object_name");
			String linename = templinename.replaceAll("\\d+", "").replace("-", "").replace("_", "");;// 去掉数字，虚层产线名称
			if (statistics2.containsKey(linename)) {
				Integer[] instr = statistics2.get(linename);
				setStringCellAndStyle(sheet1, linename + "焊点总数", 0, 7 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet1, linename + "人工", 0, 8 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet1, linename + "机器人", 0, 9 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
				setStringCellAndStyle(sheet1, instr[0].toString(), 1, 7 + 3 * colnum, style, 10);
				setStringCellAndStyle(sheet1, instr[1].toString(), 1, 8 + 3 * colnum, style, 10);
				setStringCellAndStyle(sheet1, instr[2].toString(), 1, 9 + 3 * colnum, style2, 10);
				colnum++;
			}
			if (sheet2 != null) {
				if (map.containsKey(linename)) {
					setStringCellAndStyle(sheet2, linename, 0, 0 + 4 * colnum2, style4, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, "RH", 0, 1 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, "LH", 0, 2 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, "", 0, 3 + 4 * colnum2, style3, Cell.CELL_TYPE_STRING);

					ArrayList list = map.get(linename);
					String begin = "";
					int beginrow = 1;// 起始行
					int endrow = 0;// 终止行
					int n = 0;// 标记行
					for (int j = 0; j < list.size(); j++) {
						String[] values = (String[]) list.get(j);
						if (values[0].equals("总计")) {
							setStringCellAndStyle(sheet2, values[0], 1 + j, 0 + 4 * colnum2, style2,
									Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sheet2, values[1], 1 + j, 1 + 4 * colnum2, style2, 10);
							setStringCellAndStyle(sheet2, values[2], 1 + j, 2 + 4 * colnum2, style2, 10);
						} else {
							setStringCellAndStyle(sheet2, values[0], 1 + j, 0 + 4 * colnum2, style,
									Cell.CELL_TYPE_STRING);
							setStringCellAndStyle(sheet2, values[1], 1 + j, 1 + 4 * colnum2, style, 10);
							setStringCellAndStyle(sheet2, values[2], 1 + j, 2 + 4 * colnum2, style, 10);
						}
						setStringCellAndStyle(sheet2, "", 1 + j, 3 + 4 * colnum2, style3, Cell.CELL_TYPE_STRING);

						if (j == 0) {
							begin = values[0];
						} else {
							if (!begin.equals(values[0])) {
								endrow = beginrow + n;
								region = new CellRangeAddress(beginrow, endrow, (short) (0 + 4 * colnum2),
										(short) (0 + 4 * colnum2));
								sheet2.addMergedRegion(region);
								begin = values[0];
								beginrow = endrow + 1;
								n = 0;
							} else {
								n++;
							}
						}
						if (j == list.size() - 1) {
							endrow = beginrow + n;
							region = new CellRangeAddress(beginrow, endrow, (short) (0 + 4 * colnum2),
									(short) (0 + 4 * colnum2));
							sheet2.addMergedRegion(region);
						}

					}
					colnum2++;
				}

			}
		}
		if (statistics2.containsKey("其它")) {
			Integer[] instr = statistics2.get("其它");
			setStringCellAndStyle(sheet1, "其它" + "焊点总数", 0, 7 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet1, "其它" + "人工", 0, 8 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet1, "其它" + "机器人", 0, 9 + 3 * colnum, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet1, instr[0].toString(), 1, 7 + 3 * colnum, style, 10);
			setStringCellAndStyle(sheet1, instr[1].toString(), 1, 8 + 3 * colnum, style, 10);
			setStringCellAndStyle(sheet1, instr[2].toString(), 1, 9 + 3 * colnum, style2, 10);

		}
		if (map.containsKey("其它")) {
			setStringCellAndStyle(sheet2, "其它", 0, 0 + 4 * colnum2, style4, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet2, "RH", 0, 1 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet2, "LH", 0, 2 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
			setStringCellAndStyle(sheet2, "", 0, 3 + 4 * colnum2, style3, Cell.CELL_TYPE_STRING);

			ArrayList list = map.get("其它");
			String begin = "";
			int beginrow = 1;// 起始行
			int endrow = 0;// 终止行
			int n = 0;// 标记行

			for (int j = 0; j < list.size(); j++) {
				String[] values = (String[]) list.get(j);
				if (values[0].equals("总计")) {
					setStringCellAndStyle(sheet2, values[0], 1 + j, 0 + 4 * colnum2, style2, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, values[1], 1 + j, 1 + 4 * colnum2, style2, 10);
					setStringCellAndStyle(sheet2, values[2], 1 + j, 2 + 4 * colnum2, style2, 10);
				} else {
					setStringCellAndStyle(sheet2, values[0], 1 + j, 0 + 4 * colnum2, style, Cell.CELL_TYPE_STRING);
					setStringCellAndStyle(sheet2, values[1], 1 + j, 1 + 4 * colnum2, style, 10);
					setStringCellAndStyle(sheet2, values[2], 1 + j, 2 + 4 * colnum2, style, 10);
				}
				setStringCellAndStyle(sheet2, "", 1 + j, 3 + 4 * colnum2, style3, Cell.CELL_TYPE_STRING);

				if (j == 0) {
					begin = values[0];
				} else {
					if (!begin.equals(values[0])) {
						endrow = beginrow + n;
						region = new CellRangeAddress(beginrow, endrow, (short) (0 + 4 * colnum2),
								(short) (0 + 4 * colnum2));
						sheet2.addMergedRegion(region);
						begin = values[0];
						beginrow = endrow + 1;
						n = 0;
					} else {
						n++;
					}
				}
				if (j == list.size() - 1) {
					endrow = beginrow + n;
					region = new CellRangeAddress(beginrow, endrow, (short) (0 + 4 * colnum2),
							(short) (0 + 4 * colnum2));
					sheet2.addMergedRegion(region);
				}
			}
		}
//		//设置自动列宽
//        for (int i = 0; i < 3*colnum; i++) {
//            sheet1.autoSizeColumn(i);
//        }
//        // 处理中文不能自动调整列宽的问题
//        this.setSizeColumn(sheet1, 3*colnum);
//        for (int i = 0; i < 3*colnum2; i++) {
//            sheet2.autoSizeColumn(i);
//        }
//        this.setSizeColumn(sheet2, 3*colnum2);
	}
	// 自适应宽度(中文支持)
    private void setSizeColumn(XSSFSheet sheet, int size) {
        for (int columnNum = 0; columnNum < size; columnNum++) {
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
	public void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex, XSSFCellStyle Style,
			int celltype) {

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

	// 统计点焊工序下的焊点数
	private Map<String, ArrayList> getAllWeldnumInfo(ArrayList asahi) throws TCException {
		// TODO Auto-generated method stub
		Map<String, ArrayList> map = new HashMap<String, ArrayList>();
		if (asahi != null && asahi.size() > 0) {
			for (int i = 0; i < asahi.size(); i++) {
				ArrayList dataList = new ArrayList();
				int psw = 0;
				int rsw = 0;
				TCComponentBOMLine bl = (TCComponentBOMLine) asahi.get(i);
				String templinename = Util.getProperty(bl, "bl_rev_object_name");
				String linename = templinename.replaceAll("\\d+", "").replace("-", "").replace("_", "");// 去掉数字，虚层产线名称
				ArrayList linelist = Util.getChildrenByBOMLine(bl, "B8_BIWMEProcLineRevision");// 实际焊装产线
				if (linelist != null && linelist.size() > 0) {
					for (int j = 0; j < linelist.size(); j++) {
						boolean flag = false;// 判断实际产线是否有工位，没有不输出总计
						TCComponentBOMLine linebl = (TCComponentBOMLine) linelist.get(j);
						int rhsum = 0;// RH 总计数量
						int lhsum = 0;// LH 总计数量
						boolean flag2 = Util.getIsMEProcStat(linebl);
						String lname = Util.getProperty(linebl.getItemRevision(), "b8_LineType") + Util.getProperty(linebl, "bl_rev_object_name"); // 实际产线名称 需要拼接线体式样输出   20191202修改
						ArrayList statelist = Util.getChildrenByBOMLine(linebl, "B8_BIWMEProcStatRevision");// 焊装工位工艺
						if (statelist != null && statelist.size() > 0) {
							flag = true;
							for (int k = 0; k < statelist.size(); k++) {
								TCComponentBOMLine statebl = (TCComponentBOMLine) statelist.get(k);
								String statename = Util.getProperty(statebl, "bl_rev_object_name"); // 工位名称
								ArrayList dislist = Util.getChildrenByBOMLine(statebl, "B8_BIWDiscreteOPRevision");// 点焊工序
								if (dislist != null && dislist.size() > 0) {
									for (int m = 0; m < dislist.size(); m++) {
										TCComponentBOMLine diebl = (TCComponentBOMLine) dislist.get(m);
										String diename = Util.getProperty(diebl, "bl_rev_object_name");
										String[] strVal = new String[3];
										int weldnum = 0;
										ArrayList weldlist = Util.getChildrenByBOMLine(diebl, "WeldPointRevision");
										if (weldlist != null && weldlist.size() > 0) {
											weldnum = weldlist.size();
										}
										// 根据产线名称判断是否为RH列，还是LH列
										if (lname.length() > 2
												&& lname.substring(lname.length() - 2).equals("LH")) {
											lhsum = lhsum + weldnum;
											strVal[2] = Integer.toString(weldnum);// LH
										} else {
											rhsum = rhsum + weldnum;
											strVal[1] = Integer.toString(weldnum);// RH
										}
										// 根据点焊工序里名称为R和M开头的统计
										if (diename.length() > 1 && (diename.substring(0, 1).equals("R")
												|| diename.substring(0, 1).equals("M"))) {
											rsw = rsw + weldnum;
											rswsum = rswsum + weldnum;
										} else {
											psw = psw + weldnum;
										}
										allsum = allsum + weldnum;
                                        //如果产线下只有一个工位输出线体名称
										if(flag2) {
											strVal[0] = lname + " " + statename;
										}else {
											strVal[0] = lname ;
										}
										
										dataList.add(strVal);

									}
								} else {
									String[] strVal = new String[3];
									 //如果产线下只有一个工位输出线体名称
									if(flag2) {
										strVal[0] = lname + " " + statename;
									}else {
										strVal[0] = lname ;
									}
									dataList.add(strVal);
								}
							}
						}
						if (flag) {
							String[] strVal = new String[3];
							strVal[0] = "总计";
							strVal[1] = Integer.toString(rhsum);
							strVal[2] = Integer.toString(lhsum);
							dataList.add(strVal);
						}
					}
				}
				map.put(linename, dataList);
				Integer[] strValue = new Integer[3];
				strValue[0] = psw + rsw;
				strValue[1] = psw;
				strValue[2] = rsw;
				statistics.put(linename, strValue);
			}
		}
		// 其他产线处理
		if (other != null && other.size() > 0) {
			ArrayList dataList = new ArrayList();
			int psw = 0;
			int rsw = 0;
			String linename = "其它";//
			for (int j = 0; j < other.size(); j++) {
				TCComponentBOMLine linebl = (TCComponentBOMLine) other.get(j);
				int rhsum = 0;// RH 总计数量
				int lhsum = 0;// LH 总计数量
				boolean flag2 = Util.getIsMEProcStat(linebl);
				String lname = Util.getProperty(linebl.getItemRevision(), "b8_LineType") + Util.getProperty(linebl, "bl_rev_object_name"); // 实际产线名称
				ArrayList statelist = Util.getChildrenByBOMLine(linebl, "B8_BIWMEProcStatRevision");// 焊装工位工艺
				if (statelist != null && statelist.size() > 0) {
					for (int k = 0; k < statelist.size(); k++) {
						TCComponentBOMLine statebl = (TCComponentBOMLine) statelist.get(k);
						String statename = Util.getProperty(statebl.parent(), "bl_rev_object_name"); // 工位名称
						ArrayList dislist = Util.getChildrenByBOMLine(statebl, "B8_BIWDiscreteOPRevision");// 点焊工序
						if (dislist != null && dislist.size() > 0) {
							for (int m = 0; m < dislist.size(); m++) {
								TCComponentBOMLine diebl = (TCComponentBOMLine) dislist.get(m);
								String diename = Util.getProperty(diebl, "bl_rev_object_name");
								String[] strVal = new String[3];
								int weldnum = 0;
								ArrayList weldlist = Util.getChildrenByBOMLine(diebl, "WeldPointRevision");
								if (weldlist != null && weldlist.size() > 0) {
									weldnum = weldlist.size();
								}
								// 根据产线名称判断是否为RH列，还是LH列
								if (diename.length() > 2 && diename.substring(diename.length() - 2).equals("LH")) {
									lhsum = lhsum + weldnum;
									strVal[2] = Integer.toString(weldnum);// LH
								} else {
									rhsum = rhsum + weldnum;
									strVal[1] = Integer.toString(weldnum);// RH
								}
								// 根据产线里名称为R和M开头的统计
								if (diename.length() > 1 && (diename.substring(0, 1).equals("R")
										|| diename.substring(0, 1).equals("M"))) {
									rsw = rsw + weldnum;
									rswsum = rswsum + weldnum;
								} else {
									psw = psw + weldnum;
								}
								allsum = allsum + weldnum;

								 //如果产线下只有一个工位输出线体名称
								if(flag2) {
									strVal[0] = lname + " " + statename;
								}else {
									strVal[0] = lname ;
								}
								dataList.add(strVal);

							}
						} else {
							String[] strVal = new String[3];
							 //如果产线下只有一个工位输出线体名称
							if(flag2) {
								strVal[0] = lname + " " + statename;
							}else {
								strVal[0] = lname ;
							}
							dataList.add(strVal);
						}
					}
					String[] strVal = new String[3];
					strVal[0] = "总计";
					strVal[1] = Integer.toString(rhsum);
					strVal[2] = Integer.toString(lhsum);
					dataList.add(strVal);
				}
			}

			map.put(linename, dataList);
			Integer[] strValue = new Integer[3];
			strValue[0] = psw + rsw;
			strValue[1] = psw;
			strValue[2] = rsw;
			statistics.put(linename, strValue);
		}

		return map;
	}

	// 遍历BOP顶层，获取所有的虚层产线，如果产线没有放在虚层产线，归为其他
	private ArrayList getAsahiLine(TCComponentBOMLine topbl) throws TCException {
		// TODO Auto-generated method stub
		ArrayList list = new ArrayList();
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

}
