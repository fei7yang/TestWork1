package com.dfl.report.util;

import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.examples.CellTypes;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.Region;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFPrintSetup;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.ExcelReader.BoardInformation;

public class NewOutputDataToExcel {

	private static XSSFCellStyle cellStyle;

	public NewOutputDataToExcel() {
		// TODO Auto-generated constructor stub

	}

	// 根据模板创建Excel空模板
	public static XSSFWorkbook creatXSSFWorkbook(InputStream input, ArrayList list, ArrayList errornamelist) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);

			XSSFSheet sheet1 = book.getSheetAt(0);
			System.out.println("sheet名称：" + sheet1.getSheetName());
			int num = book.getNumberOfSheets();
			int deletenum = num - 1;
			if (deletenum > 0) {
				for (int i = 0; i < deletenum; i++) {
					XSSFSheet sheet2 = book.getSheetAt(1);
					// TC添加的Excel模板数据集，会有三个页签，需要把另外两个先删除
					if (sheet2 != null) {
						book.removeSheetAt(1);
					}
				}
			}
			System.out.println("删除sheet后：" + book.getNumberOfSheets());
			//////////// 设置分组显示上方/下方
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);

			cellStyle = book.createCellStyle();
			// 边框
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

			// 点焊工序名称数据集
			List nameList = new ArrayList();
			List renameList = new ArrayList();
			Map<String, List<String>> map = new LinkedHashMap();
			for (int i = 0; i < list.size(); i++) {
				String[] str = (String[]) list.get(i);
				String[] statename = str[0].toString().split("\\\\");
				String prestatename = statename[0].toString();
				String prename = str[8].toString() + "_" + prestatename;
				if (nameList.contains(prename)) {
					if (!renameList.contains(prename)) {
						renameList.add(prename);
					}
				}
				if (!map.containsKey(prename)) {
					List<String> numlist = new ArrayList();
					numlist.add(str[10].toString());
					map.put(prename, numlist);
				} else {
					List<String> numlist = map.get(prename);
					numlist.add(str[10].toString());
					map.put(prename, numlist);
				}
				nameList.add(prename);
			}
			// 对相同工位相同点焊工序名称按照查找编号排序
			Map<String, Integer> sortmap = new HashMap<>();
			for (Map.Entry<String, List<String>> entry : map.entrySet()) {
				String key = entry.getKey();
				List<String> valuelist = entry.getValue();
				for (int i = 0; i < valuelist.size(); i++) {
					String sortnum = valuelist.get(i);
					String sortkey = key + "_" + sortnum;
					sortmap.put(sortkey, i + 1);
				}
			}
			System.out.println("有重复名称的：" + renameList);
			System.out.println("sheet名称map：" + sortmap);

			for (int i = 0; i < list.size(); i++) {
				String[] str = (String[]) list.get(i);
				System.out.println("点焊工序：" + i + " " + str[10].toString());
				String[] statename = str[0].toString().split("\\\\");
				String prename = statename[0].toString();
				// String compareshname = str[8].toString() + "_" + str[0].toString();
				String shname = str[8].toString() + "_" + prename;
				if (renameList.contains(shname)) {
					String sortshname = str[8].toString() + "_" + prename + "_" + str[10].toString();
					shname = str[8].toString() + "_" + prename + "_" + sortmap.get(sortshname);
				}
				System.out.println("sheet名称：" + shname);
				String name = FilterSpecialCharacters(shname);
				System.out.println("去除特殊字符后的sheet名称：" + name);

				// 如果sheet名称超过31个字符长度，给出提示
				if (name.length() > 31) {
					errornamelist.add(shname);
				}
			}
			if (errornamelist.size() > 0) {
				return null;
			}
			// 循环数据集，添加sheet
			for (int i = 0; i < list.size(); i++) {
				String[] str = (String[]) list.get(i);
				String[] statename = str[0].toString().split("\\\\");
				String prename = statename[0].toString();
				// String compareshname = str[8].toString() + "_" + str[0].toString();
				String shname = str[8].toString() + "_" + prename;
				if (renameList.contains(shname)) {
					String sortshname = str[8].toString() + "_" + prename + "_" + str[10].toString();
					shname = str[8].toString() + "_" + prename + "_" + sortmap.get(sortshname);
				}
				String name = FilterSpecialCharacters(shname);

				if (i == 0) {
					// 把模板中的第一个sheet修改一下
					book.setSheetName(0, name);
				} else {
					// 根据集合有多少数据添加sheet
					XSSFSheet sheet = book.cloneSheet(0);
					book.setSheetName(i, name);
					//////////// 设置分组显示上方/下方
					sheet.setRowSumsBelow(false);
					sheet.setRowSumsRight(false);
					sheet.setRowSumsBelow(false);
					sheet.setRowSumsRight(false);
				}
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	// 根据数据添加多个sheet页
	public static void creatXSSFWorkbookByData(XSSFWorkbook book, int sheetnum) {

		try {
			XSSFSheet sheet1 = book.getSheetAt(0);
			System.out.println("sheet名称：" + sheet1.getSheetName());
			int num = book.getNumberOfSheets();
			int deletenum = num - 2;
			if (deletenum > 0) {
				for (int i = 0; i < deletenum; i++) {
					XSSFSheet sheet2 = book.getSheetAt(2);
					// 如果模板多余2个sheet页，需要删除
					if (sheet2 != null) {
						book.removeSheetAt(2);
					}
				}
			}
			System.out.println("删除sheet后：" + book.getNumberOfSheets());
			//////////// 设置分组显示上方/下方
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);

			cellStyle = book.createCellStyle();
			// 边框
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

			// 循环数据集，添加sheet
			for (int i = 0; i < sheetnum; i++) {
				// 第一个sheet不用增加
				if (i != 0) {
					// 根据集合有多少数据添加sheet
					XSSFSheet sheet = book.cloneSheet(1);
					String sheetname = "直材清单" + Integer.toString(i + 1);
					int sheetIndex = i + 1;
					book.setSheetName(sheetIndex, sheetname);
					//////////// 设置分组显示上方/下方
					sheet.setRowSumsBelow(false);
					sheet.setRowSumsRight(false);
					sheet.setRowSumsBelow(false);
					sheet.setRowSumsRight(false);
				}
			}
			System.out.println("sheet数量：" + book.getNumberOfSheets());
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return;
	}

	// 根据模板创建Excel空模板
	public static XSSFWorkbook creatXSSFWorkbook(InputStream input) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);
			XSSFSheet sheet1 = book.getSheetAt(0);
			//////////// 设置分组显示上方/下方
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);

			cellStyle = book.createCellStyle();
			// 边框
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	public static XSSFWorkbook creatXSSFWorkbook2(InputStream input) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);
			XSSFSheet sheet1 = book.getSheetAt(0);
			//////////// 设置分组显示上方/下方
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);

			cellStyle = book.createCellStyle();
			// 边框
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle.setWrapText(true);
			cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
			cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// 左对齐
			cellStyle.setWrapText(true);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	// 根据模板更新Excel内容
	public static XSSFWorkbook updateXSSFWorkbook(InputStream input, ArrayList list, ArrayList errorname) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);

			cellStyle = book.createCellStyle();
			// 边框
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

			// 循环所有sheet页，根据名称去找，如果没有找到移除sheet
			System.out.println("移除前数量：" + book.getNumberOfSheets());

			// 点焊工序名称数据集
			List nameList = new ArrayList();
			List renameList = new ArrayList();
			Map<String, List<String>> map = new LinkedHashMap();
			for (int i = 0; i < list.size(); i++) {
				String[] str = (String[]) list.get(i);
				String[] statename = str[0].toString().split("\\\\");
				String prestatename = statename[0].toString();
				String prename = str[8].toString() + "_" + prestatename;
				if (nameList.contains(prename)) {
					if (!renameList.contains(prename)) {
						renameList.add(prename);
					}
				}
				if (!map.containsKey(prename)) {
					List<String> numlist = new ArrayList();
					numlist.add(str[10].toString());
					map.put(prename, numlist);
				} else {
					List<String> numlist = map.get(prename);
					numlist.add(str[10].toString());
					map.put(prename, numlist);
				}
				nameList.add(prename);
			}
			// 对相同工位相同点焊工序名称按照查找编号排序
			Map<String, Integer> sortmap = new HashMap<>();
			for (Map.Entry<String, List<String>> entry : map.entrySet()) {
				String key = entry.getKey();
				List<String> valuelist = entry.getValue();
				for (int i = 0; i < valuelist.size(); i++) {
					String sortnum = valuelist.get(i);
					String sortkey = key + "_" + sortnum;
					sortmap.put(sortkey, i + 1);
				}
			}
			System.out.println("有重复名称的：" + renameList);
			System.out.println("sheet名称map：" + sortmap);
			for (int i = 0; i < list.size(); i++) {
				String[] str = (String[]) list.get(i);
				System.out.println("点焊工序：" + i + " " + str[10].toString());
				String[] statename = str[0].toString().split("\\\\");
				String prename = statename[0].toString();
				String shname1 = str[8].toString() + "_" + prename;
				String sortname = str[8].toString() + "_" + prename + "_" + str[10].toString();
				String shname2 = str[8].toString() + "_" + prename + "_" + sortmap.get(sortname);
				String name = "";
				if (renameList.contains(shname1)) {
					name = FilterSpecialCharacters(shname2);
					if (name.length() > 31) {
						errorname.add(shname2);
					}
				} else {
					name = FilterSpecialCharacters(shname1);
					if (name.length() > 31) {
						errorname.add(shname1);
					}
				}
			}
			if (errorname.size() > 0) {
				return null;
			}
			// 避免sheet命名冲突，先重新命名一次
			for (int j = 0; j < book.getNumberOfSheets(); j++) {
				book.setSheetName(j, "A" + j);
			}

			for (int j = 0; j < book.getNumberOfSheets(); j++) {
				boolean flag = false;
				String sheetname = book.getSheetName(j);
				XSSFSheet sheet = book.getSheetAt(j);
				String dreaid = "";
				XSSFRow row = sheet.getRow(0);
				if (row != null) {
					XSSFCell cell = row.getCell(26);
					if (cell != null) {
						dreaid = cell.getStringCellValue().trim();
					}
				}
				System.out.println("获取的点焊工序ID：" + dreaid);
				for (int k = 0; k < list.size(); k++) {
					String[] str = (String[]) list.get(k);
					String factdreaid = str[10];
					// 先通过点焊工序ID取匹配，如果未匹配到 再根据点焊工序名称匹配（为了解决目前系统已产生的以点焊工序名称命名的数据）
					if (dreaid.equals(factdreaid)) {
						String[] statename = str[0].toString().split("\\\\");
						String prename = statename[0].toString();
						// String compareshname = str[8].toString() + "_" + str[0].toString();
						String shname1 = str[8].toString() + "_" + prename;
						String sortname = str[8].toString() + "_" + prename + "_" + str[10].toString();
						String shname2 = str[8].toString() + "_" + prename + "_" + sortmap.get(sortname);
						String name = "";
						if (renameList.contains(shname1)) {
							name = FilterSpecialCharacters(shname2);
						} else {
							name = FilterSpecialCharacters(shname1);
						}
						int sheetIndex = book.getSheetIndex(sheet);
						if (!sheet.getSheetName().equals(name)) {
							book.setSheetName(sheetIndex, name);
						}
						flag = true;
						continue;
					} else {
						String shname = str[0].toString();
						if (shname.equals(sheetname)) {
							String[] statename = str[0].toString().split("\\\\");
							String prename = statename[0].toString();
							// String compareshname = str[8].toString() + "_" + str[0].toString();
							String shname1 = str[8].toString() + "_" + prename;
							String sortname = str[8].toString() + "_" + prename + "_" + str[10].toString();
							String shname2 = str[8].toString() + "_" + prename + "_" + sortmap.get(sortname);
							String name = "";
							if (renameList.contains(shname1)) {
								name = FilterSpecialCharacters(shname2);
							} else {
								name = FilterSpecialCharacters(shname1);
							}
							int sheetIndex = book.getSheetIndex(sheet);
							book.setSheetName(sheetIndex, name);
							flag = true;
							continue;
						}

					}
				}
				if (!flag) {
					book.removeSheetAt(j);
				}
			}
			int st = book.getNumberOfSheets();
			System.out.println("移除后数量：" + st);

			// 循环数据集，添加sheet
			for (int i = 0; i < list.size(); i++) {
				String[] str = (String[]) list.get(i);
				String[] statename = str[0].toString().split("\\\\");
				String prename = statename[0].toString();
				// String compareshname = str[8].toString() + "_" + str[0].toString();
				String shname1 = str[8].toString() + "_" + prename;
				String sortname = str[8].toString() + "_" + prename + "_" + str[10].toString();
				String shname2 = str[8].toString() + "_" + prename + "_" + sortmap.get(sortname);
				XSSFSheet sheet = null;
				if (renameList.contains(shname1)) {
					sheet = book.getSheet(FilterSpecialCharacters(shname1));
					if (sheet == null) {
						sheet = book.getSheet(FilterSpecialCharacters(shname2));
					}
					if (sheet != null) {
						int sheetIndex = book.getSheetIndex(sheet);
						book.setSheetName(sheetIndex, FilterSpecialCharacters(shname2));
					}
				} else {
					sheet = book.getSheet(FilterSpecialCharacters(shname1));
					if (sheet == null) {
						sheet = book.getSheet(FilterSpecialCharacters(shname2));
						if (sheet != null) {
							int sheetIndex = book.getSheetIndex(sheet);
							book.setSheetName(sheetIndex, FilterSpecialCharacters(shname1));
						}
					}
				}
				// 如果不存在则创建
				if (sheet == null) {
					sheet = book.cloneSheet(0);
					if (renameList.contains(shname1)) {
						book.setSheetName(st, FilterSpecialCharacters(shname2));
					} else {
						book.setSheetName(st, FilterSpecialCharacters(shname1));
					}
					//////////// 设置分组显示上方/下方
					sheet.setRowSumsBelow(false);
					sheet.setRowSumsRight(false);
					sheet.setRowSumsBelow(false);
					sheet.setRowSumsRight(false);
					st++;
				}
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	// 写sheet数据，点焊工序数据写入
	public static void writeDiscreteDataToSheet(XSSFWorkbook book, ArrayList list, ArrayList weldlist,
			boolean setGroup) {
		// TODO Auto-generated method stub
		// int row = 3;

		int len = list.size();
		// 点焊工序名称数据集
		List nameList = new ArrayList();
		List renameList = new ArrayList();
		Map<String, List<String>> map = new LinkedHashMap();
		for (int i = 0; i < list.size(); i++) {
			String[] str = (String[]) list.get(i);
			String[] statename = str[0].toString().split("\\\\");
			String prestatename = statename[0].toString();
			String prename = str[8].toString() + "_" + prestatename;
			if (nameList.contains(prename)) {
				if (!renameList.contains(prename)) {
					renameList.add(prename);
				}
			}
			if (!map.containsKey(prename)) {
				List<String> numlist = new ArrayList();
				numlist.add(str[10].toString());
				map.put(prename, numlist);
			} else {
				List<String> numlist = map.get(prename);
				numlist.add(str[10].toString());
				map.put(prename, numlist);
			}
			nameList.add(prename);
		}
		// 对相同工位相同点焊工序名称按照查找编号排序
		Map<String, Integer> sortmap = new HashMap<>();
		for (Map.Entry<String, List<String>> entry : map.entrySet()) {
			String key = entry.getKey();
			List<String> valuelist = entry.getValue();
			for (int i = 0; i < valuelist.size(); i++) {
				String sortnum = valuelist.get(i);
				String sortkey = key + "_" + sortnum;
				sortmap.put(sortkey, i + 1);
			}
		}
		System.out.println("有重复名称的：" + renameList);
		System.out.println("sheet名称map：" + sortmap);

		for (int i = 0; i < len; i++) {
			String[] values = (String[]) list.get(i);
			XSSFSheet sheet = null;
			String[] statename = values[0].toString().split("\\\\");
			String prename = statename[0].toString();
			// String compareshname = values[8].toString() + "_" + values[0].toString();
			String shname1 = values[8].toString() + "_" + prename;
			if (renameList.contains(shname1)) {
				String sortshname = values[8].toString() + "_" + prename + "_" + values[10].toString();
				String shname2 = values[8].toString() + "_" + prename + "_" + sortmap.get(sortshname);
				sheet = book.getSheet(FilterSpecialCharacters(shname2));
			} else {
				sheet = book.getSheet(FilterSpecialCharacters(shname1));
			}

			// 机器人信息改为放在头两行，点焊信息数据 {点焊工序名称，日期，机器人编号，机器人型号，焊枪枪号，焊装工位，车型}
			setCell(sheet, values[1], 1, 11, Cell.CELL_TYPE_STRING);// 日期
			setCell(sheet, values[2], 0, 6, Cell.CELL_TYPE_STRING);// 机器人编号
			setCell(sheet, values[3], 0, 13, Cell.CELL_TYPE_STRING);// 机器人型号
			setCell(sheet, values[4], 1, 6, Cell.CELL_TYPE_STRING);// 焊枪枪号
			setCell(sheet, values[5], 1, 2, Cell.CELL_TYPE_STRING);// 焊装工位
			setCell(sheet, values[6], 0, 2, Cell.CELL_TYPE_STRING);// 车型
			setCell(sheet, i + 1 + "/" + len, 1, 14, Cell.CELL_TYPE_STRING);// 页数

			// 点焊工序ID
			setCell(sheet, values[10], 0, 26, Cell.CELL_TYPE_STRING);// 车型
			// row++;
			int witd = 17;
			if (weldlist.size() > 15) {
				witd = witd + weldlist.size() - 15;
			}
			// 设置打印区域
			book.setPrintArea(i, 0, 14, 0, witd);
			PrintSetup printSetup = sheet.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 94);// 自定义缩放，此处100为无缩放
			printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
		}

	}

	// 写sheet数据，点焊工序数据写入
	public static void UpdateDiscreteDataToSheet(XSSFWorkbook book, ArrayList list, ArrayList weldlist,
			boolean setGroup) {
		// TODO Auto-generated method stub
		// int row = 3;

		int len = list.size();
		List nameList = new ArrayList();
		List renameList = new ArrayList();
		Map<String, List<String>> map = new LinkedHashMap();
		for (int i = 0; i < list.size(); i++) {
			String[] str = (String[]) list.get(i);
			String[] statename = str[0].toString().split("\\\\");
			String prestatename = statename[0].toString();
			String prename = str[8].toString() + "_" + prestatename;
			if (nameList.contains(prename)) {
				if (!renameList.contains(prename)) {
					renameList.add(prename);
				}
			}
			if (!map.containsKey(prename)) {
				List<String> numlist = new ArrayList();
				numlist.add(str[10].toString());
				map.put(prename, numlist);
			} else {
				List<String> numlist = map.get(prename);
				numlist.add(str[10].toString());
				map.put(prename, numlist);
			}
			nameList.add(prename);
		}
		// 对相同工位相同点焊工序名称按照查找编号排序
		Map<String, Integer> sortmap = new HashMap<>();
		for (Map.Entry<String, List<String>> entry : map.entrySet()) {
			String key = entry.getKey();
			List<String> valuelist = entry.getValue();
			for (int i = 0; i < valuelist.size(); i++) {
				String sortnum = valuelist.get(i);
				String sortkey = key + "_" + sortnum;
				sortmap.put(sortkey, i + 1);
			}
		}
		System.out.println("有重复名称的：" + renameList);
		System.out.println("sheet名称map：" + sortmap);

		//

		for (int i = 0; i < len; i++) {
			String[] values = (String[]) list.get(i);
			XSSFSheet sheet = null;
			String direid = values[10].toString();
			for (int j = 0; j < book.getNumberOfSheets(); j++) {
				XSSFSheet sh = book.getSheetAt(j);
				String factid = "";
				XSSFRow row = sh.getRow(0);
				if (row != null) {
					XSSFCell cell = row.getCell(26);
					if (cell != null) {
						factid = cell.getStringCellValue().trim();
					}
				}
				if (factid.equals(direid)) {
					sheet = sh;
				}
			}
			if (sheet == null) // 处理之前数据没有写点焊工序ID的数据
			{
				String[] statename = values[0].toString().split("\\\\");
				String prename = statename[0].toString();
				// String compareshname = values[8].toString() + "_" + values[0].toString();
				String shname1 = values[8].toString() + "_" + prename;
				String sortname = values[8].toString() + "_" + prename + "_" + values[10].toString();
				String shname2 = values[8].toString() + "_" + prename + "_" + sortmap.get(sortname);
				if (renameList.contains(shname1)) {
					sheet = book.getSheet(FilterSpecialCharacters(shname2));
				} else {
					sheet = book.getSheet(FilterSpecialCharacters(shname1));
				}
			}

			// 机器人信息改为放在头两行，点焊信息数据 {点焊工序名称，日期，机器人编号，机器人型号，焊枪枪号，焊装工位，车型}
			setCell(sheet, values[1], 1, 11, Cell.CELL_TYPE_STRING);// 日期
			setCell(sheet, values[2], 0, 6, Cell.CELL_TYPE_STRING);// 机器人编号
			setCell(sheet, values[3], 0, 13, Cell.CELL_TYPE_STRING);// 机器人型号
			setCell(sheet, values[4], 1, 6, Cell.CELL_TYPE_STRING);// 焊枪枪号
			setCell(sheet, values[5], 1, 2, Cell.CELL_TYPE_STRING);// 焊装工位
			setCell(sheet, values[6], 0, 2, Cell.CELL_TYPE_STRING);// 车型
			setCell(sheet, i + 1 + "/" + len, 1, 14, Cell.CELL_TYPE_STRING);// 页数
			// 点焊工序ID
			setCell(sheet, values[10], 0, 26, Cell.CELL_TYPE_STRING);// 车型
			// row++;

			int witd = 17;
			if (weldlist.size() > 15) {
				witd = witd + weldlist.size() - 15;
			}
			// 设置打印区域
			book.setPrintArea(i, 0, 14, 0, witd);
			PrintSetup printSetup = sheet.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 94);// 自定义缩放，此处100为无缩放
			printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
		}

	}

	// 写sheet数据，焊点数据写入
	public static void writeWeldDataToSheet(XSSFWorkbook book, ArrayList list, ArrayList partlist, boolean setGroup) {
		// TODO Auto-generated method stub
		// int row = 3;
		// 设置字体为微软雅黑

		CellStyle style = book.createCellStyle();// 新建样式对象
		XSSFFont font = (XSSFFont) book.createFont();// 创建字体对象
		font.setFontName("微软雅黑");
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);
		style.setFont(font);

		int len = list.size();

		// 点焊工序名称数据集
		List nameList = new ArrayList();
		List renameList = new ArrayList();
		Map<String, List<String>> map = new LinkedHashMap();
		for (int i = 0; i < partlist.size(); i++) {
			String[] str = (String[]) partlist.get(i);
			String[] statename = str[0].toString().split("\\\\");
			String prestatename = statename[0].toString();
			String prename = str[8].toString() + "_" + prestatename;
			if (nameList.contains(prename)) {
				if (!renameList.contains(prename)) {
					renameList.add(prename);
				}
			}
			if (!map.containsKey(prename)) {
				List<String> numlist = new ArrayList();
				numlist.add(str[10].toString());
				map.put(prename, numlist);
			} else {
				List<String> numlist = map.get(prename);
				numlist.add(str[10].toString());
				map.put(prename, numlist);
			}
			nameList.add(prename);
		}
		// 对相同工位相同点焊工序名称按照查找编号排序
		Map<String, Integer> sortmap = new HashMap<>();
		for (Map.Entry<String, List<String>> entry : map.entrySet()) {
			String key = entry.getKey();
			List<String> valuelist = entry.getValue();
			for (int i = 0; i < valuelist.size(); i++) {
				String sortnum = valuelist.get(i);
				String sortkey = key + "_" + sortnum;
				sortmap.put(sortkey, i + 1);
			}
		}
		System.out.println("有重复名称的：" + renameList);
		System.out.println("sheet名称map：" + sortmap);

		for (int i = 0; i < len; i++) {
			String[] values = (String[]) list.get(i);
			XSSFSheet sheet = null;
			String[] statename = values[0].toString().split("\\\\");
			String prename = statename[0].toString();
			// String compareshname = values[5].toString() + "_" + values[0].toString();
			String shname1 = values[5].toString() + "_" + prename;

			if (renameList.contains(shname1)) {
				String sortname = values[5].toString() + "_" + prename + "_" + values[7].toString();
				String shname2 = values[5].toString() + "_" + prename + "_" + sortmap.get(sortname);
				sheet = book.getSheet(FilterSpecialCharacters(shname2));
			} else {
				sheet = book.getSheet(FilterSpecialCharacters(shname1));
			}
			// 焊点信息数组{点焊工序名称，序号，焊点ID}
			for (int j = 1; j < values.length - 3; j++) {

				setCellwithCellStyle(sheet, values[j], Integer.parseInt(values[1]) + 2, j + 10, Cell.CELL_TYPE_STRING,
						style);
			}
			// row++;
		}

	}

	// 更新sheet数据，焊点数据更新
	public static void updateWeldDataToSheet(XSSFWorkbook book, ArrayList list, ArrayList partlist, boolean setGroup) {
		// TODO Auto-generated method stub

		// 设置字体为微软雅黑

		CellStyle style = book.createCellStyle();// 新建样式对象
		XSSFFont font = (XSSFFont) book.createFont();// 创建字体对象
		font.setFontName("微软雅黑");
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);
		style.setFont(font);

		int len = list.size();

		List nameList = new ArrayList();
		List renameList = new ArrayList();
		Map<String, List<String>> map = new LinkedHashMap();
		for (int i = 0; i < partlist.size(); i++) {
			String[] str = (String[]) partlist.get(i);
			String[] statename = str[0].toString().split("\\\\");
			String prestatename = statename[0].toString();
			String prename = str[8].toString() + "_" + prestatename;
			if (nameList.contains(prename)) {
				if (!renameList.contains(prename)) {
					renameList.add(prename);
				}
			}
			if (!map.containsKey(prename)) {
				List<String> numlist = new ArrayList();
				numlist.add(str[10].toString());
				map.put(prename, numlist);
			} else {
				List<String> numlist = map.get(prename);
				numlist.add(str[10].toString());
				map.put(prename, numlist);
			}
			nameList.add(prename);
		}
		// 对相同工位相同点焊工序名称按照查找编号排序
		Map<String, Integer> sortmap = new HashMap<>();
		for (Map.Entry<String, List<String>> entry : map.entrySet()) {
			String key = entry.getKey();
			List<String> valuelist = entry.getValue();
			for (int i = 0; i < valuelist.size(); i++) {
				String sortnum = valuelist.get(i);
				String sortkey = key + "_" + sortnum;
				sortmap.put(sortkey, i + 1);
			}
		}
		System.out.println("有重复名称的：" + renameList);
		System.out.println("sheet名称map：" + sortmap);

		// 循环sheet页
		for (int i = 0; i < partlist.size(); i++) {
			String[] partstr = (String[]) partlist.get(i);
			String dieaid = partstr[10].toString();
			String sheetname = "";
			String[] statename = partstr[0].toString().split("\\\\");
			String prename = statename[0].toString();
			// String compareshname = partstr[8].toString() + "_" + partstr[0].toString();
			String shname1 = partstr[8].toString() + "_" + prename;
			String sortname = partstr[8].toString() + "_" + prename + "_" + partstr[10].toString();
			String shname2 = partstr[8].toString() + "_" + prename + "_" + sortmap.get(sortname);
			if (renameList.contains(shname1)) {
				sheetname = FilterSpecialCharacters(shname2);
			} else {
				sheetname = FilterSpecialCharacters(shname1);
			}
			ArrayList hdlist = new ArrayList();
			for (int j = 0; j < list.size(); j++) {
				String[] str = (String[]) list.get(j);
				String welddreaid = str[7].toString();
				if (dieaid.equals(welddreaid)) {
					hdlist.add(str);
				}
			}
			// 根据点焊工序ID匹配sheet页
			XSSFSheet sheet = null;
			for (int j = 0; j < book.getNumberOfSheets(); j++) {
				XSSFSheet sh = book.getSheetAt(j);
				String factid = "";
				XSSFRow row = sh.getRow(0);
				if (row != null) {
					XSSFCell cell = row.getCell(26);
					if (cell != null) {
						factid = cell.getStringCellValue().trim();
					}
				}
				if (factid.equals(dieaid)) {
					sheet = sh;
				}
			}
			if (sheet == null) {
				sheet = book.getSheet(FilterSpecialCharacters(sheetname));
			}
			int oldnum = sheet.getPhysicalNumberOfRows();
			int newnum = hdlist.size();
			for (int k = 0; k < newnum; k++) {
				// 判断新的焊点信息是否在旧的焊点数据中，如果存在就记录焊点信息：P=和S=
				String[] hdstr = (String[]) hdlist.get(k);
				boolean flag = false;
				for (int m = 3; m < oldnum; m++) {
					XSSFRow row = sheet.getRow(m);
					if (row != null) {
						// 比较weid_id
						System.out.println(hdstr[2]);
						XSSFCell cell = row.getCell(12);
						if (cell == null) {
							cell = row.createCell(12);
						}
						if (hdstr[2].equals(getCellValue(row.getCell(12)))) {
							flag = true;
							continue;
						}

					}
				}
				XSSFRow hzrow = sheet.getRow(Integer.parseInt(hdstr[1]) + 2);
				// 如果存在就保存P、S值
				if (flag) {
					if (hzrow != null) {
						hdstr[3] = getCellValue(hzrow.getCell(13));// P=
						hdstr[4] = getCellValue(hzrow.getCell(14));// S=
					}
				}
				// 焊点信息数组{点焊工序名称，序号，焊点ID}
				for (int n = 1; n < hdstr.length - 3; n++) {
					setCellwithCellStyle(sheet, hdstr[n], Integer.parseInt(hdstr[1]) + 2, n + 10, Cell.CELL_TYPE_STRING,
							style);
				}
			}
			// 如果就老数据比新数据多需要把多余的行内容置空
			System.out.println("老数据：" + oldnum + " 新数据" + newnum + 3);
			if (oldnum > newnum + 3) {
				for (int s = oldnum; s > newnum + 3; s--) {
					XSSFRow oldrow = sheet.getRow(s - 1);
					// sheet.removeRow(oldrow);
					if (oldrow != null) {
						for (int k = 0; k < 4; k++) {
							XSSFCell ce = oldrow.getCell(11 + k);
							if (ce != null) {
								ce.setCellType(XSSFCell.CELL_TYPE_BLANK);
							}
						}
					}
				}
			}
		}
	}

	// 写直劳清单sheet数据
	public static void writeDataToSheet(XSSFWorkbook book, ArrayList list, int num, int num2, int rownum,
			ReportViwePanel viewPanel, int ajm) {
		// TODO Auto-generated method stub
		// int row = 3;
		// XSSFFormulaEvaluator evaluator=new XSSFFormulaEvaluator(book);
		viewPanel.addInfomation("", 40, 100);

		XSSFSheet sheet = book.getSheetAt(0);
		System.out.println("添加前行数：" + sheet.getPhysicalNumberOfRows());
		// 先写两个汇总数据
		setStringCell(sheet, "总焊点数" + num, 1, 4, false);
		setStringCell(sheet, "RSW总数" + num2, 1, 10, false);
		// 将汇总行向下移动list.get(list.size()-2)
		int downnum = (int) list.get(list.size() - 2);
		if (rownum == 11) {
			downnum = downnum - 1; // 如果是第一次，已经存在一行，所以减去
		}
		// 解决只有metal且只有一条数据的情况下，出现报错问题
		if (downnum > 0) {
			sheet.shiftRows(rownum, rownum + 12, downnum, true, false);
		}

		System.out.println("添加后行数：" + sheet.getPhysicalNumberOfRows());

		// 通过模板行复制行用于写数据，最后的时候在删除这个模板行
		String formula = "";
		for (int i = 0; i < downnum; i++) {
			viewPanel.addInfomation("", 40, 100);
			copyRows(sheet, 11, 11, rownum + i);
			// I列公式 0.22*D11+0.11*E11+0.08*F11+0.05*G11+0.04*H11
			String cs = Integer.toString((rownum + i + 1));
			formula = "0.22*D" + cs + "+0.11*E" + cs + "+0.08*F" + cs + "+0.05*G" + cs + "+0.04*H" + cs;
			setCellFormula(sheet, formula, rownum + i, 8);
			System.out.println(formula);
			// M列公式 0.06*J11+0.022*L11
			formula = "0.06*J11+0.022*L11";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 12);
			// X列公式
			// IF(ISBLANK(N11),0,0.055*N11+0.11)+IF(ISBLANK(O11),0,0.03*O11+0.2)+IF(ISBLANK(P11),0,0.035*P11+0.12)+IF(ISBLANK(Q11),0,0.03*Q11+0.4)+IF(ISBLANK(R11),0,0.04*R11+0.11)+IF(ISBLANK(S11),0,0.04*S11+0.11)+IF(ISBLANK(T11),0,0.04*T11+0.11)+IF(ISBLANK(V11),0,0.03*V11+0.07)+IF(ISBLANK(W11),0,0.2*W11+0.07)
			formula = "IF(ISBLANK(N" + cs + "),0,0.055*N" + cs + "+0.11)+IF(ISBLANK(O" + cs + "),0,0.03*O" + cs
					+ "+0.2)+IF(ISBLANK(P" + cs + "),0,0.035*P" + cs + "+0.12)+IF(ISBLANK(Q" + cs + "),0,0.03*Q" + cs
					+ "+0.4)+IF(ISBLANK(R" + cs + "),0,0.04*R" + cs + "+0.11)+IF(ISBLANK(S" + cs + "),0,0.04*S" + cs
					+ "+0.11)+IF(ISBLANK(T" + cs + "),0,0.04*T" + cs + "+0.11)+IF(ISBLANK(V" + cs + "),0,0.03*V" + cs
					+ "+0.07)+IF(ISBLANK(W" + cs + "),0,0.2*W" + cs + "+0.07)";
			setCellFormula(sheet, formula, rownum + i, 23);
			// AG列公式
			// IF(ISBLANK(Y11),0,0.06*Y11+0.06)+IF(ISBLANK(Z11),0,0.05*Z11+0.22)+IF(ISBLANK(AA11),0,0.05*AA11)+IF(ISBLANK(AB11),0,0.5*AB11+0.18)+IF(ISBLANK(AC11),0,0.025*AC11+0.35)+IF(ISBLANK(AD11),0,0.01*AD11+0.21)+IF(ISBLANK(AE11),0,0.5*AE11+0.21)+IF(ISBLANK(AF11),0,2*AF11)
			formula = "IF(ISBLANK(Y11),0,0.06*Y11+0.06)+IF(ISBLANK(Z11),0,0.05*Z11+0.22)+IF(ISBLANK(AA11),0,0.05*AA11)+IF(ISBLANK(AB11),0,0.5*AB11+0.18)+IF(ISBLANK(AC11),0,0.025*AC11+0.35)+IF(ISBLANK(AD11),0,0.01*AD11+0.21)+IF(ISBLANK(AE11),0,0.5*AE11+0.21)+IF(ISBLANK(AF11),0,2*AF11)";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 32);
			// AH列公式 (I11+M11+X11+AG11)
			formula = "(I11+M11+X11+AG11)";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 33);
			// AJ列公式 AH11*AI11
			formula = "AH11*AI11";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 35);
			// AL列公式 AH11*AI11*AK11
			formula = "AH11*AI11*AK11";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 37);
			// AM列公式 ($AG$6*60*$AG$5)/$AG$1

			// AN列公式 AL11/AM11
			formula = "AL11/AM11";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 39);

			// AP列公式 AO11*AG$3/AG$3
			formula = "AO11*AG$3/AG$3";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 41);
			// AQ列公式 AP11*$AG$6/AG$3
			formula = "AP11*$AG$6/AG$3";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 42);
			// }
		}
		System.out.println("复制后行数：" + sheet.getPhysicalNumberOfRows());
		// 清空后两行内容

		// 设置焊装产线统计行的单元格公式
		int rowIndex = downnum + rownum - 1;
		// 开始和结束行统计
		int start = rownum + 1;
		int end = downnum + rownum - 1;
		if (rownum == 11) {
			start = start - 1;
		}
		String[] strcell = { "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
				"T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK",
				"AL", "AM", "AN", "AO", "AP", "AQ" };
		// 工艺焊不添加check等3行
		if (ajm != 2) {
			for (int j = 0; j < strcell.length; j++) {

				viewPanel.addInfomation("", 40, 100);

				// 设置产线合计行
				if (strcell[j] != "B" && strcell[j] != "C" && strcell[j] != "AI" && strcell[j] != "AK"
						&& strcell[j] != "AM" && strcell[j] != "AP") {
					formula = "SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")";
					System.out.println("产线合计行：" + formula);
					setCellFormula(sheet, formula, rowIndex, 1 + j);
				} else {
					setStringCell(sheet, "", rowIndex, 1 + j, true);
				}
				// 设置产线合计行的上一行数据都为空
				setStringCell(sheet, "", rowIndex - 1, 1 + j, true);

				// 设置WELD CHECKER行导出数据列为空
				if (strcell[j] == "D" || strcell[j] == "E" || strcell[j] == "F" || strcell[j] == "G"
						|| strcell[j] == "H" || strcell[j] == "J" || strcell[j] == "K" || strcell[j] == "L") {
					setIntCell(sheet, null, rowIndex - 2, 1 + j, true);
				}

			}
			// 设置产线合计行的上上行
			setStringCell(sheet, list.get(list.size() - 1).toString() + " WELD CHECKER", rowIndex - 2, 2, true);
			setIntCell(sheet, Integer.toString((int) list.get(list.size() - 2) - 2), rowIndex - 2, 1, true);
			// 设置产线合计行
			setStringCell(sheet, list.get(list.size() - 1).toString() + " SUB TOTAL", rowIndex, 2, true);

		} else { // 设置METAL统计，跟业务已确认，METAL产线只有一个且在最后面
			for (int j = 0; j < strcell.length; j++) {
				// 设置METAL统计行
				if (strcell[j] != "B" && strcell[j] != "C" && strcell[j] != "AI" && strcell[j] != "AK"
						&& strcell[j] != "AM" && strcell[j] != "AP") {
					formula = "SUM(" + strcell[j] + start + ":" + strcell[j] + (end + 5) + ")";
					System.out.println("METAL统计行行：" + formula);
					setCellFormula(sheet, formula, rowIndex + 5, 1 + j);
				}
			}
		}

		int len = list.size() - 2;

		if (rownum == 11)
			for (int i = 0; i < len; i++) {
				String[] values = (String[]) list.get(i);
				setCell2(sheet, values, rownum + i - 1);
			}
		else {
			for (int i = 0; i < len; i++) {
				String[] values = (String[]) list.get(i);
				setCell2(sheet, values, rownum + i);
			}
		}

	}

	// 写工程作业表-封面sheet数据
	public static void writeDataToSheet(XSSFWorkbook book, String[] values) {
		XSSFSheet sheet = book.getSheetAt(0);

		// 设置字体
//		Font font = book.createFont();
//		font.setFontName("新宋体");
//		font.setBoldweight(Font.BOLDWEIGHT_BOLD);// 加粗
//		font.setFontHeightInPoints((short) 16);
//		// 创建一个样式
//
//		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_NONE); // 下边框
//		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_NONE);// 左边框
//		cellStyle.setBorderTop(XSSFCellStyle.BORDER_NONE);// 上边框
//		cellStyle.setBorderRight(XSSFCellStyle.BORDER_NONE);// 右边框
//		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// 左对齐
//		cellStyle.setFont(font);
		cellStyle = null;

		setStringCell(sheet, values[4], 5, 2, true);
		setStringCell(sheet, values[0], 6, 2, true);
		setStringCell(sheet, values[1], 7, 2, true);
		setStringCell(sheet, values[2], 8, 2, true);
		setStringCell(sheet, values[3], 10, 2, true);

	}

	// 写封面sheet数据
	public static void writeDataToSheetByGeneral(XSSFWorkbook book, String[] values) {
		XSSFSheet sheet = book.getSheetAt(0);

		// 设置字体
		Font font = book.createFont();
		font.setFontName("新宋体");
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);// 加粗
		font.setFontHeightInPoints((short) 16);
		// 创建一个样式

		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_NONE); // 下边框
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_NONE);// 左边框
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_NONE);// 上边框
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_NONE);// 右边框
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// 左对齐
		cellStyle.setFont(font);

		setStringCell(sheet, values[0], 9, 3, true);
		setStringCell(sheet, values[1], 10, 3, true);
		setStringCell(sheet, values[2], 12, 1, false);
	}

	// 写直材消耗定额清单数据
	public static void writeStraightDataToSheet(XSSFWorkbook book, int sheetnum, String[] values, ArrayList list) {

		// 设置字体
		Font font = book.createFont();
		font.setColor((short) 12);
		font.setFontHeightInPoints((short) 12);
		// 创建一个样式
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中

		cellStyle.setFont(font);

		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i + 1);

			// 写编制信息、编制日期、当前页、总页
			setStringCell(sheet, values[2], 2, 2, true);
			setStringCell(sheet, values[3], 2, 7, true);
			setStringCell(sheet, Integer.toString(i + 1), 25, 23, false);
			setStringCell(sheet, Integer.toString(sheetnum), 26, 23, false);

			int n = 0;
			for (int j = 18 * i; j < list.size(); j++) {
				if (n == 18) {
					break;
				}
				String[] value = (String[]) list.get(j);
				setIntCell(sheet, value[0], 5 + n, 1, false);
				setStringCell(sheet, value[1], 5 + n, 2, false);
				setStringCell(sheet, value[2], 5 + n, 4, false);
				setStringCell(sheet, value[3], 5 + n, 6, false);
				setStringCell(sheet, value[4], 5 + n, 8, false);
				setStringCell(sheet, value[5], 5 + n, 10, false);

				setDoubleCell(sheet, value[6], 5 + n, 11, false);
				setDoubleCell(sheet, value[7], 5 + n, 13, false);
				setDoubleCell(sheet, value[8], 5 + n, 15, false);
				setDoubleCell(sheet, value[9], 5 + n, 17, false);
				setDoubleCell(sheet, value[10], 5 + n, 18, false);
				setDoubleCell(sheet, value[11], 5 + n, 19, false);

				setStringCell(sheet, value[12], 5 + n, 20, false);
				setStringCell(sheet, value[13], 5 + n, 21, false);
				setStringCell(sheet, value[14], 5 + n, 23, false);
				n++;

			}

		}

	}

	// 写工程作业表-目录sheet数据
	public static XSSFWorkbook creatXSSFWorkbook(InputStream input, ArrayList list, List bzlist) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);
			String sheetname = "";
			int page = 1; // 用于标记sheet序号
			// 根据目录数据判断是否分多个sheet页，每40条数据一个sheet
			int mlnum = list.size();
			int mlpage = mlnum / 40 + 1;
			if (mlpage > 1) {
				sheetname = "01目录1";
				book.setSheetName(0, sheetname);
			}
			// 目录分sheet页逻辑
			for (int i = 1; i < mlpage; i++) {
				XSSFSheet sheet = book.cloneSheet(0);
				sheetname = String.format("%02d", i + 1) + "目录" + (i + 1);
				book.setSheetName(book.getSheetIndex(sheet), sheetname);
				book.setSheetOrder(sheetname, i); // 调整sheet顺序
				// 设置打印区域
				int sheetIndex = book.getSheetIndex(sheet);
				book.setPrintArea(sheetIndex, 0, 115, 0, 51);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 65);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)

				page = i + 1;
			}
			int bznnum = 0;
			// 板组清单sheet分页逻辑
			if (bzlist != null) {
				bznnum = bzlist.size();
			}
			int bzpage = bznnum / 40 + 1;
			if (bzpage > 1) {
				sheetname = String.format("%02d", page + 1) + "附录-板组清单" + 1;
				book.setSheetName(page, sheetname);
			} else {
				sheetname = String.format("%02d", page + 1) + "附录-板组清单";
				page = page + 1;
				book.setSheetName(page - 1, sheetname);
			}

			int cnum = page;
			for (int j = 1; j < bzpage; j++) {
				XSSFSheet sheet = book.cloneSheet(cnum);
				sheetname = String.format("%02d", cnum + j + 1) + "附录-板组清单" + (j + 1);
				book.setSheetName(book.getSheetIndex(sheet), sheetname);
				book.setSheetOrder(sheetname, cnum + j); // 调整sheet顺序
				// 设置打印区域
				int sheetIndex = book.getSheetIndex(sheet);
				book.setPrintArea(sheetIndex, 0, 115, 0, 51);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 65);// 自定义缩放，此处100为无缩放
				printSetup.setLandscape(true); // 打印方向，true：横向，false：纵向(默认)
				page = cnum + j + 1;
			}
			// 处理固定模板的sheet命名
			int sheetnum = book.getNumberOfSheets();
			sheetname = String.format("%02d", page + 1) + "附录-255参数序列1";
			book.setSheetName(sheetnum - 5, sheetname); // 03附录-255参数序列1
			sheetname = String.format("%02d", page + 2) + "附录-255参数序列2";
			book.setSheetName(sheetnum - 4, sheetname);// 04附录-255参数序列2
			sheetname = String.format("%02d", page + 3) + "附录-255参数序列3";
			book.setSheetName(sheetnum - 3, sheetname);// 05附录-255参数序列3
			sheetname = String.format("%02d", page + 4) + "附录-24参数序列";
			book.setSheetName(sheetnum - 2, sheetname);// 06附录-24参数序列
			sheetname = String.format("%02d", page + 5) + "附录-序列对照表";
			book.setSheetName(sheetnum - 1, sheetname);// 07附录-序列对照表)

			cellStyle = book.createCellStyle();
			// 边框
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle.setWrapText(true);
			cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
			cellStyle.setAlignment(XSSFCellStyle.VERTICAL_CENTER);
			// cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// 左对齐

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	// 写工程作业表-目录sheet数据
	public static void writeDataToSheet(XSSFWorkbook book, ArrayList plist, ArrayList list, List bzlist) {

		// 设置字体
//		Font font = book.createFont();
//		font.setColor((short) 12);
//		font.setFontHeightInPoints((short) 16);
		// 创建一个样式
		cellStyle = null;
//		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle.setFont(font);
//		Font font2 = book.createFont();
//		font2.setFontHeightInPoints((short) 18);
//		font2.setFontName("MS PGothic");
//		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
//		XSSFCellStyle cellStyle1 = book.createCellStyle();
		XSSFCellStyle cellStyle1 = null;
//		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_NONE);
//		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle1.setFont(font2);

		// 循环所有sheet页，把公共部分内容写入
		int sheetnum = book.getNumberOfSheets();
		for (int n = 0; n < sheetnum; n++) {
			XSSFSheet sh = book.getSheetAt(n);
			setStringCell(sh, plist.get(4).toString(), 2, 0, true);
			setStringCell(sh, plist.get(0).toString(), 2, 6, true);
			setStringCell(sh, plist.get(1).toString(), 2, 30, true);
			setStringCell(sh, plist.get(2).toString(), 2, 90, true);
			if (n < sheetnum - 5) {
				setStringCellAndStyle(sh, plist.get(3).toString(), 48, 110, cellStyle1);
			} else {
				setStringCellAndStyle(sh, plist.get(3).toString(), 48, 108, cellStyle1);
			}

		}

		// 写目录数据
		int mlpage = list.size() / 40 + 1; // 目录数据sheet数
		int totalpages = 0;// 合计总页数
		for (int i = 0; i < mlpage; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			sheet.removeColumnBreak(55);
			if (i == mlpage - 1) {
				for (int j = 0; j + 40 * i < list.size(); j++) {
					String[] value = (String[]) list.get(j + 40 * i);
					String rowNo = Integer.toString(j + 1);
					setIntCell(sheet, rowNo, 7 + j, 1, false);// 序号
					setStringCell(sheet, value[1], 7 + j, 4, false);// 工序编号
					setStringCell(sheet, value[2], 7 + j, 14, false);// 工序中文名称
					setStringCell(sheet, value[3], 7 + j, 44, false);// 工序英文名称
					setStringCell(sheet, value[4], 7 + j, 71, false);// 版次
					setIntCell(sheet, value[5], 7 + j, 76, false);// 页数
					setStringCell(sheet, value[6], 7 + j, 81, false);// 备注
					if (value[5] != null && !value[5].isEmpty()) {
						totalpages = totalpages + Integer.parseInt(value[5]);
					}
				}
				setStringCell(sheet, "合计", 46, 4, false);// 合计
				setIntCell(sheet, Integer.toString(totalpages), 46, 76, false);// 备注
			} else {
				for (int j = 0; j + 40 * i < 40 + 40 * i; j++) {
					String rowNo = Integer.toString(j + 1);
					String[] value = (String[]) list.get(j + 40 * i);
					setIntCell(sheet, rowNo, 7 + j, 1, false);// 序号
					setStringCell(sheet, value[1], 7 + j, 4, false);// 工序编号
					setStringCell(sheet, value[2], 7 + j, 14, false);// 工序中文名称
					setStringCell(sheet, value[3], 7 + j, 44, false);// 工序英文名称
					setStringCell(sheet, value[4], 7 + j, 71, false);// 版次
					setStringCell(sheet, value[5], 7 + j, 76, false);// 页数
					setStringCell(sheet, value[6], 7 + j, 81, false);// 备注
					if (value[5] != null && !value[5].isEmpty()) {
						totalpages = totalpages + Integer.parseInt(value[5]);
					}
				}
			}
			setIntCell(sheet, Integer.toString(i + 1), 50, 109, false);// 第几页码
			setIntCell(sheet, Integer.toString(mlpage), 50, 112, false);// 总页码
		}
		XSSFCellStyle style = null;
//		style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		style.setAlignment(XSSFCellStyle.ALIGN_LEFT);

		XSSFCellStyle style2 = null;
//		style2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		style2.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		style2.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		style2.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);

		XSSFCellStyle style3 = null;
//		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// 内容垂直居中
//		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);

		// 写板组清单数据
		if (bzlist == null) {
			bzlist = new ArrayList();
		}
		int bzpage = bzlist.size() / 40 + 1; // 板组数据sheet数
		for (int i = 0; i < bzpage; i++) {
			XSSFSheet sheet = book.getSheetAt(mlpage + i);
			sheet.setColumnBreak(115);
			if (i == bzpage - 1) {
				for (int j = 0; j + 40 * i < bzlist.size(); j++) {
					BoardInformation value = (BoardInformation) bzlist.get(j + 40 * i);
					// Object[] value = (Object[]) bzlist.get(j + 40 * i);
					String rowNum = value.getRowNum();
					String rowNo = Integer.toString(j + 1);
					String boardnumber = value.getBoardnumber();
					String partn = value.getPartn();
					String boardname = value.getBoardname();
					String partmaterial = value.getPartmaterial();
					String partthickness = value.getPartthickness();
					String sheetstrength = value.getSheetstrength();
					String gagi = value.getGagi();
					String maunit = value.getMaunit();
					String thunit = value.getThunit();

					if (rowNum != null) {
						setIntCellBydouble(sheet, rowNo, 7 + j, 1, style2);// 序号
					}
					if (partn != null) {
						setStringCellAndStyle(sheet, partn, 7 + j, 8, style);// 零件号
					}
					if (boardnumber != null) {
						setStringCellAndStyle(sheet, boardnumber, 7 + j, 25, style);// 板组编号
					}
					if (boardname != null) {
						setStringCellAndStyle(sheet, boardname, 7 + j, 38, style);// 板组名称
					}
					if (partmaterial != null) {
						setStringCellAndStyle(sheet, partmaterial, 7 + j, 68, style);// 材质
					}
					if (partthickness != null) {
						setDoubleCellAndStyle(sheet, partthickness, 7 + j, 81, style3);// 板厚
					}
					if (maunit != null) {
						setStringCellAndStyle(sheet, maunit, 7 + j, 87, style);// 板厚单位
					}
					if (sheetstrength != null) {
						setStringCellAndStyle(sheet, sheetstrength, 7 + j, 91, style3, 10);// 强度
					}
					if (thunit != null) {
						setStringCellAndStyle(sheet, thunit, 7 + j, 96, style);// 强度单位
					}
					if (gagi != null) {
						setStringCellAndStyle(sheet, gagi, 7 + j, 104, style);// GA/GI
					}
				}
			} else {
				for (int j = 0; j + 40 * i < 40 + 40 * i; j++) {
					BoardInformation value = (BoardInformation) bzlist.get(j + 40 * i);
					// Object[] value = (Object[]) bzlist.get(j + 40 * i);
					String rowNum = value.getRowNum();
					String rowNo = Integer.toString(j + 1);
					String boardnumber = value.getBoardnumber();
					String partn = value.getPartn();
					String boardname = value.getBoardname();
					String partmaterial = value.getPartmaterial();
					String partthickness = value.getPartthickness();
					String sheetstrength = value.getSheetstrength();
					String gagi = value.getGagi();
					String maunit = value.getMaunit();
					String thunit = value.getThunit();

					if (rowNum != null) {
						setIntCellBydouble(sheet, rowNo, 7 + j, 1, style2);// 序号
					}
					if (partn != null) {
						setStringCellAndStyle(sheet, partn, 7 + j, 8, style);// 零件号
					}
					if (boardnumber != null) {
						setStringCellAndStyle(sheet, boardnumber, 7 + j, 25, style);// 板组编号
					}
					if (boardname != null) {
						setStringCellAndStyle(sheet, boardname, 7 + j, 38, style);// 板组名称
					}
					if (partmaterial != null) {
						setStringCellAndStyle(sheet, partmaterial, 7 + j, 68, style);// 材质
					}
					if (partthickness != null) {
						setDoubleCellAndStyle(sheet, partthickness, 7 + j, 81, style3);// 板厚
					}
					if (maunit != null) {
						setStringCellAndStyle(sheet, maunit, 7 + j, 87, style);// 板厚单位
					}
					if (sheetstrength != null) {
						setStringCellAndStyle(sheet, sheetstrength, 7 + j, 91, style3, 10);// 强度
					}
					if (thunit != null) {
						setStringCellAndStyle(sheet, thunit, 7 + j, 96, style);// 强度单位
					}
					if (gagi != null) {
						setStringCellAndStyle(sheet, gagi, 7 + j, 104, style);// GA/GI
					}
				}
			}
			setIntCell(sheet, Integer.toString(i + 1), 50, 109, false);// 第几页码
			setIntCell(sheet, Integer.toString(bzpage), 50, 112, false);// 总页码
		}
	}

	// 输出文件
	public static void exportFile(XSSFWorkbook book, String reportname) {
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

			// 打开excel
			// openFile(fullFileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void openFile(String fullFileName) {
		// TODO Auto-generated method stub
		try {
			System.out.println("cmd /c call " + '"' + fullFileName + '"');
			Runtime.getRuntime().exec("cmd /c call " + '"' + fullFileName + '"');
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	// 对单元格赋值，强制使用文本格式
	public static void setCell(XSSFSheet sheet, String value, int rowIndex, int cellIndex, int celltype) {
		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		cell.setCellType(celltype);

		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue(value);
		// cell.setCellStyle(cellStyle);
	}

	// 对单元格赋值，强制使用文本格式
	public static void setCellwithCellStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex, int celltype,
			CellStyle style) {
		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		cell.setCellType(celltype);
		if (style != null) {
			cell.setCellStyle(style);
		}
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		if (celltype == Cell.CELL_TYPE_NUMERIC) {
			cell.setCellValue(Integer.parseInt(value));
		} else {
			cell.setCellValue(value);
		}
		// cell.setCellStyle(cellStyle);

	}

	// 对单元格赋值，强制使用文本格式
	public static void setStringCell(XSSFSheet sheet, String value, int rowIndex, int cellIndex, boolean flag) {

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
//		cell.setCellType(Cell.CELL_TYPE_STRING);
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		// 去掉删除线
		if (flag) {
			// cell.getCellStyle().getFont().getStrikeout();
			cell.getCellStyle().getFont().setStrikeout(false);
		}
		if (value == null) {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		} else {
			cell.setCellValue(value);
		}

		if (flag) {
			if (cellStyle != null) {
				cell.setCellStyle(cellStyle);
			}
		}
	}

	public static void setStringCellwithcellStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			CellStyle style) {

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		cell.setCellType(Cell.CELL_TYPE_STRING);

		cell.setCellValue(value);

		if (style != null) {
			cell.setCellStyle(style);
		}

	}

	// 对单元格赋值，强制使用整形
	public static void setIntCell(XSSFSheet sheet, String value, int rowIndex, int cellIndex, boolean flag) {

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		if (value != null && !value.isEmpty()) {
			cell.setCellValue(Integer.parseInt(value));
		} else {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		}
		if (flag) {
			if (cellStyle != null) {
				cell.setCellStyle(cellStyle);
			}
		}
	}

	// 对单元格赋值，强制使用整形，通过double类型强转
	public static void setIntCellBydouble(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style) {

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		if (value != null && !value.isEmpty()) {
			cell.setCellValue((int) (Double.parseDouble(value)));
		} else {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		}
		if (Style != null) {
			cell.setCellStyle(Style);
		}

	}

	// 对单元格赋值，强制使用double
	public static void setDoubleCell(XSSFSheet sheet, String value, int rowIndex, int cellIndex, boolean flag) {

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		cell.setCellType(Cell.CELL_TYPE_NUMERIC);
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		if (value != null) {
			cell.setCellValue(Double.parseDouble(value));
		} else {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		}
		if (flag) {
			cell.setCellStyle(cellStyle);
		}

	}

	// 对单元格赋值
	public static void setCell2(XSSFSheet sheet, String[] values, int num) {
		try {
			if (values[0] == null) {
				return;
			}
			int rownum = num;
			setIntCell(sheet, values[0], rownum, 1, true);
			setStringCell(sheet, values[1], rownum, 2, true);
			setIntCell(sheet, values[2], rownum, 3, true);
			setIntCell(sheet, values[3], rownum, 4, true);
			setIntCell(sheet, values[4], rownum, 5, true);
			setIntCell(sheet, values[5], rownum, 6, true);
			setIntCell(sheet, values[6], rownum, 7, true);
			setIntCell(sheet, values[7], rownum, 9, true);
			setIntCell(sheet, values[8], rownum, 10, true);
			setIntCell(sheet, values[9], rownum, 11, true);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 根据指定单元格的值，查找是否在excel中存在
	private static int findRow(XSSFSheet sheet, String cellContent) {
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
					if (cell.getRichStringCellValue().getString().trim().equals(cellContent)) {
						return row.getRowNum();
					}
				}
			}
		}
		return 0;
	}

	// Excel的sheet名称不能包含特殊字符，过滤特殊字符
	private static String FilterSpecialCharacters(String str) {
		String name = null;
		name = str.replaceAll("[*|':;',\\\\\\\\[\\\\\\\\].<>/?~！@￥%……&*（）——+|{}【】‘；：”“’。，、？]", "");
		return name;
		// return null;
	}

	/**
	 * 复制行
	 * 
	 * @param startRowIndex 起始行
	 * @param endRowIndex   结束行
	 * @param pPosition     目标起始行位置
	 */
	public static void copyRows(XSSFSheet currentSheet, int startRow, int endRow, int pPosition) {
		int pStartRow = startRow - 1;
		int pEndRow = endRow - 1;
		int targetRowFrom;
		int targetRowTo;
		int columnCount;
		CellRangeAddress region = null;
		int i;
		int j;
		if (pStartRow == -1 || pEndRow == -1) {
			return;
		}
		System.out.println(currentSheet.getNumMergedRegions());
		for (i = 0; i < currentSheet.getNumMergedRegions(); i++) {
			region = currentSheet.getMergedRegion(i);
			System.out.println("FirstRow=" + region.getFirstRow());
			System.out.println("LastRow=" + region.getLastRow());
			if ((region.getFirstRow() >= pStartRow) && (region.getLastRow() <= pEndRow)) {
				targetRowFrom = region.getFirstRow() - pStartRow + pPosition;
				targetRowTo = region.getLastRow() - pStartRow + pPosition;
				CellRangeAddress newRegion = region.copy();
				newRegion.setFirstRow(targetRowFrom);
				newRegion.setFirstColumn(region.getFirstColumn());
				newRegion.setLastRow(targetRowTo);
				newRegion.setLastColumn(region.getLastColumn());
				currentSheet.addMergedRegion(newRegion);
			}
		}
		for (i = pStartRow; i <= pEndRow; i++) {
			XSSFRow sourceRow = currentSheet.getRow(i);
			columnCount = sourceRow.getLastCellNum();
			if (sourceRow != null) {
				XSSFRow newRow = currentSheet.createRow(pPosition - pStartRow + i);
				newRow.setHeight(sourceRow.getHeight());
				for (j = 0; j < columnCount; j++) {
					XSSFCell templateCell = sourceRow.getCell(j);
					if (templateCell != null) {
						XSSFCell newCell = newRow.createCell(j);
						copyCell(templateCell, newCell);
					}
				}
			}
		}
	}

	private static void copyCell(XSSFCell srcCell, XSSFCell distCell) {
		distCell.setCellStyle(srcCell.getCellStyle());
		if (srcCell.getCellComment() != null) {
			distCell.setCellComment(srcCell.getCellComment());
		}

//		if (!getCellValue(srcCell).isEmpty()) {
//			System.out.println("单元格值：" + getCellValue(srcCell));
//			String Formula = srcCell.getCellFormula();
//			System.out.println("单元格公式：" + Formula);
//		}
		int srcCellType = srcCell.getCellType();
		distCell.setCellType(srcCellType);
		if (srcCellType == XSSFCell.CELL_TYPE_NUMERIC) {
			if (DateUtil.isCellDateFormatted(srcCell)) {
				distCell.setCellValue(srcCell.getDateCellValue());
			} else {
				distCell.setCellValue(srcCell.getNumericCellValue());
			}
		} else if (srcCellType == XSSFCell.CELL_TYPE_STRING) {
			distCell.setCellValue(srcCell.getRichStringCellValue());
		} else if (srcCellType == XSSFCell.CELL_TYPE_BLANK) {
			// nothing21
		} else if (srcCellType == XSSFCell.CELL_TYPE_BOOLEAN) {
			distCell.setCellValue(srcCell.getBooleanCellValue());
		} else if (srcCellType == XSSFCell.CELL_TYPE_ERROR) {
			distCell.setCellErrorValue(srcCell.getErrorCellValue());
		} else if (srcCellType == XSSFCell.CELL_TYPE_FORMULA) {
			distCell.setCellFormula(srcCell.getCellFormula());
		} else { // nothing29

		}
	}

	private static void copyCell(XSSFWorkbook book, XSSFCell srcCell, XSSFCell distCell) {
		distCell.setCellStyle(srcCell.getCellStyle());
		if (srcCell.getCellComment() != null) {
			distCell.setCellComment(srcCell.getCellComment());
		}
		CellStyle newStyle = book.createCellStyle();
		CellStyle srcStyle = srcCell.getCellStyle();
		newStyle.cloneStyleFrom(srcStyle);
		newStyle.setFont(book.getFontAt(srcStyle.getFontIndex()));

		XSSFSheet sheet = book.getSheetAt(0);
		int colWidth = sheet.getColumnWidth(srcCell.getColumnIndex());
		sheet.setColumnWidth(distCell.getColumnIndex(), colWidth);
		// 样式
		distCell.setCellStyle(newStyle);

		int srcCellType = srcCell.getCellType();
		distCell.setCellType(srcCellType);
		if (srcCellType == XSSFCell.CELL_TYPE_NUMERIC) {
			if (DateUtil.isCellDateFormatted(srcCell)) {
				distCell.setCellValue(srcCell.getDateCellValue());
			} else {
				distCell.setCellValue(srcCell.getNumericCellValue());
			}
		} else if (srcCellType == XSSFCell.CELL_TYPE_STRING) {
			distCell.setCellValue(srcCell.getRichStringCellValue());
		} else if (srcCellType == XSSFCell.CELL_TYPE_BLANK) {
			// nothing21
		} else if (srcCellType == XSSFCell.CELL_TYPE_BOOLEAN) {
			distCell.setCellValue(srcCell.getBooleanCellValue());
		} else if (srcCellType == XSSFCell.CELL_TYPE_ERROR) {
			distCell.setCellErrorValue(srcCell.getErrorCellValue());
		} else if (srcCellType == XSSFCell.CELL_TYPE_FORMULA) {
			distCell.setCellFormula(srcCell.getCellFormula());
		} else { // nothing29

		}
	}

	// 设置单元格公式
	public static void setCellFormula(XSSFSheet Sheet, String formula, int rowIndex, int cellIndex) {
		try {
			XSSFRow row = Sheet.getRow(rowIndex);
			Cell cell = row.getCell(cellIndex);
			cell.setCellFormula(formula);
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	// 最后删除模板行并设置汇总行的公式
	public static void dealTotalRowFormula(XSSFWorkbook book, ReportViwePanel viewPanel) throws IOException {
		XSSFSheet sheet = book.getSheetAt(0);

		// 设置焊装产线统计行的单元格公式
		int total_rownum = sheet.getPhysicalNumberOfRows();
		int rowIndex = total_rownum - 9;
		// 开始和结束行统计
		int start = 11;
		int end = total_rownum - 8;
		String[] strcell = { "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
				"V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM",
				"AN", "AO", "AP", "AQ" };
		for (int j = 0; j < strcell.length; j++) {
			viewPanel.addInfomation("", 70, 100);
			// 小计列
			if (strcell[j] == "AN" || strcell[j] == "AO") {
				setCellFormula(sheet, "SUM(" + strcell[j] + (rowIndex + 2) + ":" + strcell[j] + (rowIndex + 4) + ")",
						rowIndex + 4, 3 + j);
				System.out.println(
						"小计列:" + "SUM(" + strcell[j] + (rowIndex + 1) + ":" + strcell[j] + (rowIndex + 3) + ")");
			}
			// 合计列
			if (strcell[j] != "AI" && strcell[j] != "AK" && strcell[j] != "AM" && strcell[j] != "AP") {
				String formula = "SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")/2";
				setCellFormula(sheet, formula, rowIndex + 5, 3 + j);
				System.out.println("合计列:" + formula);
			}
			// 有些列需要保留两位小数
//			if (strcell[j] == "I" || strcell[j] == "M" || strcell[j] == "X" || strcell[j] == "AG" || strcell[j] == "AH"
//					|| strcell[j] == "AJ" || strcell[j] == "AL" || strcell[j] == "AN" || strcell[j] == "AP") {
//
//				String formula = "ROUND(SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")/2,2)";
//				setCellFormula(sheet, formula, rowIndex + 5, 3 + j);
//				System.out.println("合计列:" + formula);
//			}

			// 倒数第三行
			if (strcell[j] == "AH") {
				setCellFormula(sheet, "SUM(" + strcell[j] + (rowIndex + 6) + "+" + strcell[j] + (rowIndex + 4) + ")/60",
						rowIndex + 6, 3 + j);
				System.out.println(
						"倒数第三行:" + "(" + strcell[j] + (rowIndex + 3) + ":" + strcell[j] + (rowIndex + 5) + ")/60");
			}
			// 倒数第二行
			if (strcell[j] == "AH" || strcell[j] == "AJ" || strcell[j] == "AL" || strcell[j] == "AN"
					|| strcell[j] == "AO" || strcell[j] == "AQ") {
				setCellFormula(sheet, "SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")/2", rowIndex + 7,
						3 + j);
				System.out.println("倒数第二行:" + "SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")/2");
			}
			// 最后一行
			if (strcell[j] == "AH" || strcell[j] == "AJ" || strcell[j] == "AL") {
				setCellFormula(sheet, strcell[j] + (rowIndex + 8) + "/60", rowIndex + 8, 3 + j);
				System.out.println("最后一行1:" + strcell[j] + (rowIndex + 8) + "/60");
			}
			// 最后一行
			if (strcell[j] == "AO") {
				setCellFormula(sheet, strcell[j - 1] + (rowIndex + 8) + "/" + strcell[j] + (rowIndex + 8), rowIndex + 8,
						3 + j);
				System.out.println("小最后一行2:" + strcell[j - 1] + (rowIndex + 8) + "/" + strcell[j] + (rowIndex + 8));
			}
			// 最后一行
			if (strcell[j] == "AQ") {
				setCellFormula(sheet, strcell[j] + (rowIndex + 8) + "/" + strcell[j - 9] + (rowIndex + 9), rowIndex + 8,
						3 + j);
				System.out.println("最后一行3:" + strcell[j] + (rowIndex + 8) + "/" + strcell[j - 9] + (rowIndex + 9));
			}

			// excel设置公式自动计算
			sheet.setForceFormulaRecalculation(true);
		}
	}

	// 获取单元格内容
	public static String getCellValue(XSSFCell cell) {
		String cellValue = "";
		DecimalFormat df = new DecimalFormat("#.##");
		if (cell == null) {
			return cellValue;
		}
		switch (cell.getCellType()) {
		case XSSFCell.CELL_TYPE_STRING:
			cellValue = cell.getRichStringCellValue().getString().trim();
			break;
		case XSSFCell.CELL_TYPE_NUMERIC:
			cellValue = df.format(cell.getNumericCellValue()).toString();
			break;
		case XSSFCell.CELL_TYPE_BOOLEAN:
			cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
			break;
		case XSSFCell.CELL_TYPE_FORMULA:
			cellValue = cell.getCellFormula();
			break;
		default:
			cellValue = "";
		}
		return cellValue;
	}

	/**
	 * Remove a row by its index
	 * 
	 * @param sheet    a Excel sheet
	 * @param rowIndex a 0 based index of removing row
	 */
	public static void removeRow(XSSFSheet sheet, int rowIndex) {
		int lastRowNum = sheet.getLastRowNum();
		if (rowIndex >= 0 && rowIndex < lastRowNum)
			sheet.shiftRows(rowIndex + 1, lastRowNum, -1);// 将行号为rowIndex+1一直到行号为lastRowNum的单元格全部上移一行，以便删除rowIndex行
		if (rowIndex == lastRowNum) {
			XSSFRow removingRow = sheet.getRow(rowIndex);
			if (removingRow != null)
				sheet.removeRow(removingRow);
		}
	}

	// 画单元格斜线
	public static void CellSlash(XSSFSheet sheet, int x1, int x2, int y1, int y2) {

		return;
	}

	// 根据阶段种类，初始化表头
	public static HashMap InitializeHeader(XSSFWorkbook book, ArrayList<String> stage) {
		// TODO Auto-generated method stub
		HashMap map = new HashMap();
		if (stage.size() > 0) {

			XSSFSheet sheet = book.getSheetAt(0);
			XSSFRow row = sheet.getRow(15);
			XSSFRow row1 = sheet.getRow(16);
			for (int i = 0; i < stage.size(); i++) {
				// 第一钟阶段，不需要复制，模板已经存在列了，直接赋值阶段名称就行
				XSSFCell cell = row.getCell(17); // 対応
				XSSFCell cell1 = row1.getCell(8); // 判定
				XSSFCell cell2 = row1.getCell(9); // 内容
				XSSFCell cell4 = row1.getCell(18); // 部署
				XSSFCell cell5 = row1.getCell(19); // 日程
				if (i == 0) {
					setStringCell(sheet, stage.get(i), 14, 8, false);
					map.put(stage.get(i), 8);
				} else {
					setStringCell(sheet, stage.get(i), 14, 8 + 13 * i, false);

					XSSFCell newcell = row.createCell(17 + 13 * i); // 対応
					copyCell(book, cell, newcell);

					CellRangeAddress region1 = new CellRangeAddress(15, 15, (short) (17 + 13 * i),
							(short) (17 + 13 * i + 2)); // 参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
					sheet.addMergedRegion(region1);

					XSSFCell newcell1 = row1.createCell(8 + 13 * i); // 判定
					copyCell(book, cell1, newcell1);
					for (int j = 0; j < 9; j++) {
						XSSFCell newcell2 = row1.createCell(9 + j + 13 * i); // 内容
						copyCell(book, cell2, newcell2);
					}
					CellRangeAddress region2 = new CellRangeAddress(16, 16, (short) (9 + 13 * i),
							(short) (9 + 13 * i + 8)); // 参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
					sheet.addMergedRegion(region2);

					XSSFCell newcell4 = row1.createCell(18 + 13 * i); // 部署
					copyCell(book, cell4, newcell4);
					XSSFCell newcell5 = row1.createCell(19 + 13 * i); // 日程
					copyCell(book, cell5, newcell5);

					map.put(stage.get(i), 8 + 13 * i);
				}
			}
		}
		return map;
	}

	// 写返修要件检查表数据，支持多个阶段输出
	public static void writeRequirementsDataToSheet(XSSFWorkbook book, ArrayList datalist, HashMap map) {
		// TODO Auto-generated method stub
		XSSFSheet sheet = book.getSheetAt(0);

		// 对于内容列，文本要左上角显示
		CellStyle style = book.createCellStyle();

		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);

		for (int i = 0; i < datalist.size(); i++) {
			Object[] obj = (Object[]) datalist.get(i);

			// 根据阶段获取起始列
			int startColum = (int) map.get(obj[7].toString());

			setStringCell(sheet, obj[7].toString(), 14, startColum, false); // 阶段

			for (int k = 0; k < 9; k++) {
				setStringCell(sheet, obj[0].toString(), 17 + 9 * i + k, 0, true); // 番号
				setStringCell(sheet, obj[1].toString(), 17 + 9 * i + k, 1, true); // 項目
				setStringCell(sheet, obj[2].toString(), 17 + 9 * i + k, 2, true); // 場所
				setStringCell(sheet, obj[3].toString(), 17 + 9 * i + k, 4, true); // 過去問題点
				setStringCell(sheet, obj[4].toString(), 17 + 9 * i + k, 6, true); // 補足

				setStringCell(sheet, "", 17 + 9 * i + k, 3, true); // 過去問題点
				setStringCell(sheet, "", 17 + 9 * i + k, 5, true); // 補足

				Iterator iter = map.entrySet().iterator();
				while (iter.hasNext()) {
					Map.Entry entry = (Entry) iter.next();
					int val = (int) entry.getValue();
					if (val == startColum) {
						setStringCell(sheet, obj[8].toString(), 17 + 9 * i + k, val, true); // 判定

						for (int n = 0; n < 9; n++) {
							setStringCellwithcellStyle(sheet, obj[9].toString(), 17 + 9 * i + k, val + n + 1, style); // 内容
						}

						setStringCell(sheet, obj[11].toString(), 17 + 9 * i + k, val + 10, true); // 部署
						setStringCell(sheet, obj[12].toString(), 17 + 9 * i + k, val + 11, true); // 日程
					} else {
						setStringCell(sheet, "", 17 + 9 * i + k, val, true); // 判定

						for (int n = 0; n < 9; n++) {
							setStringCellwithcellStyle(sheet, "", 17 + 9 * i + k, val + n + 1, style); // 内容
						}

						setStringCell(sheet, "", 17 + 9 * i + k, val + 10, true); // 部署
						setStringCell(sheet, "", 17 + 9 * i + k, val + 11, true); // 日程
					}
				}

			}
			if (((ArrayList) obj[5]).size() > 0) {
				writepicturetosheet(book, sheet, (ArrayList) obj[5], 17 + 9 * i, 3, false); // 点
			}
			if (((ArrayList) obj[6]).size() > 0) {
				writepicturetosheet(book, sheet, (ArrayList) obj[6], 17 + 9 * i, 5, false); // 图纸表記方法
			}
			if (((ArrayList) obj[10]).size() > 0) {
				writepicturetosheets(book, sheet, (ArrayList) obj[10], 18 + 9 * i, startColum + 1);
			}

			// 合并单元格
			CellRangeAddress region1;
			int[] colum = { 0, 1, 2, 3, 4, 5, 6 };
			for (int j = 0; j < colum.length; j++) {

				// 参数1：起始行;参数2：终止行;参数3：起始列; 参数4：终止列
				region1 = new CellRangeAddress(17 + 9 * i, 25 + 9 * i, (short) colum[j], (short) colum[j]);

				sheet.addMergedRegion(region1);
			}
			Iterator iter = map.entrySet().iterator();
			while (iter.hasNext()) {
				Map.Entry entry = (Entry) iter.next();
				int val = (int) entry.getValue();
				int[] colum1 = { val, val + 1, val + 10, val + 11 };
				for (int j = 0; j < colum1.length; j++) {
					if (colum1[j] == val + 1) {
						// 参数1：起始行;参数2：终止行;参数3：起始列; 参数4：终止列
						region1 = new CellRangeAddress(17 + 9 * i, 25 + 9 * i, (short) (val + 1), (short) (val + 9));
					} else {
						// 参数1：起始行;参数2：终止行;参数3：起始列; 参数4：终止列
						region1 = new CellRangeAddress(17 + 9 * i, 25 + 9 * i, (short) colum1[j], (short) colum1[j]);
					}
					sheet.addMergedRegion(region1);
				}
			}
		}
	}

	// 写返修要件检查表数据，只输出当前阶段
	public static void writeRequirementsDataToSheet(XSSFWorkbook book, ArrayList datalist) {
		// TODO Auto-generated method stub
		XSSFSheet sheet = book.getSheetAt(0);

		// 对于内容列，文本要左上角显示
		CellStyle style = book.createCellStyle();

		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // 下边框
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// 左边框
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// 上边框
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// 右边框
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
		style.setWrapText(true);

		for (int i = 0; i < datalist.size(); i++) {
			Object[] obj = (Object[]) datalist.get(i);

			setStringCell(sheet, obj[7].toString(), 14, 8, false); // 阶段

			for (int k = 0; k < 9; k++) {
				setStringCell(sheet, obj[0].toString(), 17 + 9 * i + k, 0, true); // 番号
				setStringCell(sheet, obj[1].toString(), 17 + 9 * i + k, 1, true); // 項目
				setStringCell(sheet, obj[2].toString(), 17 + 9 * i + k, 2, true); // 場所
				setStringCell(sheet, obj[3].toString(), 17 + 9 * i + k, 4, true); // 過去問題点
				setStringCell(sheet, obj[4].toString(), 17 + 9 * i + k, 6, true); // 補足

//				setStringCell(sheet, "", 17 + 9 * i + k, 3, true); // 過去問題点
//				setStringCell(sheet, "", 17 + 9 * i + k, 5, true); // 補足
				setStringCellwithcellStyle(sheet, obj[13].toString(), 17 + 9 * i + k, 3, style); // 点
				setStringCellwithcellStyle(sheet, obj[14].toString(), 17 + 9 * i + k, 5, style); // 图纸表記方法

				setStringCell(sheet, obj[8].toString(), 17 + 9 * i + k, 8, true); // 判定
				for (int m = 0; m < 9; m++) {
					setStringCellwithcellStyle(sheet, obj[9].toString(), 17 + 9 * i + k, 9 + m, style); // 内容
				}

				setStringCell(sheet, obj[11].toString(), 17 + 9 * i + k, 18, true); // 部署
				setStringCell(sheet, obj[12].toString(), 17 + 9 * i + k, 19, true); // 日程
			}
			if (((ArrayList) obj[5]).size() > 0) {
				if (obj[13].toString().isEmpty()) {
					writepicturetosheet(book, sheet, (ArrayList) obj[5], 17 + 9 * i, 3, false); // 点
				} else {
					writepicturetosheet(book, sheet, (ArrayList) obj[5], 18 + 9 * i, 3, true); // 点
				}
			}
			if (((ArrayList) obj[6]).size() > 0) {
				if (obj[14].toString().isEmpty()) {
					writepicturetosheet(book, sheet, (ArrayList) obj[6], 17 + 9 * i, 5, false); // 图纸表記方法
				} else {
					writepicturetosheet(book, sheet, (ArrayList) obj[6], 18 + 9 * i, 5, true); // 图纸表記方法
				}

			}
			if (((ArrayList) obj[10]).size() > 0) {
				writepicturetosheets(book, sheet, (ArrayList) obj[10], 18 + 9 * i, 9);
			}

			// 合并单元格
			CellRangeAddress region1;
			int[] colum = { 0, 1, 2, 3, 4, 5, 6, 8, 9, 18, 19 };
			for (int j = 0; j < colum.length; j++) {
				if (colum[j] == 9) {
					region1 = new CellRangeAddress(17 + 9 * i, 25 + 9 * i, (short) 9, (short) 17); // 参数1：起始行 参数2：终止行
																									// 参数3：起始列 参数4：终止列
				} else {
					region1 = new CellRangeAddress(17 + 9 * i, 25 + 9 * i, (short) colum[j], (short) colum[j]); // 参数1：起始行
																												// 参数2：终止行
																												// 参数3：起始列
																												// 参数4：终止列
				}
				sheet.addMergedRegion(region1);
			}
		}
	}

	// 根据单个文件写图片到excel
	private static void writepicturetosheet(XSSFWorkbook book, XSSFSheet sheet, ArrayList obj, int rowindex,
			int colindex, boolean flag) {
		// 先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		BufferedImage bufferImg;
		int rowNum = 9;
		if (flag) {
			rowNum = 8;
		}
		try {
			File file = (File) obj.get(0);
			bufferImg = ImageIO.read(file);
			ImageIO.write(bufferImg, "png", byteArrayOut);
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) colindex, rowindex,
					(short) (colindex + 1), rowindex + rowNum);
			anchor.setAnchorType(2);
			// 插入图片
			patriarch.createPicture(anchor,
					book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	// 根据多个文件写图片到excel
	private static void writepicturetosheets(XSSFWorkbook book, XSSFSheet sheet, ArrayList list, int rowindex,
			int colindex) {
		// 先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray

		int num = list.size(); // 图片数量
		int x1 = 0;
		int x2 = 0;
		int y1 = 0;
		int y2 = 0;
		// 记录需要显示多少行
		int rows = (num + 2) / 3;
		// 每行的高度
		int hight = 8 / rows;
		// 每列的宽度
		int width = 0;

		if (num < 4) {
			width = 9 / num;
		} else {
			width = 9 / 3;
		}

		for (int i = 0; i < num; i++) {
			File file = (File) list.get(i);

			ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
			BufferedImage bufferImg;

			int number = (i + 3) / 3;// 处于第几行
			int number2 = (i + 1) % 3;// 处于第几列
			if (number2 == 0) {
				number2 = 3;
			}
			x1 = rowindex + hight * (number - 1);
			y1 = colindex + width * (number2 - 1);
			x2 = rowindex + hight * number;
			y2 = colindex + width * number2;
			System.out.println("y1:" + y1 + "x1:" + x1 + "y2" + y2 + "x2" + x2);
			try {
				bufferImg = ImageIO.read(file);
				ImageIO.write(bufferImg, "png", byteArrayOut);
				XSSFDrawing patriarch = sheet.createDrawingPatriarch();
				XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) y1, x1, (short) y2,
						x2);
				anchor.setAnchorType(2);
				// 插入图片
				patriarch.createPicture(anchor,
						book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}

	// 对单元格赋值，强制使用文本格式
	public static void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style) {

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
//		cell.setCellType(Cell.CELL_TYPE_STRING);
		if (value == null) {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		} else {
			cell.setCellValue(value);
		}
		if (Style != null) {
			cell.setCellStyle(Style);
		}

	}

	// 对单元格赋值，强制使用文本格式
	public static void setDoubleCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style) {

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
//			cell.setCellType(Cell.CELL_TYPE_STRING);
		if (value == null) {
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		} else {
			if(Util.isNumber(value))
			{
				//cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(BigDecimal.valueOf(Double.parseDouble(value)).doubleValue());
			}
			else
			{
				cell.setCellValue(value);
			}
			
		}
		if (Style != null) {
			cell.setCellStyle(Style);
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
				double math = Double.parseDouble(value);
				cell.setCellValue((int) math);
			} else if (celltype == 11) {
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(Double.parseDouble(value));
			} else {
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(value);
			}
		}
		if (Style != null) {
			cell.setCellStyle(Style);
		}

	}

	/**
	 * 设置单元格字体大小
	 */
	public static void setFontSize(XSSFWorkbook book, Cell cell, short num) {
		Font font = book.createFont();
		font.setFontName("宋体");
		// font.setStrikeout(true);//删除线
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBoldweight(Font.BOLDWEIGHT_BOLD); // 字体加粗
		for (short k = num; k >= 9; k--) {
			font.setFontHeightInPoints(k);
			if (checkCellReasonable(cell, k)) {
				break;
			}
		}
		// 解决单元格样式覆盖的问题
		CellStyle cStyle = book.createCellStyle();
		cStyle.cloneStyleFrom(cell.getCellStyle());
		// cStyle.setWrapText(true);
		cStyle.setFont(font);
		if (cStyle != null) {
			cell.setCellStyle(cStyle);
		}

	}

	/**
	 * 校验单元格中的字体大小是否合理
	 */
	public static boolean checkCellReasonable(Cell cell, short fontSize) {
		int sum = cell.getStringCellValue().length();
		double cellWidth = getTotalWidth(cell);
		double fontWidth = (double) fontSize / 72 * 96 * 2;
		double cellHeight = cell.getRow().getHeightInPoints();
		return fontSize + sum < cellWidth;
//	    double rows1 = fontWidth * sum / cellWidth;
//	    double rows2 = cellHeight / fontSize;
//	    return rows2 >= rows1;
	}

	/**
	 * 获取单元格的总宽度（单位：像素）
	 */
	public static double getTotalWidth(Cell cell) {
		int x = getColNum(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex());
		double totalWidthInPixels = 0;
		for (int i = 0; i < x; i++) {
			double d = cell.getSheet().getColumnWidth(i + cell.getColumnIndex()) / 256.0;
			totalWidthInPixels += cell.getSheet().getColumnWidth(i + cell.getColumnIndex()) / 256.0;
		}
		return totalWidthInPixels;
	}

	/**
	 * 获取单元格的列数，如果是合并单元格，就获取总的列数
	 */
	public static int getColNum(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		// 判断该单元格是否是合并区域的内容
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn) {
				return lastColumn - firstColumn + 1;
			}
		}
		return 1;
	}
}
