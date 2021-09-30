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

	// ����ģ�崴��Excel��ģ��
	public static XSSFWorkbook creatXSSFWorkbook(InputStream input, ArrayList list, ArrayList errornamelist) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);

			XSSFSheet sheet1 = book.getSheetAt(0);
			System.out.println("sheet���ƣ�" + sheet1.getSheetName());
			int num = book.getNumberOfSheets();
			int deletenum = num - 1;
			if (deletenum > 0) {
				for (int i = 0; i < deletenum; i++) {
					XSSFSheet sheet2 = book.getSheetAt(1);
					// TC��ӵ�Excelģ�����ݼ�����������ҳǩ����Ҫ������������ɾ��
					if (sheet2 != null) {
						book.removeSheetAt(1);
					}
				}
			}
			System.out.println("ɾ��sheet��" + book.getNumberOfSheets());
			//////////// ���÷�����ʾ�Ϸ�/�·�
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);

			cellStyle = book.createCellStyle();
			// �߿�
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

			// �㺸�����������ݼ�
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
			// ����ͬ��λ��ͬ�㺸�������ư��ղ��ұ������
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
			System.out.println("���ظ����Ƶģ�" + renameList);
			System.out.println("sheet����map��" + sortmap);

			for (int i = 0; i < list.size(); i++) {
				String[] str = (String[]) list.get(i);
				System.out.println("�㺸����" + i + " " + str[10].toString());
				String[] statename = str[0].toString().split("\\\\");
				String prename = statename[0].toString();
				// String compareshname = str[8].toString() + "_" + str[0].toString();
				String shname = str[8].toString() + "_" + prename;
				if (renameList.contains(shname)) {
					String sortshname = str[8].toString() + "_" + prename + "_" + str[10].toString();
					shname = str[8].toString() + "_" + prename + "_" + sortmap.get(sortshname);
				}
				System.out.println("sheet���ƣ�" + shname);
				String name = FilterSpecialCharacters(shname);
				System.out.println("ȥ�������ַ����sheet���ƣ�" + name);

				// ���sheet���Ƴ���31���ַ����ȣ�������ʾ
				if (name.length() > 31) {
					errornamelist.add(shname);
				}
			}
			if (errornamelist.size() > 0) {
				return null;
			}
			// ѭ�����ݼ������sheet
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
					// ��ģ���еĵ�һ��sheet�޸�һ��
					book.setSheetName(0, name);
				} else {
					// ���ݼ����ж����������sheet
					XSSFSheet sheet = book.cloneSheet(0);
					book.setSheetName(i, name);
					//////////// ���÷�����ʾ�Ϸ�/�·�
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

	// ����������Ӷ��sheetҳ
	public static void creatXSSFWorkbookByData(XSSFWorkbook book, int sheetnum) {

		try {
			XSSFSheet sheet1 = book.getSheetAt(0);
			System.out.println("sheet���ƣ�" + sheet1.getSheetName());
			int num = book.getNumberOfSheets();
			int deletenum = num - 2;
			if (deletenum > 0) {
				for (int i = 0; i < deletenum; i++) {
					XSSFSheet sheet2 = book.getSheetAt(2);
					// ���ģ�����2��sheetҳ����Ҫɾ��
					if (sheet2 != null) {
						book.removeSheetAt(2);
					}
				}
			}
			System.out.println("ɾ��sheet��" + book.getNumberOfSheets());
			//////////// ���÷�����ʾ�Ϸ�/�·�
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);

			cellStyle = book.createCellStyle();
			// �߿�
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

			// ѭ�����ݼ������sheet
			for (int i = 0; i < sheetnum; i++) {
				// ��һ��sheet��������
				if (i != 0) {
					// ���ݼ����ж����������sheet
					XSSFSheet sheet = book.cloneSheet(1);
					String sheetname = "ֱ���嵥" + Integer.toString(i + 1);
					int sheetIndex = i + 1;
					book.setSheetName(sheetIndex, sheetname);
					//////////// ���÷�����ʾ�Ϸ�/�·�
					sheet.setRowSumsBelow(false);
					sheet.setRowSumsRight(false);
					sheet.setRowSumsBelow(false);
					sheet.setRowSumsRight(false);
				}
			}
			System.out.println("sheet������" + book.getNumberOfSheets());
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return;
	}

	// ����ģ�崴��Excel��ģ��
	public static XSSFWorkbook creatXSSFWorkbook(InputStream input) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);
			XSSFSheet sheet1 = book.getSheetAt(0);
			//////////// ���÷�����ʾ�Ϸ�/�·�
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);

			cellStyle = book.createCellStyle();
			// �߿�
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
			//////////// ���÷�����ʾ�Ϸ�/�·�
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);
			sheet1.setRowSumsBelow(false);
			sheet1.setRowSumsRight(false);

			cellStyle = book.createCellStyle();
			// �߿�
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle.setWrapText(true);
			cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
			cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// �����
			cellStyle.setWrapText(true);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	// ����ģ�����Excel����
	public static XSSFWorkbook updateXSSFWorkbook(InputStream input, ArrayList list, ArrayList errorname) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);

			cellStyle = book.createCellStyle();
			// �߿�
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

			// ѭ������sheetҳ����������ȥ�ң����û���ҵ��Ƴ�sheet
			System.out.println("�Ƴ�ǰ������" + book.getNumberOfSheets());

			// �㺸�����������ݼ�
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
			// ����ͬ��λ��ͬ�㺸�������ư��ղ��ұ������
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
			System.out.println("���ظ����Ƶģ�" + renameList);
			System.out.println("sheet����map��" + sortmap);
			for (int i = 0; i < list.size(); i++) {
				String[] str = (String[]) list.get(i);
				System.out.println("�㺸����" + i + " " + str[10].toString());
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
			// ����sheet������ͻ������������һ��
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
				System.out.println("��ȡ�ĵ㺸����ID��" + dreaid);
				for (int k = 0; k < list.size(); k++) {
					String[] str = (String[]) list.get(k);
					String factdreaid = str[10];
					// ��ͨ���㺸����IDȡƥ�䣬���δƥ�䵽 �ٸ��ݵ㺸��������ƥ�䣨Ϊ�˽��Ŀǰϵͳ�Ѳ������Ե㺸�����������������ݣ�
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
			System.out.println("�Ƴ���������" + st);

			// ѭ�����ݼ������sheet
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
				// ����������򴴽�
				if (sheet == null) {
					sheet = book.cloneSheet(0);
					if (renameList.contains(shname1)) {
						book.setSheetName(st, FilterSpecialCharacters(shname2));
					} else {
						book.setSheetName(st, FilterSpecialCharacters(shname1));
					}
					//////////// ���÷�����ʾ�Ϸ�/�·�
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

	// дsheet���ݣ��㺸��������д��
	public static void writeDiscreteDataToSheet(XSSFWorkbook book, ArrayList list, ArrayList weldlist,
			boolean setGroup) {
		// TODO Auto-generated method stub
		// int row = 3;

		int len = list.size();
		// �㺸�����������ݼ�
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
		// ����ͬ��λ��ͬ�㺸�������ư��ղ��ұ������
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
		System.out.println("���ظ����Ƶģ�" + renameList);
		System.out.println("sheet����map��" + sortmap);

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

			// ��������Ϣ��Ϊ����ͷ���У��㺸��Ϣ���� {�㺸�������ƣ����ڣ������˱�ţ��������ͺţ���ǹǹ�ţ���װ��λ������}
			setCell(sheet, values[1], 1, 11, Cell.CELL_TYPE_STRING);// ����
			setCell(sheet, values[2], 0, 6, Cell.CELL_TYPE_STRING);// �����˱��
			setCell(sheet, values[3], 0, 13, Cell.CELL_TYPE_STRING);// �������ͺ�
			setCell(sheet, values[4], 1, 6, Cell.CELL_TYPE_STRING);// ��ǹǹ��
			setCell(sheet, values[5], 1, 2, Cell.CELL_TYPE_STRING);// ��װ��λ
			setCell(sheet, values[6], 0, 2, Cell.CELL_TYPE_STRING);// ����
			setCell(sheet, i + 1 + "/" + len, 1, 14, Cell.CELL_TYPE_STRING);// ҳ��

			// �㺸����ID
			setCell(sheet, values[10], 0, 26, Cell.CELL_TYPE_STRING);// ����
			// row++;
			int witd = 17;
			if (weldlist.size() > 15) {
				witd = witd + weldlist.size() - 15;
			}
			// ���ô�ӡ����
			book.setPrintArea(i, 0, 14, 0, witd);
			PrintSetup printSetup = sheet.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 94);// �Զ������ţ��˴�100Ϊ������
			printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)
		}

	}

	// дsheet���ݣ��㺸��������д��
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
		// ����ͬ��λ��ͬ�㺸�������ư��ղ��ұ������
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
		System.out.println("���ظ����Ƶģ�" + renameList);
		System.out.println("sheet����map��" + sortmap);

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
			if (sheet == null) // ����֮ǰ����û��д�㺸����ID������
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

			// ��������Ϣ��Ϊ����ͷ���У��㺸��Ϣ���� {�㺸�������ƣ����ڣ������˱�ţ��������ͺţ���ǹǹ�ţ���װ��λ������}
			setCell(sheet, values[1], 1, 11, Cell.CELL_TYPE_STRING);// ����
			setCell(sheet, values[2], 0, 6, Cell.CELL_TYPE_STRING);// �����˱��
			setCell(sheet, values[3], 0, 13, Cell.CELL_TYPE_STRING);// �������ͺ�
			setCell(sheet, values[4], 1, 6, Cell.CELL_TYPE_STRING);// ��ǹǹ��
			setCell(sheet, values[5], 1, 2, Cell.CELL_TYPE_STRING);// ��װ��λ
			setCell(sheet, values[6], 0, 2, Cell.CELL_TYPE_STRING);// ����
			setCell(sheet, i + 1 + "/" + len, 1, 14, Cell.CELL_TYPE_STRING);// ҳ��
			// �㺸����ID
			setCell(sheet, values[10], 0, 26, Cell.CELL_TYPE_STRING);// ����
			// row++;

			int witd = 17;
			if (weldlist.size() > 15) {
				witd = witd + weldlist.size() - 15;
			}
			// ���ô�ӡ����
			book.setPrintArea(i, 0, 14, 0, witd);
			PrintSetup printSetup = sheet.getPrintSetup();
			printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
			printSetup.setScale((short) 94);// �Զ������ţ��˴�100Ϊ������
			printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)
		}

	}

	// дsheet���ݣ���������д��
	public static void writeWeldDataToSheet(XSSFWorkbook book, ArrayList list, ArrayList partlist, boolean setGroup) {
		// TODO Auto-generated method stub
		// int row = 3;
		// ��������Ϊ΢���ź�

		CellStyle style = book.createCellStyle();// �½���ʽ����
		XSSFFont font = (XSSFFont) book.createFont();// �����������
		font.setFontName("΢���ź�");
		style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);
		style.setFont(font);

		int len = list.size();

		// �㺸�����������ݼ�
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
		// ����ͬ��λ��ͬ�㺸�������ư��ղ��ұ������
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
		System.out.println("���ظ����Ƶģ�" + renameList);
		System.out.println("sheet����map��" + sortmap);

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
			// ������Ϣ����{�㺸�������ƣ���ţ�����ID}
			for (int j = 1; j < values.length - 3; j++) {

				setCellwithCellStyle(sheet, values[j], Integer.parseInt(values[1]) + 2, j + 10, Cell.CELL_TYPE_STRING,
						style);
			}
			// row++;
		}

	}

	// ����sheet���ݣ��������ݸ���
	public static void updateWeldDataToSheet(XSSFWorkbook book, ArrayList list, ArrayList partlist, boolean setGroup) {
		// TODO Auto-generated method stub

		// ��������Ϊ΢���ź�

		CellStyle style = book.createCellStyle();// �½���ʽ����
		XSSFFont font = (XSSFFont) book.createFont();// �����������
		font.setFontName("΢���ź�");
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
		// ����ͬ��λ��ͬ�㺸�������ư��ղ��ұ������
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
		System.out.println("���ظ����Ƶģ�" + renameList);
		System.out.println("sheet����map��" + sortmap);

		// ѭ��sheetҳ
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
			// ���ݵ㺸����IDƥ��sheetҳ
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
				// �ж��µĺ�����Ϣ�Ƿ��ھɵĺ��������У�������ھͼ�¼������Ϣ��P=��S=
				String[] hdstr = (String[]) hdlist.get(k);
				boolean flag = false;
				for (int m = 3; m < oldnum; m++) {
					XSSFRow row = sheet.getRow(m);
					if (row != null) {
						// �Ƚ�weid_id
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
				// ������ھͱ���P��Sֵ
				if (flag) {
					if (hzrow != null) {
						hdstr[3] = getCellValue(hzrow.getCell(13));// P=
						hdstr[4] = getCellValue(hzrow.getCell(14));// S=
					}
				}
				// ������Ϣ����{�㺸�������ƣ���ţ�����ID}
				for (int n = 1; n < hdstr.length - 3; n++) {
					setCellwithCellStyle(sheet, hdstr[n], Integer.parseInt(hdstr[1]) + 2, n + 10, Cell.CELL_TYPE_STRING,
							style);
				}
			}
			// ����������ݱ������ݶ���Ҫ�Ѷ�����������ÿ�
			System.out.println("�����ݣ�" + oldnum + " ������" + newnum + 3);
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

	// дֱ���嵥sheet����
	public static void writeDataToSheet(XSSFWorkbook book, ArrayList list, int num, int num2, int rownum,
			ReportViwePanel viewPanel, int ajm) {
		// TODO Auto-generated method stub
		// int row = 3;
		// XSSFFormulaEvaluator evaluator=new XSSFFormulaEvaluator(book);
		viewPanel.addInfomation("", 40, 100);

		XSSFSheet sheet = book.getSheetAt(0);
		System.out.println("���ǰ������" + sheet.getPhysicalNumberOfRows());
		// ��д������������
		setStringCell(sheet, "�ܺ�����" + num, 1, 4, false);
		setStringCell(sheet, "RSW����" + num2, 1, 10, false);
		// �������������ƶ�list.get(list.size()-2)
		int downnum = (int) list.get(list.size() - 2);
		if (rownum == 11) {
			downnum = downnum - 1; // ����ǵ�һ�Σ��Ѿ�����һ�У����Լ�ȥ
		}
		// ���ֻ��metal��ֻ��һ�����ݵ�����£����ֱ�������
		if (downnum > 0) {
			sheet.shiftRows(rownum, rownum + 12, downnum, true, false);
		}

		System.out.println("��Ӻ�������" + sheet.getPhysicalNumberOfRows());

		// ͨ��ģ���и���������д���ݣ�����ʱ����ɾ�����ģ����
		String formula = "";
		for (int i = 0; i < downnum; i++) {
			viewPanel.addInfomation("", 40, 100);
			copyRows(sheet, 11, 11, rownum + i);
			// I�й�ʽ 0.22*D11+0.11*E11+0.08*F11+0.05*G11+0.04*H11
			String cs = Integer.toString((rownum + i + 1));
			formula = "0.22*D" + cs + "+0.11*E" + cs + "+0.08*F" + cs + "+0.05*G" + cs + "+0.04*H" + cs;
			setCellFormula(sheet, formula, rownum + i, 8);
			System.out.println(formula);
			// M�й�ʽ 0.06*J11+0.022*L11
			formula = "0.06*J11+0.022*L11";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 12);
			// X�й�ʽ
			// IF(ISBLANK(N11),0,0.055*N11+0.11)+IF(ISBLANK(O11),0,0.03*O11+0.2)+IF(ISBLANK(P11),0,0.035*P11+0.12)+IF(ISBLANK(Q11),0,0.03*Q11+0.4)+IF(ISBLANK(R11),0,0.04*R11+0.11)+IF(ISBLANK(S11),0,0.04*S11+0.11)+IF(ISBLANK(T11),0,0.04*T11+0.11)+IF(ISBLANK(V11),0,0.03*V11+0.07)+IF(ISBLANK(W11),0,0.2*W11+0.07)
			formula = "IF(ISBLANK(N" + cs + "),0,0.055*N" + cs + "+0.11)+IF(ISBLANK(O" + cs + "),0,0.03*O" + cs
					+ "+0.2)+IF(ISBLANK(P" + cs + "),0,0.035*P" + cs + "+0.12)+IF(ISBLANK(Q" + cs + "),0,0.03*Q" + cs
					+ "+0.4)+IF(ISBLANK(R" + cs + "),0,0.04*R" + cs + "+0.11)+IF(ISBLANK(S" + cs + "),0,0.04*S" + cs
					+ "+0.11)+IF(ISBLANK(T" + cs + "),0,0.04*T" + cs + "+0.11)+IF(ISBLANK(V" + cs + "),0,0.03*V" + cs
					+ "+0.07)+IF(ISBLANK(W" + cs + "),0,0.2*W" + cs + "+0.07)";
			setCellFormula(sheet, formula, rownum + i, 23);
			// AG�й�ʽ
			// IF(ISBLANK(Y11),0,0.06*Y11+0.06)+IF(ISBLANK(Z11),0,0.05*Z11+0.22)+IF(ISBLANK(AA11),0,0.05*AA11)+IF(ISBLANK(AB11),0,0.5*AB11+0.18)+IF(ISBLANK(AC11),0,0.025*AC11+0.35)+IF(ISBLANK(AD11),0,0.01*AD11+0.21)+IF(ISBLANK(AE11),0,0.5*AE11+0.21)+IF(ISBLANK(AF11),0,2*AF11)
			formula = "IF(ISBLANK(Y11),0,0.06*Y11+0.06)+IF(ISBLANK(Z11),0,0.05*Z11+0.22)+IF(ISBLANK(AA11),0,0.05*AA11)+IF(ISBLANK(AB11),0,0.5*AB11+0.18)+IF(ISBLANK(AC11),0,0.025*AC11+0.35)+IF(ISBLANK(AD11),0,0.01*AD11+0.21)+IF(ISBLANK(AE11),0,0.5*AE11+0.21)+IF(ISBLANK(AF11),0,2*AF11)";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 32);
			// AH�й�ʽ (I11+M11+X11+AG11)
			formula = "(I11+M11+X11+AG11)";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 33);
			// AJ�й�ʽ AH11*AI11
			formula = "AH11*AI11";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 35);
			// AL�й�ʽ AH11*AI11*AK11
			formula = "AH11*AI11*AK11";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 37);
			// AM�й�ʽ ($AG$6*60*$AG$5)/$AG$1

			// AN�й�ʽ AL11/AM11
			formula = "AL11/AM11";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 39);

			// AP�й�ʽ AO11*AG$3/AG$3
			formula = "AO11*AG$3/AG$3";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 41);
			// AQ�й�ʽ AP11*$AG$6/AG$3
			formula = "AP11*$AG$6/AG$3";
			setCellFormula(sheet, formula.replace("11", cs), rownum + i, 42);
			// }
		}
		System.out.println("���ƺ�������" + sheet.getPhysicalNumberOfRows());
		// ��պ���������

		// ���ú�װ����ͳ���еĵ�Ԫ��ʽ
		int rowIndex = downnum + rownum - 1;
		// ��ʼ�ͽ�����ͳ��
		int start = rownum + 1;
		int end = downnum + rownum - 1;
		if (rownum == 11) {
			start = start - 1;
		}
		String[] strcell = { "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
				"T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK",
				"AL", "AM", "AN", "AO", "AP", "AQ" };
		// ���պ������check��3��
		if (ajm != 2) {
			for (int j = 0; j < strcell.length; j++) {

				viewPanel.addInfomation("", 40, 100);

				// ���ò��ߺϼ���
				if (strcell[j] != "B" && strcell[j] != "C" && strcell[j] != "AI" && strcell[j] != "AK"
						&& strcell[j] != "AM" && strcell[j] != "AP") {
					formula = "SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")";
					System.out.println("���ߺϼ��У�" + formula);
					setCellFormula(sheet, formula, rowIndex, 1 + j);
				} else {
					setStringCell(sheet, "", rowIndex, 1 + j, true);
				}
				// ���ò��ߺϼ��е���һ�����ݶ�Ϊ��
				setStringCell(sheet, "", rowIndex - 1, 1 + j, true);

				// ����WELD CHECKER�е���������Ϊ��
				if (strcell[j] == "D" || strcell[j] == "E" || strcell[j] == "F" || strcell[j] == "G"
						|| strcell[j] == "H" || strcell[j] == "J" || strcell[j] == "K" || strcell[j] == "L") {
					setIntCell(sheet, null, rowIndex - 2, 1 + j, true);
				}

			}
			// ���ò��ߺϼ��е�������
			setStringCell(sheet, list.get(list.size() - 1).toString() + " WELD CHECKER", rowIndex - 2, 2, true);
			setIntCell(sheet, Integer.toString((int) list.get(list.size() - 2) - 2), rowIndex - 2, 1, true);
			// ���ò��ߺϼ���
			setStringCell(sheet, list.get(list.size() - 1).toString() + " SUB TOTAL", rowIndex, 2, true);

		} else { // ����METALͳ�ƣ���ҵ����ȷ�ϣ�METAL����ֻ��һ�����������
			for (int j = 0; j < strcell.length; j++) {
				// ����METALͳ����
				if (strcell[j] != "B" && strcell[j] != "C" && strcell[j] != "AI" && strcell[j] != "AK"
						&& strcell[j] != "AM" && strcell[j] != "AP") {
					formula = "SUM(" + strcell[j] + start + ":" + strcell[j] + (end + 5) + ")";
					System.out.println("METALͳ�����У�" + formula);
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

	// д������ҵ��-����sheet����
	public static void writeDataToSheet(XSSFWorkbook book, String[] values) {
		XSSFSheet sheet = book.getSheetAt(0);

		// ��������
//		Font font = book.createFont();
//		font.setFontName("������");
//		font.setBoldweight(Font.BOLDWEIGHT_BOLD);// �Ӵ�
//		font.setFontHeightInPoints((short) 16);
//		// ����һ����ʽ
//
//		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_NONE); // �±߿�
//		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_NONE);// ��߿�
//		cellStyle.setBorderTop(XSSFCellStyle.BORDER_NONE);// �ϱ߿�
//		cellStyle.setBorderRight(XSSFCellStyle.BORDER_NONE);// �ұ߿�
//		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// �����
//		cellStyle.setFont(font);
		cellStyle = null;

		setStringCell(sheet, values[4], 5, 2, true);
		setStringCell(sheet, values[0], 6, 2, true);
		setStringCell(sheet, values[1], 7, 2, true);
		setStringCell(sheet, values[2], 8, 2, true);
		setStringCell(sheet, values[3], 10, 2, true);

	}

	// д����sheet����
	public static void writeDataToSheetByGeneral(XSSFWorkbook book, String[] values) {
		XSSFSheet sheet = book.getSheetAt(0);

		// ��������
		Font font = book.createFont();
		font.setFontName("������");
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);// �Ӵ�
		font.setFontHeightInPoints((short) 16);
		// ����һ����ʽ

		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_NONE); // �±߿�
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_NONE);// ��߿�
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_NONE);// �ϱ߿�
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_NONE);// �ұ߿�
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// �����
		cellStyle.setFont(font);

		setStringCell(sheet, values[0], 9, 3, true);
		setStringCell(sheet, values[1], 10, 3, true);
		setStringCell(sheet, values[2], 12, 1, false);
	}

	// дֱ�����Ķ����嵥����
	public static void writeStraightDataToSheet(XSSFWorkbook book, int sheetnum, String[] values, ArrayList list) {

		// ��������
		Font font = book.createFont();
		font.setColor((short) 12);
		font.setFontHeightInPoints((short) 12);
		// ����һ����ʽ
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����

		cellStyle.setFont(font);

		for (int i = 0; i < sheetnum; i++) {
			XSSFSheet sheet = book.getSheetAt(i + 1);

			// д������Ϣ���������ڡ���ǰҳ����ҳ
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

	// д������ҵ��-Ŀ¼sheet����
	public static XSSFWorkbook creatXSSFWorkbook(InputStream input, ArrayList list, List bzlist) {
		XSSFWorkbook book = null;

		try {
			book = new XSSFWorkbook(input);
			String sheetname = "";
			int page = 1; // ���ڱ��sheet���
			// ����Ŀ¼�����ж��Ƿ�ֶ��sheetҳ��ÿ40������һ��sheet
			int mlnum = list.size();
			int mlpage = mlnum / 40 + 1;
			if (mlpage > 1) {
				sheetname = "01Ŀ¼1";
				book.setSheetName(0, sheetname);
			}
			// Ŀ¼��sheetҳ�߼�
			for (int i = 1; i < mlpage; i++) {
				XSSFSheet sheet = book.cloneSheet(0);
				sheetname = String.format("%02d", i + 1) + "Ŀ¼" + (i + 1);
				book.setSheetName(book.getSheetIndex(sheet), sheetname);
				book.setSheetOrder(sheetname, i); // ����sheet˳��
				// ���ô�ӡ����
				int sheetIndex = book.getSheetIndex(sheet);
				book.setPrintArea(sheetIndex, 0, 115, 0, 51);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 65);// �Զ������ţ��˴�100Ϊ������
				printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)

				page = i + 1;
			}
			int bznnum = 0;
			// �����嵥sheet��ҳ�߼�
			if (bzlist != null) {
				bznnum = bzlist.size();
			}
			int bzpage = bznnum / 40 + 1;
			if (bzpage > 1) {
				sheetname = String.format("%02d", page + 1) + "��¼-�����嵥" + 1;
				book.setSheetName(page, sheetname);
			} else {
				sheetname = String.format("%02d", page + 1) + "��¼-�����嵥";
				page = page + 1;
				book.setSheetName(page - 1, sheetname);
			}

			int cnum = page;
			for (int j = 1; j < bzpage; j++) {
				XSSFSheet sheet = book.cloneSheet(cnum);
				sheetname = String.format("%02d", cnum + j + 1) + "��¼-�����嵥" + (j + 1);
				book.setSheetName(book.getSheetIndex(sheet), sheetname);
				book.setSheetOrder(sheetname, cnum + j); // ����sheet˳��
				// ���ô�ӡ����
				int sheetIndex = book.getSheetIndex(sheet);
				book.setPrintArea(sheetIndex, 0, 115, 0, 51);
				PrintSetup printSetup = sheet.getPrintSetup();
				printSetup.setPaperSize(XSSFPrintSetup.A4_PAPERSIZE);
				printSetup.setScale((short) 65);// �Զ������ţ��˴�100Ϊ������
				printSetup.setLandscape(true); // ��ӡ����true������false������(Ĭ��)
				page = cnum + j + 1;
			}
			// ����̶�ģ���sheet����
			int sheetnum = book.getNumberOfSheets();
			sheetname = String.format("%02d", page + 1) + "��¼-255��������1";
			book.setSheetName(sheetnum - 5, sheetname); // 03��¼-255��������1
			sheetname = String.format("%02d", page + 2) + "��¼-255��������2";
			book.setSheetName(sheetnum - 4, sheetname);// 04��¼-255��������2
			sheetname = String.format("%02d", page + 3) + "��¼-255��������3";
			book.setSheetName(sheetnum - 3, sheetname);// 05��¼-255��������3
			sheetname = String.format("%02d", page + 4) + "��¼-24��������";
			book.setSheetName(sheetnum - 2, sheetname);// 06��¼-24��������
			sheetname = String.format("%02d", page + 5) + "��¼-���ж��ձ�";
			book.setSheetName(sheetnum - 1, sheetname);// 07��¼-���ж��ձ�)

			cellStyle = book.createCellStyle();
			// �߿�
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
			cellStyle.setWrapText(true);
			cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
			cellStyle.setAlignment(XSSFCellStyle.VERTICAL_CENTER);
			// cellStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);// �����

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}

	// д������ҵ��-Ŀ¼sheet����
	public static void writeDataToSheet(XSSFWorkbook book, ArrayList plist, ArrayList list, List bzlist) {

		// ��������
//		Font font = book.createFont();
//		font.setColor((short) 12);
//		font.setFontHeightInPoints((short) 16);
		// ����һ����ʽ
		cellStyle = null;
//		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle.setFont(font);
//		Font font2 = book.createFont();
//		font2.setFontHeightInPoints((short) 18);
//		font2.setFontName("MS PGothic");
//		font2.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
//		XSSFCellStyle cellStyle1 = book.createCellStyle();
		XSSFCellStyle cellStyle1 = null;
//		cellStyle1.setBorderLeft(XSSFCellStyle.BORDER_NONE);
//		cellStyle1.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		cellStyle1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		cellStyle1.setFont(font2);

		// ѭ������sheetҳ���ѹ�����������д��
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

		// дĿ¼����
		int mlpage = list.size() / 40 + 1; // Ŀ¼����sheet��
		int totalpages = 0;// �ϼ���ҳ��
		for (int i = 0; i < mlpage; i++) {
			XSSFSheet sheet = book.getSheetAt(i);
			sheet.removeColumnBreak(55);
			if (i == mlpage - 1) {
				for (int j = 0; j + 40 * i < list.size(); j++) {
					String[] value = (String[]) list.get(j + 40 * i);
					String rowNo = Integer.toString(j + 1);
					setIntCell(sheet, rowNo, 7 + j, 1, false);// ���
					setStringCell(sheet, value[1], 7 + j, 4, false);// ������
					setStringCell(sheet, value[2], 7 + j, 14, false);// ������������
					setStringCell(sheet, value[3], 7 + j, 44, false);// ����Ӣ������
					setStringCell(sheet, value[4], 7 + j, 71, false);// ���
					setIntCell(sheet, value[5], 7 + j, 76, false);// ҳ��
					setStringCell(sheet, value[6], 7 + j, 81, false);// ��ע
					if (value[5] != null && !value[5].isEmpty()) {
						totalpages = totalpages + Integer.parseInt(value[5]);
					}
				}
				setStringCell(sheet, "�ϼ�", 46, 4, false);// �ϼ�
				setIntCell(sheet, Integer.toString(totalpages), 46, 76, false);// ��ע
			} else {
				for (int j = 0; j + 40 * i < 40 + 40 * i; j++) {
					String rowNo = Integer.toString(j + 1);
					String[] value = (String[]) list.get(j + 40 * i);
					setIntCell(sheet, rowNo, 7 + j, 1, false);// ���
					setStringCell(sheet, value[1], 7 + j, 4, false);// ������
					setStringCell(sheet, value[2], 7 + j, 14, false);// ������������
					setStringCell(sheet, value[3], 7 + j, 44, false);// ����Ӣ������
					setStringCell(sheet, value[4], 7 + j, 71, false);// ���
					setStringCell(sheet, value[5], 7 + j, 76, false);// ҳ��
					setStringCell(sheet, value[6], 7 + j, 81, false);// ��ע
					if (value[5] != null && !value[5].isEmpty()) {
						totalpages = totalpages + Integer.parseInt(value[5]);
					}
				}
			}
			setIntCell(sheet, Integer.toString(i + 1), 50, 109, false);// �ڼ�ҳ��
			setIntCell(sheet, Integer.toString(mlpage), 50, 112, false);// ��ҳ��
		}
		XSSFCellStyle style = null;
//		style.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		style.setAlignment(XSSFCellStyle.ALIGN_LEFT);

		XSSFCellStyle style2 = null;
//		style2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		style2.setBorderLeft(XSSFCellStyle.BORDER_THIN);
//		style2.setBorderRight(XSSFCellStyle.BORDER_THIN);
//		style2.setBorderTop(XSSFCellStyle.BORDER_THIN);
//		style2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		style2.setAlignment(XSSFCellStyle.ALIGN_CENTER);

		XSSFCellStyle style3 = null;
//		style3.setBorderBottom(XSSFCellStyle.BORDER_THIN);
//		style3.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);// ���ݴ�ֱ����
//		style3.setAlignment(XSSFCellStyle.ALIGN_CENTER);

		// д�����嵥����
		if (bzlist == null) {
			bzlist = new ArrayList();
		}
		int bzpage = bzlist.size() / 40 + 1; // ��������sheet��
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
						setIntCellBydouble(sheet, rowNo, 7 + j, 1, style2);// ���
					}
					if (partn != null) {
						setStringCellAndStyle(sheet, partn, 7 + j, 8, style);// �����
					}
					if (boardnumber != null) {
						setStringCellAndStyle(sheet, boardnumber, 7 + j, 25, style);// ������
					}
					if (boardname != null) {
						setStringCellAndStyle(sheet, boardname, 7 + j, 38, style);// ��������
					}
					if (partmaterial != null) {
						setStringCellAndStyle(sheet, partmaterial, 7 + j, 68, style);// ����
					}
					if (partthickness != null) {
						setDoubleCellAndStyle(sheet, partthickness, 7 + j, 81, style3);// ���
					}
					if (maunit != null) {
						setStringCellAndStyle(sheet, maunit, 7 + j, 87, style);// ���λ
					}
					if (sheetstrength != null) {
						setStringCellAndStyle(sheet, sheetstrength, 7 + j, 91, style3, 10);// ǿ��
					}
					if (thunit != null) {
						setStringCellAndStyle(sheet, thunit, 7 + j, 96, style);// ǿ�ȵ�λ
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
						setIntCellBydouble(sheet, rowNo, 7 + j, 1, style2);// ���
					}
					if (partn != null) {
						setStringCellAndStyle(sheet, partn, 7 + j, 8, style);// �����
					}
					if (boardnumber != null) {
						setStringCellAndStyle(sheet, boardnumber, 7 + j, 25, style);// ������
					}
					if (boardname != null) {
						setStringCellAndStyle(sheet, boardname, 7 + j, 38, style);// ��������
					}
					if (partmaterial != null) {
						setStringCellAndStyle(sheet, partmaterial, 7 + j, 68, style);// ����
					}
					if (partthickness != null) {
						setDoubleCellAndStyle(sheet, partthickness, 7 + j, 81, style3);// ���
					}
					if (maunit != null) {
						setStringCellAndStyle(sheet, maunit, 7 + j, 87, style);// ���λ
					}
					if (sheetstrength != null) {
						setStringCellAndStyle(sheet, sheetstrength, 7 + j, 91, style3, 10);// ǿ��
					}
					if (thunit != null) {
						setStringCellAndStyle(sheet, thunit, 7 + j, 96, style);// ǿ�ȵ�λ
					}
					if (gagi != null) {
						setStringCellAndStyle(sheet, gagi, 7 + j, 104, style);// GA/GI
					}
				}
			}
			setIntCell(sheet, Integer.toString(i + 1), 50, 109, false);// �ڼ�ҳ��
			setIntCell(sheet, Integer.toString(bzpage), 50, 112, false);// ��ҳ��
		}
	}

	// ����ļ�
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

			// ��excel
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

	// �Ե�Ԫ��ֵ��ǿ��ʹ���ı���ʽ
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

	// �Ե�Ԫ��ֵ��ǿ��ʹ���ı���ʽ
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

	// �Ե�Ԫ��ֵ��ǿ��ʹ���ı���ʽ
	public static void setStringCell(XSSFSheet sheet, String value, int rowIndex, int cellIndex, boolean flag) {

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
//		cell.setCellType(Cell.CELL_TYPE_STRING);
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		// ȥ��ɾ����
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

	// �Ե�Ԫ��ֵ��ǿ��ʹ������
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

	// �Ե�Ԫ��ֵ��ǿ��ʹ�����Σ�ͨ��double����ǿת
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

	// �Ե�Ԫ��ֵ��ǿ��ʹ��double
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

	// �Ե�Ԫ��ֵ
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

	// ����ָ����Ԫ���ֵ�������Ƿ���excel�д���
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

	// Excel��sheet���Ʋ��ܰ��������ַ������������ַ�
	private static String FilterSpecialCharacters(String str) {
		String name = null;
		name = str.replaceAll("[*|':;',\\\\\\\\[\\\\\\\\].<>/?~��@��%����&*��������+|{}������������������������]", "");
		return name;
		// return null;
	}

	/**
	 * ������
	 * 
	 * @param startRowIndex ��ʼ��
	 * @param endRowIndex   ������
	 * @param pPosition     Ŀ����ʼ��λ��
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
//			System.out.println("��Ԫ��ֵ��" + getCellValue(srcCell));
//			String Formula = srcCell.getCellFormula();
//			System.out.println("��Ԫ��ʽ��" + Formula);
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
		// ��ʽ
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

	// ���õ�Ԫ��ʽ
	public static void setCellFormula(XSSFSheet Sheet, String formula, int rowIndex, int cellIndex) {
		try {
			XSSFRow row = Sheet.getRow(rowIndex);
			Cell cell = row.getCell(cellIndex);
			cell.setCellFormula(formula);
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	// ���ɾ��ģ���в����û����еĹ�ʽ
	public static void dealTotalRowFormula(XSSFWorkbook book, ReportViwePanel viewPanel) throws IOException {
		XSSFSheet sheet = book.getSheetAt(0);

		// ���ú�װ����ͳ���еĵ�Ԫ��ʽ
		int total_rownum = sheet.getPhysicalNumberOfRows();
		int rowIndex = total_rownum - 9;
		// ��ʼ�ͽ�����ͳ��
		int start = 11;
		int end = total_rownum - 8;
		String[] strcell = { "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
				"V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM",
				"AN", "AO", "AP", "AQ" };
		for (int j = 0; j < strcell.length; j++) {
			viewPanel.addInfomation("", 70, 100);
			// С����
			if (strcell[j] == "AN" || strcell[j] == "AO") {
				setCellFormula(sheet, "SUM(" + strcell[j] + (rowIndex + 2) + ":" + strcell[j] + (rowIndex + 4) + ")",
						rowIndex + 4, 3 + j);
				System.out.println(
						"С����:" + "SUM(" + strcell[j] + (rowIndex + 1) + ":" + strcell[j] + (rowIndex + 3) + ")");
			}
			// �ϼ���
			if (strcell[j] != "AI" && strcell[j] != "AK" && strcell[j] != "AM" && strcell[j] != "AP") {
				String formula = "SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")/2";
				setCellFormula(sheet, formula, rowIndex + 5, 3 + j);
				System.out.println("�ϼ���:" + formula);
			}
			// ��Щ����Ҫ������λС��
//			if (strcell[j] == "I" || strcell[j] == "M" || strcell[j] == "X" || strcell[j] == "AG" || strcell[j] == "AH"
//					|| strcell[j] == "AJ" || strcell[j] == "AL" || strcell[j] == "AN" || strcell[j] == "AP") {
//
//				String formula = "ROUND(SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")/2,2)";
//				setCellFormula(sheet, formula, rowIndex + 5, 3 + j);
//				System.out.println("�ϼ���:" + formula);
//			}

			// ����������
			if (strcell[j] == "AH") {
				setCellFormula(sheet, "SUM(" + strcell[j] + (rowIndex + 6) + "+" + strcell[j] + (rowIndex + 4) + ")/60",
						rowIndex + 6, 3 + j);
				System.out.println(
						"����������:" + "(" + strcell[j] + (rowIndex + 3) + ":" + strcell[j] + (rowIndex + 5) + ")/60");
			}
			// �����ڶ���
			if (strcell[j] == "AH" || strcell[j] == "AJ" || strcell[j] == "AL" || strcell[j] == "AN"
					|| strcell[j] == "AO" || strcell[j] == "AQ") {
				setCellFormula(sheet, "SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")/2", rowIndex + 7,
						3 + j);
				System.out.println("�����ڶ���:" + "SUM(" + strcell[j] + start + ":" + strcell[j] + end + ")/2");
			}
			// ���һ��
			if (strcell[j] == "AH" || strcell[j] == "AJ" || strcell[j] == "AL") {
				setCellFormula(sheet, strcell[j] + (rowIndex + 8) + "/60", rowIndex + 8, 3 + j);
				System.out.println("���һ��1:" + strcell[j] + (rowIndex + 8) + "/60");
			}
			// ���һ��
			if (strcell[j] == "AO") {
				setCellFormula(sheet, strcell[j - 1] + (rowIndex + 8) + "/" + strcell[j] + (rowIndex + 8), rowIndex + 8,
						3 + j);
				System.out.println("С���һ��2:" + strcell[j - 1] + (rowIndex + 8) + "/" + strcell[j] + (rowIndex + 8));
			}
			// ���һ��
			if (strcell[j] == "AQ") {
				setCellFormula(sheet, strcell[j] + (rowIndex + 8) + "/" + strcell[j - 9] + (rowIndex + 9), rowIndex + 8,
						3 + j);
				System.out.println("���һ��3:" + strcell[j] + (rowIndex + 8) + "/" + strcell[j - 9] + (rowIndex + 9));
			}

			// excel���ù�ʽ�Զ�����
			sheet.setForceFormulaRecalculation(true);
		}
	}

	// ��ȡ��Ԫ������
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
			sheet.shiftRows(rowIndex + 1, lastRowNum, -1);// ���к�ΪrowIndex+1һֱ���к�ΪlastRowNum�ĵ�Ԫ��ȫ������һ�У��Ա�ɾ��rowIndex��
		if (rowIndex == lastRowNum) {
			XSSFRow removingRow = sheet.getRow(rowIndex);
			if (removingRow != null)
				sheet.removeRow(removingRow);
		}
	}

	// ����Ԫ��б��
	public static void CellSlash(XSSFSheet sheet, int x1, int x2, int y1, int y2) {

		return;
	}

	// ���ݽ׶����࣬��ʼ����ͷ
	public static HashMap InitializeHeader(XSSFWorkbook book, ArrayList<String> stage) {
		// TODO Auto-generated method stub
		HashMap map = new HashMap();
		if (stage.size() > 0) {

			XSSFSheet sheet = book.getSheetAt(0);
			XSSFRow row = sheet.getRow(15);
			XSSFRow row1 = sheet.getRow(16);
			for (int i = 0; i < stage.size(); i++) {
				// ��һ�ӽ׶Σ�����Ҫ���ƣ�ģ���Ѿ��������ˣ�ֱ�Ӹ�ֵ�׶����ƾ���
				XSSFCell cell = row.getCell(17); // ����
				XSSFCell cell1 = row1.getCell(8); // �ж�
				XSSFCell cell2 = row1.getCell(9); // ����
				XSSFCell cell4 = row1.getCell(18); // ����
				XSSFCell cell5 = row1.getCell(19); // �ճ�
				if (i == 0) {
					setStringCell(sheet, stage.get(i), 14, 8, false);
					map.put(stage.get(i), 8);
				} else {
					setStringCell(sheet, stage.get(i), 14, 8 + 13 * i, false);

					XSSFCell newcell = row.createCell(17 + 13 * i); // ����
					copyCell(book, cell, newcell);

					CellRangeAddress region1 = new CellRangeAddress(15, 15, (short) (17 + 13 * i),
							(short) (17 + 13 * i + 2)); // ����1����ʼ�� ����2����ֹ�� ����3����ʼ�� ����4����ֹ��
					sheet.addMergedRegion(region1);

					XSSFCell newcell1 = row1.createCell(8 + 13 * i); // �ж�
					copyCell(book, cell1, newcell1);
					for (int j = 0; j < 9; j++) {
						XSSFCell newcell2 = row1.createCell(9 + j + 13 * i); // ����
						copyCell(book, cell2, newcell2);
					}
					CellRangeAddress region2 = new CellRangeAddress(16, 16, (short) (9 + 13 * i),
							(short) (9 + 13 * i + 8)); // ����1����ʼ�� ����2����ֹ�� ����3����ʼ�� ����4����ֹ��
					sheet.addMergedRegion(region2);

					XSSFCell newcell4 = row1.createCell(18 + 13 * i); // ����
					copyCell(book, cell4, newcell4);
					XSSFCell newcell5 = row1.createCell(19 + 13 * i); // �ճ�
					copyCell(book, cell5, newcell5);

					map.put(stage.get(i), 8 + 13 * i);
				}
			}
		}
		return map;
	}

	// д����Ҫ���������ݣ�֧�ֶ���׶����
	public static void writeRequirementsDataToSheet(XSSFWorkbook book, ArrayList datalist, HashMap map) {
		// TODO Auto-generated method stub
		XSSFSheet sheet = book.getSheetAt(0);

		// ���������У��ı�Ҫ���Ͻ���ʾ
		CellStyle style = book.createCellStyle();

		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);

		for (int i = 0; i < datalist.size(); i++) {
			Object[] obj = (Object[]) datalist.get(i);

			// ���ݽ׶λ�ȡ��ʼ��
			int startColum = (int) map.get(obj[7].toString());

			setStringCell(sheet, obj[7].toString(), 14, startColum, false); // �׶�

			for (int k = 0; k < 9; k++) {
				setStringCell(sheet, obj[0].toString(), 17 + 9 * i + k, 0, true); // ����
				setStringCell(sheet, obj[1].toString(), 17 + 9 * i + k, 1, true); // �Ŀ
				setStringCell(sheet, obj[2].toString(), 17 + 9 * i + k, 2, true); // ����
				setStringCell(sheet, obj[3].toString(), 17 + 9 * i + k, 4, true); // �^ȥ���}��
				setStringCell(sheet, obj[4].toString(), 17 + 9 * i + k, 6, true); // �a��

				setStringCell(sheet, "", 17 + 9 * i + k, 3, true); // �^ȥ���}��
				setStringCell(sheet, "", 17 + 9 * i + k, 5, true); // �a��

				Iterator iter = map.entrySet().iterator();
				while (iter.hasNext()) {
					Map.Entry entry = (Entry) iter.next();
					int val = (int) entry.getValue();
					if (val == startColum) {
						setStringCell(sheet, obj[8].toString(), 17 + 9 * i + k, val, true); // �ж�

						for (int n = 0; n < 9; n++) {
							setStringCellwithcellStyle(sheet, obj[9].toString(), 17 + 9 * i + k, val + n + 1, style); // ����
						}

						setStringCell(sheet, obj[11].toString(), 17 + 9 * i + k, val + 10, true); // ����
						setStringCell(sheet, obj[12].toString(), 17 + 9 * i + k, val + 11, true); // �ճ�
					} else {
						setStringCell(sheet, "", 17 + 9 * i + k, val, true); // �ж�

						for (int n = 0; n < 9; n++) {
							setStringCellwithcellStyle(sheet, "", 17 + 9 * i + k, val + n + 1, style); // ����
						}

						setStringCell(sheet, "", 17 + 9 * i + k, val + 10, true); // ����
						setStringCell(sheet, "", 17 + 9 * i + k, val + 11, true); // �ճ�
					}
				}

			}
			if (((ArrayList) obj[5]).size() > 0) {
				writepicturetosheet(book, sheet, (ArrayList) obj[5], 17 + 9 * i, 3, false); // ��
			}
			if (((ArrayList) obj[6]).size() > 0) {
				writepicturetosheet(book, sheet, (ArrayList) obj[6], 17 + 9 * i, 5, false); // ͼֽ��ӛ����
			}
			if (((ArrayList) obj[10]).size() > 0) {
				writepicturetosheets(book, sheet, (ArrayList) obj[10], 18 + 9 * i, startColum + 1);
			}

			// �ϲ���Ԫ��
			CellRangeAddress region1;
			int[] colum = { 0, 1, 2, 3, 4, 5, 6 };
			for (int j = 0; j < colum.length; j++) {

				// ����1����ʼ��;����2����ֹ��;����3����ʼ��; ����4����ֹ��
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
						// ����1����ʼ��;����2����ֹ��;����3����ʼ��; ����4����ֹ��
						region1 = new CellRangeAddress(17 + 9 * i, 25 + 9 * i, (short) (val + 1), (short) (val + 9));
					} else {
						// ����1����ʼ��;����2����ֹ��;����3����ʼ��; ����4����ֹ��
						region1 = new CellRangeAddress(17 + 9 * i, 25 + 9 * i, (short) colum1[j], (short) colum1[j]);
					}
					sheet.addMergedRegion(region1);
				}
			}
		}
	}

	// д����Ҫ���������ݣ�ֻ�����ǰ�׶�
	public static void writeRequirementsDataToSheet(XSSFWorkbook book, ArrayList datalist) {
		// TODO Auto-generated method stub
		XSSFSheet sheet = book.getSheetAt(0);

		// ���������У��ı�Ҫ���Ͻ���ʾ
		CellStyle style = book.createCellStyle();

		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_TOP);
		style.setWrapText(true);

		for (int i = 0; i < datalist.size(); i++) {
			Object[] obj = (Object[]) datalist.get(i);

			setStringCell(sheet, obj[7].toString(), 14, 8, false); // �׶�

			for (int k = 0; k < 9; k++) {
				setStringCell(sheet, obj[0].toString(), 17 + 9 * i + k, 0, true); // ����
				setStringCell(sheet, obj[1].toString(), 17 + 9 * i + k, 1, true); // �Ŀ
				setStringCell(sheet, obj[2].toString(), 17 + 9 * i + k, 2, true); // ����
				setStringCell(sheet, obj[3].toString(), 17 + 9 * i + k, 4, true); // �^ȥ���}��
				setStringCell(sheet, obj[4].toString(), 17 + 9 * i + k, 6, true); // �a��

//				setStringCell(sheet, "", 17 + 9 * i + k, 3, true); // �^ȥ���}��
//				setStringCell(sheet, "", 17 + 9 * i + k, 5, true); // �a��
				setStringCellwithcellStyle(sheet, obj[13].toString(), 17 + 9 * i + k, 3, style); // ��
				setStringCellwithcellStyle(sheet, obj[14].toString(), 17 + 9 * i + k, 5, style); // ͼֽ��ӛ����

				setStringCell(sheet, obj[8].toString(), 17 + 9 * i + k, 8, true); // �ж�
				for (int m = 0; m < 9; m++) {
					setStringCellwithcellStyle(sheet, obj[9].toString(), 17 + 9 * i + k, 9 + m, style); // ����
				}

				setStringCell(sheet, obj[11].toString(), 17 + 9 * i + k, 18, true); // ����
				setStringCell(sheet, obj[12].toString(), 17 + 9 * i + k, 19, true); // �ճ�
			}
			if (((ArrayList) obj[5]).size() > 0) {
				if (obj[13].toString().isEmpty()) {
					writepicturetosheet(book, sheet, (ArrayList) obj[5], 17 + 9 * i, 3, false); // ��
				} else {
					writepicturetosheet(book, sheet, (ArrayList) obj[5], 18 + 9 * i, 3, true); // ��
				}
			}
			if (((ArrayList) obj[6]).size() > 0) {
				if (obj[14].toString().isEmpty()) {
					writepicturetosheet(book, sheet, (ArrayList) obj[6], 17 + 9 * i, 5, false); // ͼֽ��ӛ����
				} else {
					writepicturetosheet(book, sheet, (ArrayList) obj[6], 18 + 9 * i, 5, true); // ͼֽ��ӛ����
				}

			}
			if (((ArrayList) obj[10]).size() > 0) {
				writepicturetosheets(book, sheet, (ArrayList) obj[10], 18 + 9 * i, 9);
			}

			// �ϲ���Ԫ��
			CellRangeAddress region1;
			int[] colum = { 0, 1, 2, 3, 4, 5, 6, 8, 9, 18, 19 };
			for (int j = 0; j < colum.length; j++) {
				if (colum[j] == 9) {
					region1 = new CellRangeAddress(17 + 9 * i, 25 + 9 * i, (short) 9, (short) 17); // ����1����ʼ�� ����2����ֹ��
																									// ����3����ʼ�� ����4����ֹ��
				} else {
					region1 = new CellRangeAddress(17 + 9 * i, 25 + 9 * i, (short) colum[j], (short) colum[j]); // ����1����ʼ��
																												// ����2����ֹ��
																												// ����3����ʼ��
																												// ����4����ֹ��
				}
				sheet.addMergedRegion(region1);
			}
		}
	}

	// ���ݵ����ļ�дͼƬ��excel
	private static void writepicturetosheet(XSSFWorkbook book, XSSFSheet sheet, ArrayList obj, int rowindex,
			int colindex, boolean flag) {
		// �ȰѶ�������ͼƬ�ŵ�һ��ByteArrayOutputStream�У��Ա����ByteArray
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
			// ����ͼƬ
			patriarch.createPicture(anchor,
					book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	// ���ݶ���ļ�дͼƬ��excel
	private static void writepicturetosheets(XSSFWorkbook book, XSSFSheet sheet, ArrayList list, int rowindex,
			int colindex) {
		// �ȰѶ�������ͼƬ�ŵ�һ��ByteArrayOutputStream�У��Ա����ByteArray

		int num = list.size(); // ͼƬ����
		int x1 = 0;
		int x2 = 0;
		int y1 = 0;
		int y2 = 0;
		// ��¼��Ҫ��ʾ������
		int rows = (num + 2) / 3;
		// ÿ�еĸ߶�
		int hight = 8 / rows;
		// ÿ�еĿ��
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

			int number = (i + 3) / 3;// ���ڵڼ���
			int number2 = (i + 1) % 3;// ���ڵڼ���
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
				// ����ͼƬ
				patriarch.createPicture(anchor,
						book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}

	// �Ե�Ԫ��ֵ��ǿ��ʹ���ı���ʽ
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

	// �Ե�Ԫ��ֵ��ǿ��ʹ���ı���ʽ
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

	// �Ե�Ԫ��ֵ
	public static void setStringCellAndStyle(XSSFSheet sheet, String value, int rowIndex, int cellIndex,
			XSSFCellStyle Style, int celltype) {

		// �����������ַ��͵����� 10Ϊ���ͣ�11Ϊdouble��

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
	 * ���õ�Ԫ�������С
	 */
	public static void setFontSize(XSSFWorkbook book, Cell cell, short num) {
		Font font = book.createFont();
		font.setFontName("����");
		// font.setStrikeout(true);//ɾ����
		font.setColor(IndexedColors.BLUE.getIndex());
		font.setBoldweight(Font.BOLDWEIGHT_BOLD); // ����Ӵ�
		for (short k = num; k >= 9; k--) {
			font.setFontHeightInPoints(k);
			if (checkCellReasonable(cell, k)) {
				break;
			}
		}
		// �����Ԫ����ʽ���ǵ�����
		CellStyle cStyle = book.createCellStyle();
		cStyle.cloneStyleFrom(cell.getCellStyle());
		// cStyle.setWrapText(true);
		cStyle.setFont(font);
		if (cStyle != null) {
			cell.setCellStyle(cStyle);
		}

	}

	/**
	 * У�鵥Ԫ���е������С�Ƿ����
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
	 * ��ȡ��Ԫ����ܿ�ȣ���λ�����أ�
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
	 * ��ȡ��Ԫ�������������Ǻϲ���Ԫ�񣬾ͻ�ȡ�ܵ�����
	 */
	public static int getColNum(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		// �жϸõ�Ԫ���Ƿ��Ǻϲ����������
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
