package com.dfl.report.util;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.util.FileUtil;
import com.teamcenter.rac.kernel.TCComponent;

public class OutputDataToExcel3 {

	//	private static XSSFWorkbook book;
	private static XSSFCellStyle cellStyle;


	public OutputDataToExcel3() {
		// TODO Auto-generated constructor stub
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
			//�߿�
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
	// ����ģ�崴��Excel��ģ��
	public static XSSFWorkbook creatXSSFWorkbook(InputStream input, ArrayList uplist, ArrayList downlist, ArrayList coverlist,ArrayList commonlist) {
		XSSFWorkbook book = null;
		try {
			book = new XSSFWorkbook(input);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		if (uplist != null && uplist.size() > 0) {

			XSSFSheet sheet1 = book.getSheetAt(0);
			book.setSheetName(0, "����");
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

		}
		if (downlist != null && downlist.size() > 0) {
			// book = new XSSFWorkbook(input);
			XSSFSheet sheet2 = book.getSheetAt(1);
			book.setSheetName(1, "����");
			//////////// ���÷�����ʾ�Ϸ�/�·�
			sheet2.setRowSumsBelow(false);
			sheet2.setRowSumsRight(false);
			sheet2.setRowSumsBelow(false);
			sheet2.setRowSumsRight(false);
			cellStyle = book.createCellStyle();
			// �߿�
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		}
		if (coverlist != null && coverlist.size() > 0) {

			// book = new XSSFWorkbook(input);
			XSSFSheet sheet3 = book.getSheetAt(2);
			book.setSheetName(2, "COVER");
			//////////// ���÷�����ʾ�Ϸ�/�·�
			sheet3.setRowSumsBelow(false);
			sheet3.setRowSumsRight(false);
			sheet3.setRowSumsBelow(false);
			sheet3.setRowSumsRight(false);
			cellStyle = book.createCellStyle();
			// �߿�
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		}
		if (commonlist != null && commonlist.size() > 0) {

			// book = new XSSFWorkbook(input);
			XSSFSheet sheet4 = book.getSheetAt(3);
			book.setSheetName(3, "COMMON");
			//////////// ���÷�����ʾ�Ϸ�/�·�
			sheet4.setRowSumsBelow(false);
			sheet4.setRowSumsRight(false);
			sheet4.setRowSumsBelow(false);
			sheet4.setRowSumsRight(false);
			cellStyle = book.createCellStyle();
			// �߿�
			cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
			cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		}

		return book;
	}

	// ɾ����Sheetҳ
	@SuppressWarnings("null")
	public static XSSFWorkbook deleteXSSFWorkbook(XSSFWorkbook book, ArrayList uplist, ArrayList downlist,
			ArrayList coverlist, ArrayList commonlist) {

		if (uplist.isEmpty()) {
			// ɾ��Sheet1
			book.removeSheetAt(book.getSheetIndex("Sheet1"));
		}
		if (downlist.isEmpty()) {
			// ɾ��Sheet2
			book.removeSheetAt(book.getSheetIndex("Sheet2"));
		}
		if (coverlist.isEmpty()) {
			// ɾ��Sheet3
			book.removeSheetAt(book.getSheetIndex("Sheet3"));
		}
		if (commonlist.isEmpty()) {
			// ɾ��Sheet4
			book.removeSheetAt(book.getSheetIndex("Sheet4"));			
		}
		return book;
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
			//openFile(fullFileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}



	// ���ļ�
	public static void openFile(String fullFileName) {
		// TODO Auto-generated method stub
		try {
			System.out.println("cmd /c call " +'"'+fullFileName+'"');
			Runtime.getRuntime().exec("cmd /c call " +'"'+fullFileName+'"');
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	
	// д����
	public static void writeDataToSheet1(XSSFWorkbook book, String[][] data, HashMap<String, ArrayList> picmap,
			HashMap<String, ArrayList> phomap, boolean setGroup) {
		// TODO Auto-generated method stub

		// ��ֱ����
		cellStyle = book.createCellStyle();
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER_SELECTION);// ����һ�����и�ʽ
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		XSSFSheet sheet = book.getSheetAt(0);

		int row = 1;

		int len = data.length;
		for (int i = 0; i < len; i++) {

			if ((i + 10) % 10 == 0) {
				int rownum = Integer.parseInt(data[i][1]);
				System.out.println(rownum);
				int a = picmap.get(Integer.toString(rownum)).size();
				//д����״ͼʾ
				writePicture(book, sheet, picmap.get(Integer.toString(rownum)),
						picmap.get(Integer.toString(rownum)).size(), rownum - 1);
				//д��Ҫ��ͼʾ
				writePicture1(book, sheet, phomap.get(Integer.toString(rownum)),
						phomap.get(Integer.toString(rownum)).size(), rownum - 1);
			}
			String[] values = data[i];
			for (int j = 0; j < values.length; j++) {
				setCell(sheet, values[j], row, j);
				if(j==6) {
					sheet.addMergedRegion(new CellRangeAddress(10 * i + 1, 10 * i + 10, 6,15));// �ϲ���״ͼʾ��Ԫ��
					//setCell(sheet, values[j], row, j);
				}
				if(j<6||j>15){
					sheet.addMergedRegion(new CellRangeAddress(10 * i + 1, 10 * i + 10, j, j));// �ϲ���Ԫ��
					//setCell(sheet, values[j], row, j);
				}
			}
			row++;

		}

	}
	public static void writeDataToSheet2(XSSFWorkbook book, String[][] data1, HashMap<String, ArrayList> picmap1,
			HashMap<String, ArrayList> phomap1, boolean setGroup) {
		// TODO Auto-generated method stub

		// ��ֱ����
		cellStyle = book.createCellStyle();
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER_SELECTION);// ����һ�����и�ʽ
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		XSSFSheet sheet = book.getSheetAt(1);

		int row = 1;

		int len = data1.length;
		for (int i = 0; i < len; i++) {

			if ((i + 10) % 10 == 0) {
				int rownum = Integer.parseInt(data1[i][1]);
				int a = picmap1.get(Integer.toString(rownum)).size();
				//д����״ͼʾ
				writePicture(book, sheet, picmap1.get(Integer.toString(rownum)),
						picmap1.get(Integer.toString(rownum)).size(), rownum - 1);
				//д��Ҫ��ͼʾ
				writePicture1(book, sheet, phomap1.get(Integer.toString(rownum)),
						phomap1.get(Integer.toString(rownum)).size(), rownum - 1);
			}
			String[] values = data1[i];
			for (int j = 0; j < values.length; j++) {
				setCell(sheet, values[j], row, j);
				if(j==6) {
					sheet.addMergedRegion(new CellRangeAddress(10 * i + 1, 10 * i + 10, 6,15));// �ϲ���״ͼʾ��Ԫ��
					//setCell(sheet, values[j], row, j);
				}
				if(j<6||j>15){
					sheet.addMergedRegion(new CellRangeAddress(10 * i + 1, 10 * i + 10, j, j));// �ϲ���Ԫ��
					//setCell(sheet, values[j], row, j);
				}
			}
			row++;

		}

	}

	public static void writeDataToSheet3(XSSFWorkbook book, String[][] data2, HashMap<String, ArrayList> picmap2,
			HashMap<String, ArrayList> phomap2, boolean setGroup) {
		// TODO Auto-generated method stub

		// ��ֱ����
		cellStyle = book.createCellStyle();
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER_SELECTION);// ����һ�����и�ʽ
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		XSSFSheet sheet = book.getSheetAt(2);

		int row = 1;

		int len = data2.length;
		for (int i = 0; i < len; i++) {

			if ((i + 10) % 10 == 0) {
				int rownum = Integer.parseInt(data2[i][1]);
				int a = picmap2.get(Integer.toString(rownum)).size();
				//д����״ͼʾ
				writePicture(book, sheet, picmap2.get(Integer.toString(rownum)),
						picmap2.get(Integer.toString(rownum)).size(), rownum - 1);
				//д��Ҫ��ͼʾ
				writePicture1(book, sheet, phomap2.get(Integer.toString(rownum)),
						phomap2.get(Integer.toString(rownum)).size(), rownum - 1);
			}
			String[] values = data2[i];
			for (int j = 0; j < values.length; j++) {
				setCell(sheet, values[j], row, j);
				if(j==6) {
					sheet.addMergedRegion(new CellRangeAddress(10 * i + 1, 10 * i + 10, 6,15));// �ϲ���״ͼʾ��Ԫ��
					
				}
				if(j<6||j>15){
					sheet.addMergedRegion(new CellRangeAddress(10 * i + 1, 10 * i + 10, j, j));// �ϲ���Ԫ��
					//setCell(sheet, values[j], row, j);
				}
			}
			row++;

		}

	}
	public static void writeDataToSheet4(XSSFWorkbook book, String[][] data3, HashMap<String, ArrayList> picmap3,
			HashMap<String, ArrayList> phomap3, boolean setGroup) {
		// TODO Auto-generated method stub

		// ��ֱ����
		cellStyle = book.createCellStyle();
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER_SELECTION);// ����һ�����и�ʽ
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		XSSFSheet sheet = book.getSheetAt(3);

		int row = 1;

		int len = data3.length;
		for (int i = 0; i < len; i++) {

			if ((i + 10) % 10 == 0) {
				int rownum = Integer.parseInt(data3[i][1]);
				int a = picmap3.get(Integer.toString(rownum)).size();
				//д����״ͼʾ
				writePicture(book, sheet, picmap3.get(Integer.toString(rownum)),
						picmap3.get(Integer.toString(rownum)).size(), rownum - 1);
				//д��Ҫ��ͼʾ
				writePicture1(book, sheet, phomap3.get(Integer.toString(rownum)),
						phomap3.get(Integer.toString(rownum)).size(), rownum - 1);
			}
			String[] values = data3[i];
			for (int j = 0; j < values.length; j++) {
				setCell(sheet, values[j], row, j);
				if(j==6) {
					sheet.addMergedRegion(new CellRangeAddress(10 * i + 1, 10 * i + 10, 6,15));// �ϲ���״ͼʾ��Ԫ��
					
				}
				if(j<6||j>15){
					sheet.addMergedRegion(new CellRangeAddress(10 * i + 1, 10 * i + 10, j, j));// �ϲ���Ԫ��
				//	setCell(sheet, values[j], row, j);
				}

			}
			row++;

		}

	}

	public static void writeExcelHeader(XSSFSheet sheet, String[] title, int[] cloumnWidth) {
		// TODO Auto-generated method stub
		for (int i = 0; i < title.length; i++) {
			setCell(sheet, title[i], 0, i);
		}

		for (int i = 0; i < cloumnWidth.length; i++) {
			sheet.setColumnWidth(i, cloumnWidth[i]);
		}
	}

	// �Ե�Ԫ��ֵ��ǿ��ʹ���ı���ʽ
	public static void setCell(XSSFSheet sheet, String value, int rowIndex, int cellIndex) {
		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null) {
			cell = row.createCell(cellIndex);

		}
		cell.setCellType(Cell.CELL_TYPE_STRING);
		
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue(value);
		//cell.setCellStyle(cellStyle);
	}
	// �Ե�Ԫ��ֵ��ǿ��ʹ���ı���ʽ
	public static void setStringCell(XSSFSheet sheet, String value, int rowIndex, int cellIndex, boolean flag) {

		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		// ȥ��ɾ����
		if (flag) {
			// cell.getCellStyle().getFont().getStrikeout();
			cell.getCellStyle().getFont().setStrikeout(false);
		}
		cell.setCellValue(value);
		if (flag) {
			cell.setCellStyle(cellStyle);
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

		cell.setCellStyle(style);

	}

	// д������Ҫ����״ͼƬ
	public static void writePicture(XSSFWorkbook book, XSSFSheet sheet, ArrayList piclist, int pnumber, int rowIndex) {
		int x1;
		int y1;
		int x2;
		int y2;
		try {
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			BufferedImage bufferImg = null;// ͼƬ
			// �����ж�����ͼƬ��ѭ�����
			for (int i = 0; i < pnumber; i++) {
				// TODO Auto-generated method stub

				// �ȰѶ�������ͼƬ�ŵ�һ��ByteArrayOutputStream�У��Ա����ByteArray
				ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
				// ��ͼƬ����BufferedImage
				bufferImg = ImageIO.read(new File(piclist.get(i).toString()));
				// ��ͼƬд������
				ImageIO.write(bufferImg, "png", byteArrayOut);

				// ��¼��Ҫ��ʾ������
				int num = (pnumber + 1) / 2;
				// ÿ�еĿ��
				int width = 10 / num;
				// int hight = 10 / num;

				// ͼƬΪһ�ŵ��������ʾһ�е����
				if (pnumber == 1) {

					XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) 6, 1 + 10 * rowIndex, (short) 16,
							11 + 10 * rowIndex);
					anchor.setAnchorType(0);
					patriarch.createPicture(anchor,
							book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
				}
				// ͼƬΪ���ŵ��������ʾһ�е����
				if (pnumber == 2) {
					if (i == 0) {
						XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) 6, 1 + 10 * rowIndex,
								(short) 11, 11 + 10 * rowIndex);
						anchor.setAnchorType(0);
						patriarch.createPicture(anchor,
								book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
					} else {
						XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) 11, 1 + 10 * rowIndex,
								(short) 16, 11 + 10 * rowIndex);
						anchor.setAnchorType(0);
						patriarch.createPicture(anchor,
								book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
					}
				}
				// ͼƬ�������ŵ��������ʾ���е����
				if (pnumber > 2) {

					// ��¼��ǰͼƬ���ڵڼ���
					int row = (i + 2) / 2;

					if ((i + 1) % 2 == 0) {
						// ���i+1Ϊż��
						x1 = width * (row - 1) + 6;
						x2 = width * row + 6;
						y1 = 6 + 10 * rowIndex;
						y2 = 11 + 10 * rowIndex;
					} else {
						// ���i+1Ϊ����
						x1 = width * (row - 1) + 6;
						x2 = width * row + 6;
						y1 = 1 + 10 * rowIndex;
						y2 = 6 + 10 * rowIndex;
					}
					XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) x1, y1, (short) x2, y2);
					anchor.setAnchorType(0);
					patriarch.createPicture(anchor,
							book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
				}
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	//��������Ҫ��Ҫ��ͼʾ
	public static void writePicture1(XSSFWorkbook book, XSSFSheet sheet, ArrayList pholist, int phonumber, int rowIndex) {
		// TODO Auto-generated method stub
		try {
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			BufferedImage bufferImg = null;// ͼƬ
			// �����ж�����ͼƬ��ѭ�����
			for (int i = 0; i < phonumber; i++) {
				// TODO Auto-generated method stub

				// �ȰѶ�������ͼƬ�ŵ�һ��ByteArrayOutputStream�У��Ա����ByteArray
				ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
				// ��ͼƬ����BufferedImage
				bufferImg = ImageIO.read(new File(pholist.get(i).toString()));
				// ��ͼƬд������
				ImageIO.write(bufferImg, "png", byteArrayOut);

				// ͼƬΪһ��
				if (phonumber == 1) {

					XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) 19, 1 + 10 * rowIndex, (short) 20,
							11 + 10 * rowIndex);
					anchor.setAnchorType(0);
					patriarch.createPicture(anchor,
							book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
				}

			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	//д��ˮ��Ҫ������
	public static void writeDataToSheet(XSSFWorkbook book, ArrayList datalist) {
		// TODO Auto-generated method stub
		XSSFSheet sheet = book.getSheetAt(0);
		CellStyle cellstyle = book.createCellStyle();	
		// ��ֱ����
		cellStyle = book.createCellStyle();
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER_SELECTION);// ����һ�����и�ʽ
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);
		// ���ڼ����Ŀ�У��ı�Ҫ��ֱ���У�����룬�Զ�����
		CellStyle style = book.createCellStyle();

		style.setBorderBottom(XSSFCellStyle.BORDER_THIN); // �±߿�
		style.setBorderLeft(XSSFCellStyle.BORDER_THIN);// ��߿�
		style.setBorderTop(XSSFCellStyle.BORDER_THIN);// �ϱ߿�
		style.setBorderRight(XSSFCellStyle.BORDER_THIN);// �ұ߿�
		style.setVerticalAlignment(XSSFCellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);	
		style.setWrapText(true);//�����Զ�����
		for (int i = 0; i < datalist.size(); i++) {
			Object[] obj = (Object[]) datalist.get(i);
			setStringCell(sheet, obj[8].toString(), 47, 26, false); // �׶�     ����Ҫ�ϲ���26��33��
			for (int k = 0; k < 9; k++) {
				setStringCell(sheet, obj[0].toString(), 50 + 9 * i + k, 0, true); // NO
				setStringCell(sheet, obj[1].toString(), 50 + 9 * i + k, 1, true); // ����
				setStringCell(sheet, obj[2].toString(), 50 + 9 * i , 2, true); // ��λ���ϲ�����50��52��2��6��
				setStringCell(sheet, obj[2].toString(), 51 + 9 * i , 2, true); // ��λ���ϲ�����50��52��2��6��
				setStringCell(sheet, obj[2].toString(), 52 + 9 * i , 2, true); // ��λ���ϲ�����50��52��2��6��
				setStringCell(sheet, obj[3].toString(), 53 + 9 * i , 2, true); // ���ù��죨�ϲ�����53��55��2��6��
				setStringCell(sheet, obj[3].toString(), 54 + 9 * i , 2, true); // ��λ���ϲ�����50��52��2��6��
				setStringCell(sheet, obj[3].toString(), 55 + 9 * i , 2, true); // ��λ���ϲ�����50��52��2��6��
				setStringCell(sheet, obj[4].toString(), 56 + 9 * i , 2, true); // ��λ���ϲ�����50��52��2��6��
				setStringCell(sheet, obj[4].toString(), 57 + 9 * i , 2, true); // ��λ���ϲ�����50��52��2��6��
				setStringCell(sheet, obj[4].toString(), 58 + 9 * i , 2, true); // ���ߺ�Ҫ�򣨺ϲ�����56��58��2��6��
				setStringCell(sheet, "", 50 + 9 * i + k, 3, true); // ���ߺ�Ҫ�򣨺ϲ�����56��58��2��6��
				setStringCell(sheet, "", 50 + 9 * i + k, 4, true); // ���ߺ�Ҫ�򣨺ϲ�����56��58��2��6��
				setStringCell(sheet, "", 50 + 9 * i + k, 5, true); // ���ߺ�Ҫ�򣨺ϲ�����56��58��2��6��
				setStringCell(sheet, "", 50 + 9 * i + k, 6, true); // ���ߺ�Ҫ�򣨺ϲ�����56��58��2��6��
				setStringCellwithcellStyle(sheet, obj[6].toString(), 50 + 9 * i + k, 17, style); // �����Ŀ
				setStringCellwithcellStyle(sheet, "", 50 + 9 * i + k, 18, style); // �����Ŀ
				setStringCellwithcellStyle(sheet, "", 50 + 9 * i + k, 19, style); // �����Ŀ
				setStringCellwithcellStyle(sheet, "", 50 + 9 * i + k, 20, style); // �����Ŀ
				setStringCellwithcellStyle(sheet, "", 50 + 9 * i + k, 21, style); // �����Ŀ
				setStringCellwithcellStyle(sheet, "", 50 + 9 * i + k, 22, style); // �����Ŀ
				setStringCellwithcellStyle(sheet, "", 50 + 9 * i + k, 23, style); // �����Ŀ
				
				setStringCell(sheet, obj[7].toString(), 50 + 9 * i + k, 24, true); // ��ע
				setStringCell(sheet, "", 50 + 9 * i + k, 25, true); // ��ע
				setStringCell(sheet, obj[9].toString(), 50 + 9 * i + k, 26, true); // ���
				setStringCellwithcellStyle(sheet, obj[11].toString(), 50 + 9 * i + k, 33, style); // ��������
				
				setStringCell(sheet, "", 50 + 9 * i + k, 7, true); // ˮ����Ҫע�ⲿλ
				setStringCell(sheet, "", 50 + 9 * i + k, 8, true); // ˮ����Ҫע�ⲿλ
				setStringCell(sheet, "", 50 + 9 * i + k, 9, true); // ˮ����Ҫע�ⲿλ
				setStringCell(sheet, "", 50 + 9 * i + k, 10, true); // ˮ����Ҫע�ⲿλ
				setStringCell(sheet, "", 50 + 9 * i + k, 11, true); // ˮ����Ҫע�ⲿλ
				setStringCell(sheet, "", 50 + 9 * i + k, 12, true); // ˮ����Ҫע�ⲿλ
				setStringCell(sheet, "", 50 + 9 * i + k, 13, true); // ˮ����Ҫע�ⲿλ
				setStringCell(sheet, "", 50 + 9 * i + k, 14, true); // ˮ����Ҫע�ⲿλ
				setStringCell(sheet, "", 50 + 9 * i + k, 15, true); // ˮ����Ҫע�ⲿλ
				setStringCell(sheet, "", 50 + 9 * i + k, 16, true); // �հ���
				setStringCell(sheet, "", 50 + 9 * i + k, 27, true); // ��״ͼʾ
				setStringCell(sheet, "", 50 + 9 * i + k, 28, true); // ��״ͼʾ
				setStringCell(sheet, "", 50 + 9 * i + k, 29, true); // ��״ͼʾ
				setStringCell(sheet, "", 50 + 9 * i + k, 30, true); // ��״ͼʾ
				setStringCell(sheet, "", 50 + 9 * i + k, 31, true); // ��״ͼʾ
				setStringCell(sheet, "", 50 + 9 * i + k, 32, true); // ��״ͼʾ
	
			}
			if (((ArrayList) obj[5]).size() > 0) {
				writepicturetosheet(book, sheet, (ArrayList) obj[5], 50 + 9 * i,7); // ע�ⲿλͼƬ
			}
			if (((ArrayList) obj[10]).size() > 0) {
				writepicturetosheets(book, sheet, (ArrayList) obj[10], 50 + 9 * i, 27);//��״ͼƬ���ϲ���10��15��
			}
			if (i == 0 && ((ArrayList) obj[12]).size() > 0) {

				writepicturetosheet1(book, sheet, (ArrayList) obj[12],6,2); // Ҫ��λ��ͼƬ
			}
			//�ϲ���Ԫ��
			CellRangeAddress region1;
			int[] colum = {0,1,2,3,4,7,16,17,24,26,27, 33};
			for (int j = 0; j < colum.length; j++) {
				if (colum[j] == 2) {//��λ
					region1 = new CellRangeAddress(50 + 9 * i, 52 + 9 * i, (short) 2, (short) 6); // ����1����ʼ�� ����2����ֹ��
																									// ����3����ʼ�� ����4����ֹ��
				} 
				else if (colum[j] == 3) {//���ù���
					region1 = new CellRangeAddress(53 + 9 * i, 55 + 9 * i, (short) 2, (short) 6); // ����1����ʼ�� ����2����ֹ��
																									// ����3����ʼ�� ����4����ֹ��
				}
				else if (colum[j] == 4) {//���ߺ�Ҫ��
					region1 = new CellRangeAddress(56 + 9 * i, 58 + 9 * i, (short) 2, (short) 6); // ����1����ʼ�� ����2����ֹ��
																									// ����3����ʼ�� ����4����ֹ��
				}
				else
					if (colum[j] == 27) {//��״ͼʾ
					region1 = new CellRangeAddress(50 + 9 * i, 58 + 9 * i, (short) 27, (short) 32); // ����1����ʼ�� ����2����ֹ��
																									// ����3����ʼ�� ����4����ֹ��
				}else if(colum[j] == 7) {
					region1 = new CellRangeAddress(50 + 9 * i, 58 + 9 * i, (short) 7, (short) 15); // ����1����ʼ�� ����2����ֹ��
																									// ����3����ʼ�� ����4����ֹ��
				}
				else if(colum[j] == 17) {
					region1 = new CellRangeAddress(50 + 9 * i, 58 + 9 * i, (short) 17, (short) 23); // ����1����ʼ�� ����2����ֹ��
																									// ����3����ʼ�� ����4����ֹ��
				}
				else if(colum[j] == 24) 
					{
					region1 = new CellRangeAddress(50 + 9 * i, 58 + 9 * i, (short) 24, (short) 25); // ����1����ʼ�� ����2����ֹ��
																									// ����3����ʼ�� ����4����ֹ��
				}
				else {
					region1 = new CellRangeAddress(50 + 9 * i, 58 + 9 * i, (short) colum[j], (short) colum[j]); // ����1����ʼ��// ����2����ֹ��
																												// ����3����ʼ��// ����4����ֹ��
				}				
				sheet.addMergedRegion(region1);
			}
		}

	}
	//д��ˮ��Ҫ��λ��ͼƬ
	private static void writepicturetosheet1(XSSFWorkbook book, XSSFSheet sheet, ArrayList obj, int rowindex, int colindex) {
		// TODO Auto-generated method stub
		// �ȰѶ�������ͼƬ�ŵ�һ��ByteArrayOutputStream�У��Ա����ByteArray
				ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
				BufferedImage bufferImg;
				try {
					File file = (File) obj.get(0);
					bufferImg = ImageIO.read(file);
					ImageIO.write(bufferImg, "png", byteArrayOut);
					XSSFDrawing patriarch = sheet.createDrawingPatriarch();
					XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) colindex, rowindex,
							(short) (colindex)+22, rowindex+40);
					anchor.setAnchorType(0);
					// ����ͼƬ
					patriarch.createPicture(anchor,
							book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
	}
	// ���ݵ����ļ�дͼƬ��excel
	private static void writepicturetosheet(XSSFWorkbook book, XSSFSheet sheet, ArrayList obj, int rowindex,
			int colindex) {
		// �ȰѶ�������ͼƬ�ŵ�һ��ByteArrayOutputStream�У��Ա����ByteArray
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		BufferedImage bufferImg;
		try {
			File file = (File) obj.get(0);
			bufferImg = ImageIO.read(file);
			ImageIO.write(bufferImg, "png", byteArrayOut);
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) colindex, rowindex,
					(short) (colindex + 9), rowindex + 9);
			anchor.setAnchorType(0);
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
			int hight = 9 / rows;
			// ÿ�еĿ��
			int width = 0;

			if (num < 4) {
				width = 6 / num;
			} else {
				width = 6 / 3;
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
					XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) y1, x1, (short) y2, x2);
					anchor.setAnchorType(0);
					// ����ͼƬ
					patriarch.createPicture(anchor,
							book.addPicture(byteArrayOut.toByteArray(), XSSFWorkbook.PICTURE_TYPE_JPEG));
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}

		}

}