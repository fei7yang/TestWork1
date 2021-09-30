package com.dfl.report.util;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
//import org.apache.poi.hssf.usermodel.HSSFCell;
//import org.apache.poi.hssf.usermodel.HSSFRow;
//import org.apache.poi.hssf.usermodel.HSSFSheet;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.eclipse.swt.widgets.TableItem;

public class OutputDataToExcel {
	private XSSFWorkbook book;
	private ArrayList result1 = new ArrayList();
	// private String[] title;//��̬
	// private int[] cloumnWidth;// �����п�
	private String reportname;// ��������
	private InputStream inputStream;
	private XSSFCellStyle cellStyle;
	private ReportViwePanel viewPanel;

	public OutputDataToExcel(ArrayList result1, InputStream inputStream, String reportname,ReportViwePanel viewPanel) {
		// TODO Auto-generated constructor stub
		this.result1 = result1;
		// this.title = title;
		this.reportname = reportname;
		this.inputStream = inputStream;
		this.viewPanel = viewPanel;
		executeOperation();
	}

	public void executeOperation() {
		// TODO Auto-generated method stub
		try {
			book = new XSSFWorkbook(inputStream);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		XSSFSheet sheet1 = book.getSheetAt(0);

		// XSSFSheet sheet2 = book.cloneSheet(1);
		// HSSFSheet sheet1=book.createSheet("Sheet1");

		//////////// ���÷�����ʾ�Ϸ�/�·�
		sheet1.setRowSumsBelow(false);
		sheet1.setRowSumsRight(false);
		sheet1.setRowSumsBelow(false);
		sheet1.setRowSumsRight(false);

		cellStyle = sheet1.getWorkbook().createCellStyle();
		// �߿�
		cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);
		cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);

		// ��ֱ����
		cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//		  cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER_SELECTION); // ����һ�����и�ʽ

//		String[] title = new String[]{"���","����","����"};//�����ͷ
//		
//		int[] cloumnWidth = new int[]{1200,11000,6000};//�����п�
		// writeExcelHeader(sheet1,title,cloumnWidth);//д��ͷ
        if(result1.size()>0) {
        	writeDataToSheet(sheet1, result1, true);
        }		

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

	private void openFile(String fullFileName) {
		// TODO Auto-generated method stub
		try {
			System.out.println("cmd /c " + fullFileName);
			Runtime.getRuntime().exec("cmd /c " + fullFileName);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void writeDataToSheet(XSSFSheet sheet, ArrayList result, boolean setGroup) {
		// TODO Auto-generated method stub
		int row = 3;

		int len = result.size();

		String[] val = (String[]) result.get(0);
		String begin = val[1];
		String begin2 = val[2];//���
		int beginrow = 3;// ��ʼ��
		int endrow = 0;// ��ֹ��
		int beginrow2 = 3;// ��ʼ��
		int endrow2 = 0;// ��ֹ��
		int n = 0;// �����
		int n2 = 0;
		// �ϲ���Ԫ��
		CellRangeAddress region1;

		for (int i = 0; i < len; i++) {
			viewPanel.addInfomation("", 80, 100);
			String[] values = (String[]) result.get(i);
			for (int j = 0; j < values.length; j++) {
				setCell(sheet, values[j], row, j);
			}
			// ����1����ʼ�� ����2����ֹ�� ����3����ʼ�� ����4����ֹ��
			// �ӵ�һ�����ݿ�ʼ�����±����������λ��ͬ��ϲ���Ԫ��
            if(i!=0) {
            	/** ********************************
            	 * ��λ�кϲ�
            	 */
            	if (!values[1].equals(begin)) {
    				endrow = beginrow + n;				
    				region1 = new CellRangeAddress(beginrow, endrow, (short) 1, (short) 1);								
    				sheet.addMergedRegion(region1);
    				begin = values[1].toString();
    				beginrow = endrow+1;
    				n = 0;
    			} else {
    				n++;
    			}
    			if(i==len-1) {
    				endrow = beginrow + n;
    				region1 = new CellRangeAddress(beginrow, endrow, (short) 1, (short) 1);
    				sheet.addMergedRegion(region1);
    			}
    			
    			/** ********************************
            	 * �����˵�������
            	 */
    			if (!values[2].equals(begin2)) {
    				endrow2 = beginrow2 + n2;				
    				region1 = new CellRangeAddress(beginrow2, endrow2, (short) 2, (short) 2);								
    				sheet.addMergedRegion(region1);
    				region1 = new CellRangeAddress(beginrow2, endrow2, (short) 3, (short) 3);								
    				sheet.addMergedRegion(region1);
    				region1 = new CellRangeAddress(beginrow2, endrow2, (short) 4, (short) 4);								
    				sheet.addMergedRegion(region1);
    				region1 = new CellRangeAddress(beginrow2, endrow2, (short) 5, (short) 5);								
    				sheet.addMergedRegion(region1);
    				begin2 = values[2].toString();
    				beginrow2 = endrow2+1;
    				n2 = 0;
    			} else {
    				n2++;
    			}
    			if(i==len-1) {
    				endrow2 = beginrow2 + n2;
    				region1 = new CellRangeAddress(beginrow2, endrow2, (short) 2, (short) 2);								
    				sheet.addMergedRegion(region1);
    				region1 = new CellRangeAddress(beginrow2, endrow2, (short) 3, (short) 3);								
    				sheet.addMergedRegion(region1);
    				region1 = new CellRangeAddress(beginrow2, endrow2, (short) 4, (short) 4);								
    				sheet.addMergedRegion(region1);
    				region1 = new CellRangeAddress(beginrow2, endrow2, (short) 5, (short) 5);								
    				sheet.addMergedRegion(region1);
    			}
            }
			

			row++;
		}

	}
	
	private void writeExcelHeader(XSSFSheet sheet, String[] title, int[] cloumnWidth) {
		// TODO Auto-generated method stub
		for (int i = 0; i < title.length; i++) {
			setCell(sheet, title[i], 0, i);
		}

		for (int i = 0; i < cloumnWidth.length; i++) {
			sheet.setColumnWidth(i, cloumnWidth[i]);
		}
	}

	// �Ե�Ԫ��ֵ��ǿ��ʹ���ı���ʽ
	protected void setCell(XSSFSheet sheet, String value, int rowIndex, int cellIndex) {
		XSSFRow row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		XSSFCell cell = row.getCell(cellIndex);
		if (cell == null)
			cell = row.createCell(cellIndex);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		if (cellIndex == 0) {
			cell.setCellType(Cell.CELL_TYPE_NUMERIC);
			if (value != null) {
				cell.setCellValue(Integer.parseInt(value));
			} else {
				cell.setCellType(Cell.CELL_TYPE_BLANK);
			}
		} else {
			cell.setCellValue(value);
		}
		cell.setCellStyle(cellStyle);
	}

}
