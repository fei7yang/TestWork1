package com.dfl.report.util;

import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;  
import java.util.Set;  
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.Region;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
  
/** 
 *  
 * @author trx 
 */  
public final class CopySheetUtil {  
  
    public CopySheetUtil() {  
    }  
  
    /**
	 * ����һ����Ԫ����ʽ��Ŀ�ĵ�Ԫ����ʽ
	 * @param fromStyle
	 * @param toStyle
	 */
	public static void copyCellStyle(XSSFCellStyle fromStyle,
			XSSFCellStyle toStyle) {
		toStyle.setAlignment(fromStyle.getAlignment());
		//�߿�ͱ߿���ɫ
		toStyle.setBorderBottom(fromStyle.getBorderBottom());
		toStyle.setBorderLeft(fromStyle.getBorderLeft());
		toStyle.setBorderRight(fromStyle.getBorderRight());
		toStyle.setBorderTop(fromStyle.getBorderTop());
		toStyle.setTopBorderColor(fromStyle.getTopBorderColor());
		toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());
		toStyle.setRightBorderColor(fromStyle.getRightBorderColor());
		toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());
		
		//������ǰ��
		toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());
		toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());
		
		toStyle.setDataFormat(fromStyle.getDataFormat());
		toStyle.setFillPattern(fromStyle.getFillPattern());
//		toStyle.setFont(fromStyle.getFont(null));
		toStyle.setHidden(fromStyle.getHidden());
		toStyle.setIndention(fromStyle.getIndention());//��������
		toStyle.setLocked(fromStyle.getLocked());
		toStyle.setRotation(fromStyle.getRotation());//��ת
		toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());
		toStyle.setWrapText(fromStyle.getWrapText());
		
	}
	/**
	 * Sheet����
	 * @param fromSheet
	 * @param toSheet
	 * @param copyValueFlag
	 */
	public static void copySheet(XSSFWorkbook wb,XSSFSheet fromSheet, XSSFSheet toSheet,
			boolean copyValueFlag) {
		//�ϲ�������
		mergerRegion(fromSheet, toSheet);
		for (Iterator rowIt = fromSheet.rowIterator(); rowIt.hasNext();) {
			XSSFRow tmpRow = (XSSFRow) rowIt.next();
			XSSFRow newRow = toSheet.createRow(tmpRow.getRowNum());
			//�и���
			copyRow(wb,tmpRow,newRow,copyValueFlag);
		}
	}
	/**
	 * �и��ƹ���
	 * @param fromRow
	 * @param toRow
	 */
	public static void copyRow(XSSFWorkbook wb,XSSFRow fromRow,XSSFRow toRow,boolean copyValueFlag){
		for (Iterator cellIt = fromRow.cellIterator(); cellIt.hasNext();) {
			XSSFCell tmpCell = (XSSFCell) cellIt.next();
			XSSFCell newCell = toRow.createCell(tmpCell.getColumnIndex());
			copyCell(wb,tmpCell, newCell, copyValueFlag);
		}
	}
	/**
	* ����ԭ��sheet�ĺϲ���Ԫ���´�����sheet
	* 
	* @param sheetCreat �´���sheet
	* @param sheet      ԭ�е�sheet
	*/
	public static void mergerRegion(XSSFSheet fromSheet, XSSFSheet toSheet) {
	   int sheetMergerCount = fromSheet.getNumMergedRegions();
	   for (int i = 0; i < sheetMergerCount; i++) {
	    CellRangeAddress mergedRegionAt = fromSheet.getMergedRegion(i);
	    toSheet.addMergedRegion(mergedRegionAt);
	   }
	}
	/**
	 * ���Ƶ�Ԫ��
	 * 
	 * @param srcCell
	 * @param distCell
	 * @param copyValueFlag
	 *            true����ͬcell������һ����
	 */
	public static void copyCell(XSSFWorkbook wb,XSSFCell srcCell, XSSFCell distCell,
			boolean copyValueFlag) {
		XSSFCellStyle newstyle=wb.createCellStyle();
		copyCellStyle(srcCell.getCellStyle(), newstyle);
		//distCell.setEncoding(srcCell.get);
		//��ʽ
		distCell.setCellStyle(newstyle);
		//����
		if (srcCell.getCellComment() != null) {
			distCell.setCellComment(srcCell.getCellComment());
		}
		// ��ͬ�������ʹ���
		int srcCellType = srcCell.getCellType();
		distCell.setCellType(srcCellType);
		if (copyValueFlag) {
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
	}
}  
