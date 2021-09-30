package com.dfl.report.handlers;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.dfl.report.util.FileUtil;
import com.dfl.report.util.Util;
public class POItest {

	public static void main(String[] args) throws Exception {
		// ����һ��������
		File oldfile = new File("C://Users//Administrator//Desktop//test.xlsx");
		FileInputStream filein=new FileInputStream(oldfile);
				Workbook workbook = new XSSFWorkbook(filein);
				// ����һ��sheet
				Sheet sheet = workbook.createSheet();
				// ��ͼ�Ķ�������������HSSFPatriarch, һ��sheetֻ�ܻ�ȡһ��
				XSSFDrawing hssfPatriarch = (XSSFDrawing) sheet.createDrawingPatriarch();
						
//				XSSFClientAnchor anchor1 = new XSSFClientAnchor(0, 0, 0, 0, (short)1, 2, (short)3, 10);
//				// û��ͼ��
//				HSSFSimpleShape notPrimitive = hssfPatriarch.createSimpleShape(anchor1);
//				notPrimitive.setShapeType(HSSFShapeTypes.NotPrimitive);
//				
//				HSSFClientAnchor anchor2 = new HSSFClientAnchor(0, 0, 0, 0, (short)4, 2, (short)6, 10);
//				// ����һ��û�����ɫ�ľ���
//				HSSFSimpleShape rect = hssfPatriarch.createSimpleShape(anchor2);
//				rect.setShapeType(HSSFShapeTypes.Rectangle);
//				rect.setNoFill(true);
//				
//				HSSFClientAnchor anchor3 = new HSSFClientAnchor(0, 0, 0, 0, (short)7, 2, (short)9, 10);
//				// ����һ��Բ�Ǿ���
//				HSSFSimpleShape roundRectangle = hssfPatriarch.createSimpleShape(anchor3);
//				roundRectangle.setShapeType(HSSFShapeTypes.RoundRectangle);
				
				// ����һ����Բ
				XSSFClientAnchor anchor1 = new XSSFClientAnchor(0, 0, -300000, 0, (short)5, 25, (short)6, 26);
				anchor1.setAnchorType(2);			
				XSSFSimpleShape ellipse = hssfPatriarch.createSimpleShape(anchor1);
				XSSFSimpleShape a = (XSSFSimpleShape) hssfPatriarch.getShapes().get(0);
			
				XSSFClientAnchor anchor = (XSSFClientAnchor) a.getAnchor();
				
				ellipse.setShapeType(ShapeTypes.ELLIPSE);
				ellipse.setNoFill(false);
				XSSFClientAnchor anchor2 = new XSSFClientAnchor(0, 0, -100000, 0, (short)4, 25, (short)5, 26);
				anchor1.setAnchorType(2);			
				XSSFSimpleShape ellipse2 = hssfPatriarch.createSimpleShape(anchor2);
				ellipse2.setShapeType(ShapeTypes.ELLIPSE);
				ellipse2.setNoFill(false);
				
				
//				// ����һ������
//				XSSFClientAnchor anchor2 = new XSSFClientAnchor(300000, 0, 0, 0, (short)1, 3, (short)2, 4);
//				anchor2.setAnchorType(2);			
//				XSSFSimpleShape ellipse2 = hssfPatriarch.createSimpleShape(anchor2);
//				ellipse2.setShapeType(ShapeTypes.LINE);
//				//ellipse2.setNoFill(false);
//				//��������
				XSSFClientAnchor anchor3 = new XSSFClientAnchor(0, 0, 0, 0, (short)1, 2, (short)2, 3);
				anchor1.setAnchorType(2);
				XSSFSimpleShape rect = hssfPatriarch.createSimpleShape(anchor3);
				
				XSSFRichTextString str = new XSSFRichTextString();
				Font font = workbook.createFont();
				font.setColor((short)2);
				str.setString("����");
				str.applyFont(font);
				rect.setShapeType(ShapeTypes.RECT);
				rect.setNoFill(false);
				rect.setText(str);
//				HSSFClientAnchor anchor5 = new HSSFClientAnchor(0, 0, 0, 0, (short)13, 2, (short)15, 10);
//				// ����һ������
//				HSSFSimpleShape diamond = hssfPatriarch.createSimpleShape(anchor5);
//				diamond.setShapeType(HSSFShapeTypes.Diamond);
//				
//				HSSFClientAnchor anchor6 = new HSSFClientAnchor(0, 0, 0, 0, (short)16, 2, (short)18, 10);
//				// ����һ������������
//				HSSFSimpleShape isocelesTriangle = hssfPatriarch.createSimpleShape(anchor6);
//				isocelesTriangle.setShapeType(HSSFShapeTypes.IsocelesTriangle);
//				
//				HSSFClientAnchor anchor21 = new HSSFClientAnchor(0, 0, 0, 0, (short)1, 16, (short)3, 24);
//				// ����һ��ֱ��������
//				HSSFSimpleShape rightTriangle = hssfPatriarch.createSimpleShape(anchor21);
//				rightTriangle.setShapeType(HSSFShapeTypes.RightTriangle);
				
//				HSSFClientAnchor anchor22 = new HSSFClientAnchor(0, 0, 0, 0, (short)4, 16, (short)6, 24);
//				// ����һ��ƽ���ı���
//				HSSFSimpleShape parallelogram = hssfPatriarch.createSimpleShape(anchor22);
//				parallelogram.setShapeType(HSSFShapeTypes.Parallelogram);
//				
////				HSSFClientAnchor anchor23 = new HSSFClientAnchor(0, 0, 0, 0, (short)7, 16, (short)9, 24);
////				// ����һ������ - ��֧��
////				HSSFSimpleShape trapezoid = hssfPatriarch.createSimpleShape(anchor23);
////				trapezoid.setShapeType(HSSFShapeTypes.Trapezoid);
//				
//				HSSFClientAnchor anchor24 = new HSSFClientAnchor(0, 0, 0, 0, (short)10, 16, (short)12, 24);
//				// ����һ��������
//				HSSFSimpleShape hexagon = hssfPatriarch.createSimpleShape(anchor24);
//				hexagon.setShapeType(HSSFShapeTypes.Hexagon);
//				
//				HSSFClientAnchor anchor25 = new HSSFClientAnchor(0, 0, 0, 0, (short)13, 16, (short)15, 24);
//				// ����һ���˱���
//				HSSFSimpleShape octagon = hssfPatriarch.createSimpleShape(anchor25);
//				octagon.setShapeType(HSSFShapeTypes.Octagon);
//				
//				HSSFClientAnchor anchor26 = new HSSFClientAnchor(0, 0, 0, 0, (short)16, 16, (short)18, 24);
//				// ����һ��ʮ����
//				HSSFSimpleShape plus = hssfPatriarch.createSimpleShape(anchor26);
//				plus.setShapeType(HSSFShapeTypes.Plus);
				//String fullFileName = FileUtil.getReportFileName(reportname);
//				File file1 = new File("C://Users//Administrator//Desktop//test.xlsx");
//				if (file1.exists()) {
//					file1.delete();
//					file1 = new File("C://Users//Administrator//Desktop//test.xlsx");
//				}
//				FileOutputStream file = new FileOutputStream(file1);
//				workbook.write(file);
//				file.close();
//				try {
//					Runtime.getRuntime().exec("cmd /c C:\\Users\\Administrator\\Desktop\\test.xlsx");
//				} catch (IOException e) {
//					// TODO Auto-generated catch block
//					e.printStackTrace();
//				}
				String str123 = "10.1";
				System.out.println(Util.isNumber(str123));
	}

}

 
