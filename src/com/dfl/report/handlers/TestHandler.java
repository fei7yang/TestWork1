package com.dfl.report.handlers;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.eclipse.core.commands.AbstractHandler;
import org.eclipse.core.commands.ExecutionEvent;
import org.eclipse.core.commands.ExecutionException;

import com.dfl.report.util.Util;
import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.aif.kernel.InterfaceAIFComponent;
import com.teamcenter.rac.aifrcp.AIFUtility;
import com.teamcenter.rac.kernel.TCComponent;
import com.teamcenter.rac.kernel.TCComponentItemRevision;
import com.teamcenter.rac.kernel.TCException;
import com.teamcenter.rac.kernel.TCSession;

public class TestHandler extends AbstractHandler {

	private AbstractAIFUIApplication application;
	private TCSession session;

	@Override
	public Object execute(ExecutionEvent arg0) throws ExecutionException {
		// TODO Auto-generated method stub
		application = AIFUtility.getCurrentApplication();
		session = (TCSession) application.getSession();

		//InterfaceAIFComponent[] target = application.getTargetComponents();
		// 获取选择的对象
		InterfaceAIFComponent aifComponent = application.getTargetComponent();
		TCComponentItemRevision rev = (TCComponentItemRevision) aifComponent;
		System.out.println(rev);
		try {
			TCComponent[] tccdata = rev.getRelatedComponents("IMAN_3D_snap_shot");
			if (tccdata != null && tccdata.length > 0) {
				File pngfile = Util.downLoadPicture(tccdata[0]);
				System.out.println(pngfile);
				File oldfile = new File("C://重要文件//Desktop//test.xlsx");
				FileInputStream filein = new FileInputStream(oldfile);
				XSSFWorkbook workbook = new XSSFWorkbook(filein);
				XSSFSheet sheet = workbook.createSheet();

				writepicturetosheet(workbook, sheet, pngfile, 1, 1, false);

				File file1 = new File("C://重要文件//Desktop//test.xlsx");
				if (file1.exists()) {
					file1.delete();
					file1 = new File("C://重要文件//Desktop//test.xlsx");
				}
				FileOutputStream file = new FileOutputStream(file1);
				workbook.write(file);
				file.close();
				try {
					Runtime.getRuntime().exec("cmd /c C:\\重要文件\\Desktop\\test.xlsx");
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}

			}
		} catch (TCException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		return null;
	}

	// 根据单个文件写图片到excel
	private void writepicturetosheet(XSSFWorkbook book, XSSFSheet sheet, Object obj, int rowindex, int colindex,
			boolean flag) {
		// 先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray
		ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
		BufferedImage bufferImg;
		int rowNum = 9;
		if (flag) {
			rowNum = 8;
		}
		try {
			File file = (File) obj;
			bufferImg = ImageIO.read(file);
			ImageIO.write(bufferImg, "png", byteArrayOut);
			XSSFDrawing patriarch = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = new XSSFClientAnchor(50000, 50000, -50000, -50000, (short) colindex, rowindex,
					(short) (colindex + 6), rowindex + rowNum);
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
