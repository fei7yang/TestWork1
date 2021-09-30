package com.dfl.report.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import com.teamcenter.rac.kernel.TCComponentDataset;
import com.teamcenter.rac.kernel.TCComponentDatasetType;
import com.teamcenter.rac.kernel.TCSession;
import com.teamcenter.rac.ui.common.RACUIUtil;


public class FileUtil {
	public static String getReportFileName(String name) {
		String fullFileName = null;
		Calendar cal = Calendar.getInstance();
		//SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd_HHmm");

		String tempfilename = null;
		try {

			String fileName = name + ".xlsx";
			tempfilename = getReportDirName();

			if (System.getProperty("os.name").startsWith("Windows")) {
				fullFileName = tempfilename + "\\" + fileName;
			} else {
				fullFileName = tempfilename + "/" + fileName;
			}
		} catch (Exception e) {
//			writeFlag = false;
			e.printStackTrace();
		}
		return fullFileName;
	}
	
	
	public static String getReportDirName() {
		String tempfilename = null;
		try {
			tempfilename = System.getenv("tmp");
			if (tempfilename == null || "".compareTo(tempfilename) == 0) {
				tempfilename = System.getenv("temp");
			}

			if (tempfilename == null || "".compareTo(tempfilename) == 0) {
				tempfilename = System.getenv("TMP");
			}

			if (tempfilename == null || "".compareTo(tempfilename) == 0) {
				tempfilename = System.getenv("TEMP");
			}

			if (tempfilename == null || "".compareTo(tempfilename) == 0) {
				tempfilename = "C:\\Temp";
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return tempfilename;
	}

	/**
	 * 通过名称获取Word模板文件
	 * @param name
	 * @return 数据集对象
	 */
	public static InputStream getTemplateFile(String name) {

		File downloadFile=null;
		InputStream filein=null;
		try {
			TCSession session=RACUIUtil.getTCSession();
			TCComponentDatasetType datasettype=(TCComponentDatasetType )session.getTypeComponent("Dataset");
			TCComponentDataset dataset=	datasettype.find(name);
			if(dataset!=null)
			{
				System.out.println(dataset.getType());
				if(dataset.getType().equals("MSExcelX"))
				{
					String filepath = System.getProperty("java.io.tmpdir");
					File tf=new File(filepath+name);
					if(tf.exists())
						tf.delete();
					File files[]=dataset.getFiles("excel");
					downloadFile=files[0];
					System.err.println(downloadFile.getPath());
					filein=new FileInputStream(downloadFile);
				}
				if(dataset.getType().equals("MSWord"))
				{
					String filepath = System.getProperty("java.io.tmpdir");
					File tf=new File(filepath+name);
					if(tf.exists())
						tf.delete();
					File files[]=dataset.getFiles("word");
					downloadFile=files[0];
					System.err.println(downloadFile.getPath());
					filein=new FileInputStream(downloadFile);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

		if(filein==null)
		{
			filein = FileUtil.class.getResourceAsStream(name);
		}
		return filein;
	}
	/**
	 * 通过名称获取Word模板文件
	 * @param name
	 * @return 数据集对象
	 */
	public static TCComponentDataset getDatasetFile(String name) {
		
		try {
			TCSession session=RACUIUtil.getTCSession();
			TCComponentDatasetType datasettype=(TCComponentDatasetType )session.getTypeComponent("Dataset");
			TCComponentDataset dataset=	datasettype.find(name);
			if(dataset!=null)
			{
				System.out.println(dataset.getType());
				return dataset;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	
}
