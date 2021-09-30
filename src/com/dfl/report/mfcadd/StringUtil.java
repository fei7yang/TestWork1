package com.dfl.report.mfcadd;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.util.Enumeration;
import java.util.Random;
import java.util.Vector;

// Referenced classes of package ufc:
//            DateUtil

public class StringUtil {

	private static int		i			= 0;
	private static String	today;
	private static boolean	isEncoding	= true;
	public static boolean isStringInStrings(String type, String[] types) {
		boolean flag = false;
		int i = 0;
		int len = types.length;
		for(i = 0; i < len ; i ++) {
			if(types[i].equals(type)) {
				flag = true;
				break;
			}
		}
		return flag;
	}
	public StringUtil() {}

	public static String inputStream2String(InputStream is) throws IOException {
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		for (int i = -1; (i = is.read()) != -1;) {
			baos.write(i);
		}

		return baos.toString();
	}

	public static String convertStreamToString(InputStream is) throws IOException {
		StringBuilder sb = null;
		try {
			if (is != null) {
				System.out.println(is.toString());
				sb = new StringBuilder();
				BufferedReader reader = new BufferedReader(new InputStreamReader(is, "UTF-8"));
				String line;
				while ((line = reader.readLine()) != null) {
					sb.append(line).append("\n");
				}
			}
		}
		catch (Exception ex) {
			is.close();
			ex.printStackTrace();
		}
		is.close();
		if (sb != null) {
			return sb.toString();
		} else {
			return "";
		}
	}

	public static String nullTo(String str) {
		if (str == null) {
			return "";
		} else {
			return str;
		}
	}

	

	public static String RandowStr(int srcLength) {
		String s = "";
		String arrykey[] = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };
		Random x = new Random();
		for (int i = 1; i <= srcLength; i++) {
			String temp = null;
			int tempnum = x.nextInt(35);
			temp = arrykey[tempnum];
			s = (new StringBuilder(String.valueOf(s))).append(temp).toString();
		}

		return s;
	}

	public static String convert(String src, String getCode, String outCode) {
		try {
			return new String(src.getBytes(getCode), outCode);
		}
		catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		return null;
	}

	public static String webEncode(String src) {
		if (isEncoding) {
			return src != null ? convert(src, "iso8859-1", "UTF-8") : "";
		} else {
			return src != null ? src : "";
		}
	}

	public static String webDecode(String src) {
		if (isEncoding) {
			return src != null ? convert(src, "iso8859-1", "UTF-8") : "";
		} else {
			return src != null ? src : "";
		}
	}

	public static String[] split(String str, String separator) {
		if (str == null || str.length() == 0 || separator == null || separator.equals("")) { return null; }
		Vector v = new Vector();
		String temp = new String(str);
		int len = separator.length();
		for (int pos = temp.indexOf(separator); pos != -1; pos = temp.indexOf(separator)) {
			if (pos > 0) {
				v.add(temp.substring(0, pos));
			}
			temp = temp.substring(pos + len, temp.length());
		}

		if (temp.length() > 0) {
			v.add(temp);
		}
		String result[] = new String[v.size()];
		int i = 0;
		for (Enumeration e = v.elements(); e.hasMoreElements();) {
			result[i] = (String) e.nextElement();
			i++;
		}

		return result;
	}

	public static String txt2html(String src) {
		src = replace(replace(src, "<", "&lt;"), ">", "&gt;");
		src = replace(replace(src, " ", " "), "\r\n", "<br>");
		src = replace(replace(src, "'", "&#39;"), "\"", "&quot;");
		return src;
	}

	public static String replaceStr(String strSrc, String strFind, String strReplace) {
		if (strSrc == null) { return null; }
		if (strFind == null || strReplace == null) { return strSrc; }
		String tmp[] = split(strSrc, strFind);
		if (tmp == null || tmp.length <= 1) { return strSrc; }
		String ret = "";
		ret = tmp[0];
		for (int i = 1; i < tmp.length; i++) {
			ret = (new StringBuilder(String.valueOf(ret))).append(strReplace).append(tmp[i]).toString();
		}

		return ret;
	}

	public static String replace(String strSrc, String strFind, String strReplace) {
		if (strSrc == null) { return null; }
		if (strFind == null || strReplace == null) { return strSrc; }
		StringBuffer dst = new StringBuffer();
		int lngLength = strFind.length();
		int lngBeginPos;
		int lngCurrentPos;
		for (lngBeginPos = 0; (lngCurrentPos = strSrc.indexOf(strFind, lngBeginPos)) >= lngBeginPos; lngBeginPos = lngCurrentPos + lngLength) {
			dst.append(strSrc.substring(lngBeginPos, lngCurrentPos));
			dst.append(strReplace);
		}

		if (lngBeginPos < strSrc.length()) {
			dst.append(strSrc.substring(lngBeginPos));
		}
		return dst.toString();
	}
	/**
	 * 去掉小数点后面多余的零
	 * 
	 * @param strWithZero
	 * @return
	 */
	public static String getStringCutZero(String strWithZero) {
		if (strWithZero != null && strWithZero.indexOf(".") > -1) {
			while (true) {
				if (strWithZero.lastIndexOf("0") == (strWithZero.length() - 1)) {
					strWithZero = strWithZero.substring(0, strWithZero
							.lastIndexOf("0"));
				} else {
					break;
				}
			}
			if (strWithZero.lastIndexOf(".") == (strWithZero.length() - 1)) {
				strWithZero = strWithZero.substring(0, strWithZero
						.lastIndexOf("."));
			}
		}
		return strWithZero;
	}
	public static boolean isEmpty(Object str){
		if(str == null || str.toString().trim().length() == 0){
			return true;
		}
		return false;
	}
	public static boolean isStrLenOver40(String str) {
		int length = 0;
		if(str == null || str.length() == 0) {
			return false;
		}
		char[] chars = str.toCharArray();
		length = chars.length;
		System.out.println("length := " + length);
		return length > 40;
	}
	public static String leftStrcat(String oraStr, int length, String catChar) {
		if(oraStr == null) {
			oraStr = "";
		}
		if(oraStr.length() > length) {
			return oraStr;
		}
		StringBuffer sbNew = new StringBuffer();
		for(int i = 0; i < length - oraStr.length() ; i ++) {
			sbNew.append(catChar);
		}
		sbNew.append(oraStr);
		return sbNew.toString();
	}
	public static void main(String[] args) {
		String saleMan = "3282|董欢";
		if(!StringUtil.isEmpty(saleMan) ) {
			if( saleMan.contains("|")) {
				saleMan = saleMan.substring(saleMan.indexOf("|") + 1);
			}
		}
		System.out.println(saleMan);
	}
}
