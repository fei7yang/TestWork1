package com.dfl.report.util;

import org.apache.log4j.Logger;

public class LogTest {
	private static Logger logger = Logger.getLogger(LogTest.class.getName()); // 日志打印类
	public static void main(String[] args) {
		// TODO Auto-generated method stub
        System.out.println("dhfs ");
        //logger.warning("warn....");
        logger.info("info....");
        logger.debug("debug...");
        logger.error("error...");
        logger.warn("warn...");
	}

}
