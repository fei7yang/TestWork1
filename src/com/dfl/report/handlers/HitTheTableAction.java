package com.dfl.report.handlers;

import java.io.InputStream;

import com.teamcenter.rac.aif.AbstractAIFUIApplication;
import com.teamcenter.rac.kernel.TCComponentItem;

public class HitTheTableAction implements Runnable {

	private AbstractAIFUIApplication app;
	private InputStream inputStream;
	private boolean Isupdateflag;
	private TCComponentItem tcc = null;
	public HitTheTableAction(AbstractAIFUIApplication app, Object object, String string, TCComponentItem tcc, InputStream inputStream, boolean isupdateflag) {
		// TODO Auto-generated constructor stub
		super();
		this.app=app;
		this.inputStream = inputStream;
		this.Isupdateflag = isupdateflag;
		this.tcc = tcc;
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub
		System.out.println("更新标识：" + Isupdateflag);
		HitTheTableExportOp dlg = new HitTheTableExportOp(app,inputStream,Isupdateflag,tcc);
	}

}
