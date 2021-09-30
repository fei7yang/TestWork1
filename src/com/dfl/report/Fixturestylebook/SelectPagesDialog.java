package com.dfl.report.Fixturestylebook;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.wb.swt.SWTResourceManager;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.layout.FormLayout;
import org.eclipse.swt.layout.FormData;
import org.eclipse.swt.layout.FormAttachment;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.VerifyListener;
import org.eclipse.swt.events.VerifyEvent;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;

public class SelectPagesDialog extends Dialog {

	protected Object result;
	protected Shell shell;
	private Text text_1;
	//public String page1;
	public String page2;
	private Label label;

	/**
	 * Create the dialog.
	 * @param parent
	 * @param style
	 */
	public SelectPagesDialog(Shell parent, int style) {
		super(parent, style);
		setText("SWT Dialog");
	}

	/**
	 * Open the dialog.
	 * @return the result
	 */
	public Object open() {
		createContents();
		shell.open();
		shell.layout();
		Display display = getParent().getDisplay();
		centerToScreen(display);
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
		return result;
	}
	
	protected void centerToScreen(Display display) {
		int nLocationX = display.getClientArea().width / 2 - shell.getSize().x / 2;
		int nLocationY = display.getClientArea().height / 2 - shell.getSize().y / 2;
		shell.setLocation(nLocationX, nLocationY);
	}

	/**
	 * Create contents of the dialog.
	 */
	private void createContents() {
		shell = new Shell(getParent(), getStyle());
		shell.setImage(SWTResourceManager.getImage(SelectPagesDialog.class, "/com/dfl/report/dialog/defaultapplication_16.png"));
		shell.setSize(438, 174);
		shell.setText("\u9875\u6570\u9009\u62E9");
		shell.setLayout(new FillLayout(SWT.HORIZONTAL));
		
		Composite composite = new Composite(shell, SWT.NONE);
		composite.setLayout(null);
		
		Label lblNewLabel_1 = new Label(composite, SWT.RIGHT);
		lblNewLabel_1.setBounds(10, 10, 161, 24);
		lblNewLabel_1.setText("10_weld layout\uFF1A");
		
		text_1 = new Text(composite, SWT.BORDER);
		text_1.setText("1");
		text_1.addVerifyListener(new VerifyListener() {
			public void verifyText(VerifyEvent arg0) {
				arg0.doit = "0123456789".indexOf(arg0.text) >= 0;
			}
		});
		text_1.setBounds(194, 7, 119, 23);
		
		Button btnNewButton = new Button(composite, SWT.NONE);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				Okkey();
			}
		});
		btnNewButton.setBounds(224, 96, 80, 27);
		btnNewButton.setText("\u786E\u5B9A");
		
		Button button = new Button(composite, SWT.NONE);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shell.dispose();
			}
		});
		button.setText("\u53D6\u6D88");
		button.setBounds(310, 96, 80, 27);
		
		label = new Label(composite, SWT.NONE);
		label.setForeground(SWTResourceManager.getColor(SWT.COLOR_RED));
		label.setBounds(20, 40, 370, 23);

	}

	protected void Okkey() {
		// TODO Auto-generated method stub				
		//page1 = text.getText();
		page2 = text_1.getText();
		
//		if(page1 == null || page1.isEmpty()) {
//			label.setText("请填写断面定位形状仕様的sheet页数！");
//			return;
//		}
		if(page2 == null || page2.isEmpty()) {
			label.setText("请填写weld layout的sheet页数！");
			return;
		}	
//		if(Integer.parseInt(page1)<=0) {
//			label.setText("sheet页数必须大于0！");
//			return;
//		}
		if(Integer.parseInt(page2)<=0) {
			label.setText("sheet页数必须大于0！");
			return;
		}
		
		shell.dispose();
	}
	public static void main(String[] args) {
		SelectPagesDialog dialog1 = new SelectPagesDialog(new Shell(),SWT.SHELL_TRIM);
		dialog1.open();
	}
}
