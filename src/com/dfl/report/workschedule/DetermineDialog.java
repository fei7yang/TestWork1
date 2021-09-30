package com.dfl.report.workschedule;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.wb.swt.SWTResourceManager;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;

public class DetermineDialog extends Dialog {

	protected Object result;
	protected Shell shell;
	private Combo combo;
	private String message = "";

	/**
	 * Create the dialog.
	 * @param parent
	 * @param style
	 */
	public DetermineDialog(Shell parent, int style) {
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
		shell.setImage(SWTResourceManager.getImage(DetermineDialog.class, "/com/dfl/report/imags/defaultapplication_16.png"));
		shell.setSize(424, 153);
		shell.setText("\u6E29\u99A8\u63D0\u793A");
		
		Label lblNewLabel = new Label(shell, SWT.NONE);
		lblNewLabel.setBounds(10, 10, 302, 22);
		lblNewLabel.setText("\u662F\u5426\u9700\u8981\u91CD\u65B0\u8BA1\u7B97\u7EFC\u5408\u53C2\u6570");
		
		combo = new Combo(shell, SWT.NONE);
		combo.setBounds(56, 38, 53, 25);
		combo.add("ÊÇ");
		combo.add("·ñ");
		combo.select(1);
		
		Button btnNewButton = new Button(shell, SWT.NONE);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shell.dispose();
			}
		});
		btnNewButton.setBounds(261, 76, 59, 27);
		btnNewButton.setText("\u53D6\u6D88");
		
		Button btnNewButton_1 = new Button(shell, SWT.NONE);
		btnNewButton_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				message = combo.getText();
				shell.dispose();
			}
		});
		btnNewButton_1.setBounds(326, 76, 59, 27);
		btnNewButton_1.setText("\u786E\u5B9A");
	}
	
	public String getMessage()
	{
		return message;
	}
	public static void main(String[] args) {
		DetermineDialog dialog1 = new DetermineDialog(new Shell(), SWT.SHELL_TRIM);
		dialog1.open();
	}
}
