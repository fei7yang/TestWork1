package com.dfl.report.dialog;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.wb.swt.SWTResourceManager;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.FillLayout;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;

public class SelectionSiteNameDialog extends Dialog {

	protected Object result;
	protected Shell shell;
	private Combo combo;
	public String name;

	/**
	 * Create the dialog.
	 * @param parent
	 * @param style
	 */
	public SelectionSiteNameDialog(Shell parent, int style) {
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
	
	protected void centerToScreen(Display display)
	{
		int nLocationX=display.getClientArea().width/2-shell.getSize().x/2;
		int nLocationY=display.getClientArea().height/2-shell.getSize().y/2;
		shell.setLocation(nLocationX,nLocationY);
	}

	/**
	 * Create contents of the dialog.
	 */
	private void createContents() {
		shell = new Shell(getParent(), getStyle());
		shell.setImage(SWTResourceManager.getImage(SelectionSiteNameDialog.class, "/com/dfl/report/dialog/defaultapplication_16.png"));
		shell.setSize(450, 171);
		shell.setText("\u547D\u540D\u9009\u62E9");
		shell.setLayout(new FillLayout(SWT.HORIZONTAL));
		
		Composite composite = new Composite(shell, SWT.NONE);
		composite.setLayout(null);
		
		Label label = new Label(composite, SWT.NONE);
		label.setBounds(5, 9, 60, 17);
		label.setText("\u90E8\u4F4D\u540D\u79F0\uFF1A");
		
		combo = new Combo(composite, SWT.DROP_DOWN | SWT.READ_ONLY);
		combo.setBounds(70, 5, 233, 25);
		combo.add("上屋");
		combo.add("下屋");
		combo.add("COVER");
		combo.add("整个车身");
		combo.select(0);
		
		Button button = new Button(composite, SWT.NONE);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				Okevent();
			}
		});
		button.setBounds(260, 79, 80, 27);
		button.setText("\u786E\u5B9A");
		
		Button btnNewButton = new Button(composite, SWT.NONE);
		btnNewButton.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {				
				shell.dispose();
			}
		});
		btnNewButton.setBounds(346, 79, 80, 27);
		btnNewButton.setText("\u53D6\u6D88");

	}
	protected void Okevent() {
		// TODO Auto-generated method stub
		if(combo.getSelectionIndex()== -1) {
			name = null;
		}else {
			name = combo.getItem(combo.getSelectionIndex());
			System.out.println("name:"+name);
			shell.dispose();
		}
		
	}
	public static void main(String[] args) {
		SelectionSiteNameDialog dialog = new SelectionSiteNameDialog(new Shell(),SWT.SHELL_TRIM);
		dialog.open();
	}
}
