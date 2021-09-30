package com.dfl.report.mfcadd;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Combo;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;

public class VersionSelectionDialog extends Dialog {

	protected Object result;
	protected Shell shell;
	private Combo cbVersion;
	public String version = "";
	/**
	 * Create the dialog.
	 * @param parent
	 * @param style
	 */
	public VersionSelectionDialog(Shell parent, int style) {
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
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
		return result;
	}

	/**
	 * Create contents of the dialog.
	 */
	private void createContents() {
		shell = new Shell(getParent(), SWT.DIALOG_TRIM | SWT.MIN | SWT.MAX);
		shell.setSize(450, 179);
		shell.setText("\u9009\u62E9\u7248\u6B21");
		MFCUtility.setSWTCenter(shell);
		Label label = new Label(shell, SWT.NONE);
		label.setBounds(10, 22, 50, 20);
		label.setText("\u7248\u6B21\uFF1A");
		
		cbVersion = new Combo(shell, SWT.BORDER);
		cbVersion.setBounds(70, 20, 365, 26);
//		cbVersion.add("A");
//		cbVersion.add("B");
//		cbVersion.add("C");
		
		Label label_1 = new Label(shell, SWT.SEPARATOR | SWT.HORIZONTAL);
		label_1.setBounds(10, 75, 425, 2);
		
		Button button = new Button(shell, SWT.NONE);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				version = cbVersion.getText();
				System.out.println(cbVersion.getText());
				if(StringUtil.isEmpty(version)) {
					MFCUtility.errorMassges("±ÿ–ÎÃÓ–¥∞Ê¥Œ£°");
					return;
				}
				shell.dispose();
			}
		});
		button.setBounds(52, 93, 98, 30);
		button.setText("\u786E\u5B9A");
		
		Button button_1 = new Button(shell, SWT.NONE);
		button_1.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shell.dispose();
			}
		});
		button_1.setText("\u53D6\u6D88");
		button_1.setBounds(245, 93, 98, 30);

	}
	
	public static void main(String[] args) {
		VersionSelectionDialog dialog = new VersionSelectionDialog(new Shell(), SWT.SHELL_TRIM);
		dialog.open();
		System.out.println("1111");
	}
}
