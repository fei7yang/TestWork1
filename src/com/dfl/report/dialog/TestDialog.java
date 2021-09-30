package com.dfl.report.dialog;

import org.eclipse.swt.widgets.Dialog;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.ProgressBar;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.custom.StackLayout;
import org.eclipse.swt.layout.FillLayout;
import swing2swt.layout.BoxLayout;
import org.eclipse.swt.layout.RowLayout;
import swing2swt.layout.BorderLayout;

public class TestDialog extends Dialog {

	protected Object result;
	protected Shell shell;
	private Text text;
	private ProgressBar progressBar;

	/**
	 * Create the dialog.
	 * @param parent
	 * @param style
	 */
	public TestDialog(Shell parent, int style) {
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
		shell = new Shell(getParent(), getStyle());
		shell.setSize(453, 305);
		shell.setText(getText());
		shell.setLayout(null);
		
		progressBar = new ProgressBar(shell, SWT.NONE);
		progressBar.setBounds(5, 5, 442, 17);
		
		text = new Text(shell, SWT.BORDER);
		text.setBounds(5, 27, 442, 206);
		
		Button button = new Button(shell, SWT.NONE);
		button.setBounds(208, 239, 57, 27);
		button.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shell.dispose();
			}
		});
		button.setText("\u5173\u95ED");

	}
	/*
	 * ½ø¶ÈÌõ
	 */
	public void addinfomation(String message,int parno) {
		if(message!=null && !message.isEmpty()) {
			text.append("dsdfs");
		}
		progressBar.setSelection(10);
	}
	public static void main(String[] args) {
		TestDialog dialog = new TestDialog(new Shell(),SWT.SHELL_TRIM);
		dialog.open();
		dialog.addinfomation("sddfs", 10);
		
	}
}
