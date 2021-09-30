package com.dfl.report.util;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Insets;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.SwingUtilities;
import javax.swing.Timer;
import javax.swing.border.EmptyBorder;

import org.eclipse.swt.widgets.Display;

import com.dfl.report.dialog.SelectionStageDialog;
import com.teamcenter.rac.util.ButtonLayout;
import com.teamcenter.rac.util.PropertyLayout;
import com.teamcenter.rac.util.Separator;

public class ReportViwePanel extends JFrame implements ActionListener {
	private static final long serialVersionUID = 1L;
	protected JTextArea taskOutput;
	public JProgressBar progressBar;
	public JButton cancelBtn;
	public JPanel attrLayOut;
	public JPanel btnPanel;
	public JScrollPane crolpn;
	public JPanel basicInfoPanel;
	public String title;
	public ReportViwePanel(String title) {
		this.title = title;

	
		openUI();
	}
	
	
	/**
	 * UI界面初始化
	 * @param title
	 */
	public void openUI() {
		setIconImage(Toolkit.getDefaultToolkit().getImage(SelectionStageDialog.class.getResource("/com/dfl/report/dialog/defaultapplication_16.png")));
		setTitle(title);
		setAlwaysOnTop(true);
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setBackground(Color.WHITE);
		basicInfoPanel = new JPanel(new PropertyLayout());
		basicInfoPanel.setBackground(Color.white);
		basicInfoPanel.setVisible(false);
		getContentPane().add(BorderLayout.NORTH, basicInfoPanel);
		progressBar = new JProgressBar(0, 100);
		progressBar.setBackground(Color.white);
		progressBar.setForeground(Color.blue);
		progressBar.setIndeterminate(false);
		progressBar.setStringPainted(true);
		progressBar.setValue(1);
		progressBar.setMaximumSize(new Dimension(200, 25));

		taskOutput = new JTextArea(5, 80);
		taskOutput.setMargin(new Insets(5, 5, 5, 5));
		taskOutput.setLineWrap(true);
		taskOutput.setEditable(false);
		progressBar.setValue(0);
		attrLayOut = new JPanel(new BorderLayout(5, 5));
		attrLayOut.add(BorderLayout.NORTH, progressBar);

		crolpn = new JScrollPane(taskOutput);
		crolpn.setBackground(Color.WHITE);
		attrLayOut.add(BorderLayout.CENTER, crolpn);
		attrLayOut.setBackground(Color.WHITE);

		getContentPane().add(BorderLayout.CENTER, attrLayOut);

		btnPanel = new JPanel(new ButtonLayout(ButtonLayout.HORIZONTAL,
				ButtonLayout.CENTER, 7));

		cancelBtn = new JButton("关闭");
		cancelBtn.setEnabled(false);
		cancelBtn.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				dispose();
			}
		});
		btnPanel.add(cancelBtn);
		JPanel btnOutter = new JPanel(new BorderLayout());
		btnOutter.setBorder(new EmptyBorder(5, 10, 5, 10));
		btnOutter.add(BorderLayout.NORTH, new Separator());
		btnOutter.add(BorderLayout.CENTER, btnPanel);
		getContentPane().add(BorderLayout.SOUTH, btnOutter);


		//居中显示
		int width = 450;
		int height = 280;
		Dimension dimension1 = Toolkit.getDefaultToolkit().getScreenSize();
		int i = (dimension1.width - width) / 2;
		int j = (dimension1.height - height) / 2;
		setBounds(i, j, width, height);
	}


	/**
	 * 进度条计算
	 * @param value
	 * @param step
	 * @param totalStep
	 */
	public void addInfomation(String value, int step, int totalStep) {
		
		final String value1 = value;
		final int step1 = step;
		final int totalStep1 = totalStep;

		SwingUtilities.invokeLater(new Runnable() {
		
		@Override
		public void run() {
			// TODO Auto-generated method stub
			update(value1,step1,totalStep1);
		}
	    });
//		Thread thread = new Thread(new Runnable() {
//			
//			@Override
//			public void run() {
//				// TODO Auto-generated method stub
//				update(value1,step1,totalStep1);
//			}
//		});
//		thread.run();
		
	}

	protected void update(String value, int step, int totalStep) {
		// TODO Auto-generated method stub
		if (value != null&&!value.equals("")) {
			// taskOutput.setText("");
			taskOutput.append(value);
			taskOutput.setCaretPosition(taskOutput.getText().length());   
		}
		/*int currentPoint = (int) (((step + 0.0) / totalStep) * 10);
		if (currentPoint >= 0) {
			progressBar.setValue(currentPoint);
		}
		if (currentPoint >= 10) {
			//			endOperation();
		}*/
		
		progressBar.setValue(step);
		if(step==100)
		{
			cancelBtn.setEnabled(true);
		}
	}


	/**
	 * 清除输出值
	 */
	public void clear() {
		taskOutput.setText("");
	}

	public void setFinalFilePath(String value) {
	}

	/**
	 * 设置进度
	 * @param value
	 */
	public void setOutput(String value) {
		if (value != null) {
			if (taskOutput.getText() != null
					&& taskOutput.getText().trim().length() > 0) {
				taskOutput.append("\n");
			}
			taskOutput.append(value);
		}
	}


	@Override
	public void actionPerformed(ActionEvent arg0) {
		// TODO Auto-generated method stub
		
	}


}
