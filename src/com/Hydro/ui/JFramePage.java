package com.Hydro.ui;

import java.awt.Color;
import java.awt.FileDialog;
import java.awt.Frame;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

import com.Hydro.service.ReportServiceBorr;
import com.Hydro.service.ReportServiceDept;

/**
 * Excel操作界面
 */
public class JFramePage extends JFrame {
	private static final long serialVersionUID = 1L;
	JButton btnSourceFile = new JButton();
	JButton btnTemplateFile = new JButton();
	JButton exportFiledept = new JButton();
	JButton exportFileborr = new JButton();
	public JLabel jLable = new JLabel();
	public JTextField jTextField = new JTextField();
	public JTextField jTextField1 = new JTextField();
	public JTextField jTextField2 = new JTextField();
	String jfPath = null;
	String templetPath = null;

	@SuppressWarnings("deprecation")
	public static void run() {
		JFramePage jf = new JFramePage();
		jf.setTitle("导出部门报表或总汇总表");
		jf.show();

	}

	public static void main(String[] args) {
		JFramePage.run();
	}

	public JFramePage() {
		try {
			framePageInit();
			this.setDefaultCloseOperation(3);
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	// 画页面
	private void framePageInit() throws Exception {
		this.getContentPane().setLayout(null);
		this.setSize(500, 300);// UI长宽
		this.setLocation(300, 300);// UI在显示器中的位置坐标

		btnSourceFile.addActionListener(new ActionListener() {

			@SuppressWarnings("deprecation")
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				FileDialog fd = new FileDialog(new Frame(), "上传文件源文件", FileDialog.LOAD);
				fd.show();
				jfPath = fd.getDirectory() + fd.getFile();
				if (fd.getDirectory() != null && fd.getFile() != null) {
					jTextField1.setText(jfPath);
				}
			}

		});// 源文件按钮事件
		btnTemplateFile.addActionListener(new ActionListener() {

			@SuppressWarnings("deprecation")
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				FileDialog fd = new FileDialog(new Frame(), "上传模板文件", FileDialog.LOAD);
				fd.show();
				templetPath = fd.getDirectory() + fd.getFile();
				if (fd.getDirectory() != null && fd.getFile() != null) {
					jTextField2.setText(templetPath);
				}
			}
		});// 模板按钮事件

		exportFiledept.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				exportFiledept();
			}

		});// 按部门导出按钮事件

		exportFileborr.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				exportFileborr();
			}

		});// 导出总汇总表

		// Rectangle（x,y,xl,y1） 按钮坐标 ,(x:起始x轴坐标，y:起始y轴坐标，xl:x轴向右偏移量，y1:y轴下偏移量)
		btnSourceFile.setBounds(new Rectangle(50, 80, 80, 30));
		btnSourceFile.setText("源文件");

		btnTemplateFile.setBounds(new Rectangle(50, 120, 80, 30));
		btnTemplateFile.setText("模板");

		jLable.setBounds(50, 10, 350, 30);
		jLable.setSize(400, 20);
		jLable.setText("说明=> 源文件：核销余额 Excle表, 模板：备用金反馈单Excel表");
		jLable.setForeground(Color.red);
		jLable.setVisible(true);// 设为可见

		jTextField1.setBounds(150, 80, 300, 30);
		jTextField1.setEditable(false);
		jTextField1.setVisible(true);

		jTextField2.setBounds(150, 120, 300, 30);
		jTextField2.setEditable(false);
		jTextField2.setVisible(true);

		exportFiledept.setBounds(new Rectangle(150, 200, 90, 30));
		exportFiledept.setText("部门导出");

		exportFileborr.setBounds(new Rectangle(260, 200, 90, 30));
		exportFileborr.setText("汇总导出");

		this.getContentPane().add(jLable);
		this.getContentPane().add(jTextField1);
		this.getContentPane().add(jTextField2);
		this.getContentPane().add(btnSourceFile);
		this.getContentPane().add(btnTemplateFile);
		this.getContentPane().add(exportFiledept);
		this.getContentPane().add(exportFileborr);
	}

	private void exportFiledept() {
		JFileChooser fileChooser = new JFileChooser(".");
		fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		fileChooser.setDialogTitle("打开文件夹");
		int ret = fileChooser.showOpenDialog(null);
		if (ret == JFileChooser.APPROVE_OPTION) {
			final String filePath = fileChooser.getSelectedFile().getAbsolutePath();
			// 文件夹路径
			System.out.println("报表生成路径：" + fileChooser.getSelectedFile().getAbsolutePath());
			Thread t = new Thread() {
				public void run() {
					boolean flag = true;
					try {

						ReportServiceDept app = new ReportServiceDept();
						app.readExcel(jfPath);
						app.produceExcel(templetPath, filePath);// 生成Excel，必须为.xls格式的Excel
						System.out.println("");
						System.out.println("完成");
					} catch (Exception e2) {
						flag = false;
						e2.printStackTrace();
					}
					new Dialog(flag, filePath);
				}
			};
			t.start();
		}

	}

	private void exportFileborr() {
		JFileChooser fileChooser = new JFileChooser(".");
		fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		fileChooser.setDialogTitle("打开文件夹");
		int ret = fileChooser.showOpenDialog(null);
		if (ret == JFileChooser.APPROVE_OPTION) {
			final String filePath = fileChooser.getSelectedFile().getAbsolutePath();
			// 文件夹路径
			System.out.println("报表生成路径：" + fileChooser.getSelectedFile().getAbsolutePath());
			Thread t = new Thread() {
				public void run() {
					boolean flag = true;
					try {

						ReportServiceBorr app = new ReportServiceBorr();
						app.readExcel(jfPath);
						app.produceExcel(templetPath, filePath);// 生成Excel，必须为.xls格式的Excel
						System.out.println("");
						System.out.println("完成");
					} catch (Exception e2) {
						flag = false;
						e2.printStackTrace();
					}
					new Dialog(flag, filePath);
				}
			};
			t.start();
		}

	}
}

/**
 * 导出完成后的提示框
 */
class Dialog extends JOptionPane {
	private static final long serialVersionUID = 1L;

	@SuppressWarnings("static-access")
	public Dialog(boolean flag, String info) {
		if (flag == true) {
			this.showMessageDialog(null, "操作成功： 保存路径：" + info, "操作信息", JOptionPane.PLAIN_MESSAGE);
		} else {
			this.showMessageDialog(null, "操作失败：导入的表错误，请重新检查。", "操作信息", JOptionPane.ERROR_MESSAGE);
		}
		this.setVisible(true);
	}
}
