package com.Hydro.service;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.Hydro.model.Borrower;
import com.Hydro.utils.CellStyle;
import com.Hydro.utils.ExcelUtils;

@SuppressWarnings("deprecation")
public class ReportServiceborr implements ReportServiceInterface {
	public Map<String, ArrayList<Borrower>> DEPTMAP = new HashMap<String, ArrayList<Borrower>>();// <部门,<借款集合>>
	public String path = System.getProperty("user.dir") + File.separator + "file";// 声明file文件夹在项目中的相对路径
	public HashMap<String, String> BORRID_DEPTID_RELATION = new HashMap<String, String>();

	public ArrayList<Borrower> borrList = new ArrayList<Borrower>();// 所有借款人
	public double[] total = { 0.0, 0.0, 0.0 };

	public CellStyle cellStyle;
	public ExcelUtils excelUtils;

	@Override
	public void readExcel(String fname) throws Exception {
		// TODO 自动生成的方法存根
		Object co = null;
		try {
			System.out.println(path + File.separator + fname);
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fname));
			HSSFSheet sheet = workbook.getSheetAt(0);// 取Excel第一个sheet表
			Iterator<Row> rows = sheet.rowIterator();// 行迭代器
			ExcelUtils excelUtils = new ExcelUtils();
			int row_i = 0;
			while (rows.hasNext()) {// 判断是否存在下一行
				HSSFRow row = (HSSFRow) rows.next();// 获得下一行。
				if (row_i >= 5) {// 从第6行开始读
					Iterator<Cell> cells = row.cellIterator();// 单元格迭代器
					int cell_i = 0;
					Borrower borrower = new Borrower();
					while (cells.hasNext()) {// 判断是否存在下一单元格
						HSSFCell cell = (HSSFCell) cells.next();// //获得下一单元格。
						try {
							co = excelUtils.getCellValue(cell);// 获取单元格的值
							if (co instanceof String) {// 当单元格的值是String类型并且含有"制表人"字段时，则认为到达最后一行，迭代退出。
								if (((String) co).indexOf("制表人") != -1) {
									break;
								}
							}
							if (cell_i == 1) {
								borrower.setBorrId((String) co);
							}
							if (cell_i == 2) {
								borrower.setDeptId((String) co);
							}
							if (cell_i == 5) {
								borrower.setBorrDate((String) co);
							}
							if (cell_i == 7) {
								borrower.setPurpose((String) co);
							}
							if (cell_i == 9) {
								if (co == null) {
									co = 0.0;
								}
								borrower.setOriginalCurrency((Double) co);
							}
							if (cell_i == 10) {
								if (co == null) {
									co = 0.0;
								}
								borrower.setOriginalCurrencyBalance((Double) co);
							}
							if (cell_i == 11) {
								if (co == null) {
									co = 0.0;
								}
								borrower.setAging((Double) co);
							}
						} catch (Exception e) {
							e.printStackTrace();
						}
						cell_i++;
					}
					borrower.setVerification(borrower.getOriginalCurrency() - borrower.getOriginalCurrencyBalance());// 已核销=原币-原币余额
					if (!"小计".equals(borrower.getPurpose())) {
						borrList.add(borrower);
					}
				}
				row_i++;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings("unused")
	@Override
	public void produceExcel(String fname, String savaPath) throws Exception {
		// TODO 自动生成的方法存根
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fname));// 创建一个Excel表对象
		cellStyle = new CellStyle(workbook);
		excelUtils = new ExcelUtils();
		HSSFSheet sheet = workbook.getSheetAt(0);// 获取第一个sheet表
		Collections.sort(borrList);// 集合中借款信息按照账龄升序排序，注释掉就是不排序

		int rid = 3, rid1 = 0, rid2 = 0, rid3 = 0;
		rid1 = excelByAgingSort(0, 30, rid, borrList, workbook, sheet);
		rid2 = excelByAgingSort(30, 60, rid1, borrList, workbook, sheet);
		rid3 = excelByAgingSort(60, Integer.MAX_VALUE, rid2, borrList, workbook, sheet);

		HSSFRow row = sheet.createRow(rid3++);// 新建一个空白行
		setSubtotalRow("总计", rid3++, total, workbook, sheet);// 总计行
		leaderSign(rid3, workbook, sheet);// 领导签字行

		// 自动调节每列宽度
		for (int i = 0; i < 10; i++) {
			sheet.autoSizeColumn(i);
		}

		excelUtils.exportExcel(workbook, savaPath + File.separator + "备用资金反馈表.xls");// 生成Excel
	}

	public void leaderSign(int rid, HSSFWorkbook workbook, HSSFSheet sheet) {
		rid += 2;
		sheet.addMergedRegion(new CellRangeAddress(rid, rid, 0, 9));
		HSSFRow row = sheet.createRow(rid);
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("主管签字：");
		cell.setCellStyle(cellStyle.cellFontBoldCenterNoBorderStyle);// 样式赋给单元格
	}

	public void setSubtotalRow(String countType, int rid, double[] subtotal, HSSFWorkbook workbook, HSSFSheet sheet) {
		HSSFRow row = sheet.createRow(rid);
		for (int k = 0; k < 10; k++) {
			try {
				switch (k) {
				case 3:
					excelUtils.setCellValue(subtotal[0], workbook, k, row, cellStyle.cellBorderStyle);// 原币
					break;
				case 4:
					excelUtils.setCellValue(subtotal[1], workbook, k, row, cellStyle.cellBorderStyle);// 原币
					break;
				case 5:
					excelUtils.setCellValue(subtotal[2], workbook, k, row, cellStyle.cellBorderStyle);// 原币
					break;
				default:
					excelUtils.setCellValue(null, workbook, k, row, cellStyle.cellBorderStyle);// 其它赋空
					break;
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		sheet.addMergedRegion(new CellRangeAddress(rid, rid, 0, 2));// 合并单元格
		HSSFCell cell = row.getCell(0);
		cell.setCellValue(countType);
		cell.setCellStyle(cellStyle.cellFontBoldCenterBorderStyle);// 样式赋给单元格
	}

	private int excelByAgingSort(int start, int end, int rid, ArrayList<Borrower> list, HSSFWorkbook workbook, HSSFSheet sheet) {
		int _rid = rid;
		// 生成一行显示“多少天账龄”提示,可以注释掉，下面代码可注释
		sheet.addMergedRegion(new CellRangeAddress(_rid, _rid, 0, 9));
		HSSFRow row = sheet.createRow(_rid);
		HSSFCell cell = row.createCell(0);
		if (start >= 60) {
			cell.setCellValue("大于" + start + "天账龄");
		} else {
			cell.setCellValue(start + "~" + end + "天账龄");
		}
		cell.setCellStyle(cellStyle.cellFontBoldCenterBorderStyle);

		_rid++;

		double[] subtotal = { 0.0, 0.0, 0.0 };// 小计统计
		// 上面代码可注释
		for (int i = 0; i < list.size(); i++) {// 借款信息集合
			try {
				Borrower borr = (Borrower) list.get(i);// 获取每一条借款信息
				if (borr.getAging() != 0 && (borr.getAging() >= start && borr.getAging() < end)) {
					System.out.println(borr);
					row = sheet.createRow(_rid);
					int j = 0;
					excelUtils.setCellValue(borr.getBorrId(), workbook, j++, row, cellStyle.cellBorderStyle);// 借款人
					excelUtils.setCellValue(borr.getBorrDate(), workbook, j++, row, cellStyle.cellBorderStyle);// 日期
					excelUtils.setCellValue(borr.getPurpose(), workbook, j++, row, cellStyle.cellBorderStyle);// 借款用途
					excelUtils.setCellValue(borr.getOriginalCurrency(), workbook, j++, row, cellStyle.cellBorderStyle);// 原币
					excelUtils.setCellValue(borr.getVerification(), workbook, j++, row, cellStyle.cellBorderStyle);// 已核销
					excelUtils.setCellValue(borr.getOriginalCurrencyBalance(), workbook, j++, row, cellStyle.cellBorderStyle);// 原币余额
					excelUtils.setCellValue(borr.getAging(), workbook, j++, row, cellStyle.cellBorderStyle);
					excelUtils.setCellValue(null, workbook, j++, row, cellStyle.cellBorderStyle);// 情况说明单元格赋空值
					excelUtils.setCellValue(null, workbook, j++, row, cellStyle.cellBorderStyle);
					excelUtils.setCellValue(null, workbook, j++, row, cellStyle.cellBorderStyle);

					subtotal[0] += borr.getOriginalCurrency();// 累计原币
					subtotal[1] += borr.getVerification();// 累计已核销
					subtotal[2] += borr.getOriginalCurrencyBalance();// 累计原币余额

					_rid++;
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		// 总计
		total[0] += subtotal[0];
		total[1] += subtotal[1];
		total[2] += subtotal[2];

		setSubtotalRow("小计", _rid++, subtotal, workbook, sheet);// 小计行
		return _rid++;
	}

}
