package com.Hydro.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.Hydro.model.Borrower;

@SuppressWarnings("deprecation")
public class ReportServiceborr implements ReportServiceInterface{
	public Map<String, ArrayList<Borrower>> DEPTMAP = new HashMap<String, ArrayList<Borrower>>();// <部门,<借款集合>>
	public String path = System.getProperty("user.dir") + File.separator + "file";// 声明file文件夹在项目中的相对路径
	public HashMap<String, String> BORRID_DEPTID_RELATION = new HashMap<String, String>();

	public ArrayList<Borrower> borrList = new ArrayList<Borrower>();// 所有借款人
	public double[] total = { 0.0, 0.0, 0.0 };

	@Override
	public void readExcel(String fname) throws Exception {
		// TODO 自动生成的方法存根
		Object co = null;
		try {
			System.out.println(path + File.separator + fname);
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fname));
			HSSFSheet sheet = workbook.getSheetAt(0);// 取Excel第一个sheet表
			Iterator<Row> rows = sheet.rowIterator();// 行迭代器
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
							co = getCellValue(cell);// 获取单元格的值
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

		
	private Object getCellValue(HSSFCell cell) {
		Object o = null;
		switch (cell.getCellType()) {
		case HSSFCell.CELL_TYPE_NUMERIC:
			o = cell.getNumericCellValue();
			break;
		case HSSFCell.CELL_TYPE_STRING:
			o = cell.getStringCellValue();
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			o = cell.getBooleanCellValue();
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			o = cell.getCellFormula();
			break;
		default:
			break;
		}
		return o;
	}


	

	@SuppressWarnings("unused")
	@Override
	public void produceExcel(String fname, String savapath) throws Exception {
		// TODO 自动生成的方法存根
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fname));// 创建一个Excel表对象
		HSSFCellStyle borderStyle = workbook.createCellStyle();// 单元格加边框样式
		HSSFCellStyle fontStyle = workbook.createCellStyle();// 单元格样式
		HSSFSheet sheet = workbook.getSheetAt(0);// 获取第一个sheet表
		Collections.sort(borrList);// 集合中借款信息按照账龄升序排序，注释掉就是不排序

		int rid = 3, rid1 = 0, rid2 = 0, rid3 = 0;
		rid1 = excelByAgingSort(0, 30, rid, borrList, workbook, sheet, borderStyle, fontStyle);
		rid2 = excelByAgingSort(30, 60, rid1, borrList, workbook, sheet, borderStyle, fontStyle);
		rid3 = excelByAgingSort(60, Integer.MAX_VALUE, rid2, borrList, workbook, sheet, borderStyle, fontStyle);

		HSSFRow row = sheet.createRow(rid3++);// 新建一个空白行
		setSubtotalRow("总计", rid3++, total, workbook, sheet, borderStyle, fontStyle);// 总计行
		leaderSign(rid3, workbook, sheet);// 领导签字行

		// 自动调节每列宽度
		for (int i = 0; i < 10; i++) {
			sheet.autoSizeColumn(i);
		}

		FileOutputStream out = null;
		try {
			out = new FileOutputStream(savapath + File.separator + "资金备用反馈表.xls");
			workbook.write(out);// workbook对象数据持久化，创建Excel
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
			System.out.println("备用资金反馈单.xls 生成");
		}
	}


	public void leaderSign(int rid, HSSFWorkbook workbook, HSSFSheet sheet) {
		rid += 2;
		sheet.addMergedRegion(new CellRangeAddress(rid, rid, 0, 9));
		HSSFRow row = sheet.createRow(rid);
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("主管签字：");
		HSSFCellStyle fontStyle = workbook.createCellStyle();
		setCellFontBoldCenter(workbook, fontStyle);
		cell.setCellStyle(fontStyle);// 样式赋给单元格
	}


	public void setSubtotalRow(String countType, int rid, double[] subtotal, HSSFWorkbook workbook, HSSFSheet sheet, HSSFCellStyle borderStyle, HSSFCellStyle fontStyle) {
		HSSFRow row = sheet.createRow(rid);
		for (int k = 0; k < 10; k++) {
			try {
				switch (k) {
				case 3:
					this.setCellValue(subtotal[0], workbook, k, row);// 原币
					break;
				case 4:
					this.setCellValue(subtotal[1], workbook, k, row);// 原币
					break;
				case 5:
					this.setCellValue(subtotal[2], workbook, k, row);// 原币
					break;
				default:
					this.setCellValue(null, workbook, k, row);// 其它赋空
					break;
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		for (int z = 0; z < 10; z++) {
			setCellBorder(workbook, row, borderStyle, z);// 设置小计行边框
		}
		sheet.addMergedRegion(new CellRangeAddress(rid, rid, 0, 2));// 合并单元格
		HSSFCell cell = row.getCell(0);
		cell.setCellValue(countType);
		setCellBorder(workbook, row, fontStyle, 0);
		setCellFontBoldCenter(workbook, fontStyle);
		cell.setCellStyle(fontStyle);// 样式赋给单元格
	}



	private int excelByAgingSort(int start, int end, int rid, ArrayList<Borrower> list, HSSFWorkbook workbook, HSSFSheet sheet, HSSFCellStyle borderStyle, HSSFCellStyle fontStyle) {
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
		setCellFontBoldCenter(workbook, fontStyle);
		cell.setCellStyle(fontStyle);

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
					this.setCellValue(borr.getBorrId(), workbook, j++, row);// 借款人
					this.setCellValue(borr.getBorrDate(), workbook, j++, row);// 日期
					this.setCellValue(borr.getPurpose(), workbook, j++, row);// 借款用途
					this.setCellValue(borr.getOriginalCurrency(), workbook, j++, row);// 原币
					this.setCellValue(borr.getVerification(), workbook, j++, row);// 已核销
					this.setCellValue(borr.getOriginalCurrencyBalance(), workbook, j++, row);// 原币余额
					this.setCellValue(borr.getAging(), workbook, j++, row);
					this.setCellValue(null, workbook, j++, row);// 情况说明单元格赋空值
					this.setCellValue(null, workbook, j++, row);
					this.setCellValue(null, workbook, j++, row);
					// 设置行边框
					for (int z = 0; z < 10; z++) {
						setCellBorder(workbook, row, borderStyle, z);// 设置小计行边框
					}
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

		setSubtotalRow("小计", _rid++, subtotal, workbook, sheet, borderStyle, fontStyle);// 小计行
		return _rid++;
	}


	public void setCellFontBoldCenter(HSSFWorkbook workbook, HSSFCellStyle fontStyle) {
		HSSFFont font = workbook.createFont(); // 创建字体对象
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 字体加粗
		fontStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 字体居中
		fontStyle.setFont(font);
	}

	public void setCellBorder(HSSFWorkbook workbook, HSSFRow row, HSSFCellStyle borderStyle, int i) {
		HSSFCell cellBorder = row.getCell(i);
		borderStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		borderStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		borderStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		borderStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		borderStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		cellBorder.setCellStyle(borderStyle);
	}

	private void setCellValue(Object value, HSSFWorkbook workbook, int nIndex, HSSFRow row) throws Exception {
		HSSFCell cell = row.createCell(nIndex);// 创建单元格
		if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if (value instanceof Number) {// 单元格为数值时，格式化数值保留两位小数
			HSSFCellStyle cellStyle = workbook.createCellStyle();// 单元格样式
			cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));// 价格保留两位小数
			cell.setCellStyle(cellStyle);
			cell.setCellValue((Double) value);
		}
	}
		
	
}
