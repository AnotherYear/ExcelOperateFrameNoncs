package com.Hydro.utils;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelUtils {

	/**
	 * 获取源Excel表中每行每个单元格的值，根据单个元数值类型，返回响应的数值对象
	 * 
	 * @param cell
	 *            单元格对象
	 * @return
	 */
	public Object getCellValue(HSSFCell cell) {
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

	/**
	 * 创建单元格并赋值
	 * 
	 * @param value
	 *            数值
	 * @param workbook
	 *            Excel表对象
	 * @param nIndex
	 *            单元格索引
	 * @param row
	 *            行对象
	 * @throws Exception
	 */
	public void setCellValue(Object value, HSSFWorkbook workbook, int nIndex, HSSFRow row, HSSFCellStyle style) throws Exception {
		HSSFCell cell = row.createCell(nIndex);// 创建单元格
		if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if (value instanceof Number) {// 单元格为数值时，格式化数值保留两位小数
			cell.setCellValue((Double) value);
		}
		if (style != null) {
			cell.setCellStyle(style);
		}
	}

	/**
	 * Excel持久化
	 * 
	 * @param workbook
	 *            Excel对象
	 * @param path
	 *            持久化路径
	 */
	public void exportExcel(HSSFWorkbook workbook, String path) {
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(path);
			workbook.write(out);// workbook对象数据持久化，创建Excel
			out.close();
			if (path != null) {
				path = path.substring(path.lastIndexOf(File.separator) + 1);
			}
			System.out.println(path + "生成");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
