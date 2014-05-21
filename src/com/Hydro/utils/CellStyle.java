package com.Hydro.utils;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CellStyle {

	public HSSFCellStyle cellNumberFormatStyle = null;// 单元格数字精确度样式
	public HSSFCellStyle cellBorderStyle = null;// 单元格边框样式
	public HSSFCellStyle cellFontBoldCenterBorderStyle = null;// 单元格字体加粗居中有边框样式
	public HSSFCellStyle cellFontBoldCenterNoBorderStyle = null;// 单元格字体加粗居中无边框样式
	public HSSFCellStyle cellFontBoldBorderStyle = null;// 单元格字体加粗有边框样式

	public CellStyle(HSSFWorkbook workbook) {
		cellNumberFormatStyle = getCellNumberFormat(workbook);
		cellBorderStyle = getCellBorder(workbook);
		cellFontBoldCenterBorderStyle = getCellFontBoldCenterBorder(workbook);
		cellFontBoldCenterNoBorderStyle = getCellFontBoldCenterNoBorder(workbook);
		cellFontBoldBorderStyle = getCellFontBoldBorder(workbook);
	}

	/**
	 * 单元格数字精确度样式
	 */
	public HSSFCellStyle getCellNumberFormat(HSSFWorkbook workbook) {
		HSSFCellStyle style = workbook.createCellStyle();// 单元格样式
		style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));// 价格保留两位小数
		return style;
	}

	/**
	 * 单元格边框样式
	 */
	public HSSFCellStyle getCellBorder(HSSFWorkbook workbook) {
		HSSFCellStyle style = workbook.createCellStyle();// 单元格加边框样式
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		return style;
	}

	/**
	 * 单元格字体加粗居中无边框样式
	 */
	public HSSFCellStyle getCellFontBoldCenterNoBorder(HSSFWorkbook workbook) {
		HSSFCellStyle style = workbook.createCellStyle();// 单元格加边框样式
		HSSFFont font = workbook.createFont(); // 创建字体对象
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 字体加粗
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 字体居中
		style.setFont(font);

		return style;
	}

	/**
	 * 单元格字体加粗居中有边框样式
	 */
	public HSSFCellStyle getCellFontBoldCenterBorder(HSSFWorkbook workbook) {
		HSSFCellStyle style = workbook.createCellStyle();// 单元格加边框样式
		HSSFFont font = workbook.createFont(); // 创建字体对象
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 字体加粗
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);// 字体居中

		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);

		style.setFont(font);

		return style;
	}

	/**
	 * 单元格字体加粗有边框样式
	 */
	public HSSFCellStyle getCellFontBoldBorder(HSSFWorkbook workbook) {
		HSSFCellStyle style = workbook.createCellStyle();// 单元格加边框样式
		HSSFFont font = workbook.createFont(); // 创建字体对象
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// 字体加粗

		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);

		style.setFont(font);

		return style;
	}

}
