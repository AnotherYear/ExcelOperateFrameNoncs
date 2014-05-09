package com.Hydro.service;

public interface ReportServiceInterface {
	public void readExcel(String fname) throws Exception;
	public void produceExcel(String fname,String savapath) throws Exception;

}
