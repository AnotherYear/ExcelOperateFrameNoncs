package com.Hydro.model;
/**
 * 借款信息
 * 实现Comparable排序接口
 */
public class Borrower implements Comparable<Borrower> {
	private String borrId;// 借款人 
	private String deptId;// 所属部门
	private String borrDate;// 借款日期
	private String purpose;// 借款用途
	private double originalCurrency;// 原币
	private double verification;// 已核销
	private double originalCurrencyBalance;// 原币余额
	private double aging;// 账龄

	public String getBorrId() {
		return borrId;
	}
 
	public void setBorrId(String borrId) {
		this.borrId = borrId;
	}

	public String getDeptId() {
		return deptId;
	}

	public void setDeptId(String deptId) {
		this.deptId = deptId;
	}

	public String getBorrDate() {
		return borrDate;
	}

	public void setBorrDate(String borrDate) {
		this.borrDate = borrDate;
	}

	public String getPurpose() {
		return purpose;
	}

	public void setPurpose(String purpose) {
		this.purpose = purpose;
	}

	public double getOriginalCurrency() {
		return originalCurrency;
	}

	public void setOriginalCurrency(double originalCurrency) {
		this.originalCurrency = originalCurrency;
	}

	public double getVerification() {
		return verification;
	}

	public void setVerification(double verification) {
		this.verification = verification;
	}

	public double getOriginalCurrencyBalance() {
		return originalCurrencyBalance;
	}

	public void setOriginalCurrencyBalance(double originalCurrencyBalance) {
		this.originalCurrencyBalance = originalCurrencyBalance;
	}

	public double getAging() {
		return aging;
	}

	public void setAging(double aging) {
		this.aging = aging;
	}

	public int compareTo(Borrower o) {
		return (int) (this.getAging() - o.getAging());
	}

	public String toString() {

		return this.getBorrId() + ",  deptId:" + this.getDeptId() + ", BorrDate:" + this.getBorrDate() + ",  Purpose:" + this.getPurpose() + ",  originalCurrency:" + this.getOriginalCurrency()
				+ ",  verification:" + this.getVerification() + ",  originalCurrencyBalance:" + this.getOriginalCurrencyBalance() + ",  aging:" + this.getAging();
	}

}
