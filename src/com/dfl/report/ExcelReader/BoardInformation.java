package com.dfl.report.ExcelReader;

/*
 * 基本信息中的板组信息
 */
public class BoardInformation {
	    public String getBoardnumber() {
		return boardnumber;
	}
	public void setBoardnumber(String boardnumber) {
		this.boardnumber = boardnumber;
	}
	public String getPartn() {
		return partn;
	}
	public void setPartn(String partn) {
		this.partn = partn;
	}
	public String getBoardname() {
		return boardname;
	}
	public void setBoardname(String boardname) {
		this.boardname = boardname;
	}
	public String getPartmaterial() {
		return partmaterial;
	}
	public void setPartmaterial(String partmaterial) {
		this.partmaterial = partmaterial;
	}
	public String getPartthickness() {
		return partthickness;
	}
	public void setPartthickness(String partthickness) {
		this.partthickness = partthickness;
	}
	public String getSheetstrength() {
		return sheetstrength;
	}
	public void setSheetstrength(String sheetstrength) {
		this.sheetstrength = sheetstrength;
	}
	public String getGagi() {
		return gagi;
	}
	public void setGagi(String gagi) {
		this.gagi = gagi;
	}
	public String getRowNum() {
		return rowNum;
	}
	public void setRowNum(String rowNum) {
		this.rowNum = rowNum;
	}
    public String getMaunit() {
		return maunit;
	}
	public void setMaunit(String maunit) {
		this.maunit = maunit;
	}
	public String getThunit() {
		return thunit;
	}
	public void setThunit(String thunit) {
		this.thunit = thunit;
	}
		private String boardnumber; // 板材编号
	    private String partn; // 零件编号
	    private String boardname; // 板材名称
	    private String partmaterial; // 板材材质
	    private String partthickness; // 板材板厚
	    private String sheetstrength; // 材料强度(Mpa)
	    private String gagi; // GA　/GI
	    private String rowNum; //序号
		private String maunit; //板材板厚单位
	    private String thunit; //材料强度(Mpa)单位
	    
}
