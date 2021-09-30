package com.dfl.report.ExcelReader;

public class WeldPointInfo {
	public String getPartno() {
		return partno;
	}
	public void setPartno(String partno) {
		this.partno = partno;
	}
	public String getBoardnumber() {
		return boardnumber;
	}
	public void setBoardnumber(String boardnumber) {
		this.boardnumber = boardnumber;
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
	public String getGagi() {
		return gagi;
	}
	public void setGagi(String gagi) {
		this.gagi = gagi;
	}
	public String getSheetstrength() {
		return sheetstrength;
	}
	public void setSheetstrength(String sheetstrength) {
		this.sheetstrength = sheetstrength;
	}
	private String partno; // Áã¼şºÅ
    private String boardnumber; // °å²Ä±àºÅ
    private String boardname; // °å²ÄÃû³Æ
    private String partmaterial; // °å²Ä²ÄÖÊ
    private String partthickness; // °å²Ä°åºñ   
    private String gagi; // GA¡¡/GI
    private String sheetstrength; // ²ÄÁÏÇ¿¶È
}
