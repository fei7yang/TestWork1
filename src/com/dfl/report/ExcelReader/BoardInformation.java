package com.dfl.report.ExcelReader;

/*
 * ������Ϣ�еİ�����Ϣ
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
		private String boardnumber; // ��ı��
	    private String partn; // ������
	    private String boardname; // �������
	    private String partmaterial; // ��Ĳ���
	    private String partthickness; // ��İ��
	    private String sheetstrength; // ����ǿ��(Mpa)
	    private String gagi; // GA��/GI
	    private String rowNum; //���
		private String maunit; //��İ��λ
	    private String thunit; //����ǿ��(Mpa)��λ
	    
}
