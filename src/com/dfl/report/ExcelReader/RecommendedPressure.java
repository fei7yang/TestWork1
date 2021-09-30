package com.dfl.report.ExcelReader;
/* *******************************
 * 推荐加压力
 * *********************************/
public class RecommendedPressure {
	  public String getBasethickness() {
			return basethickness;
		}
		public void setBasethickness(String basethickness) {
			this.basethickness = basethickness;
		}
		public String getBvalue() {
			return Bvalue;
		}
		public void setBvalue(String bvalue) {
			Bvalue = bvalue;
		}
		public String getCvalue() {
			return Cvalue;
		}
		public void setCvalue(String cvalue) {
			Cvalue = cvalue;
		}
		public String getDvalue() {
			return Dvalue;
		}
		public void setDvalue(String dvalue) {
			Dvalue = dvalue;
		}
		public String getEvalue() {
			return Evalue;
		}
		public void setEvalue(String evalue) {
			Evalue = evalue;
		}
		public String getFvalue() {
			return Fvalue;
		}
		public void setFvalue(String fvalue) {
			Fvalue = fvalue;
		}
	
		private String basethickness; //基准板厚
	     private String Bvalue; //B列
	     private String Cvalue; //C列
	     private String Dvalue; //D列
	     private String Evalue; //E列
	     private String Fvalue; //F列
	     
}
