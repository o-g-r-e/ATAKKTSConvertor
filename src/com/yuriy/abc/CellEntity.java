package com.yuriy.abc;

import org.apache.poi.ss.usermodel.CellStyle;

public class CellEntity {
	private String data;
	private CellStyle style;
	
	public CellEntity(String data, CellStyle style) {
		this.data = data;
		this.style = style;
	}

	public String getData() {
		return data;
	}
	
	public void setData(String data) {
		this.data = data;
	}

	public CellStyle getStyle() {
		return style;
	}
}