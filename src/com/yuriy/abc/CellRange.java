package com.yuriy.abc;

public class CellRange
{
	private int firstRow;
	private int lastRow;
	private int firstCol;
	private int lastCol;
	
	public CellRange(int firstRow, int lastRow, int firstCol, int lastCol) {
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.firstCol = firstCol;
		this.lastCol = lastCol;
	}

	public int getFirstRow() {
		return firstRow;
	}

	public int getLastRow() {
		return lastRow;
	}

	public void setFirstRow(int firstRow) {
		this.firstRow = firstRow;
	}

	public void setLastRow(int lastRow) {
		this.lastRow = lastRow;
	}

	public int getFirstCol() {
		return firstCol;
	}

	public int getLastCol() {
		return lastCol;
	}
}