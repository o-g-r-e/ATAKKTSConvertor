package com.yuriy.abc;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;

public class RowEntity {
	private CellEntity[] cells;
	private CellRange[] areas;
	
	public RowEntity(CellEntity[] cells, CellRange[] areas) {
		this.cells = cells;
		this.areas = areas;
	}

	public CellEntity[] getCells() {
		return cells;
	}

	public CellRange[] getAreas() {
		return areas;
	}
}