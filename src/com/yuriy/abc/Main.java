package com.yuriy.abc;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class Main {
	
	
	 
	public static void main(String[] args) throws FileNotFoundException, IOException {
		String srcFileName = null;
		String defaultFileName = "atak.xls";
		
		if(args.length == 0 || args[0].length() == 0)
		{
			File f = new File(defaultFileName);
			if(f.exists() && !f.isDirectory()) { 
				srcFileName = "atak.xls";
			}
		}
		else
		{
			srcFileName = args[0];
		}
		if(srcFileName != null)
		{
			/*ReportProcessor rp = new ReportProcessor(srcFileName);
			rp.doFirstPhase();
			rp.doSecondPhase();
			rp.doThirdPhase();
			rp.removingLastRows();
			rp.writeChanges();*/
			HSSFWorkbook workBook = new HSSFWorkbook(new FileInputStream(srcFileName));
			Sheet srcSheet = workBook.getSheetAt(0);
			DOMReportProcessor proc = new DOMReportProcessor();
			proc.process(srcSheet);
			proc.removeAllRows(srcSheet);
			proc.pasteRows(srcSheet, proc.rows.toArray(new RowEntity[0]));
			
			FileOutputStream out = new FileOutputStream(srcFileName);
			workBook.write(out);
			out.close();
		}
	}
	
//	public static boolean isGrayTable(Sheet sheet, int  rowIndex)
//	{
//		if(sheet.getRow(rowIndex).getCell(0).getStringCellValue().contains("Дата / Время") && 
//				sheet.getRow(rowIndex).getCell(2).getStringCellValue().contains("Код") && 
//				sheet.getRow(rowIndex).getCell(4).getStringCellValue().contains("Класс события") && 
//				sheet.getRow(rowIndex).getCell(6).getStringCellValue().contains("ШП") && 
//				sheet.getRow(rowIndex).getCell(7).getStringCellValue().contains("Описание события") &&
//				sheet.getRow(rowIndex).getCell(10).getStringCellValue().contains("Канал"))
//			return true;
//		return false;
//	}
	
//	public static void deleteRow(Sheet sheet, Row rowToDelete) {
//
//		int lastRowNum = sheet.getLastRowNum();
//		for (int i = rowToDelete.getRowNum(); i < lastRowNum; i++) {
//			Row rowToDelete2 = sheet.getRow(i);
//		
//	        // if the row contains merged regions, delete them
//	        List<Integer> mergedRegionsToDelete = new ArrayList<>();
//	        int numberMergedRegions = sheet.getNumMergedRegions();
//	        for (int j = 0; j < numberMergedRegions; j++) {
//	            CellRangeAddress mergedRegion = sheet.getMergedRegion(j);
//	
//	            if (mergedRegion.getFirstRow() == rowToDelete2.getRowNum()
//	                    && mergedRegion.getLastRow() == rowToDelete2.getRowNum()) {
//	                // this region is within the row - so mark it for deletion
//	                mergedRegionsToDelete.add(j);
//	            }
//	        }
//	
//	        // now that we know all regions to delete just do it
//	        for (Integer indexToDelete : mergedRegionsToDelete) {
//	            sheet.removeMergedRegion(indexToDelete);
//	        }
//	
//	        int rowIndex = rowToDelete2.getRowNum();
//	
//	        // this only removes the content of the row
//	        sheet.removeRow(rowToDelete2);
//	
//	        
//	
//	        // shift the rest of the sheet one index down
//	        if (rowIndex >= 0 && rowIndex < lastRowNum) {
//	            sheet.shiftRows(rowIndex + 1, rowIndex + 1, -1, true, false);
//	            
//	        }
//		}
//    }
	
//	private static void copyRow(HSSFWorkbook workbook, Sheet srcSheet, Sheet destSheet, int sourceRowNum, int destinationRowNum) {
//        // Get the source / new row
//        Row newRow = destSheet.getRow(destinationRowNum);
//        Row sourceRow = srcSheet.getRow(sourceRowNum);
//
//        // If the row exist in destination, push down all rows by 1 else create a new row
//       // if (newRow != null) {
//           // srcSheet.shiftRows(destinationRowNum, srcSheet.getLastRowNum(), 1);
//       //} else {
//            newRow = srcSheet.createRow(destinationRowNum);
//        //}
//
//        // Loop through source columns to add to new row
//        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
//            // Grab a copy of the old/new cell
//            Cell oldCell = sourceRow.getCell(i);
//            Cell newCell = newRow.createCell(i);
//
//            // If the old cell is null jump to next cell
//            if (oldCell == null) {
//                newCell = null;
//                continue;
//            }
//
//            // Copy style from old cell and apply to new cell
//            HSSFCellStyle newCellStyle = workbook.createCellStyle();
//            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
//            
//            newCell.setCellStyle(newCellStyle);
//
//            // If there is a cell comment, copy
//            if (oldCell.getCellComment() != null) {
//                newCell.setCellComment(oldCell.getCellComment());
//            }
//
//            // If there is a cell hyperlink, copy
//            if (oldCell.getHyperlink() != null) {
//                newCell.setHyperlink(oldCell.getHyperlink());
//            }
//
//            // Set the cell data type
//            newCell.setCellType(oldCell.getCellType());
//
//            // Set the cell data value
//            switch (oldCell.getCellType()) {
//                case Cell.CELL_TYPE_BLANK:
//                    newCell.setCellValue(oldCell.getStringCellValue());
//                    break;
//                case Cell.CELL_TYPE_BOOLEAN:
//                    newCell.setCellValue(oldCell.getBooleanCellValue());
//                    break;
//                case Cell.CELL_TYPE_ERROR:
//                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
//                    break;
//                case Cell.CELL_TYPE_FORMULA:
//                    newCell.setCellFormula(oldCell.getCellFormula());
//                    break;
//                case Cell.CELL_TYPE_NUMERIC:
//                    newCell.setCellValue(oldCell.getNumericCellValue());
//                    break;
//                case Cell.CELL_TYPE_STRING:
//                    newCell.setCellValue(oldCell.getRichStringCellValue());
//                    break;
//            }
//        }
//
//        // If there are are any merged regions in the source row, copy to new row
//        for (int i = 0; i < srcSheet.getNumMergedRegions(); i++) {
//            CellRangeAddress cellRangeAddress = srcSheet.getMergedRegion(i);
//            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
//                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
//                        (newRow.getRowNum() + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow())),
//                        cellRangeAddress.getFirstColumn(),
//                        cellRangeAddress.getLastColumn());
//                srcSheet.addMergedRegion(newCellRangeAddress);
//            }
//        }
//    }
}