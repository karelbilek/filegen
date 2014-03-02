package filegen;

import java.util.List;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelSerialiser {

    private Sheet vzor;
    private List<Sheet> pasteSheets;
    private List<Sheet> sheets;
    //private HSSFWorkbook pasteSheetWorkbook;
    private int maxRows;

    public ExcelSerialiser(Sheet vzor,
            List<Sheet> sheets,
            List<Sheet> pasteSheets,
            int maxRows) {
        this.vzor = vzor;
        this.sheets = sheets;
        this.pasteSheets = pasteSheets;
        this.maxRows = maxRows;
        //pasteSheetWorkbook = (HSSFWorkbook) pasteSheets.get(0).getWorkbook();
    }

    static String getValueFromCell(Cell cell){
        try {
            String s = cell.getStringCellValue();
            return s;
        } catch (IllegalStateException E) {
            //this is fine - it's just number...
            return "";
        }

    }
    public void serialise() {


        Sheet pasteSheet = null;
        boolean begin = true;
        int lowestRow = 0;
        int index = 0;
        int lowestPageRow = 0;
        for (Sheet copySheet : sheets) {
            if (copySheet == null || begin) {

                pasteSheet = pasteSheets.get(index);
                index++;
                for (int j = 0; j < 50; j++) {
                    
                    if (vzor == null) {
                        System.out.println("watlord");
                    }
                    pasteSheet.setColumnWidth(j, vzor.getColumnWidth(j));
                }
                lowestRow = 0;
                if (!begin) {
                    continue;
                } else {
                    begin = false;
                }
            }



            for (int j = 0; j < copySheet.getNumMergedRegions(); j++) {
                CellRangeAddress region = copySheet.getMergedRegion(j);
                CellRangeAddress newRegion = new CellRangeAddress(
                        region.getFirstRow() + lowestRow,
                        region.getLastRow() + lowestRow,
                        region.getFirstColumn(),
                        region.getLastColumn());
                pasteSheet.addMergedRegion(newRegion);
            }

            int lowestCurrentRow = -1;

            for (int currentRow = 0; true; currentRow++) {

                lowestPageRow++;

                Row pasteRow = pasteSheet.getRow(currentRow + lowestRow);
                Row copyRow = copySheet.getRow(currentRow);
                if (copyRow == null) {
                    continue;
                }
                if (pasteRow == null) {
                    pasteRow = pasteSheet.createRow(currentRow + lowestRow);
                }

                Cell testC = copyRow.getCell(0);
                
                    if ((testC != null)
                            && (ExcelSerialiser.getValueFromCell(testC).equals("%%END%%"))) {
                        break;
                    }
                

                for (int currentColumn = 0; true; currentColumn++) {
                    Cell copyCell = copyRow.getCell(currentColumn);
                    Cell pasteCell = pasteRow.getCell(currentColumn);
                    if (copyCell == null) {
                        continue;
                    }
                    
                        if (ExcelSerialiser.getValueFromCell(copyCell).equals("PAGEBREAK")) {
                            if (maxRows != 0 && lowestPageRow > this.maxRows) {
                                pasteSheet.setRowBreak(currentRow + lowestRow);
                                lowestPageRow = 0;
                            }
                            continue;
                        }
                    
                    if (pasteCell == null) {
                        pasteCell = pasteRow.createCell(currentColumn);
                    }
                    
                        if (ExcelSerialiser.getValueFromCell(copyCell).equals("%%END%%")) {
                            break;
                        }
                    

                    CellStyle cs = copyCell.getCellStyle();
                    if (cs != null) {
                        pasteCell.setCellStyle(cs);
                    }

                    boolean numeric;
                    double del_d = 0;
                    String del_s = null;

                    try {
                        del_s = copyCell.getStringCellValue();
                        numeric = false;

                    } catch (Exception E) {
                        del_d = copyCell.getNumericCellValue();
                        numeric = true;
                    }


                    if (!numeric) {
                        if (!del_s.equals("")) {
                            pasteCell.setCellValue(del_s);
                        } else {
                            String n = null;
                            pasteCell.setCellValue(n);
                        }
                    } else {
                        pasteCell.setCellValue(del_d);
                    }


                }
                lowestCurrentRow = currentRow;
            }
            lowestRow += lowestCurrentRow;
            lowestRow++;

        }



    }
}
