package filegen;

import java.util.Iterator;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class ExcelSheetGenerator {

    private int counter = 0;
    private static Pattern patt = Pattern.compile("\\$\\$\\w+");
    private HSSFWorkbook wf;

    ExcelSheetGenerator(HSSFWorkbook wf) {
        this.wf = wf;
    }

    public Sheet genSheet(Sheet sourceSheet, Map<String,String> p) {

        int index = wf.getSheetIndex(sourceSheet);
        Sheet newS = wf.cloneSheet(index);


        for (Iterator<Row> rit = newS.rowIterator(); rit.hasNext();) {
            Row row = rit.next();
            for (Iterator<Cell> cit = row.cellIterator(); cit.hasNext();) {
                Cell sourceCell = cit.next();

                String val = sourceCell.getStringCellValue();
                if (val.contains("##")) {
                    String substr = val.substring(2);

                    double wat;
                    if (substr.equals("poradi")) {
                        ++counter;
                        wat = counter;
                    } else {
                        String watS = p.get(substr);
                        wat = Double.parseDouble(watS);
                    }
                    if (!val.contains("notNull") || wat != 0) {
                        sourceCell.setCellValue(wat);
                    } else {
                        sourceCell.setCellValue(" ");
                    }



                } else {

                    Matcher m = patt.matcher(val);

                    while (m.find()) {
                        String subst = val.substring(m.start() + 2, m.end());
                        String wat = "CHYBNY OBJECT! ";

                        wat = p.get(subst);
                        if (wat == null) {
                            System.out.println("OMG SPADNU NA " + subst);
                        }
                        val = m.replaceFirst(wat);
                        m = patt.matcher(val);
                    }


                    if (!val.equals("")) {
                        sourceCell.setCellValue(val);
                    } else {
                        String n = null;
                        sourceCell.setCellValue(n);
                    }
                }
            }
        }
        return newS;
    }
}
