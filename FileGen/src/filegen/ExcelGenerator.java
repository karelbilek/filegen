package filegen;

import com.fasterxml.jackson.core.type.TypeReference;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;


/**
 * Třídy na generování excelu jsem dělal kdysi dávno (cca 2011),
 * už vůbec nevím, jak vlastně fungují, ale je to takový "black box",
 * co mi funguje a nešahám na něj.
 *
 * Tahle vygeneruje XLS a udělá věci kolem, ale pak to pošle
 * na ExcelGenerator a ExcelSerialiser
 */
public class ExcelGenerator extends ObecnyGenerator {
    
    public ExcelGenerator(String templateName, String finalName, String finalFormat, String jsonstring) {
        super(templateName,finalName,finalFormat,jsonstring);
        System.out.println(jsonstring);
    }

    @Override
    protected String getNativeFormat() {
        return "xls";
    }

    @Override
    protected File generateNatural(File templateName, Object jsonO) throws Exception {

        ExcelData ed = (ExcelData) jsonO;
        int number_of_sheets = 1;
        for (ExcelEntity entity : ed.entities) {
            if (entity.type.equals("newpage")) {
                number_of_sheets++;
            }
        }


        InputStream formatIs = new FileInputStream(templateName);
        HSSFWorkbook mainWb = new HSSFWorkbook(formatIs);
      


        List<Sheet> ss = new ArrayList<>();
        ExcelSheetGenerator esg = new ExcelSheetGenerator(mainWb);

        List<ExcelEntity> entities = ed.entities;

        
            for (ExcelEntity entita : entities) {


                if (!entita.type.equals("newpage")) {
                    System.out.println("Hledam getclass " + entita.type);
                    Sheet sourceS = mainWb.getSheet(entita.type);
                    System.out.println("Nalezena " + sourceS.toString() + "YAY");
                    Sheet copyS = esg.genSheet(sourceS, entita.data);
                    ss.add(copyS);
                } else {
                    if (ss.isEmpty()) {
                        throw new Exception("first sheet of a result can't be null");
                    }
                    ss.add(null);
                    //null = sheet breaker
                    //first should never be null
                }
            }
        


        List<Sheet> outputSheets = new ArrayList<>();
        for (int i = 0; i < number_of_sheets; i++) {
            String c = Integer.toString(i + 1);
            Sheet outp = mainWb.createSheet("vysledek " + c);
            outputSheets.add(outp);
        }


        ExcelSerialiser es = new ExcelSerialiser(
                ss.get(0), ss, outputSheets, 40);

        es.serialise();

        List<Sheet> allSheets = new ArrayList<>();
        for (int i = 0; i < mainWb.getNumberOfSheets(); ++i) {
            allSheets.add(mainWb.getSheetAt(i));
        }


        for (Sheet s : allSheets) {
            if (!outputSheets.contains(s)) {
                int w = mainWb.getSheetIndex(s);
                mainWb.removeSheetAt(w);
            }
        }

        int j=0;
        for (Sheet s : outputSheets) {
            Header h = s.getHeader();

            h.setLeft("");
            j++;
            PrintSetup p = s.getPrintSetup();
            p.setLandscape(ed.sideways);

            System.out.println("pred");
            System.out.println(p.getFitHeight());
            System.out.println(p.getFitWidth());

            s.setMargin(Sheet.BottomMargin, 0.4);
            s.setMargin(Sheet.TopMargin, 0.4);
            s.setMargin(Sheet.RightMargin, 0.4);
            s.setMargin(Sheet.LeftMargin, 0.4);

            s.setAutobreaks(true);

            p.setPaperSize(PrintSetup.A4_PAPERSIZE);

            short height=(short) (ed.fitHeight ? 1 : 0);
            short width = (short) (ed.fitWidth ? 1 : 0);
            p.setFitHeight(height);
            p.setFitWidth(width);
            
            for (ExcelImageData image: ed.images) {
                int pictureIdx;
                try (InputStream inputStream = new FileInputStream(image.path)) {
                    byte[] bytes = IOUtils.toByteArray(inputStream);
                    pictureIdx = mainWb.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_PNG);
                }
                
                CreationHelper helper = mainWb.getCreationHelper();
                Drawing drawing = s.createDrawingPatriarch();
                ClientAnchor anchor = helper.createClientAnchor();
                anchor.setCol1(image.x);
                anchor.setRow1(image.y);
                Picture pict = drawing.createPicture(anchor, pictureIdx);
                pict.resize(image.scale);
            }


            System.out.println(p.getFitHeight());
            System.out.println(p.getFitWidth());

        }

            File f = File.createTempFile("exceltemp", ".xls");
            f.deleteOnExit();
            System.out.println(f);
            try (FileOutputStream outStream = new FileOutputStream(f)) {
                mainWb.write(outStream);
            }
            return f;




    }
    
    @Override
    protected TypeReference getType() {
        return new TypeReference<ExcelData>() {
        };
    }

}
