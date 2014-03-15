package filegen;

import com.fasterxml.jackson.core.type.TypeReference;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Map;
import net.sf.jooreports.templates.DocumentTemplate;
import net.sf.jooreports.templates.DocumentTemplateFactory;


public class WordGenerator extends ObecnyGenerator {
    
    public WordGenerator(String templateName, String finalName, String finalFormat, String jsonstring) {
        super(templateName,finalName,finalFormat,jsonstring);
    }

    protected String getNativeFormat() {
        return "odt";
    }

    protected File generateNatural(File templateName, Object mapO) throws Exception {
        DocumentTemplateFactory documentTemplateFactory =
                new DocumentTemplateFactory();

        WordData wd=(WordData) mapO;

        DocumentTemplate template;

        InputStream is = new FileInputStream(templateName);
        template = documentTemplateFactory.getTemplate(is);
        is.close();

        File ODTFile = File.createTempFile("ODTemp", ".odt");
        FileOutputStream outStream = new FileOutputStream(ODTFile);

        Map<String, Object> mapa = wd.all();
        template.createDocument(mapa, outStream);

        outStream.close();

        System.out.println(ODTFile);
        //Runtime.getRuntime().exec("open "+ODTFile);

        ODTFile.deleteOnExit();

        //AllHelper.rekniNahlas(
        //      "Soubor vygenerovan.");
        return ODTFile;


    }
    
    protected TypeReference getType() {
        return new TypeReference<WordData>() {
        };
    }

}
