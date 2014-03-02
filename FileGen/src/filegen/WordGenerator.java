package filegen;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.fasterxml.jackson.core.JsonParseException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.ConnectException;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import net.sf.jooreports.templates.DocumentTemplate;
import net.sf.jooreports.templates.DocumentTemplateFactory;


public class WordGenerator {
    public String getNativeFormat(){
        return "odt";
    }

    String sourceName;

    String finalName;
    String finalFormat;
    String jsonstring;

    public WordGenerator(String sourceName, String finalName, String finalFormat, String jsonstring) {
        this.sourceName = sourceName;
        this.finalName = finalName;
        this.finalFormat = finalFormat;
        this.jsonstring = jsonstring;
    }

    public void doIt() {
        
    }


    public HashMap getJsonMap() throws IOException {
            return new ObjectMapper().
                    readValue(jsonstring, HashMap.class);

    }

    public File generateDifferent(String endformat) throws Exception {
        File original = generateNatural();
        File convertedFile = File.createTempFile("convfile", "." + endformat);

        OpenOfficeConnection connection = new SocketOpenOfficeConnection(8100);
        connection.connect();
        DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
        converter.convert(original, convertedFile);
        connection.disconnect();
        return convertedFile;
    }

    public File generateNatural() throws Exception {
        DocumentTemplateFactory documentTemplateFactory =
                new DocumentTemplateFactory();
        DocumentTemplate template;

            InputStream is = new FileInputStream(sourceName);
            template = documentTemplateFactory.getTemplate(is);
            is.close();

            File ODTFile = File.createTempFile("ODTemp", ".odt");
            FileOutputStream outStream = new FileOutputStream(ODTFile);


            template.createDocument(getJsonMap(), outStream);

            outStream.close();

            System.out.println(ODTFile);
            //Runtime.getRuntime().exec("open "+ODTFile);

            ODTFile.deleteOnExit();

            //AllHelper.rekniNahlas(
              //      "Soubor vygenerovan.");
            return ODTFile;


    }
}
