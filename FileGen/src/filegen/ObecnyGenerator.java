package filegen;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.io.File;
import java.io.IOException;
import org.apache.commons.io.FileUtils;


public abstract class ObecnyGenerator {
    abstract protected String getNativeFormat();

    private final String templateName;
    private final String finalName;
    private final String finalFormat;
    private final String jsonstring;
    

    public ObecnyGenerator(String templateName, String finalName, String finalFormat, String jsonstring) {
        this.templateName = templateName;
        this.finalName = finalName;
        this.finalFormat = finalFormat;
        this.jsonstring = jsonstring;
    }


    final public void doIt() throws Exception {
        File template = new File(templateName);
        if (!template.exists()) {
            throw new Exception("No template.");
        }
        File r = this.generateNatural(template, this.getJsonMap());
        if (!finalFormat.equals(this.getNativeFormat())) {
            r=this.generateDifferent(finalFormat, r);
        }
        File cil = new File(finalName+"."+finalFormat);
        if (cil.exists()) {
            throw new Exception("Soubor uz existuje.");
        }
        FileUtils.copyFile(r, cil);
    }

    abstract protected TypeReference getType();

    private Object getJsonMap() throws IOException {
        ObjectMapper mapper = new ObjectMapper();

        Object mapObject = mapper.readValue(jsonstring, this.getType());
        return mapObject;

    }

    private File generateDifferent(String endformat, File original) throws Exception {
        File convertedFile = File.createTempFile("convfile", "." + endformat);

        OpenOfficeConnection connection = new SocketOpenOfficeConnection(8101);
        connection.connect();
        DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
        converter.convert(original, convertedFile);
        connection.disconnect();
        convertedFile.deleteOnExit();

        return convertedFile;
    }

    abstract protected File generateNatural(File templateName, Object jsonObject) throws Exception;

  
}
