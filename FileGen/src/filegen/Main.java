package filegen;


public class Main {


    public static void main(String[] args) throws Exception {
        String isWord = args[0];
        String address=args[1];
        String result=args[2];
        String resultType=args[3];
        String json=args[4];
        if (!json.startsWith("{")) {

            byte[] encoded = java.nio.file.Files.readAllBytes(java.nio.file.Paths.get(json));//je to soubor
            json = new String(encoded, java.nio.charset.Charset.forName( "UTF-8" ) );
        }
        ObecnyGenerator dd;
        if (isWord.equals("1")) {
            dd= new WordGenerator(address, result, resultType, json);
        } else {
            dd = new ExcelGenerator(address, result, resultType, json);
        }
        dd.doIt();
        
    }


}
