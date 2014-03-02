package filegen;


public class Main {


    public static void main(String[] args) throws Exception {
        ExcelGenerator wg = new ExcelGenerator("/Users/karelbilek/insolve/template/plachta.xls",
                "/Users/karelbilek/insolve/template/resultie",
                "xls",
                "{\"sideways\":1, \"fitWidth\":1, \"fitHeight\":1, \"hlavicky\":[\"vse\"],"
                + "\"entities\":["
                + "{\"type\":\"Kauzy\",\"data\":{\"jmenoSoudu\":\"bluargh\", \"spisovaZnacka\":\"bluargh\", \"popis\":\"bluargh\", \"celkem\":\"123.45\"}}"
                + ",{\"type\":\"Veritele\",\"data\":{\"poradCislo\":\"1\", \"dorucenoSoudu\":\"1.1.2011\",\"jmenoVeritele\":\"bluarh\", \"rcIc\":\"bigdick\",\"vysePohledavek\":\"88.48\"}}"
                + "]}");
        wg.doIt();
    }


}
