import converter.ExcelConverter;

import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        String pathIn="C:/Users/Eughen.Baldji/IdeaProjects/KOLEA/htmltableconverter/src/main/resources/test.html";
        String name="someName";
        String pathOut="C:\\Users\\Eughen.Baldji\\IdeaProjects\\KOLEA\\htmltableconverter\\src\\main\\resources\\test";

//        Converter converter = new Converter();
//        String str = converter.readFromFile("/home/kob/work/KOLEA/htmlconverter/src/main/resources/report.html");
//        String str = converter.readFromFile("C:/Users/Eughen.Baldji/IdeaProjects/KOLEA/htmltableconverter/src/main/resources/test.html");
//        converter.convertHtmlTableToExcel("someName","/home/kob/work/KOLEA/htmlconverter/src/main/resources/report",str);
//        converter.convertHtmlTableToExcel("someName","C:\\Users\\Eughen.Baldji\\IdeaProjects\\KOLEA\\htmltableconverter\\src\\main\\resources\\test",str);
//         ExcelConverter.readAndexportToExcel(name,pathIn,pathOut);
         new ExcelConverter().readAndexportToExcel(name,pathIn,pathOut);
    }
}
