package converter;

import model.ConvertModel;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.*;
import java.net.MalformedURLException;
import java.util.List;
import java.util.Scanner;
import java.util.stream.Collectors;

public class ExcelConverter {
    public  String readFromFile(String path) throws IOException {
        String str = "";
//        try {
//            File myObj = new File(path);
//            Scanner myReader = new Scanner(myObj);
//            while (myReader.hasNextLine()) {
//                String data = myReader.nextLine();
//                str += data;
//            }
//            myReader.close();
//        } catch (FileNotFoundException e) {
//            System.out.println("An error occurred.");
//            e.printStackTrace();
//        }
//        return str;
        BufferedReader br = new BufferedReader(new FileReader(path));
        try {
            StringBuilder sb = new StringBuilder();
            String line = br.readLine();

            while (line != null) {
                sb.append(line);
                sb.append(System.lineSeparator());
                line = br.readLine();
            }
            str = sb.toString();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            br.close();
        }
        return str;
    }

    public    void exportToExcel(String name, String pathOut, String str) throws IOException {
        boolean isPresent = str.indexOf("<table") != -1 ? true : false;
        if (isPresent) {
            List<Element> tables = getContentByTag(str, "table");

            HSSFWorkbook workbook = fillExcel(tables);
            FileOutputStream out =
                    new FileOutputStream(new File(pathOut + ".xls"));
            workbook.write(out);
            out.close();
        }


    }

    public  void readAndexportToExcel(String name, String pathIn, String pathOut) throws IOException {
        String str = readFromFile(pathIn);
        exportToExcel(name, pathOut, str);

    }

    private  ConvertModel fillHead(ConvertModel convertModel, Elements theadElement, Element element) {
        if (theadElement.size() != 0) {
            Elements trElemnts = theadElement.get(0).getElementsByTag("tr");
            Elements tdElements = trElemnts.get(0).getElementsByTag("td");
            Elements thElements = trElemnts.get(0).getElementsByTag("th");
            Integer tdHeadCount = 0;
            Integer position = 0;
            for (Element tdElement : tdElements) {
                Element tbodyElement = element.getElementsByTag("tbody")
                        .stream().filter(e -> e.parent()
                                .equals(element)).collect(Collectors.<Element>toList()).get(0);
                List<Element> trBodyElemnts = tbodyElement.getElementsByTag("tr")
                        .stream().filter(e -> e.parent().equals(tbodyElement)).collect(Collectors.<Element>toList());
                Integer cellColumn = checkCellColSpanBody(trBodyElemnts, position);
                String tdHeadValue = tdElement.toString().replaceAll("<[^>]*>", "").replaceAll("<[^>]*>", "");
                if (cellColumn != 1) {
                    convertModel.getSheet()
                            .addMergedRegion(new CellRangeAddress(convertModel.getRownum(), convertModel.getRownum(), tdHeadCount, tdHeadCount + cellColumn - 1));
                }
                convertModel.setCell(convertModel.getRow().createCell(tdHeadCount));
//                cell = row.createCell(tdHeadCount);
                tdHeadCount++;
                if (cellColumn != 1) {
                    tdHeadCount++;
                }
                convertModel.getCell().setCellValue(tdHeadValue);
//                cell.setCellValue(tdHeadValue);
                position++;
            }
            for (Element thElement : thElements) {
                String tdHeadValue = thElement.toString().replaceAll("<[^>]*>", "").replaceAll("<[^>]*>", "");
                convertModel.setCell(convertModel.getRow().createCell(tdHeadCount));
//                cell = row.createCell(tdHeadCount);
                convertModel.getCell().setCellValue(tdHeadValue);
//                cell.setCellValue(tdHeadValue);
                tdHeadCount++;
            }
        }
        return convertModel;
    }

    private  ConvertModel fillBody(ConvertModel convertModel, List<Element> tbodyElements, Element element) {
        if (tbodyElements.size() != 0) {
            for (Element tbodyElement : tbodyElements) {
                List<Element> trBodyElemnts = tbodyElement.getElementsByTag("tr")
                        .stream().filter(e -> e.parent().equals(tbodyElement)).collect(Collectors.<Element>toList());
                if (trBodyElemnts.size() != 0) {
                    Integer rowspanPos = 0, colspanPos = 0;
                    for (Element trBodyElement : trBodyElemnts) {
//                            rownum++;
                        convertModel.setRow(convertModel.getSheet().createRow(convertModel.getRownum()));
//                        row = sheet.createRow(rownum);
                        List<Element> tdElements = trBodyElement.getElementsByTag("td")
                                .stream().filter(e -> e.parent().equals(trBodyElement)).collect(Collectors.<Element>toList());
                        Integer tdBodyCount = 0;
                        if (tdElements.size() != 0) {
//                                if (rowspanPos != 0) {
//                                    rownum = rowspanPos;
//                                }
                            if (colspanPos != 0) {
                                tdBodyCount = colspanPos;
                            }
                            for (Element tdElement : tdElements) {

                                if (tdElement.toString().indexOf("<table") != -1 ? true : false) {

                                    String rowspan = tdElement.attr("rowspan");
                                    String colspan = tdElement.attr("colspan");
                                    Elements subTdElement = tdElement.getElementsByTag("td");
                                    String tdBodyValue = "";
                                    for (int j = 1; j < subTdElement.size(); j++) {
                                        if (rowspan != null && rowspan != "") {
                                            convertModel.getSheet().
                                                    addMergedRegion(new CellRangeAddress(convertModel.getRownum(), convertModel.getRownum() + Integer.parseInt(rowspan) - 2, tdBodyCount, tdBodyCount));
                                            rowspanPos = convertModel.getRownum() + Integer.parseInt(rowspan) - 2;
                                            colspanPos++;
                                        }
                                        if (colspan != null && colspan != "") {
                                            convertModel.getSheet()
                                                    .addMergedRegion(new CellRangeAddress(convertModel.getRownum(), convertModel.getRownum(), tdBodyCount, tdBodyCount + Integer.parseInt(colspan) - 1));
                                        }
//                                            tdBodyValue += subTdElement.get(j).text() + " ";
                                        tdBodyValue = subTdElement.get(j).text();
                                        convertModel.setCell(convertModel.getRow().createCell(tdBodyCount));
//                                                 cell = row.createCell(tdBodyCount);
                                        convertModel.getCell().setCellValue(tdBodyValue);
//                                        cell.setCellValue(tdBodyValue);
                                        tdBodyCount++;
                                    }

//                                        cell = row.createCell(tdBodyCount);
//                                        cell.setCellValue(tdBodyValue);
                                } else {
                                    String rowspan = tdElement.attr("rowspan");
                                    String colspan = tdElement.attr("colspan");
                                    if (rowspan != null && rowspan != "") {
                                        convertModel.getSheet().addMergedRegion(new CellRangeAddress(convertModel.getRownum(), convertModel.getRownum() + Integer.parseInt(rowspan) - 2, tdBodyCount, tdBodyCount));
                                        rowspanPos = convertModel.getRownum() + Integer.parseInt(rowspan) - 2;
                                        colspanPos++;
                                    }
                                    if (colspan != null && colspan != "") {
                                        convertModel.getSheet()
                                                .addMergedRegion(new CellRangeAddress(convertModel.getRownum(), convertModel.getRownum(), tdBodyCount, tdBodyCount + Integer.parseInt(colspan) - 1));
                                    }
                                    String tdBodyValue = tdElement.toString().replaceAll("<[^>]*>", "").replaceAll("<[^>]*>", "");
                                    convertModel.setCell(convertModel.getRow().createCell(tdBodyCount));
//                                                 cell = row.createCell(tdBodyCount);
                                    convertModel.getCell().setCellValue(tdBodyValue);
//                                        cell.setCellValue(tdBodyValue);


                                    tdBodyCount++;

                                }
                            }
                            System.out.println();
                        }
                        System.out.println();
                        convertModel.setRownum(convertModel.getRownum() + 1);
                    }
                }
            }
        }
        return convertModel;
    }

    private   HSSFWorkbook fillExcel(List<Element> elements) {
//        elements.remove(1);
//        HSSFWorkbook workbook = new HSSFWorkbook();
//        HSSFSheet sheet = workbook.createSheet("Report");
//        int rownum = 0;
//        Cell cell;
//        Row row;

        ConvertModel convertModel = new ConvertModel();
        convertModel.setRootElements(elements);
        for (Element element : elements) {
            convertModel.setRow(convertModel.getSheet().createRow(convertModel.getRownum()));
//            row = sheet.createRow(rownum);
            Elements theadElement = element.getElementsByTag("thead");
            fillHead(convertModel, theadElement, element);
            convertModel.setRownum(convertModel.getRownum() + 1);
//            rownum++;
            List<Element> tbodyElements = element.getElementsByTag("tbody")
                    .stream().filter(e -> e.parent()
                            .equals(element)).collect(Collectors.<Element>toList());
            fillBody(convertModel, tbodyElements, element);
            convertModel.setRow(convertModel.getSheet().createRow(convertModel.getRownum()));
//            row = sheet.createRow(rownum);
            convertModel.setRownum(convertModel.getRownum() + 1);
//            rownum++;// смотрим залетит сюда
            System.out.println(element);
        }

        return convertModel.getWorkbook();
    }

    //Посичтать сколько ячеек необходимо взять для заголовка
    private  Integer checkCellColSpanBody(List<Element> tbodyElements, Integer position) {
        for (Element tdbodyElement : tbodyElements) {
            List<Element> tdElements = tdbodyElement.getElementsByTag("td")
                    .stream().filter(e -> e.parent().equals(tdbodyElement)).collect(Collectors.<Element>toList());
            if (position < tdElements.size())
                if (tdElements.get(position).toString().indexOf("<table") != -1 ? true : false) {
                    Element element = tdElements.get(position);
                    List<Element> subTdElement = element.getElementsByTag("td");
                    return subTdElement.size() - 1;
                }
        }

        return 1;
    }

    private  List<Element> getContentByTag(String strHtml, String tag) throws MalformedURLException {
        Document doc = Jsoup.parseBodyFragment(strHtml);
        List<Element> elements = doc.select(tag).stream().filter(e -> e.parent().parent().parent()
                .equals(doc)).collect(Collectors.<Element>toList());

        return elements;
    }

}
