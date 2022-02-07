package model;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.jsoup.nodes.Element;

import java.util.List;

public class ConvertModel {
    private HSSFWorkbook workbook = null;
    private HSSFSheet sheet = null;
    private int rownum = 0;
    private Cell cell;
    private Row row;
    List<Element> rootElements;

    public ConvertModel() {
        workbook = new HSSFWorkbook();
        sheet = workbook.createSheet("Report");
    }

    public HSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(HSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public HSSFSheet getSheet() {
        return sheet;
    }

    public void setSheet(HSSFSheet sheet) {
        this.sheet = sheet;
    }

    public int getRownum() {
        return rownum;
    }

    public void setRownum(int rownum) {
        this.rownum = rownum;
    }

    public Cell getCell() {
        return cell;
    }

    public void setCell(Cell cell) {
        this.cell = cell;
    }

    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    public List<Element> getRootElements() {
        return rootElements;
    }

    public void setRootElements(List<Element> rootElements) {
        this.rootElements = rootElements;
    }
}
