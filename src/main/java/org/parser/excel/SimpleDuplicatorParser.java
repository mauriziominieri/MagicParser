package org.parser.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Data
public class SimpleDuplicatorParser<T> extends SimpleParser {
    private int lastPortionInitialIndex;
    private List<Row> rowsToDuplicate;
    int firstRowToDuplicateIndex;
    private String key;

    public void write(List<T> objectList, Direction direction, int steps, String key, int numberOfRowsToDuplicate, boolean fillDuplicateHeadersCells, int objectListPortion, String headerGroup, int headerGroupStartIndex, boolean copyStyle) throws InvocationTargetException, IllegalAccessException, ExcelException, IOException {
        if(objectList.isEmpty()) return;
        this.object = objectList.get(0);
        this.direction = direction;
        this.steps = steps;
        this.key = key;
        this.fillDuplicateHeadersCells = fillDuplicateHeadersCells; // TODO dubbi, serve veramente in questa versione SimpleDuplicatorParser?
        this.objectListPortion = objectListPortion; // TODO: devo gestirlo veramente, perchè ora come ora se setto 3 e ho una porzione da 10 voglio che riempia 3, duplichi sotto e riempie altri 3 e così via.
        this.headerGroup = headerGroup;
        this.headerGroupStartIndex = headerGroupStartIndex;
        this.copyStyle = copyStyle;

        // CREATE LIST OF ROWS TO DUPLICATE
        Workbook tempWorkbook = new XSSFWorkbook();
        Sheet tempSheet = tempWorkbook.createSheet("Sheet temporaneo");

        rowsToDuplicate = new ArrayList<>();
        firstRowToDuplicateIndex = findFirstRowToDuplicateIndex(sheet, key);
        for(int i = 0; i < numberOfRowsToDuplicate; i++) {
            Row oldRow = sheet.getRow(firstRowToDuplicateIndex + i);
            Row tempRow = tempSheet.createRow(firstRowToDuplicateIndex + i);
            for (int j = 0; j < oldRow.getLastCellNum(); j++) {
                Cell oldCell = oldRow.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                Cell newCell = tempRow.createCell(j);
                if (oldCell.getCellType() == CellType.NUMERIC) {
                    double cellValue = oldCell.getNumericCellValue();
                    newCell.setCellValue(cellValue);
                } else if (oldCell.getCellType() == CellType.STRING) {
                    String cellValue = oldCell.getStringCellValue();
                    newCell.setCellValue(cellValue);
                }
            }
            rowsToDuplicate.add(tempRow);
        }

        if(objectList.isEmpty()) return;
        List<Header> xlsxHeaders = getHeaders();
        Method[] declaredMethods = object.getClass().getDeclaredMethods(); // prendo tutti i metodi get dell'oggetto obj
        int porzionePicchetto = -1;
        for (int i = 0; i < objectList.size(); i++) {
            T obj = objectList.get(i);

            if(objectListPortion == 0)
                porzionePicchetto = 0;
            else if(i % objectListPortion == 0) // ogni n picchetti andrò alla porzione dell'excel successiva
                porzionePicchetto++;

            if(i >= objectListPortion) {
                shiftPorzione(numberOfRowsToDuplicate);
                xlsxHeaders = getHeaders();
            }
            manage(xlsxHeaders, declaredMethods, obj, 0); // TODO da gestire con objectListPortion come scritto sopra, ora come ora funziona solo con 0
        }

        tempWorkbook.close();
    }

    @Override
    protected List<Coordinate> getHeadersCoordinates(String columnTitle) {
        int rowCount = sheet.getLastRowNum();
        List<Coordinate> coordinataList = new ArrayList<>();
        for(int currentRowIndex = lastPortionInitialIndex  + 2; currentRowIndex <= rowCount; currentRowIndex++) {
            Row row = sheet.getRow(currentRowIndex);
            if (row != null) {
                for (Cell cell : row) {
                    if (CellType.STRING.equals(cell.getCellType()) && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", ""))) {
                        coordinataList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                        if(key.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", "")))
                            lastPortionInitialIndex = currentRowIndex - 1; // TODO, replaceAll toglie tutti gli spazi, devo farlo ovunque anche nel simple parser direi
                    }
                }
            }
        }
        return coordinataList;
    }

    public void shiftPorzione(int numberOfRowsToDuplicate) {
        int firstRow = lastPortionInitialIndex;
        int lastRow = firstRow + numberOfRowsToDuplicate;
        boolean isRowOfNumbers = false; // per vedere se è la riga dell'header contenente i numeri poi da incrementare alla duplicazione

        sheet.shiftRows(lastRow, sheet.getLastRowNum(), numberOfRowsToDuplicate);

        for (int j = 0; j < rowsToDuplicate.size(); j++) {
            Row copiedRow = rowsToDuplicate.get(j) == null ? sheet.createRow(j) : rowsToDuplicate.get(j);
            // non posso copiare stili tra diversi workbook quindi mi serve la riga originale da cui prenderò lo stile
            Row oldRow = rowsToDuplicate.get(j) == null ? sheet.createRow(j) : sheet.getRow(firstRow + j);
            Row newRow = sheet.createRow(firstRow + numberOfRowsToDuplicate + j);

            // copio eventuali merged region
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = sheet.getMergedRegion(i);
                if (region.getFirstRow() == copiedRow.getRowNum())
                    sheet.addMergedRegion(new CellRangeAddress(region.getFirstRow() + numberOfRowsToDuplicate, region.getLastRow() + numberOfRowsToDuplicate, region.getFirstColumn(), region.getLastColumn()));
            }

            // copio stile e valore
            for (int i = copiedRow.getFirstCellNum(); i < copiedRow.getLastCellNum(); i++) {
                Cell oldCell = copiedRow.getCell(i);
                Cell newCell = newRow.createCell(i);
                newCell.setCellStyle(oldRow.getCell(i).getCellStyle());

                if (oldCell.getCellType() == CellType.NUMERIC) {
                    double cellValue = oldCell.getNumericCellValue();
                    newCell.setCellValue(cellValue);
                    if(isRowOfNumbers)
                        newCell.setCellValue(cellValue + objectListPortion);
                } else if (oldCell.getCellType() == CellType.STRING) {
                    String cellValue = oldCell.getStringCellValue();
                    newCell.setCellValue(cellValue);
                    if(cellValue.endsWith("N.")) // TODO da rendere parametro (?), non va bene scrivere un carattere nel template, perchè poi nella parte non copiata il carattere rimarrà, credo convenga usare un parametro, per esempio N., però poi in futuro qualsiasi riga che contiene i numeri intestazione da incrementare deve terminare con N.
                        isRowOfNumbers = true;
//                    if(isRowOfNumbers)
//                        try {
//                            newCell.setCellValue(Integer.toString(Integer.valueOf(cellValue.substring(0, cellValue.length() - 2)) + objectListPortion));
//                        } catch(NumberFormatException e) { newCell.setCellValue(cellValue.substring(0, cellValue.length() - 2)); }
                }
            }
            isRowOfNumbers = false;
        }
    }

    public int findFirstRowToDuplicateIndex(Sheet sheet, String key) throws ExcelException {
        for(int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    if(CellType.STRING.equals(cell.getCellType()) && key.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", ""))) {
                        return i - 1; // TODO da gestire, dato che picchetto N non fa parte dell'oggetto in lista da inserire allora devo passare la key sostegno tipo e fare - 1, oppure mettere nPicchetto come attributo bo, semplicemente passo l'indice della prima riga della porzione da duplicare e risolvo tutto (?)
                    }
                }
            }
        }
        throw new ExcelException("La stringa " + key + " non esiste nel template");
    }
}
