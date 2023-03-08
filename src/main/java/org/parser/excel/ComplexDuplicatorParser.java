package org.parser.excel;

import lombok.Data;
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
public class ComplexDuplicatorParser<T> extends ComplexParser {
    /* TODO Ci sono 3 possibili casi di complex parser
        1. analizzo tutto il foglio (unico dto) -> COMPLEX PARSER
        2. analizzo una porzione e passando una lista di n elementi deve duplicare la porzione n volte -> (questo è sicuramente il complex duplicator parser)
        3. analizza una porzione e passando una lista di n elementi deve semplicemente riempire le n porzioni -> potrei mettere indice colonna iniziale e finale nei duplicator parser, in questo modo per i picchetti andrei a duplicare solo loro fino ad AO, e per il riepilogo andrei ad impostare la prima colonna AP, in questo modo avrei solo 4 classi principali SIMPLE PARSER, COMPLEX PARSER e le due sottoclassi SIMPLE DUPLICATOR PARSER e COMPLEX DUPLICATOR PARSER
     */
    private int firstRowIndex;
    private int lastRowIndex;
    private String firstColumn;
    private String lastColumn;
    private int gap;
    private boolean duplicate; // TODO (true per stazioni e false per elettrAereo) devo creare una versione finale del complex duplicator parser che gestisce gli shift con delle porzioni; quindi selezionando anche le colonne
    private List<Row> rowsToDuplicate;

    /**
     * Permette di scrivere su file excel
     *
     * @param objectList    Oggetti da scrivere
     * @param symbol        Il simbolo nel template che identifica la cella da sovrascrivere (formato <symbol valore>)
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    public void write(List<T> objectList, String symbol, int firstRowIndex, int lastRowIndex, String firstColumn, String lastColumn, int gap, boolean duplicate, boolean copyStyle) throws InvocationTargetException, IllegalAccessException, IOException {
        this.symbol = symbol;
        this.firstRowIndex = firstRowIndex;
        this.lastRowIndex = lastRowIndex;
        this.firstColumn = firstColumn;
        this.lastColumn = lastColumn;
        this.gap = gap;
        this.duplicate = duplicate;

        Workbook tempWorkbook = new XSSFWorkbook();
        if(duplicate)
            setRowsToDuplicate(tempWorkbook);   // salvo le rows da duplicare

        for(int i = 0; i < objectList.size(); i++) {
            this.object = objectList.get(i);
            List<Header> xlsxHeaders = getHeaders();
            Method[] declaredMethods = object.getClass().getDeclaredMethods(); // prendo tutti i metodi get dell'oggetto obj

            if(duplicate && objectList.size() > i + 1) // duplico solo se viene specificato e per tutti gli elementi della lista tranne l'ultimo
                shiftPorzione(i + 1);

            for (Header xlsxHeader : xlsxHeaders) {
                for (Coordinate coordinate : xlsxHeader.getCoordinateList()) {   // i campi duplicati hanno una lista di dimensione > 1 dell'oggetto Coordinata
                    setCellValue(coordinate, declaredMethods, xlsxHeader, object, copyStyle);
                }
            }
            this.firstRowIndex = this.lastRowIndex + gap + 1;
            this.lastRowIndex += lastRowIndex - firstRowIndex + gap + 1;
        }

        tempWorkbook.close();
    }

    /**
     * Restituisce l'header dove ogni cella ha le proprie coordinate
     *
     */
    protected List<Header> getHeaders() {
        List<Header> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : object.getClass().getDeclaredFields()) { // scorro tutti gli attributi della classe cls, il field è preso grazie alla reflection
            Field importField = field.getAnnotation(Field.class); // degli attributi nella classe prendo solo quelli taggati con @Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                List<Coordinate> coordinataList = getHeadersCoordinates(xlsxColumn);
                listHeader.add(new Header(field.getName(), xlsxColumn, -1,null, coordinataList));
            }
        }
        return listHeader;
    }

    /**
     * Cerca le celle poi da sovrascrivere (di formato <symbol VALORE>) e salva le relative coordinate
     *
     * @param columnTitle   Titolo della colonna dell'header da cercare
     */
    protected List<Coordinate> getHeadersCoordinates(String columnTitle) {
        columnTitle = (this.symbol + columnTitle).replaceAll("\\s+", ""); // elimino tutti gli spazi
        List<Coordinate> coordinataList = new ArrayList<>();
        for(int i = firstRowIndex; i <= lastRowIndex; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    if(CellType.STRING.equals(cell.getCellType())) { // con il primo if gestisco la cella unificata (solitamente) a inizio report contenente le date REPORT AVANZAMENTO
                        if((columnTitle.equals("*avanzamentoda") || columnTitle.equals("*avanzamentoa")) && cell.getStringCellValue().replaceAll("\\s+", "").equalsIgnoreCase("reportavanzamentodal*avanzamentodaal*avanzamentoa"))
                            coordinataList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                        else if(columnTitle.equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", "")))
                            coordinataList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                    }
                }
            }
        }
        return coordinataList;
    }

    public void shiftPorzione(int k) {
        int a = lastRowIndex + gap + 1;
        int b = sheet.getLastRowNum();
        int numberOfRowsToDuplicate = lastRowIndex - firstRowIndex + 1;
        int n = numberOfRowsToDuplicate + gap;
        sheet.shiftRows(a, b, n); // shifta la porzione (a, b) di n righe sotto.

//        selectRange(sheet, 1, 10, 2, 5); // seleziona le righe 2-11 (indici 1-10) e le colonne C-F (indici 2-5)
//        shiftRange(sheet, 1, 10, 2, 5, 7); // sposta la porzione selezionata di 7 righe in basso

        for (int i = 0; i < rowsToDuplicate.size(); i++) {
            Row copiedRow = rowsToDuplicate.get(i) == null ? sheet.createRow(i) : rowsToDuplicate.get(i);
            // non posso copiare stili tra diversi workbook quindi mi serve la riga originale da cui prenderò lo stile
            Row oldRow = rowsToDuplicate.get(i) == null ? sheet.createRow(i) : sheet.getRow(firstRowIndex + i);
            Row newRow = sheet.createRow(a + i);

            // copio eventuali merged region
            for (int j = 0; j < sheet.getNumMergedRegions(); j++) {
                CellRangeAddress region = sheet.getMergedRegion(j);
                if (region.getFirstRow() == copiedRow.getRowNum())
                    sheet.addMergedRegion(new CellRangeAddress(region.getFirstRow() + n * k, region.getLastRow() + n * k, region.getFirstColumn(), region.getLastColumn())); // + n perchè la porzione copiata deve essere spostata n celle sotto
            }

            // copio stile e valore
            for (int j = copiedRow.getFirstCellNum(); j < copiedRow.getLastCellNum(); j++) {
                Cell oldCell = copiedRow.getCell(j);
                Cell newCell = newRow.createCell(j);
                newCell.setCellStyle(oldRow.getCell(j).getCellStyle());
                if (oldCell.getCellType() == CellType.STRING) {
                    String cellValue = oldCell.getStringCellValue();
                    newCell.setCellValue(cellValue);
                }
            }
        }

//        firstRowToDuplicateIndex = a;
//        lastRowRoDuplicateIndex = a + numberOfRowsToDuplicate - 1;
    }

    /**
     * Salva le righe da duplicare
     *
     * @param tempWorkbook
     */
    public void setRowsToDuplicate(Workbook tempWorkbook) {
        Sheet tempSheet = tempWorkbook.createSheet("Sheet temporaneo");
        rowsToDuplicate = new ArrayList<>();
        for(int i = 0; i <= lastRowIndex - firstRowIndex; i++) {
            Row oldRow = sheet.getRow(firstRowIndex + i);
            Row tempRow = tempSheet.createRow(firstRowIndex + i);
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
    }

    public static int getColumnIndex(String columnName) {
        int columnIndex = 0;
        for (int i = 0; i < columnName.length(); i++) {
            columnIndex = columnIndex * 26 + (columnName.charAt(i) - 'A' + 1);
        }
//        CellReference.convertColStringToIndex("AP"); forse serve semplicemente questo
        return columnIndex;
    }

    public static void shiftRange(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, int shiftRows) {
        CellRangeAddress range = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.shiftRows(firstRow + shiftRows, lastRow + shiftRows, shiftRows, true, true); // sposta le righe di 7 posizioni in basso
    }

    public static void selectRange(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        CellRangeAddress range = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.setAutoFilter(range); // Opzionale: applica un filtro alla porzione selezionata
    }

}
