package org.parser.excel;

import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

@NoArgsConstructor
@AllArgsConstructor
public class ComplexParser<T> extends MagicParser {
    private T object;
    private String symbol;

    /**
     * Permette di scrivere su file excel
     *
     * @param object        Oggetto da scrivere
     * @param symbol        Il simbolo nel template che identifica la cella da sovrascrivere (formato <symbol valore>)
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    public void write(T object, String symbol, boolean copyStyle) throws InvocationTargetException, IllegalAccessException {
        this.object = object;
        this.symbol = symbol;
        List<Header> xlsxHeaders = getHeaders();
        Method[] declaredMethods = object.getClass().getDeclaredMethods(); // prendo tutti i metodi get dell'oggetto obj
        for (Header xlsxHeader : xlsxHeaders) {
            for(Coordinate coordinate : xlsxHeader.getCoordinateList()) {   // i campi duplicati hanno una lista di dimensione > 1 dell'oggetto Coordinata
                setCellValue(coordinate, declaredMethods, xlsxHeader, object, copyStyle);
            }
        }
    }

    /**
     * Restituisce l'header dove ogni cella ha le proprie coordinate
     *
     */
    protected List<Header> getHeaders() {
        List<Header> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : object.getClass().getDeclaredFields()) { // scorro tutti gli attributi della classe dell'oggetto
            Field importField = field.getAnnotation(Field.class);
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                List<Coordinate> coordinataList = getHeadersCoordinates(xlsxColumn);
                listHeader.add(new Header(field.getName(), xlsxColumn, -1, null, coordinataList)); // TODO ho fatto questa modifica, devo capire se da problemi
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
        columnTitle = this.symbol + columnTitle;
        int rowCount = sheet.getLastRowNum();
        List<Coordinate> coordinataList = new ArrayList<>();
        for(int currentRowIndex = 0; currentRowIndex <= rowCount; currentRowIndex++) {
            Row row = sheet.getRow(currentRowIndex);
            if (row != null) {
                for (Cell cell : row) {
                    if (CellType.STRING.equals(cell.getCellType()) && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", "")))
                        coordinataList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                }
            }
        }
        return coordinataList;
    }
}
