package org.parser.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

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
public class ComplexParser<T> extends MagicParser<T> {
    T object;
    String symbol;

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
        for(java.lang.reflect.Field field : object.getClass().getDeclaredFields()) { // scorro tutti gli attributi della classe cls, il field è preso grazie alla reflection
            Field importField = field.getAnnotation(Field.class); // degli attributi nella classe prendo solo quelli taggati con @Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                List<Coordinate> coordinataList = getHeadersCoordinates(xlsxColumn);
                listHeader.add(new Header(field.getName(), xlsxColumn, -1, null, coordinataList));
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
        int rowCount = sheet.getLastRowNum();
        List<Coordinate> coordinataList = new ArrayList<>();
        for(int currentRowIndex = 0; currentRowIndex <= rowCount; currentRowIndex++) {
            Row row = sheet.getRow(currentRowIndex);
            if (row != null) {
                for (Cell cell : row) {
                    if(CellType.STRING.equals(cell.getCellType())) { // con il primo if gestisco la cella unificata (solitamente) a inizio report contenente le date REPORT AVANZAMENTO
                        if((columnTitle.equalsIgnoreCase("*avanzamentoda") || columnTitle.equalsIgnoreCase("*avanzamentoa")) && cell.getStringCellValue().replaceAll("\\s+", "").equalsIgnoreCase("reportavanzamentodal*avanzamentodaal*avanzamentoa"))
                            coordinataList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                        else if(columnTitle.equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", "")))
                            coordinataList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                    }
                }
            }
        }
        return coordinataList;
    }
}