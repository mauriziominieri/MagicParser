package org.parser.excel;

import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

@NoArgsConstructor
@AllArgsConstructor
public class SimpleParser<T> extends MagicParser {

    private List<T> objectList;
    private T object;
    private boolean fillDuplicateHeadersCells;
    private int objectListPortion;
    private Direction direction;
    private int steps;
    private String headerGroup;
    private int headerGroupStartIndex;
    private boolean copyStyle;

    /**
     * Permette di scrivere su file excel
     *
     * @param objectList                Lista di oggetti da scrivere
     * @param direction                 Direzione in cui muoversi dall'header
     * @param steps                     Salti da fare in quella direzione
     * @param fillDuplicateHeadersCells Se riempire le celle di header duplicati
     * @param objectListPortion         In quante porzioni dividere la lista di oggetti
     * @param headerGroup               Gruppo logico in cui eventualmente dividere l'header (utile per layout a scacchi)
     * @param headerGroupStartIndex     Indice da cui far partire le celle appartenenti al gruppo
     * @param copyStyle                 Se copiare lo stile della cella da sovrascrivere
     */
    public void write(List<T> objectList, Direction direction, int steps, boolean fillDuplicateHeadersCells, int objectListPortion, String headerGroup, int headerGroupStartIndex, boolean copyStyle) throws InvocationTargetException, IllegalAccessException {
        if(objectList.isEmpty())
            return;
        this.object = objectList.get(0);
        this.direction = direction;
        this.steps = steps;
        this.fillDuplicateHeadersCells = fillDuplicateHeadersCells;
        this.objectListPortion = objectListPortion;
        this.headerGroup = headerGroup;
        this.headerGroupStartIndex = headerGroupStartIndex;
        this.copyStyle = copyStyle;

        List<Header> xlsxHeaders = getHeaders();
        Method[] declaredMethods = object.getClass().getDeclaredMethods(); // prendo tutti i metodi get dell'oggetto obj
        int porzionePicchetto = -1;
        for (int i = 0; i < objectList.size(); i++) {
            T obj = objectList.get(i);

            if(objectListPortion == 0)
                porzionePicchetto = 0;
            else if(i % objectListPortion == 0) // ogni n picchetti andrò alla porzione dell'excel successiva
                porzionePicchetto++;

            manage(xlsxHeaders, declaredMethods, obj, porzionePicchetto);
        }
    }

    /**
     * Restituisce l'header dove ogni cella ha le proprie coordinate
     *
     */
    protected List<Header> getHeaders() {
        List<Header> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : object.getClass().getDeclaredFields()) { // scorro tutti gli attributi della classe cls
            Field importField = field.getAnnotation(Field.class); // creoun oggetto Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                String picchettoGroup = importField.group()[0];
                List<Coordinate> coordinateList = getHeadersCoordinates(xlsxColumn);
                listHeader.add(new Header(field.getName(), xlsxColumn, -1, picchettoGroup, coordinateList)); // TODO controllare se da problemi
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

    /**
     * Gestisce headers duplicati e porzioni
     *
     * @param xlsxHeaders       Header con tutte le informazioni delle celle
     * @param declaredMethods   I metodi dell'oggetto
     * @param obj               L'oggetto da scrivere
     * @param portionIndex      Indice della porzione in cui inserire l'oggetto
     */
    protected void manage(List<Header> xlsxHeaders, Method[] declaredMethods, T obj, int portionIndex) throws InvocationTargetException, IllegalAccessException {
        for (Header xlsxHeader : xlsxHeaders) {
            if(xlsxHeader.getCoordinateList().isEmpty()) continue;  // se la colonna del field in oggetto non esiste nell'excel lo ignoro proprio
            if(fillDuplicateHeadersCells) { // con questa struttura di if ignoro il caso fillDuplicateCells = true e objectListPortion > 0
                for (Coordinate coordinate : xlsxHeader.getCoordinateList()) {
                    if(headerGroup == null)
                        setCoordinate(obj, xlsxHeader, declaredMethods, coordinate);
                    else
                        setCoordinateChess(obj, xlsxHeader, declaredMethods, coordinate);
                }
            }
            else if(objectListPortion == 0) {
                Coordinate coordinate = xlsxHeader.getCoordinateList().get(0);
                if(headerGroup == null)
                    setCoordinate(obj, xlsxHeader, declaredMethods, coordinate);
                else
                    setCoordinateChess(obj, xlsxHeader, declaredMethods, coordinate);
            }
            else {
                if(xlsxHeader.getCoordinateList().size() <= portionIndex) continue;
                Coordinate coordinate = xlsxHeader.getCoordinateList().get(portionIndex);
                if(headerGroup == null)
                    setCoordinate(obj, xlsxHeader, declaredMethods, coordinate);
                else
                    setCoordinateChess(obj, xlsxHeader, declaredMethods, coordinate);
            }
        }
    }

    /**
     * Calcola le coordinata
     *
     * @param obj               Oggetto da scrivere
     * @param xlsxHeader        Cella con tutte le informazioni
     * @param declaredMethods   I metodi dell'oggetto
     * @param coordinate        coordinata in cui scrivere l'oggetto
     */
    public void setCoordinate(T obj, Header xlsxHeader, Method[] declaredMethods, Coordinate coordinate) throws InvocationTargetException, IllegalAccessException {
        switch (this.direction) {
            case UP:
                coordinate.setRowIndex(coordinate.getRowIndex() - this.steps);
                break;
            case RIGHT:
                coordinate.setColumnIndex(coordinate.getColumnIndex() + this.steps);
                break;
            case BOTTOM:
                coordinate.setRowIndex(coordinate.getRowIndex() + this.steps);
                break;
            case LEFT:
                coordinate.setColumnIndex(coordinate.getColumnIndex() - this.steps);
                break;
        }
        setCellValue(coordinate, declaredMethods, xlsxHeader, obj, copyStyle);
    }

    /**
     * Calcola le coordinate per il layout a scacchi
     *
     * @param obj               Oggetto da scrivere
     * @param xlsxHeader        Cella con tutte le informazioni
     * @param declaredMethods   I metodi dell'oggetto
     * @param coordinate        coordinata in cui scrivere l'oggetto
     */
    public void setCoordinateChess(T obj, Header xlsxHeader, Method[] declaredMethods, Coordinate coordinate) throws InvocationTargetException, IllegalAccessException {
        switch(this.direction) {
            case UP:
                if(xlsxHeader.getPicchettoGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                    coordinate.setRowIndex(coordinate.getRowIndex() - headerGroupStartIndex);
                coordinate.setRowIndex(coordinate.getRowIndex() - this.steps);
                setCellValue(coordinate, declaredMethods, xlsxHeader, obj, copyStyle);
                if(xlsxHeader.getPicchettoGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                    coordinate.setRowIndex(coordinate.getRowIndex() + headerGroupStartIndex);
                break;
            case RIGHT:
                if(xlsxHeader.getPicchettoGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                    coordinate.setColumnIndex(coordinate.getColumnIndex() + headerGroupStartIndex);
                coordinate.setColumnIndex(coordinate.getColumnIndex() + this.steps);
                setCellValue(coordinate, declaredMethods, xlsxHeader, obj, copyStyle);
                if(xlsxHeader.getPicchettoGroup().equals(headerGroup))
                    coordinate.setColumnIndex(coordinate.getColumnIndex() - headerGroupStartIndex);
                break;
            case BOTTOM:
                if(xlsxHeader.getPicchettoGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                    coordinate.setRowIndex(coordinate.getRowIndex() + headerGroupStartIndex);
                coordinate.setRowIndex(coordinate.getRowIndex() + this.steps);
                setCellValue(coordinate, declaredMethods, xlsxHeader, obj, copyStyle);
                if(xlsxHeader.getPicchettoGroup().equals(headerGroup))
                    coordinate.setRowIndex(coordinate.getRowIndex() - headerGroupStartIndex);
                break;
            case LEFT:
                if(xlsxHeader.getPicchettoGroup().equals(headerGroup))
                    coordinate.setColumnIndex(coordinate.getColumnIndex() - headerGroupStartIndex);
                coordinate.setColumnIndex(coordinate.getColumnIndex() - this.steps);
                setCellValue(coordinate, declaredMethods, xlsxHeader, obj, copyStyle);
                if(xlsxHeader.getPicchettoGroup().equals(headerGroup))
                    coordinate.setColumnIndex(coordinate.getColumnIndex() + headerGroupStartIndex);
                break;
        }
    }
}
