package org.parser.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Data
public class SimpleParser<T> extends MagicParser {

    List<T> objectList; // package-private per comodità nella sottoclasse
    T object;
    boolean fillDuplicateHeadersCells;
    int objectListPortion;
    Direction direction;
    int steps;
    String headerGroup;
    int headerGroupStartIndex;
    boolean copyStyle;
    int maxIndex;

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
        this.maxIndex = getMaxIndex();

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
     * Restituisce l'indice massimo dei campi duplicati ma appartenenti a gruppi diversi
     */
    public int getMaxIndex() {
        int max = 0;
        for(java.lang.reflect.Field field : object.getClass().getDeclaredFields()) { // scorro tutti gli attributi della classe cls, il field è preso grazie alla reflection
            Field importField = field.getAnnotation(Field.class); // degli attributi nella classe prendo solo quelli taggati con @Field
            if(importField != null && importField.index() > max)
                max = importField.index();
        }
        return max;
    }

    /**
     * Restituisce l'header dove ogni cella ha le proprie coordinate
     */
    public List<Header> getHeaders() {
        List<Header> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : object.getClass().getDeclaredFields()) { // scorro tutti gli attributi della classe cls, il field è preso grazie alla reflection
            Field importField = field.getAnnotation(Field.class); // degli attributi nella classe prendo solo quelli taggati con @Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DA @Field
                String group = importField.group()[0];
                int index = importField.index();
                List<Coordinate> coordinateList = getHeadersCoordinates(xlsxColumn, index);
                listHeader.add(new Header(field.getName(), xlsxColumn, -1, group, coordinateList));
            }
        }
        return listHeader;
    }

    /**
     * Cerca le celle poi da sovrascrivere e salva le relative coordinate
     *
     * @param columnTitle   Titolo della colonna dell'header da cercare
     * @param index         Indice del campo (utile quando ci sono duplicati)
     */
    public List<Coordinate> getHeadersCoordinates(String columnTitle, int index) {
        int rowCount = sheet.getLastRowNum(), i = 1, n = index == maxIndex ? 0 : index;
        List<Coordinate> coordinataList = new ArrayList<>();
        for(int currentRowIndex = 0; currentRowIndex <= rowCount; currentRowIndex++) {
            Row row = sheet.getRow(currentRowIndex);
            if (row != null) {
                for (Cell cell : row) {
                    if (CellType.STRING.equals(cell.getCellType()) && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", ""))) {
                        if(index == -1 || i % maxIndex == n)  // gestisco il caso di campi duplicati ma appartenenti a gruppi diversi
                            coordinataList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                        i++;
                    }
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
    public void manage(List<Header> xlsxHeaders, Method[] declaredMethods, T obj, int portionIndex) throws InvocationTargetException, IllegalAccessException {
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
                if(xlsxHeader.getGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                    coordinate.setRowIndex(coordinate.getRowIndex() - headerGroupStartIndex);
                coordinate.setRowIndex(coordinate.getRowIndex() - this.steps);
                setCellValue(coordinate, declaredMethods, xlsxHeader, obj, copyStyle);
                if(xlsxHeader.getGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                    coordinate.setRowIndex(coordinate.getRowIndex() + headerGroupStartIndex);
                break;
            case RIGHT:
                if(xlsxHeader.getGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                    coordinate.setColumnIndex(coordinate.getColumnIndex() + headerGroupStartIndex);
                coordinate.setColumnIndex(coordinate.getColumnIndex() + this.steps);
                setCellValue(coordinate, declaredMethods, xlsxHeader, obj, copyStyle);
                if(xlsxHeader.getGroup().equals(headerGroup))
                    coordinate.setColumnIndex(coordinate.getColumnIndex() - headerGroupStartIndex);
                break;
            case BOTTOM:
                if(xlsxHeader.getGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                    coordinate.setRowIndex(coordinate.getRowIndex() + headerGroupStartIndex);
                coordinate.setRowIndex(coordinate.getRowIndex() + this.steps);
                setCellValue(coordinate, declaredMethods, xlsxHeader, obj, copyStyle);
                if(xlsxHeader.getGroup().equals(headerGroup))
                    coordinate.setRowIndex(coordinate.getRowIndex() - headerGroupStartIndex);
                break;
            case LEFT:
                if(xlsxHeader.getGroup().equals(headerGroup))
                    coordinate.setColumnIndex(coordinate.getColumnIndex() - headerGroupStartIndex);
                coordinate.setColumnIndex(coordinate.getColumnIndex() - this.steps);
                setCellValue(coordinate, declaredMethods, xlsxHeader, obj, copyStyle);
                if(xlsxHeader.getGroup().equals(headerGroup))
                    coordinate.setColumnIndex(coordinate.getColumnIndex() + headerGroupStartIndex);
                break;
        }
    }
}
