package org.parser.excel;

import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.parser.utils.PropertiesUtils;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

@NoArgsConstructor
@AllArgsConstructor
public class XLSXSimpleParser<T> extends XLSXParser2<T> {

    private XLSXSimpleTemplateSetting.Direction direction;

    private int steps;

    private String headerGroup;

    private int headerGroupStartIndex;

    private int objectListPortion;

    private boolean fillDuplicateCells;

    /**
     * Permette di scrivere su file excel
     *
     * @param objectClass   Classe dell'oggetto da scrivere
     * @param objectList    Lista di oggetti da scrivere
     * @param sheetIndex    Numero del foglio da analizzare, -1 per analizzare tutti i fogli dell'excel
     * @param direction     Direzione in cui muoversi dall'header
     * @param steps         Salti da fare in quella direzione
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */ // TODO devo fare in modo che tutte le configurazioni, come steps, direction, objectListPortionecc siano attributi della classe, non è molto furbo passarli come parametri dei metodi, in realtà potrebbe essere furbo in quanto avrei vari metodi overloadati
    public void write(int sheetIndex, Class objectClass, List<T> objectList, int objectListPortion, boolean fillDuplicateCells, XLSXSimpleTemplateSetting.Direction direction, int steps, boolean copyStyle) throws IOException, ExcelException {
        this.direction = direction;
        this.steps = steps;
        this.objectListPortion = objectListPortion;
        this.fillDuplicateCells = fillDuplicateCells;
        mela(sheetIndex, objectClass, objectList, objectListPortion, copyStyle);
    }

    /**
     * Permette di scrivere su file excel gestendo un layout "a scacchi"
     *
     * @param objectClass   Classe dell'oggetto da scrivere
     * @param objectList    Lista di oggetti da scrivere
     * @param sheetIndex    Numero del foglio da analizzare, -1 per analizzare tutti i fogli dell'excel
     * @param direction     Direzione in cui muoversi dall'header
     * @param steps         Salti da fare in quella direzione
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     * @param headerGroup
     * @param headerGroupStartIndex
     */
    public void write(int sheetIndex, Class objectClass, List<T> objectList, int objectListPortion, boolean fillDuplicateCells, XLSXSimpleTemplateSetting.Direction direction, int steps, boolean copyStyle, String headerGroup, int headerGroupStartIndex) throws IOException, ExcelException {
        this.direction = direction;
        this.steps = steps;
        this.objectListPortion = objectListPortion;
        this.fillDuplicateCells = fillDuplicateCells;
        mela(sheetIndex, objectClass, objectList, objectListPortion, copyStyle, headerGroup, headerGroupStartIndex);
    }

    /**
     * Restituisce l'header dove ogni cella ha le proprie coordinate a due dimensioni
     * Analizza tutti gli attributi annotati con @Field della classe objectClass
     * Dato che una parte del template è "a scacchi" ho dovuto differenziare alcuni field da altri
     *
     * @param objectClass   Classe dell'oggetto da scrivere
     * @param sheet         Foglio dell'excel da analizzare
     */
    @Override
    protected List<Header2> modelObjectToXLSXHeaderForWrite2(Class<?> objectClass, Sheet sheet) {
        List<Header2> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : objectClass.getDeclaredFields()) { // scorro tutti gli attributi della classe cls
            Field importField = field.getAnnotation(Field.class); // creoun oggetto Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                String picchettoGroup = importField.group()[0];
                List<Coordinata> coordinataList = findColumnCoordinates2(xlsxColumn, sheet);
                listHeader.add(new Header2(field.getName(), xlsxColumn, -1, picchettoGroup, coordinataList)); // TODO controllare se da problemi
            }
        }
        return listHeader;
    }

    /**
     * Scrive data la lista Headers e la lista di oggetti
     *
     * @param objectClass   Classe dell'oggetto da scrivere
     * @param sheet         Foglio dell'excel da analizzare
     * @param objectList    Lista di oggetti da scrivere
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    @Override
    public void mela2(Sheet sheet, Class<T> objectClass, List<T> objectList, int objectListPortion, boolean copyStyle) throws ExcelException, IOException {
        List<Header2> xlsxHeaders = modelObjectToXLSXHeaderForWrite2(objectClass, sheet);
        int porzionePicchetto = -1;
        for (int i = 0; i < objectList.size(); i++) {
            T obj = objectList.get(i);

            if(objectListPortion == 0)
                porzionePicchetto = 0;
            else if(i % objectListPortion == 0) // ogni n picchetti andrò alla porzione dell'excel successiva
                porzionePicchetto++;

            try {
                prova5(xlsxHeaders, sheet, obj, copyStyle, porzionePicchetto);
            } catch (Exception e) {
                workbook.close();
                e.printStackTrace();
                throw new ExcelException(PropertiesUtils.getMessage("message.excel.sheetName", new Object[]{sheet.getSheetName()}));
            }
        }
    }

    /**
     * Scrive data la lista Headers e la lista di oggetti gestisce "a scacchi"
     *
     * @param objectClass       Classe dell'oggetto da scrivere
     * @param sheet             Foglio dell'excel da analizzare
     * @param objectList        Lista di oggetti da scrivere
     * @param copyStyle         Se copiare lo stile della cella da sovrascrivere
     * @param objectListPortion Divide la lista di oggetti in porzioni
     */
    public void mela2(Sheet sheet, Class<T> objectClass, List<T> objectList, int objectListPortion, boolean copyStyle, String headerGroup, int headerGroupStartIndex) throws IOException, ExcelException {
        List<Header2> xlsxHeaders = modelObjectToXLSXHeaderForWrite2(objectClass, sheet);
        int porzionePicchetto = -1;
        for (int i = 0; i < objectList.size(); i++) {
            T obj = objectList.get(i);

            if(objectListPortion == 0)
                porzionePicchetto = 0;
            else if(i % objectListPortion == 0) // ogni n picchetti andrò alla porzione dell'excel successiva
                porzionePicchetto++;

            try {
                prova6(xlsxHeaders, sheet, obj, copyStyle, porzionePicchetto, headerGroup, headerGroupStartIndex);
            } catch (Exception e) {
                workbook.close();
                e.printStackTrace();
                throw new ExcelException(PropertiesUtils.getMessage("message.excel.sheetName", new Object[]{sheet.getSheetName()}));
            }
        }
    }

    /**
     * Cerca e restituisce una lista di tutte le coordinate 2D delle celle dell'header
     *
     * @param columnTitle   Titolo della colonna dell'header
     * @param sheet         Foglio dell'excel da analizzare
     */
    @Override
    protected List<Coordinata> findColumnCoordinates2(String columnTitle, Sheet sheet) {
        int rowCount = sheet.getLastRowNum();
        List<Coordinata> coordinataList = new ArrayList<>();
        for(int currentRowIndex = 0; currentRowIndex <= rowCount; currentRowIndex++) {
            Row row = sheet.getRow(currentRowIndex);
            if (row != null) {
                for (Cell cell : row) {
                    if (CellType.STRING.equals(cell.getCellType()) && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", "")))
                        coordinataList.add(new Coordinata(cell.getRowIndex(), cell.getColumnIndex()));
                }
            }
        }
        return coordinataList;
    }

    /**
     * Scrive le celle in base alla direzione e gli steps
     *
     * @param xlsxHeaders       Titolo della colonna dell'header
     * @param sheet             Foglio dell'excel da analizzare
     * @param obj               Oggetto da scrivere
     * @param copyStyle         Se copiare lo stile della cella da sovrascrivere
     * @param porzionePicchetto Nel template in questione i picchetti vanno presi a porzione, 10 alla volta
     */
    @Override
    protected void prova5(List<Header2> xlsxHeaders, Sheet sheet, T obj, boolean copyStyle, int porzionePicchetto) throws Exception {
        Method[] declaredMethods = obj.getClass().getDeclaredMethods(); // prendo tutti i metodi get dell'oggetto obj
        for (Header2 xlsxHeader : xlsxHeaders) {
            if(xlsxHeader.getCoordinataList().isEmpty()) continue;  // se la colonna del field in oggetto non esiste nell'excel lo ignoro proprio
            if(fillDuplicateCells) { // con questa struttura di if ignoro il caso fillDuplicateCells = true e objectListPortion > 0
                for (Coordinata coordinata : xlsxHeader.getCoordinataList())
                    dareUnNome(sheet, obj, xlsxHeader, declaredMethods, copyStyle, coordinata);
            }
            else if(objectListPortion == 0) {
                Coordinata coordinata = xlsxHeader.getCoordinataList().get(0);
                dareUnNome(sheet, obj, xlsxHeader, declaredMethods, copyStyle, coordinata);
            }
            else {
                if(xlsxHeader.getCoordinataList().size() <= porzionePicchetto) continue;
                Coordinata coordinata = xlsxHeader.getCoordinataList().get(porzionePicchetto);
                dareUnNome(sheet, obj, xlsxHeader, declaredMethods, copyStyle, coordinata);
            }
        }
    }

    public void dareUnNome(Sheet sheet, T obj, Header2 xlsxHeader, Method[] declaredMethods, boolean copyStyle, Coordinata coordinata) throws InvocationTargetException, IllegalAccessException {
        switch (this.direction) {
            case UP:
                coordinata.setRiga(coordinata.getRiga() - this.steps);
                break;
            case RIGHT:
                coordinata.setColonna(coordinata.getColonna() + this.steps);
                break;
            case BOTTOM:
                coordinata.setRiga(coordinata.getRiga() + this.steps);
                break;
            case LEFT:
                coordinata.setColonna(coordinata.getColonna() - this.steps);
                break;
        }
        setCellValue(coordinata, sheet, declaredMethods, xlsxHeader, obj, copyStyle);
    }

    /**
     * Scrive le celle in base alla direzione e gli steps, gestisce il layout a scacchiera specificando un gruppo e la sua partenza
     *
     * @param xlsxHeaders       Titolo della colonna dell'header
     * @param sheet             Foglio dell'excel da analizzare
     * @param obj               Oggetto da scrivere
     * @param copyStyle         Se copiare lo stile della cella da sovrascrivere
     * @param porzionePicchetto Nel template in questione i picchetti vanno presi a porzione, 10 alla volta
     */
    protected void prova6(List<Header2> xlsxHeaders, Sheet sheet, T obj, boolean copyStyle, int porzionePicchetto, String headerGroup, int headerGroupStartIndex) throws Exception {
        Method[] declaredMethods = obj.getClass().getDeclaredMethods(); // prendo tutti i metodi get dell'oggetto obj

        for (Header2 xlsxHeader : xlsxHeaders) {
            if(xlsxHeader.getCoordinataList().isEmpty()) continue;
            Coordinata coordinata = xlsxHeader.getCoordinataList().get(porzionePicchetto);
            switch(this.direction) {
                case UP:
                    if(xlsxHeader.getPicchettoGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                        coordinata.setRiga(coordinata.getRiga() - headerGroupStartIndex);
                    coordinata.setRiga(coordinata.getRiga() - this.steps);
                    setCellValue(coordinata, sheet, declaredMethods, xlsxHeader, obj, copyStyle);
                    if(xlsxHeader.getPicchettoGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                        coordinata.setRiga(coordinata.getRiga() + headerGroupStartIndex);
                    break;
                case RIGHT:
                    if(xlsxHeader.getPicchettoGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                        coordinata.setColonna(coordinata.getColonna() + headerGroupStartIndex);
                    coordinata.setColonna(coordinata.getColonna() + this.steps);
                    setCellValue(coordinata, sheet, declaredMethods, xlsxHeader, obj, copyStyle);
                    if(xlsxHeader.getPicchettoGroup().equals(headerGroup))
                        coordinata.setColonna(coordinata.getColonna() - headerGroupStartIndex);
                    break;
                case BOTTOM:
                    if(xlsxHeader.getPicchettoGroup().equals(headerGroup)) // design "a scacchi" le celle del gruppo montaggi devono essere riempite subito dopo (celle dispari), quelle della tesatura devono fare un salto (celle pari)
                        coordinata.setRiga(coordinata.getRiga() + headerGroupStartIndex);
                    coordinata.setRiga(coordinata.getRiga() + this.steps);
                    setCellValue(coordinata, sheet, declaredMethods, xlsxHeader, obj, copyStyle);
                    if(xlsxHeader.getPicchettoGroup().equals(headerGroup))
                        coordinata.setRiga(coordinata.getRiga() - headerGroupStartIndex);
                    break;
                case LEFT:
                    if(xlsxHeader.getPicchettoGroup().equals(headerGroup))
                        coordinata.setColonna(coordinata.getColonna() - headerGroupStartIndex);
                    coordinata.setColonna(coordinata.getColonna() - this.steps);
                    setCellValue(coordinata, sheet, declaredMethods, xlsxHeader, obj, copyStyle);
                    if(xlsxHeader.getPicchettoGroup().equals(headerGroup))
                        coordinata.setColonna(coordinata.getColonna() + headerGroupStartIndex);
                    break;
            }
        }
    }
}
