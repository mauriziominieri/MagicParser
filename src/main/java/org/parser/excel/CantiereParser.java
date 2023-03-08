package org.parser.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.parser.exception.ExcelException;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Data
public class CantiereParser<T> extends MagicParser {
    private int lastSubappaltatoreFieldsIndex;

    private List<Row> rowsToDuplicate;

    private Row rowToCopyStyle;

    int firstRowToDuplicateIndex;

    int numberOfRowsToDuplicate = 3; // TODO da parametrizzare

    public void write(Class<T> objectClass, List<T> objectList, String key, boolean copyStyle) throws InvocationTargetException, IllegalAccessException, IOException, ExcelException {

        Workbook tempWorkbook = new XSSFWorkbook();
        Sheet tempSheet = tempWorkbook.createSheet("Sheet temporaneo");
        firstRowToDuplicateIndex = findFirstRowToDuplicateIndex(sheet, key);
        rowsToDuplicate = new ArrayList<>();
        for(int i = 0; i < numberOfRowsToDuplicate; i++) {
            Row oldRow = sheet.getRow(firstRowToDuplicateIndex + i);
            Row tempRow = tempSheet.createRow(firstRowToDuplicateIndex + i);
            for (int j = 0; j < oldRow.getLastCellNum(); j++) {
                Cell oldCell = oldRow.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                Cell newCell = tempRow.createCell(j);
                newCell.setCellValue(oldCell.getStringCellValue());
            }
            rowsToDuplicate.add(tempRow);
        }

        // così vado a salvarmi le righe orginali, quindi se le modifico si modificano anche nell'arraylist
//        firstRowToDuplicateIndex = findFirstRowToDuplicateIndex(sheet, key);
//        rowsToDuplicate = new ArrayList<>();
//        for(int i = 0; i < numberOfRowsToDuplicate; i++) {
//            rowsToDuplicate.add(sheet.getRow(firstRowToDuplicateIndex + i)); // TODO l'indice deve essere calcolato in quache modo
//        }

        // TODO: do per scontato che la riga da cui copiare lo stile sia quella subito sotto la porzione duplicata
        if(copyStyle)
            rowToCopyStyle = sheet.getRow(firstRowToDuplicateIndex + numberOfRowsToDuplicate);

        // TODo creare una funzione per calcolare tipo lista
        Class<?> elementListClass = null;
        for(int j = 0; j < objectList.size(); j++) {
            List<T> personaleList = getListAttribute(objectList.get(j));
            if(!personaleList.isEmpty()) {
                elementListClass = personaleList.get(0).getClass();
                break;
            }
        }

        for(int j = 0; j < objectList.size(); j++) {
            List<T> personaleList = getListAttribute(objectList.get(j));
            T object = getObjectAttribute(objectList.get(j), objectClass); // todo: DEVO PRENDERE L'OGGETTO ContrattiAppaltatoriDto


            List<Header> xlsxHeadersFirstObject = findComplexCoordinates1(object.getClass(), sheet, key);
            Method[] declaredMethods = objectClass.getDeclaredMethods(); // prendo tutti i metodi get dell'oggetto obj
            for (Header xlsxHeader : xlsxHeadersFirstObject) {
                for(Coordinate coordinate : xlsxHeader.getCoordinateList()) {   // i campi duplicati hanno una lista di dimensione > 1 dell'oggetto Coordinate
                    setCellValue(coordinate, declaredMethods, xlsxHeader, object, copyStyle);
                }
            }



            List<Header> xlsxHeaders = modelObjectToXLSXHeaderForWrite2(elementListClass, sheet, key);
            for (int i = 0; i < personaleList.size(); i++) {
                T obj = personaleList.get(i);

                try {
                    riempiPersonaleSubappaltatore(sheet, xlsxHeaders, obj, i, copyStyle);
                }
                catch (Exception e) {
                    workbook.close();
                    e.printStackTrace();
                }
            }
            if(j < objectList.size() - 1)
                shiftPorzione(sheet, (j == 0 && personaleList.isEmpty()) ? 1 : personaleList.size()); // gestisco il caso in cui ci siano più oggetti e il primo oggetto ha una lista vuota
        }



        tempWorkbook.close();
    }

    protected List<Coordinate> findComplexCoordinates2(String columnTitle, Sheet sheet) {
        columnTitle = "*" + columnTitle; // TODO devo mettere il symbol
        int rowCount = sheet.getLastRowNum();
        List<Coordinate> coordinateList = new ArrayList<>();
        for(int currentRowIndex = 0; currentRowIndex <= rowCount; currentRowIndex++) {
            Row row = sheet.getRow(currentRowIndex);
            if (row != null) {
                for (Cell cell : row) {
                    if (CellType.STRING.equals(cell.getCellType()) && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", "")))
                        coordinateList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                }
            }
        }
        return coordinateList;
    }

    protected List<Header> findComplexCoordinates1(Class<?> objectClass, Sheet sheet, String key) {
        List<Header> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : objectClass.getDeclaredFields()) { // scorro tutti gli attributi della classe cls, il field è pereso grazie alla reflection
            Field importField = field.getAnnotation(Field.class); // di quelli nella classe prendo solo quelli taggati con @Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                String group = importField.group()[0];
                List<Coordinate> coordinateList = findComplexCoordinates2(xlsxColumn, sheet);
                listHeader.add(new Header(field.getName(), xlsxColumn, -1, group, coordinateList));
            }
        }
        return listHeader;
    }

    public List<Coordinate> getSubappaltatoreFieldsCoordinates(String columnTitle, Sheet sheet, String key) {
        List<Coordinate> coordinateList = new ArrayList<>();
        for(int currentRowIndex = lastSubappaltatoreFieldsIndex; currentRowIndex <= sheet.getLastRowNum(); currentRowIndex++) {
            Row row = sheet.getRow(currentRowIndex);
            Row keyRow = sheet.getRow(currentRowIndex - 2); // TODO renderlo parametro
            if (row != null && keyRow != null) {
                for (Cell cell : row) {
                    Cell keyCell = keyRow.getCell(0);
                    if (keyCell != null && CellType.STRING.equals(cell.getCellType())
                            && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", ""))
                            && key.equalsIgnoreCase(keyCell.getStringCellValue().replaceAll("\\s+", ""))
                    ) {
                        lastSubappaltatoreFieldsIndex = currentRowIndex;
                        coordinateList.add(0, new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                    }
                }
            }
        }
        return coordinateList;
    }

    public int findFirstRowToDuplicateIndex(Sheet sheet, String key) throws ExcelException {
        for(int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    if(CellType.STRING.equals(cell.getCellType()) && key.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", ""))) {
                        return i;
                    }
                }
            }
        }
        throw new ExcelException("La stringa " + key + " non esiste nel template");
    }

    protected List<Header> modelObjectToXLSXHeaderForWrite2(Class<?> objectClass, Sheet sheet, String key) {
        List<Header> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : objectClass.getDeclaredFields()) { // scorro tutti gli attributi della classe cls
            Field importField = field.getAnnotation(Field.class); // creo un oggetto Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                String group = importField.group()[0];
                List<Coordinate> coordinateList = getSubappaltatoreFieldsCoordinates(xlsxColumn, sheet, key);
                listHeader.add(new Header(field.getName(), xlsxColumn, -1, group, coordinateList)); // TODO controllare se da errore
            }
        }
        return listHeader;
    }

    // recupera l'attributo lista da un object generico
    public static <T> List<T> getListAttribute(T object) {
        java.lang.reflect.Field[] fields = object.getClass().getDeclaredFields();
        for (java.lang.reflect.Field field : fields) {
            if (List.class.isAssignableFrom(field.getType())) {
                field.setAccessible(true);
                try {
                    return (List<T>) field.get(object);
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                    return null;
                }
            }
        }
        return null;
    }

    public static <T> T getObjectAttribute(T object, Class cls) {
        try {
            java.lang.reflect.Field[] fields = object.getClass().getDeclaredFields();
            for (java.lang.reflect.Field field : fields) {
                if (field.getType().equals(cls)) {
                    field.setAccessible(true);
                    return (T) field.get(object);
                }
            }
            throw new NoSuchFieldException();
        } catch (IllegalAccessException | NoSuchFieldException e) {
            throw new RuntimeException(e);
        }
    }

    public void shiftPorzione(Sheet sheet, int nPosti) {
        int firstRow = lastSubappaltatoreFieldsIndex - 2; // TODO renderlo parametro
        int lastRow = firstRow + 2; // TODO renderlo parametro
        for (int j = 0; j < rowsToDuplicate.size(); j++) {
            Row copiedRow = rowsToDuplicate.get(j) == null ? sheet.createRow(j) : rowsToDuplicate.get(j);
            // non posso copiare stili tra diversi workbook quindi mi serve la riga originale da cui prenderò lo stile
            Row oldRow = rowsToDuplicate.get(j) == null ? sheet.createRow(j) : sheet.getRow(firstRow + j);
            Row newRow = sheet.createRow(lastRow + nPosti + j + 1);

            // copio eventuali merged region
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = sheet.getMergedRegion(i);
                if (region.getFirstRow() == copiedRow.getRowNum())
                    sheet.addMergedRegion(new CellRangeAddress(newRow.getRowNum(), newRow.getRowNum(), region.getFirstColumn(), region.getLastColumn()));
            }

            // copio stile e valore
            for (int i = copiedRow.getFirstCellNum(); i < copiedRow.getLastCellNum(); i++) {
                Cell oldCell = copiedRow.getCell(i);
                Cell newCell = newRow.createCell(i);
                newCell.setCellStyle(oldRow.getCell(i).getCellStyle());
                newCell.setCellValue(oldCell.getStringCellValue());
            }
        }
    }

    public void shift(Sheet sheet, int firstRow, int nPosti) {
        sheet.shiftRows(firstRow, sheet.getLastRowNum(), nPosti);
    }

//    public boolean isEmpty() { // TODO row is empty
//
//    }

    public void riempiPersonaleSubappaltatore(Sheet sheet, List<Header> xlsxHeaders, T obj, int i, boolean copyStyle) throws InvocationTargetException, IllegalAccessException {
        Method[] declaredMethods = obj.getClass().getDeclaredMethods();
        for (int k = 0; k < xlsxHeaders.size(); k++) { // itera per quante sono le celle della RIGA
            Header xlsxHeader = xlsxHeaders.get(k);
            if(xlsxHeader.getCoordinateList().isEmpty()) continue; // se c'è un @Field che non è presente nell'excel questo controllo previene l'indexOutOfBoundsException
            Coordinate coordinate = xlsxHeader.getCoordinateList().get(0);
            int rowIndex = coordinate.getRowIndex() + 1 + i; //todo gestirlo con la direzione

            // se la riga esiste allora la prendo, altrimenti la creo
            Row row = sheet.getRow(rowIndex) != null ? sheet.getRow(rowIndex) : sheet.createRow(rowIndex);
            // se siamo alla prima cella ed e ha un contenuto allora shifto tutta la porzione sottostante di uno step in basso TODO: da gestire con la direzione
            if (k == 0 && row.getCell(0) != null && row.getCell(0).getCellType() != CellType.BLANK) {
                shift(sheet, rowIndex, 1);
                row = sheet.createRow(rowIndex);
            }

            Cell cell = row == rowToCopyStyle ? row.getCell(coordinate.getColumnIndex()) : row.createCell(coordinate.getColumnIndex());

            if(copyStyle) {
                Cell oldCell = rowToCopyStyle.getCell(coordinate.getColumnIndex());
                cell.setCellStyle(oldCell.getCellStyle());
            }

            String field = xlsxHeader.getFieldName(); // prendo il campo @Field, es. nomeCantiere
            Optional<Method> getter = Arrays.stream(declaredMethods)
                    .filter(method -> isGetterMethod(field, method))
                    .findFirst(); // prendo il metodo get dell'attributo @Field, es. getNomeCantiere()
            if (getter.isPresent()) {
                Method getMethod = getter.get();    // es. getNomeCantiere()
                cell.setCellValue(getMethod.invoke(obj) != null ? getMethod.invoke(obj).toString() : ""); // lancio il metodo getNomeCantiere() e il risultato è il valore della cella
            }
        }
    }
}
