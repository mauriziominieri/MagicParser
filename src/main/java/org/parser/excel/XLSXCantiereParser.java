package org.parser.excel;

//import it.terna.projconf.dto.VPersonaleDto;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

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

@NoArgsConstructor
@AllArgsConstructor
public class XLSXCantiereParser<T> extends XLSXParser2<T> {

    private int lastSubappaltatoreFieldsIndex;

    private List<Row> rowsToDuplicate;

    private Row rowToCopyStyle;

    int firstRowToDuplicateIndex = 8;

    int numberOfRowsToDuplicate = 3;

    public void write(int sheetIndex, Class objectClass, List<T> objectList, int objectListPortion, XLSXCantiereTemplateSetting.Direction direction, int steps, boolean copyStyle) throws IOException, ExcelException {

        mela(sheetIndex, objectClass, objectList, objectListPortion, copyStyle);
    }

    // viene chiamato tante volte quante sono le colonne, cf, nome, cognome...
    public List<Coordinata> getSubappaltatoreFieldsCoordinates(String columnTitle, Sheet sheet, String key) {
        List<Coordinata> coordinataList = new ArrayList<>();
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
                        coordinataList.add(0, new Coordinata(cell.getRowIndex(), cell.getColumnIndex()));
                    }
                }
            }
        }
        return coordinataList;
    }

    protected List<Header2> modelObjectToXLSXHeaderForWrite2(Class<?> objectClass, Sheet sheet) {
        List<Header2> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : objectClass.getDeclaredFields()) { // scorro tutti gli attributi della classe cls
            Field importField = field.getAnnotation(Field.class); // creo un oggetto Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                String picchettoGroup = importField.group()[0];
                List<Coordinata> coordinataList = getSubappaltatoreFieldsCoordinates(xlsxColumn, sheet, "subappaltatore:");
                listHeader.add(new Header2(field.getName(), xlsxColumn, -1, picchettoGroup, coordinataList)); // TODO controllare se da errore
            }
        }
        return listHeader;
    }

    @Override
    public void mela2(Sheet sheet, Class<T> objectClass, List<T> objectList, int objectListPortion, boolean copyStyle) throws ExcelException, IOException {

        rowsToDuplicate = new ArrayList<>();
        for(int i = 0; i < numberOfRowsToDuplicate; i++)
            rowsToDuplicate.add(sheet.getRow(firstRowToDuplicateIndex + i));

        // TODO: do per scontato che la riga da cui copiare lo stile sia quella subito sotto la porzione duplicata
        rowToCopyStyle = sheet.getRow(firstRowToDuplicateIndex + numberOfRowsToDuplicate);

        //TODO da togliere
        List<List<T>> list = new ArrayList<>();
        list.add(objectList);
        list.add(objectList);
        list.add(objectList);
        list.add(objectList);
        list.add(objectList);
        list.add(objectList);
        list.add(objectList);
        list.add(objectList);
        list.add(objectList);
        list.add(objectList);

        for(int j = 0; j < list.size(); j++) {
            List<Header2> xlsxHeaders = modelObjectToXLSXHeaderForWrite2(objectClass, sheet);
            for (int i = 0; i < list.get(j).size(); i++) {
                T obj = (T) list.get(j).get(i); // avrò la prima persona del primo subappaltatore
                try {
                    riempiPersonaleSubappaltatore(sheet, xlsxHeaders, obj, i, copyStyle);
                }
                catch (Exception e) {
                    workbook.close();
                    e.printStackTrace();
                }
            }
            if(j < list.size() - 1)
                shiftSubappaltatore(sheet, list.get(j).size());
        }
    }

    public void shiftSubappaltatore(Sheet sheet, int nPosti) {
        int firstRow = lastSubappaltatoreFieldsIndex - 2; // TODO renderlo parametro
        int lastRow = firstRow + 2; // TODO renderlo parametro

        for (int j = 0; j < rowsToDuplicate.size(); j++) {
            Row oldRow = rowsToDuplicate.get(j) == null ? sheet.createRow(j) : rowsToDuplicate.get(j);
            Row newRow = sheet.createRow(lastRow + nPosti + j + 1);

            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress region = sheet.getMergedRegion(i);
                if (region.getFirstRow() == oldRow.getRowNum()) {
                    sheet.addMergedRegion(new CellRangeAddress(newRow.getRowNum(), newRow.getRowNum(), region.getFirstColumn(), region.getLastColumn()));
                }
            }

            for (int i = oldRow.getFirstCellNum(); i < oldRow.getLastCellNum(); i++) {
                Cell oldCell = oldRow.getCell(i);
                Cell newCell = newRow.createCell(i);
                newCell.setCellStyle(oldCell.getCellStyle());
                newCell.setCellValue(oldCell.getStringCellValue());
            }
        }
    }

    public void riempiPersonaleSubappaltatore(Sheet sheet, List<Header2> xlsxHeaders, T obj, int i, boolean copyStyle) throws InvocationTargetException, IllegalAccessException {
        Method[] declaredMethods = obj.getClass().getDeclaredMethods();
        for (Header2 xlsxHeader : xlsxHeaders) {
            Coordinata coordinata = xlsxHeader.getCoordinataList().get(0);
            int rowIndex = coordinata.getRiga() + 1 + i; //todo gestirlo con la direzione
            Row row = sheet.getRow(rowIndex) == null ? sheet.createRow(rowIndex) : sheet.getRow(rowIndex);
            Cell cell = row.createCell(coordinata.getColonna());

            if(copyStyle) {
                for (int j = 0; j < rowToCopyStyle.getLastCellNum(); j++) {
                    Cell oldCell = rowToCopyStyle.getCell(j);
                    cell.setCellStyle(oldCell.getCellStyle());
                }
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
