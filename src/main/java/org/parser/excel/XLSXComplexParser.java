package org.parser.excel;

import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.IOException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

@NoArgsConstructor
@AllArgsConstructor
public class XLSXComplexParser<T> extends XLSXParser2<T> {

    private String symbol;

    /**
     * Permette di scrivere su file excel
     *
     * @param objectClass   Classe dell'oggetto da scrivere
     * @param object        Oggetto da scrivere
     * @param sheetIndex    Numero del foglio da analizzare, -1 per analizzare tutti i fogli dell'excel
     * @param symbol        Il simbolo nel template che identifica la cella da sovrascrivere (formato <symbol valore>)
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    public void write(int sheetIndex, Class objectClass, T object, int objectListPortion, String symbol, boolean copyStyle) throws IOException, ExcelException {
        this.symbol = symbol;
        mela(sheetIndex, objectClass, object, objectListPortion, copyStyle);
    }

    /**
     * Cerca le celle poi da sovrascrivere (in qualsiasi posizione e anche duplicate di formato <symbol VALORE>) e salva le relative coordinate
     *
     * @param columnTitle   Titolo della colonna dell'header da cercare
     * @param sheet         Foglio da analizzare
     */
    @Override
    protected List<Coordinata> findColumnCoordinates2(String columnTitle, Sheet sheet) {
        columnTitle = this.symbol + columnTitle;
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
     * Scrive le celle
     *
     * @param xlsxHeaders   Lista delle celle dell'headers
     * @param sheet         Foglio da analizzare
     */
    @Override
    protected void pera(List<Header2> xlsxHeaders, Sheet sheet, T obj, boolean copyStyle) throws Exception {
        Method[] declaredMethods = obj.getClass().getDeclaredMethods(); // prendo tutti i metodi get dell'oggetto obj
        for (Header2 xlsxHeader : xlsxHeaders) {
            for(Coordinata coordinata : xlsxHeader.getCoordinataList()) {   // i campi duplicati hanno una lista di dimensione > 1 dell'oggetto Coordinata
                setCellValue(coordinata, sheet, declaredMethods, xlsxHeader, obj, copyStyle);
            }
        }
    }
}
