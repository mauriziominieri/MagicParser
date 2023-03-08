package org.parser.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.parser.exception.ExcelException;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Data
public class MagicParser<T> {

    static Workbook workbook;  // package-private per comodità per le sottoclassi
    static Sheet sheet;
    private int sheetIndex;
    private String avanzamentoDa = "";
    private Map<Integer, Boolean> map = new HashMap<>();

    /**
     * Stabilisce se è un metodo get
     *
     * @param field
     * @param method
     * @return
     */
    protected boolean isGetterMethod(String field, Method method) {
        return method.getName()
                .equals("get" + field.substring(0, 1)
                        .toUpperCase() + field.substring(1));
    }

    /**
     * Stabilisce se è un metodo set
     *
     * @param field
     * @param method
     * @return
     */
    protected boolean isSetterMethod(String field, Method method) {
        return method.getName()
                .equals("set" + field.substring(0, 1)
                        .toUpperCase() + field.substring(1));
    }

    /**
     * Analizza i fogli selezionati del template
     *
     * @param templatePath
     * @param sheetIndex
     * @param templateSettingList
     * @return
     * @throws IOException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws ExcelException
     */
    public ByteArrayInputStream write(String templatePath, int sheetIndex, List<TemplateSetting> templateSettingList, List<Integer> landscapePages) throws IOException, InvocationTargetException, IllegalAccessException, ExcelException {
        FileInputStream fileInputStream = new FileInputStream(templatePath);
        workbook = new XSSFWorkbook(fileInputStream);

        for(TemplateSetting setting : templateSettingList) {
            if(sheetIndex == -1) {
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    this.sheetIndex = i;
                    this.sheet = workbook.getSheetAt(i);
                    if(landscapePages.contains(i)) {
                        PrintSetup printSetup = sheet.getPrintSetup();
                        printSetup.setLandscape(true);
                    }
                    write2(setting);
                }
            }
            else {
                this.sheetIndex = sheetIndex;
                this.sheet = workbook.getSheetAt(sheetIndex);
                if(landscapePages.contains(sheetIndex)) {
                    PrintSetup printSetup = sheet.getPrintSetup();
                    printSetup.setLandscape(true);
                }
                write2(setting);
            }
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();
        return new ByteArrayInputStream(out.toByteArray());
    }

    /**
     * Gestice i vari parser da usare
     *
     * @param setting
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws ExcelException
     * @throws IOException
     */
    public void write2(TemplateSetting setting) throws InvocationTargetException, IllegalAccessException, IOException, ExcelException {
        // TODO: c'è un modo per non fare a mano la conversione? devo farla per forza causa parametri da passare...
        if(setting.getClass() == ComplexTemplateSetting.class) {
            ComplexTemplateSetting s = (ComplexTemplateSetting) setting;
            ComplexParser parser = (ComplexParser) setting.getMagicParser();
            parser.write(s.getObject(), s.getSymbol(), s.isCopyStyle());
        }
        else if(setting.getClass() == ComplexDuplicatorTemplateSetting.class && ((ComplexDuplicatorTemplateSetting) setting).getSheetIndex() == this.sheetIndex) {
            ComplexDuplicatorTemplateSetting s = (ComplexDuplicatorTemplateSetting) setting;
            ComplexDuplicatorParser parser = (ComplexDuplicatorParser) setting.getMagicParser();
            parser.write(s.getObjectList(), s.getSymbol(), s.getFirstRowIndex(), s.getLastRowIndex(), s.getFirstColumn(), s.getLastColumn(), s.getGap(), s.isDuplicate(), s.isCopyStyle());
        }
        else if(setting.getClass() == SimpleTemplateSetting.class) {
            SimpleTemplateSetting s = (SimpleTemplateSetting) setting;
            SimpleParser parser = (SimpleParser) setting.getMagicParser();
            parser.write(s.getObjectList(), s.getDirection(), s.getSteps(), s.isFillDuplicateHeadersCells(), s.getObjectListPortion(), s.getHeaderGroup(), s.getHeaderGroupStartIndex(), s.isCopyStyle());
        }
        else if(setting.getClass() == SimpleDuplicatorTemplateSetting.class && ((SimpleDuplicatorTemplateSetting) setting).getSheetIndex() == this.sheetIndex) {
            SimpleDuplicatorTemplateSetting s = (SimpleDuplicatorTemplateSetting) setting;
            SimpleDuplicatorParser parser = (SimpleDuplicatorParser) setting.getMagicParser();

            boolean firstListToDuplicate;
            if(map.get(sheetIndex) == null) {
                map.put(sheetIndex, true);
                firstListToDuplicate = true;
            }
            else
                firstListToDuplicate = false;   // TODO: se usassi una singola impostazione a scacchi per picchetti e campate non servirebbe questa logica sui fogli e sulla prima lista, in realtà forse mi basta fare prima le campate e risolvo

            parser.write(s.getObjectList(), s.getDirection(), s.getSteps(), s.getFirstRowToDuplicateIndex(), s.getLastRowToDuplicateIndex(), s.getGap(), s.isFillDuplicateHeadersCells(), s.getObjectListPortion(), s.getHeaderGroup(), s.getHeaderGroupStartIndex(), s.isCopyStyle(), firstListToDuplicate, s.getSheetIndex(), s.getMaxObjectForPage());
        }
        else if(setting instanceof CantiereTemplateSetting) {
            CantiereTemplateSetting s = (CantiereTemplateSetting) setting;
            CantiereParser parser = (CantiereParser) setting.getMagicParser();
            parser.write(s.getObjectClass(), s.getObjectList(), s.getKey(), s.isCopyStyle());
        }
    }

    /**
     * Scrive le celle
     *
     * @param coordinate        Coordinate della cella
     * @param declaredMethods   I metodi dell'oggetto
     * @param xlsxHeader        Cella con tutte le informazioni
     * @param obj               Oggetto da scrivere
     * @param copyStyle         Se copiare lo stile della cella da sovrascrivere
     */
    public void setCellValue(Coordinate coordinate, Method[] declaredMethods, Header xlsxHeader, T obj, boolean copyStyle) throws InvocationTargetException, IllegalAccessException {
        int riga = coordinate.getRowIndex();
        Row row = sheet.getRow(riga) == null ? sheet.createRow(riga) : sheet.getRow(riga);
        Cell oldCell = row.getCell(coordinate.getColumnIndex());
        CellStyle cellStyle = null;
        if(oldCell != null) cellStyle = oldCell.getCellStyle();
        Cell cell = row.createCell(coordinate.getColumnIndex());
        if(oldCell != null && copyStyle == true) {
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(cellStyle);
            cell.setCellStyle(newStyle);
        }
        String field = xlsxHeader.getFieldName(); // prendo l'attributo annotato con @Field
        String columnName = xlsxHeader.getColumnName().replaceAll(" ", "");
        Optional<Method> getter = Arrays.stream(declaredMethods)
                .filter(method -> isGetterMethod(field, method))
                .findFirst(); // prendo il metodo get dell'attributo @Field
        if (getter.isPresent()) {
            Object value = getter.get().invoke(obj); // tramite reflection lancio il metodo get del field e il risultato lo inserisco in value
            String cellValue = value instanceof Date ?
                    new SimpleDateFormat("dd/MM/yyyy").format((Date) value) : // formato di tutte le date
                    value != null ? value.toString() : "";
            if (columnName.equalsIgnoreCase("avanzamentoda")) // gestisco la cella unificata (solitamente) a inizio report contenente le date REPORT AVANZAMENTO
                avanzamentoDa = cellValue;
            else if (columnName.equalsIgnoreCase("avanzamentoa"))
                cellValue = "REPORT AVANZAMENTO DAL " + avanzamentoDa + " AL " + cellValue;
            cell.setCellValue(cellValue);
        }
    }
}
