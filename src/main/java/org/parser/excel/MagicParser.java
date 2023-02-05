package org.parser.excel;

import lombok.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.parser.utils.PropertiesUtilsParser;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@Data
public class MagicParser<T> {
    static Workbook workbook;  // package-private per comodità per le sottoclassi
    static Sheet sheet;

    // controlla se esiste il metodo get partendo dal field, es. nomeCantiere cerca getNomeCantiere
    /**
     * Permette di scrivere su file excel
     *
     * @param object        Oggetto da scrivere
     * @param symbol        Il simbolo nel template che identifica la cella da sovrascrivere (formato <symbol valore>)
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    protected boolean isGetterMethod(String field, Method method) {
        return method.getName()
                .equals("get" + field.substring(0, 1)
                        .toUpperCase() + field.substring(1));
    }

    /**
     * Permette di scrivere su file excel
     *
     * @param object        Oggetto da scrivere
     * @param symbol        Il simbolo nel template che identifica la cella da sovrascrivere (formato <symbol valore>)
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    protected boolean isSetterMethod(String field, Method method) {
        return method.getName()
                .equals("set" + field.substring(0, 1)
                        .toUpperCase() + field.substring(1));
    }

    // TODO DA QUI IMPORT
    public List<T> fromExcelToObj(MultipartFile file, Class<T> cls, int numSheet, int headerRow) throws Exception {

        List<T> outList = parse(file, cls, numSheet, headerRow);

        return outList;
    }

    protected List<T> parse(MultipartFile file, Class<T> cls, int numSheet, int headerRow) throws Exception {

        int rowIndex = 0; // TODO forse serve come attributo
        List<T> out = new ArrayList<>();

        Workbook workbook = WorkbookFactory.create(file.getInputStream());
        this.workbook = workbook;
        Sheet sheet = workbook.getSheetAt(numSheet);
        List<Header> xlsxHeaders = modelObjectToXLSXHeader(cls, sheet, headerRow);
        validateHeader(sheet.getRow(headerRow), xlsxHeaders);

        for (Row row : sheet) {
            if (row.getRowNum() > headerRow) {
                rowIndex++;
                try {
                    if (row.getCell(0) == null)
                        break;

                    out.add(createRowObject(xlsxHeaders, row, cls));

                } catch (Exception e) {
                    workbook.close();
                    e.printStackTrace();
                    //			listUploadService.save(new ListUploadFileDto(new Date(),file.getOriginalFilename(),rowIndex));
                    throw new ExcelException(PropertiesUtilsParser.getMessage("message.excel.riga", new Object[]{sheet.getSheetName(), rowIndex}));
                }
            }
        }
        workbook.close();

        //listUploadService.save(new ListUploadFileDto(new Date(),file.getOriginalFilename(),rowIndex));

        return out;
    }

    protected T createExcelRow(List<Header> xlsxHeaders, Row row, Class<T> cls, T obj, short colore) throws Exception {
        Method[] declaredMethods = obj.getClass()
                .getDeclaredMethods();

        for (Header xlsxHeader : xlsxHeaders) {
            Cell cell = row.createCell(xlsxHeader.getColumnIndex());
            String field = xlsxHeader.getFieldName();
            Optional<Method> getter = Arrays.stream(declaredMethods)
                    .filter(method -> isGetterMethod(field, method))
                    .findFirst();
            if (getter.isPresent()) {
                Method getMethod = getter.get();
                setCell(getMethod, row, cell, obj, colore);
            }
        }
        return obj;
    }

    protected void setCell(Method getMethod, Row row, Cell cell, T obj, short colore) throws Exception {

        cell.setCellValue(getMethod.invoke(obj) != null ? getMethod.invoke(obj).toString() : "");
        if (colore != 0) {
            setFontCell(cell, colore);
        }

    }

    private void setFontCell(Cell cell, short colore) {

        XSSFFont font = (XSSFFont) workbook.createFont();
        font.setColor(colore);
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        cell.setCellStyle(style);
    }

    protected void validateHeader(Row row, List<Header> xlsxHeaders) throws ExcelException {

        for (Header xlsxHeader : xlsxHeaders) {
            if (!row.getCell(xlsxHeader.getColumnIndex()).getStringCellValue().replaceAll("\\s+", "").equalsIgnoreCase(xlsxHeader.getColumnName().replaceAll("\\s+", "")))
                throw new ExcelException(PropertiesUtilsParser.getMessage("message.excel.header", new Object[]{xlsxHeader.getColumnName()}));
        }
    }

    protected List<Header> modelObjectToXLSXHeader(Class<T> cls, Sheet sheet, int headerRow) {
        return Stream.of(cls.getDeclaredFields())

                .filter(field -> field.getAnnotation(Field.class) != null && field.getAnnotation(Field.class).read())
                .map(field -> {
                    Field importField = field.getAnnotation(Field.class);
                    String xlsxColumn = importField.column()[0];
                    int columnIndex = findColumnIndex(xlsxColumn, sheet, headerRow);
                    return new Header(field.getName(), xlsxColumn, columnIndex, null, null);
                })
                .collect(Collectors.toList());
    }

    protected T createRowObject(List<Header> xlsxHeaders, Row row, Class<T> cls) throws Exception {
        T obj = cls.getDeclaredConstructor().newInstance();

        Method[] declaredMethods = obj.getClass()
                .getDeclaredMethods();

        for (Header xlsxHeader : xlsxHeaders) {
            Cell cell = row.getCell(xlsxHeader.getColumnIndex());
            String field = xlsxHeader.getFieldName();
            Optional<Method> setter = Arrays.stream(declaredMethods)
                    .filter(method -> isSetterMethod(field, method))
                    .findFirst();
            if (setter.isPresent()) {
                Method setMethod = setter.get();
                setObj(setMethod, cell, obj);
            }
        }
        return obj;
    }

    protected static final String NUMERIC = "NUMERIC";
    protected static final String STRING = "STRING";
    protected static final String BLANK = "BLANK";

    protected void setObj(Method setMethod, Cell cell, T obj) throws Exception {
        switch (cell != null && cell.getCellType() != null ? cell.getCellType().toString() : BLANK) {
            case NUMERIC:
                setMethod.invoke(obj, cell.getNumericCellValue());
                break;
            case STRING:
                setMethod.invoke(obj, cell.getStringCellValue());
                break;
            case BLANK:
                setMethod.invoke(obj, "");
                break;

        }
    }

    protected int findColumnIndex(String columnTitle, Sheet sheet, int headerRow) {
        Row row = sheet.getRow(headerRow);

        if (row != null) {
            for (Cell cell : row) {
                if (CellType.STRING.equals(cell.getCellType()) && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", ""))) {
                    return cell.getColumnIndex();
                }
            }
        }
        return 0;
    }

    // TODO FINO A QUI IMPORT


    /**
     * Permette di scrivere su file excel, scorre i vari tipi di parser e direziona
     *
     * @param templateSettingList   Lista di impostazioni da seguire
     * @param sheetIndex                Numero del foglio da analizzare, -1 per analizzare tutti i fogli dell'excel
     * @param templatePath              Path del template del foglio
     */
    public ByteArrayInputStream write(String templatePath, int sheetIndex, List<TemplateSetting> templateSettingList) throws IOException, InvocationTargetException, IllegalAccessException, ExcelException {
        FileInputStream fileInputStream = new FileInputStream(templatePath);
        workbook = new XSSFWorkbook(fileInputStream);

        for(TemplateSetting setting : templateSettingList) {
            if(sheetIndex == -1) {
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    sheet = workbook.getSheetAt(i);
                    write2(setting);
                }
            }
            else {
                sheet = workbook.getSheetAt(sheetIndex);
                write2(setting);
            }
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();
        return new ByteArrayInputStream(out.toByteArray());
    }

    /**
     * Permette di scrivere su file excel
     *
     * @param object        Oggetto da scrivere
     * @param symbol        Il simbolo nel template che identifica la cella da sovrascrivere (formato <symbol valore>)
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    public void write2(TemplateSetting setting) throws InvocationTargetException, IllegalAccessException, ExcelException, IOException {
        // TODO: c'è un modo per non fare a mano la conversione? devo farla per forza causa parametri da passare...
        if(setting instanceof ComplexTemplateSetting) {
            ComplexTemplateSetting s = (ComplexTemplateSetting) setting;
            ComplexParser parser = (ComplexParser) setting.getMagicParser();
            parser.write(s.getObject(), s.getSymbol(), s.isCopyStyle());
        }
        else if(setting instanceof SimpleTemplateSetting) {
            SimpleTemplateSetting s = (SimpleTemplateSetting) setting;
            SimpleParser parser = (SimpleParser) setting.getMagicParser();
            parser.write(s.getObjectList(), s.getDirection(), s.getSteps(), s.isFillDuplicateHeadersCells(), s.getObjectListPortion(), s.getHeaderGroup(), s.getHeaderGroupStartIndex(), s.isCopyStyle());
        }
        else if(setting instanceof DuplicatorTemplateSetting) {
            DuplicatorTemplateSetting s = (DuplicatorTemplateSetting) setting;
            DuplicatorParser parser = (DuplicatorParser) setting.getMagicParser();
            parser.write(s.getObjectList().get(0).getClass(), s.getObjectList(), s.getKey(), s.isCopyStyle());
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
