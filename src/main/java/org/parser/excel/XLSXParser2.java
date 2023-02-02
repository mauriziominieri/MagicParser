package org.parser.excel;

import lombok.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.parser.utils.PropertiesUtils;
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
//public class XLSXParser2<T,P> { // VERSIONE GENERIC TYPE
public class XLSXParser2<T> {

    public static Workbook workbook;


    // controlla se esiste il metodo get partendo dal field, es. nomeCantiere cerca getNomeCantiere
    protected boolean isGetterMethod(String field, Method method) {
        return method.getName()
            .equals("get" + field.substring(0, 1)
            .toUpperCase() + field.substring(1));
    }

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
        List<Header2> xlsxHeaders = modelObjectToXLSXHeader(cls, sheet, headerRow);
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
                    throw new ExcelException(PropertiesUtils.getMessage("message.excel.riga", new Object[]{sheet.getSheetName(), rowIndex}));
                }
            }
        }
        workbook.close();

        //listUploadService.save(new ListUploadFileDto(new Date(),file.getOriginalFilename(),rowIndex));

        return out;
    }

    protected T createExcelRow(List<Header2> xlsxHeaders, Row row, Class<T> cls, T obj, short colore) throws Exception {
        Method[] declaredMethods = obj.getClass()
                .getDeclaredMethods();

        for (Header2 xlsxHeader : xlsxHeaders) {
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

    protected void validateHeader(Row row, List<Header2> xlsxHeaders) throws ExcelException {

        for (Header2 xlsxHeader : xlsxHeaders) {
            if (!row.getCell(xlsxHeader.getColumnIndex()).getStringCellValue().replaceAll("\\s+", "").equalsIgnoreCase(xlsxHeader.getColumnName().replaceAll("\\s+", "")))
                throw new ExcelException(PropertiesUtils.getMessage("message.excel.header", new Object[]{xlsxHeader.getColumnName()}));
        }
    }

    protected List<Header2> modelObjectToXLSXHeader(Class<T> cls, Sheet sheet, int headerRow) {
        return Stream.of(cls.getDeclaredFields())

                .filter(field -> field.getAnnotation(Field.class) != null && field.getAnnotation(Field.class).read())
                .map(field -> {
                    Field importField = field.getAnnotation(Field.class);
                    String xlsxColumn = importField.column()[0];
                    int columnIndex = findColumnIndex(xlsxColumn, sheet, headerRow);
                    return new Header2(field.getName(), xlsxColumn, columnIndex, null, null);
                })
                .collect(Collectors.toList());
    }

    protected T createRowObject(List<Header2> xlsxHeaders, Row row, Class<T> cls) throws Exception {
        T obj = cls.getDeclaredConstructor().newInstance();

        Method[] declaredMethods = obj.getClass()
                .getDeclaredMethods();

        for (Header2 xlsxHeader : xlsxHeaders) {
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

    // UTILIZZANDO SOTTOCLASSI E OVERRIDANDO
    /**
     * Permette di scrivere su file excel, scorre i vari tipi di parser e direziona
     *
     * @param xlsxTemplateSettingList   Lista di impostazioni da seguire
     * @param sheetIndex                Numero del foglio da analizzare, -1 per analizzare tutti i fogli dell'excel
     * @param templatePath              Path del template del foglio
     */
    public ByteArrayInputStream write(String templatePath, int sheetIndex, List<XLSXTemplateSetting> xlsxTemplateSettingList) throws IOException, ExcelException {
        FileInputStream in = new FileInputStream(templatePath);
        Workbook workbook = new XSSFWorkbook(in);
        this.workbook = workbook;

        for(XLSXTemplateSetting setting : xlsxTemplateSettingList) {
            if(setting instanceof XLSXComplexTemplateSetting) {
                XLSXComplexTemplateSetting s = (XLSXComplexTemplateSetting) setting;
                XLSXComplexParser parser = (XLSXComplexParser) setting.getXlsxParser();
                parser.write(sheetIndex, s.getObjectClass(), s.getObject(), 10, s.getSymbol(), s.isCopyStyle());
            }
            else if(setting instanceof XLSXSimpleTemplateSetting) {
                XLSXSimpleTemplateSetting s = (XLSXSimpleTemplateSetting) setting;
                XLSXSimpleParser parser = (XLSXSimpleParser) setting.getXlsxParser();

                if(s.getHeaderGroup() == null)
                    parser.write(sheetIndex, s.getObjectClass(), s.getObjectList(), s.getObjectListPortion(), s.isFillDuplicateCells(), s.getDirection(), s.getStep(), s.isCopyStyle());
                else
                    parser.write(sheetIndex, s.getObjectClass(), s.getObjectList(), s.getObjectListPortion(), s.isFillDuplicateCells(), s.getDirection(), s.getStep(), s.isCopyStyle(), s.getHeaderGroup(), s.getHeaderGroupStartIndex());
            }
            else if(setting instanceof XLSXCantiereTemplateSetting) {
                XLSXCantiereTemplateSetting s = (XLSXCantiereTemplateSetting) setting;
                XLSXCantiereParser parser = (XLSXCantiereParser) setting.getXlsxParser();
                parser.write(sheetIndex, s.getObjectClass(), s.getObjectList(), s.getObjectListPortion(), s.getDirection(), s.getStep(), s.isCopyStyle());
            }
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();
        return new ByteArrayInputStream(out.toByteArray());
    }

    /**
     * Permette di scrivere su file excel analizzando tutti i fogli o un foglio scelto
     *
     * @param sheetIndex    Numero del foglio da analizzare, -1 per analizzare tutti i fogli dell'excel
     * @param objectClass   Classe dell'oggetto da scrivere
     * @param objectList    Lista di oggetti da scrivere
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    public void mela(int sheetIndex, Class<T> objectClass, List<T> objectList, int objectListPortion, boolean copyStyle) throws IOException, ExcelException {
        if(sheetIndex == -1) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                mela2(sheet, objectClass, objectList, objectListPortion, copyStyle);
            }
        }
        else {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            mela2(sheet, objectClass, objectList, objectListPortion, copyStyle);
        }
    }

    /**
     * Permette di scrivere su file excel gestendo un layout "a scacchi" analizzando tutti i fogli o un foglio scelto
     *
     * @param sheetIndex    Numero del foglio da analizzare, -1 per analizzare tutti i fogli dell'excel
     * @param objectClass   Classe dell'oggetto da scrivere
     * @param objectList    Lista di oggetti da scrivere
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    public void mela(int sheetIndex, Class<T> objectClass, List<T> objectList, int objectListPortion, boolean copyStyle, String headerGroup, int headerGroupStartIndex) throws IOException, ExcelException {
        if(sheetIndex == -1) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                mela2(sheet, objectClass, objectList, objectListPortion, copyStyle, headerGroup, headerGroupStartIndex);
            }
        }
        else {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            mela2(sheet, objectClass, objectList, objectListPortion, copyStyle, headerGroup, headerGroupStartIndex);
        }
    }


    // PER LA TIPOLOGIA COMPLEX, VUOOLE UN OGGETTO E NON UNA LISTA
    /**
     * Permette di scrivere su file excel analizzando tutti i fogli o un foglio scelto
     *
     * @param sheetIndex    Numero del foglio da analizzare, -1 per analizzare tutti i fogli dell'excel
     * @param objectClass   Classe dell'oggetto da scrivere
     * @param object        Oggetto da scrivere
     * @param copyStyle     Se copiare lo stile della cella da sovrascrivere
     */
    public void mela(int sheetIndex, Class<T> objectClass, T object, int objectListPortion, boolean copyStyle) throws IOException, ExcelException {
        if(sheetIndex == -1) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                mela2(sheet, objectClass, object, objectListPortion, copyStyle);
            }
        }
        else {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            mela2(sheet, objectClass, object, objectListPortion, copyStyle);
        }
    }

    public void mela2(Sheet sheet, Class<T> objectClass, List<T> objectList, int objectListPortion, boolean copyStyle) throws IOException, ExcelException {
        List<Header2> xlsxHeaders = modelObjectToXLSXHeaderForWrite2(objectClass, sheet);
        for (T obj : objectList) {
            try {
                pera(xlsxHeaders, sheet, obj, copyStyle);
            } catch (Exception e) {
                workbook.close();
                e.printStackTrace();
                throw new ExcelException(PropertiesUtils.getMessage("message.excel.sheetName", new Object[]{ sheet.getSheetName() }));
            }
        }
    }

    public void mela2(Sheet sheet, Class<T> objectClass, List<T> objectList, int objectListPortion, boolean copyStyle, String headerGroup, int headerGroupStartIndex) throws IOException, ExcelException {
        List<Header2> xlsxHeaders = modelObjectToXLSXHeaderForWrite2(objectClass, sheet);
        for (T obj : objectList) {
            try {
                pera(xlsxHeaders, sheet, obj, copyStyle, headerGroup, headerGroupStartIndex);
            } catch (Exception e) {
                workbook.close();
                e.printStackTrace();
                throw new ExcelException(PropertiesUtils.getMessage("message.excel.sheetName", new Object[]{ sheet.getSheetName() }));
            }
        }
    }


    // COMPLEX
    public void mela2(Sheet sheet, Class<T> objectClass, T object, int objectListPortion, boolean copyStyle) throws IOException, ExcelException {
        List<Header2> xlsxHeaders = modelObjectToXLSXHeaderForWrite2(objectClass, sheet);
        try {
            pera(xlsxHeaders, sheet, object, copyStyle);
        } catch (Exception e) {
            workbook.close();
            e.printStackTrace();
            throw new ExcelException(PropertiesUtils.getMessage("message.excel.sheetName", new Object[]{ sheet.getSheetName() }));
        }
    }

    /**
     * Restituisce l'header dove ogni cella ha le proprie coordinate a due dimensioni
     * Analizza tutti gli attributi annotati con @Field della classe objectClass
     *
     * @param objectClass   Classe dell'oggetto da scrivere
     * @param sheet         Foglio dell'excel da analizzare
     */
    protected List<Header2> modelObjectToXLSXHeaderForWrite2(Class<?> objectClass, Sheet sheet) {
        List<Header2> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : objectClass.getDeclaredFields()) { // scorro tutti gli attributi della classe cls
            Field importField = field.getAnnotation(Field.class); // creo proprio un oggetto Field-
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                List<Coordinata> coordinataList = findColumnCoordinates2(xlsxColumn, sheet);
                listHeader.add(new Header2(field.getName(), xlsxColumn, -1, null, coordinataList)); // TODO ho fatto questa modifica, devo capire se da problemi
            }
        }
        return listHeader;
    }

    protected List<Coordinata> findColumnCoordinates2(String columnTitle, Sheet sheet) {
        return new ArrayList<>();
    }

    protected void pera(List<Header2> xlsxHeaders, Sheet sheet, T obj, boolean copyStyle) throws Exception {
        System.out.println("NO");
    }

    protected void pera(List<Header2> xlsxHeaders, Sheet sheet, T obj, boolean copyStyle, String headerGroup, int headerGroupStartIndex) throws Exception {
        System.out.println("A SCACCHI");
    }

    protected void prova5(List<Header2> xlsxHeaders, Sheet sheet, T obj, boolean copyStyle, int porzionePicchetto) throws Exception {
        System.out.println("BO");
    }

    public void setCellValue(Coordinata coordinata, Sheet sheet, Method[] declaredMethods, Header2 xlsxHeader, T obj, boolean copyStyle) throws InvocationTargetException, IllegalAccessException {
        int riga = coordinata.getRiga();
        Row row = sheet.getRow(riga) == null ? sheet.createRow(riga) : sheet.getRow(riga);
        Cell oldCell = row.getCell(coordinata.getColonna());
        CellStyle cellStyle = null;
        if(oldCell != null) cellStyle = oldCell.getCellStyle();
        Cell cell = row.createCell(coordinata.getColonna());
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
