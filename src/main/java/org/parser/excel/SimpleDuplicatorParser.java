package org.parser.excel;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Data
public class SimpleDuplicatorParser<T> extends SimpleParser {
    private int lastPortionInitialIndex;
    private List<Row> rowsToDuplicate;
    int firstRowIndex; int firstRowIndex2;
    int gap;
    int lastRowIndex; int lastRowIndex2;
    int headerIndex;
    private boolean firstListToDuplicate;
    private int maxObjectForPage;
    private int sheetIndex;

    public void write(List<T> objectList, Direction direction, int steps, int firstRowIndex, int lastRowIndex, int gap, boolean fillDuplicateHeadersCells, int objectListPortion, String headerGroup, int headerGroupStartIndex, boolean copyStyle, boolean firstListToDuplicate, int sheetIndex, int maxObjectForPage) throws InvocationTargetException, IllegalAccessException, IOException {
        if(objectList.isEmpty()) return;
        this.object = objectList.get(0);
        this.direction = direction;
        this.lastRowIndex = lastRowIndex;
        this.lastRowIndex2 = lastRowIndex;
        this.gap = gap;
        this.steps = steps;
        this.fillDuplicateHeadersCells = fillDuplicateHeadersCells; // TODO dubbi, serve veramente in questa versione SimpleDuplicatorParser?
        this.objectListPortion = objectListPortion; // TODO: devo gestirlo veramente, perchè ora come ora se setto 3 e ho una porzione da 10 voglio che riempia 3, duplichi sotto e riempie altri 3 e così via.
        this.headerGroup = headerGroup;
        this.headerGroupStartIndex = headerGroupStartIndex;
        this.firstRowIndex = firstRowIndex; //TODO are un controllo sui parametri settati, se il parametro x è 0 allora throw exception //findFirstRowToDuplicateIndex(sheet, key);
        this.copyStyle = copyStyle;
        this.firstListToDuplicate = firstListToDuplicate; // TODO: per gestire picchetti e campate serve per forza, in quanto saranno i picchetti a spostare la porzione finale (portale esistente), voglio sicuramete migliorare la logica e usare una singola impostazione a scacchi per picchetti e campate
        this.sheetIndex = sheetIndex;
        this.maxObjectForPage = maxObjectForPage;
        this.firstRowIndex2 = firstRowIndex;

        Workbook tempWorkbook = new XSSFWorkbook();
        setRowsToDuplicate(tempWorkbook);   // salvo le rows da duplicare

        int rowIndex = 0, rowIndex2 = 0, k = 0, z = 0;
        List<Header> xlsxHeaders = getHeaders(rowIndex++);
        Method[] declaredMethods = object.getClass().getDeclaredMethods(); // prendo tutti i metodi get dell'oggetto obj

        for (int i = 0; i < objectList.size(); i++) {
            T obj = objectList.get(i);

            if(this.maxObjectForPage > 0 && i > 0 && i % this.maxObjectForPage == 0) {
                createNewSheet(sheet, ++k);
                xlsxHeaders = getHeaders2();
            }
            else if(i > 0 && i % this.objectListPortion == 0) {
                if(this.maxObjectForPage == 0 || (this.maxObjectForPage > this.objectListPortion  && this.maxObjectForPage % this.objectListPortion == 0)) {
                    shiftPorzione(++k);
                    xlsxHeaders = getHeaders(rowIndex++); // ricalcolo gli headers dalla nuova porzione duplicata, le coordinate precedenti verranno perse
                }
            }

            manage(xlsxHeaders, declaredMethods, obj, 0); // dato che le coordinata precedenti andranno perse la mia lista sarà sempre di un elemento, quindi sempre 0
        }

        tempWorkbook.close();
    }

    /**
     * Restituisce l'header dove ogni cella ha le proprie coordinate
     *
     * @param rowIndex  Gestisce l'indice della riga analizzata (se devo andare a duplicare riga per riga verso il basso l'header sarà 1 sopra, poi due sopra ecc.)
     */
    protected List<Header> getHeaders(int rowIndex) {
        List<Header> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : object.getClass().getDeclaredFields()) { // scorro tutti gli attributi della classe cls, il field è preso grazie alla reflection
            Field importField = field.getAnnotation(Field.class); // degli attributi nella classe prendo solo quelli taggati con @Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                String group = importField.group()[0];
                List<Coordinate> coordinateList = getHeadersCoordinates(xlsxColumn, group, rowIndex);
                listHeader.add(new Header(field.getName(), xlsxColumn, -1, group, coordinateList));
            }
        }
        return listHeader;
    }

    protected List<Header> getHeaders2() {
        List<Header> listHeader = new ArrayList<>();
        for(java.lang.reflect.Field field : object.getClass().getDeclaredFields()) { // scorro tutti gli attributi della classe cls, il field è preso grazie alla reflection
            Field importField = field.getAnnotation(Field.class); // degli attributi nella classe prendo solo quelli taggati con @Field
            if(importField != null) {
                String xlsxColumn = importField.column()[0]; // IL NOME DELLA COLONNA NELL'EXCEL PRESO DAL @Field
                String group = importField.group()[0];
                List<Coordinate> coordinateList = getHeadersCoordinates2(xlsxColumn, group);
                listHeader.add(new Header(field.getName(), xlsxColumn, -1, group, coordinateList));
            }
        }
        return listHeader;
    }

    protected List<Coordinate> getHeadersCoordinates2(String columnTitle, String group) {
        List<Coordinate> coordinataList = new ArrayList<>();
        if(direction == Direction.RIGHT) {
            for (int i = firstRowIndex2; i <= lastRowIndex2; i++) {
                Row row = sheet.getRow(i);
                if (row != null)
                    for (int j = 0; j < 2; j++) { //TODO do per scontato che le celle da analizzare siano sono nelle prime due colonne, (parametrizzare?)
                        Cell cell = row.getCell(j);
                        if (CellType.STRING.equals(cell.getCellType()) && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", ""))) {
                            coordinataList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                            return coordinataList; // TODO in questo modo salvo la posizione della perima cella trovata, non gestisco il caso di duplicati insomma
                        }
                    }
            }
        }
        return coordinataList;
    }

    /**
     * Restituisce le coordinate delle celle nell'header
     *
     * @param columnTitle   Titolo della colonna dell'header da cercare
     * @param group
     * @param rowIndex      Gestisce l'indice della riga analizzata (se devo andare a duplicare riga per riga verso il basso l'header sarà 1 sopra, poi due sopra ecc.)
     * @return
     */
    protected List<Coordinate> getHeadersCoordinates(String columnTitle, String group, int rowIndex) {
        List<Coordinate> coordinataList = new ArrayList<>();
        if(direction == Direction.RIGHT) {
            for (int i = firstRowIndex; i <= lastRowIndex; i++) {
                Row row = sheet.getRow(i);
                if (row != null)
                    for (int j = 0; j < 2; j++) { //TODO do per scontato che le celle da analizzare siano sono nelle prime due colonne, (parametrizzare?)
                        Cell cell = row.getCell(j);
                        if (CellType.STRING.equals(cell.getCellType()) && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", ""))) {
                            coordinataList.add(new Coordinate(cell.getRowIndex(), cell.getColumnIndex()));
                            return coordinataList; // TODO in questo modo salvo la posizione della perima cella trovata, non gestisco il caso di duplicati insomma
                        }
                    }
            }
        }
        else if(direction == Direction.BOTTOM) {
            for (int i = firstRowIndex - 1 - rowIndex; i <= lastRowIndex; i++) { // TODO do per scontato che voglio partire dalla riga superiore (quando c'è l'header e voglio copiare solo una riga alla volta mano mano)
                Row row = sheet.getRow(i);
                if (row != null)
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        Cell cell = row.getCell(j);
                        if (CellType.STRING.equals(cell.getCellType()) && columnTitle.replaceAll("\\s+", "").equalsIgnoreCase(cell.getStringCellValue().replaceAll("\\s+", ""))) {
                            if(group.isBlank()) {
                                coordinataList.add(new Coordinate(cell.getRowIndex() + rowIndex, cell.getColumnIndex()));
                                return coordinataList;
                            }
                            else {
                                Row groupRow = sheet.getRow(i - 1); // TODO do per scontato che l'header per i gruppi sia subito sopra quello delle celle
                                if (groupRow != null)
                                    for (int z = 0; z < cell.getColumnIndex(); z++) {
                                        Cell groupCell = groupRow.getCell(z);
                                        if(group.equalsIgnoreCase(groupCell.getStringCellValue().replaceAll("\\s+", ""))) {
                                            coordinataList.add(new Coordinate(cell.getRowIndex() + rowIndex, cell.getColumnIndex()));
                                            return coordinataList;
                                        }
                                    }
                            }
                        }
                    }
            }
        }
        return coordinataList;
    }

    /**
     * Shifta la porzione duplicata sotto e la copia esattamente come l'originale
     */
    public void shiftPorzione(int k) {
        int a = lastRowIndex + gap + 1;
        int b = sheet.getLastRowNum();
        int numberOfRowsToDuplicate = lastRowIndex - firstRowIndex + 1;
        int n = numberOfRowsToDuplicate + gap;
        if(firstListToDuplicate == true)
            sheet.shiftRows(a, b, n); // shifta la porzione (a, b) di c righe sotto.

        for (int i = 0; i < rowsToDuplicate.size(); i++) {
            Row copiedRow = rowsToDuplicate.get(i) == null ? sheet.createRow(i) : rowsToDuplicate.get(i);
            // non posso copiare stili tra diversi workbook quindi mi serve la riga originale da cui prenderò lo stile
            Row oldRow = rowsToDuplicate.get(i) == null ? sheet.createRow(i) : sheet.getRow(firstRowIndex + i);
            Row newRow = sheet.createRow(a + i);

            // copio eventuali merged region
            for (int j = 0; j < sheet.getNumMergedRegions(); j++) {
                CellRangeAddress region = sheet.getMergedRegion(j);
                if (region.getFirstRow() == copiedRow.getRowNum())
                    sheet.addMergedRegion(new CellRangeAddress(region.getFirstRow() + n * k, region.getLastRow() + n * k, region.getFirstColumn(), region.getLastColumn())); // + n perchè la porzione copiata deve essere spostata n celle sotto
            }

            // copio stile e valore
            newRow.setHeight(oldRow.getHeight());
            for (int j = copiedRow.getFirstCellNum(); j < copiedRow.getLastCellNum(); j++) {
                Cell oldCell = copiedRow.getCell(j);
                Cell newCell = newRow.createCell(j);
                newCell.setCellStyle(oldRow.getCell(j).getCellStyle());
                if (oldCell.getCellType() == CellType.NUMERIC) {
                    double cellValue = oldCell.getNumericCellValue();
                    newCell.setCellValue(cellValue);
                    newCell.setCellValue(cellValue + objectListPortion * k); // gestisco gli header, es. finito il primo header a 10 l'header duplicato parte 11, potrei mettere i == 0 per un controllo maggiore e assicurarmi che sia l'header, ma a volte ci sono più header...
                } else if (oldCell.getCellType() == CellType.STRING) {
                    String cellValue = oldCell.getStringCellValue();
                    newCell.setCellValue(cellValue);
                }
            }
        }

        firstRowIndex = a;
        lastRowIndex = a + numberOfRowsToDuplicate - 1;
    }

    public void createNewSheet(Sheet oldSheet, int k) {
        String name = oldSheet.getSheetName();
        Pattern pattern = Pattern.compile("\\d+$");
        Matcher matcher = pattern.matcher(name);
        if (matcher.find()) {
            int number = Integer.parseInt(matcher.group()) + 1;
            name = name.replaceAll("\\d+$", String.valueOf(number));
        } else
            name += " 1";
        if(workbook.getSheet(name) == null) {
            sheet = workbook.createSheet(name);
            workbook.setSheetOrder(name, sheetIndex + 1);
            sheet.getPrintSetup().setLandscape(oldSheet.getPrintSetup().getLandscape()); // Copia l'orientamento della pagina


            // Taglia le righe da oldSheet a sheet
            for (int i = lastRowIndex + gap + 1; i <= oldSheet.getLastRowNum(); i++) {
                Row row = sheet.createRow(i - rowsToDuplicate.size() * k + gap - 2); // Crea una nuova riga in sheet
                Row oldRow = oldSheet.getRow(i); // Prende la riga da oldSheet
                if (oldRow != null) { // Se la riga esiste in oldSheet
                    // copio eventuali merged region
                    for (int j = 0; j < oldSheet.getNumMergedRegions(); j++) {
                        CellRangeAddress region = oldSheet.getMergedRegion(j);
                        if (region.getFirstRow() == oldRow.getRowNum())
                            sheet.addMergedRegion(new CellRangeAddress(region.getFirstRow() - rowsToDuplicate.size() * k + gap - 2, region.getLastRow() - rowsToDuplicate.size() * k + gap - 2, region.getFirstColumn(), region.getLastColumn())); // + n perchè la porzione copiata deve essere spostata n celle sotto
                    }
                    // copio stile e valore
                    row.setHeight(oldRow.getHeight());
                    for (int j = 0; j < oldRow.getLastCellNum(); j++) {
                        Cell cell = row.createCell(j); // Crea una nuova cella in sheet
                        Cell oldCell = oldRow.getCell(j); // Prende la cella da oldSheet

                        if (oldCell != null) { // Se la cella esiste in oldSheet
                            cell.setCellStyle(oldCell.getCellStyle()); // Copia lo stile dalla cella di oldSheet
                            if (oldCell.getCellType() == CellType.NUMERIC) {
                                cell.setCellValue(oldCell.getNumericCellValue() + objectListPortion * k); // gestisco gli header, es. finito il primo header a 10 l'header duplicato parte 11, potrei mettere i == 0 per un controllo maggiore e assicurarmi che sia l'header, ma a volte ci sono più header...
                            } else if (oldCell.getCellType() == CellType.STRING) {
                                cell.setCellValue(oldCell.getStringCellValue()); // Copia il valore dalla cella di oldSheet
                            }
                        }
                    }
                    oldSheet.removeRow(oldRow); // Rimuove la riga da oldSheet
                }
            }
        }
        else
            sheet = workbook.getSheet(name);

        this.sheetIndex++;
        int a = firstRowIndex2;
        for (int i = 0; i < rowsToDuplicate.size(); i++) {
            Row copiedRow = rowsToDuplicate.get(i) == null ? sheet.createRow(i) : rowsToDuplicate.get(i);
            // non posso copiare stili tra diversi workbook quindi mi serve la riga originale da cui prenderò lo stile
            Row oldRow = rowsToDuplicate.get(i) == null ? sheet.createRow(i) : oldSheet.getRow(firstRowIndex2 + i);
            Row newRow = sheet.createRow(a + i);

            // copio eventuali merged region
            for (int j = 0; j < oldSheet.getNumMergedRegions(); j++) {
                CellRangeAddress region = oldSheet.getMergedRegion(j);
                if (region.getFirstRow() == copiedRow.getRowNum())
                    sheet.addMergedRegion(new CellRangeAddress(region.getFirstRow(), region.getLastRow(), region.getFirstColumn(), region.getLastColumn())); // + n perchè la porzione copiata deve essere spostata n celle sotto
            }

            // copio stile e valore
            newRow.setHeight(oldRow.getHeight());
            for (int j = copiedRow.getFirstCellNum(); j < copiedRow.getLastCellNum(); j++) {
                Cell oldCell = copiedRow.getCell(j);
                Cell newCell = newRow.createCell(j);
                newCell.setCellStyle(oldRow.getCell(j).getCellStyle());
                sheet.setColumnWidth(j, oldSheet.getColumnWidth(j)); // copio larghezza colonne
                if (oldCell.getCellType() == CellType.NUMERIC) {
                    double cellValue = oldCell.getNumericCellValue();
                    newCell.setCellValue(cellValue + objectListPortion * k); // gestisco gli header, es. finito il primo header a 10 l'header duplicato parte 11, potrei mettere i == 0 per un controllo maggiore e assicurarmi che sia l'header, ma a volte ci sono più header...
                } else if (oldCell.getCellType() == CellType.STRING) {
                    String cellValue = oldCell.getStringCellValue();
                    newCell.setCellValue(cellValue);
                }
            }
        }
    }

    /**
     * Salva le righe da duplicare
     *
     * @param tempWorkbook
     */
    public void setRowsToDuplicate(Workbook tempWorkbook) {
        Sheet tempSheet = tempWorkbook.createSheet("Sheet temporaneo");
        rowsToDuplicate = new ArrayList<>();
        for(int i = 0; i <= lastRowIndex - firstRowIndex; i++) {
            Row oldRow = sheet.getRow(firstRowIndex + i);
            Row tempRow = tempSheet.createRow(firstRowIndex + i);
            for (int j = 0; j < oldRow.getLastCellNum(); j++) {
                Cell oldCell = oldRow.getCell(j, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                Cell newCell = tempRow.createCell(j);
                if (oldCell.getCellType() == CellType.NUMERIC) {
                    double cellValue = oldCell.getNumericCellValue();
                    newCell.setCellValue(cellValue);
                } else if (oldCell.getCellType() == CellType.STRING) {
                    String cellValue = oldCell.getStringCellValue();
                    newCell.setCellValue(cellValue);
                }
            }
            rowsToDuplicate.add(tempRow);
        }
    }
}
