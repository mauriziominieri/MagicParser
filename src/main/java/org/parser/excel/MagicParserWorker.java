package org.parser.excel;

import com.aspose.cells.*;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.*;
import com.itextpdf.text.pdf.pdfcleanup.PdfCleanUpLocation;
import com.itextpdf.text.pdf.pdfcleanup.PdfCleanUpProcessor;
import org.parser.dto.AereoDto;
import org.parser.dto.MagicParserDto;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.parser.excel.Direction.BOTTOM;
import static org.parser.excel.Direction.RIGHT;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Component
public class MagicParserWorker {

    @Autowired
    MagicParserService magicParserService;

    private ResponseEntity<Resource> getResponse(InputStreamResource file, String nomeFile) {
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + nomeFile + ".xlsx" + "; readonly")
                .header("Content-Transfer-Encoding", "binary")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }

    private ResponseEntity<byte[]> getResponsePdf(InputStreamResource file, String fileName) throws Exception {

        Workbook workbook = new Workbook(file.getInputStream());
        for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
            Worksheet worksheet = workbook.getWorksheets().get(i);
            PageSetup pageSetup = worksheet.getPageSetup();
//            pageSetup.setZoom(100); // reduce or enlarge a worksheet’s size by adjusting the scaling factor
            pageSetup.setPaperSize(PaperSizeType.PAPER_A_4); // paper size that the worksheets will be printed
//            pageSetup.setPrintQuality(180); // print quality of the worksheets to be printed
            pageSetup.setFitToPagesWide(1); // imposta la larghezza della pagina in modo che il contenuto del foglio Excel venga adattato su una sola pagina in orizzontale.
            pageSetup.setLeftMargin(0.2);   // imposta il margine sinistro della pagina a 0,5 pollici.
            pageSetup.setRightMargin(0.2);  // imposta il margine destro della pagina a 0,5 pollici.
            pageSetup.setCenterHorizontally(true); // imposta l'allineamento orizzontale del foglio Excel in modo che sia centrato sulla pagina PDF.
//            pageSetup.setCenterVertically(true); // imposta l'allineamento verticale del foglio Excel in modo che sia centrato sulla pagina PDF.
//            pageSetup.setHeaderMargin(0.2);
            pageSetup.setFooterMargin(0.2);
            // Setting the current page number and page count at the right footer
            pageSetup.setFooter(1, "&P di &N"); // footer centrale, numero pagina
        }
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbook.save(baos, SaveFormat.PDF);
        PdfReader reader = new PdfReader(baos.toByteArray());
        PdfStamper stamper = new PdfStamper(reader, baos);

        PdfDictionary dict = reader.getTrailer().getAsDict(PdfName.INFO);
        dict.put(PdfName.TITLE, new PdfString(fileName));
        dict.put(PdfName.AUTHOR, new PdfString("Indra"));
        dict.put(PdfName.PRODUCER, new PdfString("Indra"));

        Rectangle rectPortrait = new Rectangle(0, 790, 600, 830);
        Rectangle rectLandscape = new Rectangle(0, 581, 841, 650);
        for (int i = 1; i <= reader.getNumberOfPages(); i++) {

            Worksheet worksheet = workbook.getWorksheets().get(i - 1);
            PageSetup pageSetup = worksheet.getPageSetup();
            Rectangle rect = pageSetup.getOrientation() == PageOrientationType.LANDSCAPE ? rectLandscape : rectPortrait;

//            Rectangle rect = landscapePages.contains(i - 1) ? rectLandscape : rectPortrait;
            PdfCleanUpLocation cleanUp = new PdfCleanUpLocation(i, rect);
            PdfCleanUpProcessor cleaner = new PdfCleanUpProcessor(Arrays.asList(cleanUp), stamper);
            cleaner.cleanUp();
        }
        stamper.close();
        reader.close();

        byte[] pdfContent = baos.toByteArray();
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", fileName + ".pdf");
        headers.setContentLength(pdfContent.length);
        return new ResponseEntity<>(pdfContent, headers, HttpStatus.OK);
    }

    /** COMPLEX
     */
    public InputStreamResource getReport0() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        String templatePath = "src/main/resources/template/Template0.xlsx";

        MagicParserDto magicParserDto = new MagicParserDto("A", "B", "C", "D", "E", "F", "G", "H", "I");

        /** IMPOSTAZIONE 1
         *  1. Seleziono il ComplexParser per poter riempire celle di formato <symbol valore>
         *  2. Elemento da inserire nel report (i valori dei suoi attributi sono il contenuto delle celle)
         *  3. Le righe che riempio avranno lo stesso stile dell'originale
         */
        ComplexTemplateSetting<MagicParserDto> templateSetting1 = new ComplexTemplateSetting<>();
        templateSetting1.setMagicParser(new ComplexParser());
        templateSetting1.setObject(magicParserDto);
//        templateSetting1.setSymbol("*"); // di default
        templateSetting1.setCopyStyle(true);

        List<TemplateSetting> templateSettingList = new ArrayList<>();
        templateSettingList.add(templateSetting1);

        InputStreamResource file = new InputStreamResource(magicParserService.exportExcel(templatePath, 0, templateSettingList, List.of()));
        return file;
    }

    public ResponseEntity<Resource> report0() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        String reportPath = "Report0";
        InputStreamResource file = getReport0();
        return getResponse(file, reportPath);
    }

    public ResponseEntity<byte[]> report0Pdf() throws Exception {
        String reportPath = "Report0";
        InputStreamResource file = getReport0();
        return getResponsePdf(file, reportPath);
    }

    /** SIMPLE
     */
    public InputStreamResource getReport1() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        String templatePath = "src/main/resources/template/Template1.xlsx";

        List<MagicParserDto> magicParserDtoList = new ArrayList<>();
        magicParserDtoList.add(new MagicParserDto("A", "B", "C", "D", "E", "F", "G", "H", "L"));
        magicParserDtoList.add(new MagicParserDto("Z", "B", "K", "D", "K", "W", "X", "P", "M"));
        magicParserDtoList.add(new MagicParserDto("A", "X", "C", "L", "E", "W", "G", "Q", "N"));

        /** IMPOSTAZIONE 1
         *  1. Seleziono il SimpleParser per poter riempire celle secondo una regola definibile (direzione e step)
         *  2. Lista degli elementi da inserire nel report
         *  3. Scrivere il valore delle celle verso il basso
         *  4. Riempie Header con colonne duplicate (ha la precedenza sulla 5.)
         *  5. Porzione da 2 -> il terzo elemento viene scritto al prossimo riferimento (ignorata se la 4. è true)
         *  6. Le righe duplicate avranno lo stesso stile dell'originale
         */
        SimpleTemplateSetting<MagicParserDto> templateSetting1 = new SimpleTemplateSetting<>();
        templateSetting1.setMagicParser(new SimpleParser());
        templateSetting1.setObjectList(magicParserDtoList);
        templateSetting1.setDirection(BOTTOM);
//        templateSetting1.setSteps(1); // di default
        templateSetting1.setFillDuplicateHeadersCells(true);
        templateSetting1.setObjectListPortion(2);
        templateSetting1.setCopyStyle(true);

        List<TemplateSetting> templateSettingList = new ArrayList<>();
        templateSettingList.add(templateSetting1);

        InputStreamResource file = new InputStreamResource(magicParserService.exportExcel(templatePath, 0, templateSettingList, List.of()));
        return file;
    }

    public ResponseEntity<Resource> report1() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        String reportPath = "Report1";
        InputStreamResource file = getReport1();
        return getResponse(file, reportPath);
    }

    public ResponseEntity<byte[]> report1Pdf() throws Exception {
        String reportPath = "Report1";
        InputStreamResource file = getReport1();
        return getResponsePdf(file, reportPath);
    }

    /** PROBLEMI
     *  1. Layout a scacchi (picchetti e campate)
     *  2. Picchetti e campate devono essere divisi in porzioni diverse (20 e 19) -> dovrei quindi utilizzare 2 liste
     *  3. Campi duplicati (Autorizzazione & Asservimento)
     */
    public InputStreamResource getReport2() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        String templatePath = "src/main/resources/template/Template2.xlsx";

        List<AereoDto> aereoDtoList = new ArrayList<>();
        aereoDtoList.add(new AereoDto("A", "B", "C", "D", "E", "F", "G", "Z", "I", "L", "M", "N", "O", "P", "Q", "R", "P123"));
        aereoDtoList.add(new AereoDto("B", "B", "C", "D", "E", "F", "G", "H", "Z", "L", "M", "S", "O", "P", "Q", "R", "P3"));
        aereoDtoList.add(new AereoDto("C", "Z", "C", "D", "E", "F", "G", "H", "I", "L", "M", "N", "O", "P", "Q", "R", "P1"));
        aereoDtoList.add(new AereoDto("D", "B", "C", "D", "E", "F", "Z", "H", "I", "L", "Z", "Z", "O", "W", "Q", "R", "P2"));
        aereoDtoList.add(new AereoDto("E", "B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "N", "O", "P", "Q", "R", "2"));
        aereoDtoList.add(new AereoDto("G", "B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "W", "X", "P", "Q", "R", "A1"));
        aereoDtoList.add(new AereoDto("F", "B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "N", "O", "P", "Q", "R", "A1"));
        aereoDtoList.add(new AereoDto("H", "B", "C", "Z", "E", "F", "G", "H", "I", "L", "M", "K", "O", "P", "P", "R", "A1"));
        aereoDtoList.add(new AereoDto("I", "B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "N", "O", "P", "Q", "Y", "A1"));
        aereoDtoList.add(new AereoDto("L", "B", "C", "D", "E", "Z", "G", "H", "I", "L", "M", "J", "O", "P", "Q", "R", "A1"));
        // 11
        aereoDtoList.add(new AereoDto("Z", "B", "C", "D", "E", "F", "G", "Z", "I", "L", "M", "W", "O", "P", "Q", "R", "A1"));
        aereoDtoList.add(new AereoDto("K", "B", "C", "D", "E", "F", "G", "H", "Z", "L", "M", "E", "O", "P", "Q", "R", "A1"));
        aereoDtoList.add(new AereoDto("L", "Z", "C", "D", "E", "F", "G", "H", "I", "L", "M", "N", "O", "P", "Q", "R", "A1"));
        aereoDtoList.add(new AereoDto("D", "B", "C", "D", "E", "F", "Z", "H", "I", "L", "Z", "Y", "O", "W", "Q", "R", "A1"));
        aereoDtoList.add(new AereoDto("E", "B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "N", "O", "P", "Q", "R", "A1"));
        aereoDtoList.add(new AereoDto("F", "B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "N", "O", "P", "Q", "R", "A1"));
        aereoDtoList.add(new AereoDto("G", "B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "W", "X", "P", "Q", "R", "A1"));
        aereoDtoList.add(new AereoDto("M", "B", "C", "Z", "E", "F", "G", "H", "I", "L", "M", "N", "O", "P", "P", "R", "A1"));
        aereoDtoList.add(new AereoDto("I", "B", "C", "D", "E", "F", "G", "H", "I", "L", "M", "T", "O", "P", "Q", "Y", "A1"));
        aereoDtoList.add(new AereoDto("L", "B", "C", "D", "E", "Z", "G", "H", "I", "L", "M", "N", "O", "P", "Q", "R", "A1"));
        // 21
        aereoDtoList.add(new AereoDto("D", "B", "C", "D", "E", "F", "Z", "H", "I", "L", "Z", "Y", "O", "W", "Q", "R", "A1"));

        /** IMPOSTAZIONE 1
         *  1. Seleziono il SimpleParser per poter riempire celle secondo una regola definibile (direzione e step)
         *  2. Lista degli elementi da inserire nel report
         *  3. Scrivere il valore delle celle verso destra
         *  4. Scrivere il valore delle celle facendo salti di 2
         *  5. Porzione da 10 -> l'undicesimo elemento viene scritto al prossimo riferimento
         *  6. Tutte le celle raggruppate logicamente come "PICCHETTO" nel DTO voglio gestirle diversamente
         *  7. Le celle del gruppo logico "PICCHETTO" le sposto una cella indietro (questo permette un layout a scacchi con le campate)
         *  8. Le righe duplicate avranno lo stesso stile dell'originale
         */
        SimpleTemplateSetting<AereoDto> templateSetting1 = new SimpleTemplateSetting<>();
        templateSetting1.setMagicParser(new SimpleParser());
        templateSetting1.setObjectList(aereoDtoList);
        templateSetting1.setDirection(RIGHT);
        templateSetting1.setSteps(2);
        templateSetting1.setObjectListPortion(20);
        templateSetting1.setHeaderGroup("PICCHETTO");
        templateSetting1.setHeaderGroupStartIndex(-1);
        templateSetting1.setCopyStyle(true);

        /** IMPOSTAZIONE 2
         *  1. Seleziono il ComplexParser per poter riempire celle di formato <symbol valore>
         *  2. Elemento da inserire nel report (i valori dei suoi attributi sono il contenuto delle celle)
         *  3. Le righe che riempio avranno lo stesso stile dell'originale
         */
        ComplexTemplateSetting<AereoDto> templateSetting2 = new ComplexTemplateSetting<>();
        templateSetting2.setMagicParser(new ComplexParser());
        templateSetting2.setObject(aereoDtoList.get(0));
        templateSetting2.setCopyStyle(true);

        List<TemplateSetting> templateSettingList = new ArrayList<>();
        templateSettingList.add(templateSetting1);
        templateSettingList.add(templateSetting2);

        InputStreamResource file = new InputStreamResource(magicParserService.exportExcel(templatePath, 0, templateSettingList, List.of()));
        return file;
    }

    public ResponseEntity<Resource> report2() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        String reportPath = "Report2";
        InputStreamResource file = getReport2();
        return getResponse(file, reportPath);
    }

    public ResponseEntity<byte[]> report2Pdf() throws Exception {
        String reportPath = "Report2";
        InputStreamResource file = getReport2();
        return getResponsePdf(file, reportPath);
    }

    /** PROBLEMI
     *  1. Layout a scacchi (picchetti e campate)
     *  2. Picchetti e campate devono essere divisi in porzioni diverse (20 e 19) -> dovrei quindi utilizzare 2 liste
     *  3. Campi duplicati (Autorizzazione & Asservimento)
     */
    public InputStreamResource getReport3() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        String templatePath = "src/main/resources/template/Template3.xlsx";

        MagicParserDto magicParserDto = new MagicParserDto("A", "B", "C", "D", "E", "F", "G", "H", "I");

        /** IMPOSTAZIONE 1
         *  1. Seleziono il ComplexParser per poter riempire celle di formato <symbol valore>
         *  2. Elemento da inserire nel report (i valori dei suoi attributi sono il contenuto delle celle)
         *  3. Le righe che riempio avranno lo stesso stile dell'originale
         */
        ComplexTemplateSetting<MagicParserDto> templateSetting1 = new ComplexTemplateSetting<>();
        templateSetting1.setMagicParser(new ComplexParser());
        templateSetting1.setObject(magicParserDto);
        templateSetting1.setCopyStyle(true);

        List<AereoDto> aereoDtoList = new ArrayList<>();
        aereoDtoList.add(new AereoDto("A", "B", "C", "D", "E", "F", "G", "Z", "I", "L", "M", "N", "O", "P", "Q", "R", "P123"));

        /** IMPOSTAZIONE 2
         *  1. Seleziono il SimpleDuplicatorParser per poter riempire celle e duplicarle se mancano
         *  2. Lista degli elementi da inserire nel report
         *  3. L'indice del foglio che dobbiamo analizzare (dato che andremo a duplicare, e quindi modificare, il foglio è necessario specificare il suo indice) // TODO: farlo funzionare senza specificare? Magari con una string key
         *  4. La prima riga della porzione da duplicare, nel nostro esempio dobbiamo partire dalla riga 10
         *  5. Solo la riga stessa da duplicare, quindi 10
         *  6. Scrivere il valore delle celle verso il basso
         *  7. Porzione da 1 -> il secondo elemento viene duplicato sotto TODO: vorrei rendere automatico il processo senza specificare
         *  8. Le righe duplicate avranno lo stesso stile dell'originale
         */
        SimpleDuplicatorTemplateSetting<AereoDto> templateSetting2 = new SimpleDuplicatorTemplateSetting<>();
        templateSetting2.setMagicParser(new SimpleDuplicatorParser());
        templateSetting2.setObjectList(aereoDtoList);
        templateSetting2.setSheetIndex(0);
        templateSetting2.setFirstRowToDuplicateIndex(26);
        templateSetting2.setLastRowToDuplicateIndex(26);
        templateSetting2.setDirection(BOTTOM);
        templateSetting2.setObjectListPortion(1);
        templateSetting2.setCopyStyle(true);

        List<TemplateSetting> templateSettingList = new ArrayList<>();
        templateSettingList.add(templateSetting1);
        templateSettingList.add(templateSetting2);

        InputStreamResource file = new InputStreamResource(magicParserService.exportExcel(templatePath, 0, templateSettingList, List.of()));
        return file;
    }

    public ResponseEntity<Resource> report3() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        String reportPath = "Report3";
        InputStreamResource file = getReport3();
        return getResponse(file, reportPath);
    }

    public ResponseEntity<byte[]> report3Pdf() throws Exception {
        String reportPath = "Report3";
        InputStreamResource file = getReport3();
        return getResponsePdf(file, reportPath);
    }

//
///*** REPORT IAC: TEMPLATE STAZIONI ***/
//    public InputStreamResource getFileStazioni() throws IOException, InvocationTargetException, IllegalAccessException, ExcelException {
//
//        String templatePath = "src/main/resources/template/Template_Report_Stazioni2.xlsx";
//
//        /** IMPOSTAZIONE 1
//         *  1. Seleziono il ComplexParser per poter riempire celle di formato <symbol valore>
//         *  2. Elemento da inserire nel report (i valori dei suoi attributi sono il contenuto delle celle)
//         *  3. Le righe che riempio avranno lo stesso stile dell'originale
//         */
//        ComplexTemplateSetting<InfoTitoloDto> templateSetting1 = new ComplexTemplateSetting<>();
//        templateSetting1.setMagicParser(new ComplexParser());
//        templateSetting1.setObject(); // templateSetting1.setObject(reportStazioniDto1);
////        templateSetting1.setSymbol("*"); // di default
//        templateSetting1.setCopyStyle(true);
//
//        /** IMPOSTAZIONE 2
//         *  1. Seleziono il SimpleDuplicatorParser per poter riempire celle e duplicarle se mancano
//         *  2. Lista degli elementi da inserire nel report
//         *  3. L'indice del foglio che dobbiamo analizzare (dato che andremo a duplicare, e quindi modificare, il foglio è necessario specificare il suo indice) // TODO: farlo funzionare senza specificare? Magari con una string key
//         *  4. La prima riga della porzione da duplicare, nel nostro esempio dobbiamo partire dalla riga 10
//         *  5. Solo la riga stessa da duplicare, quindi 10
//         *  6. Scrivere il valore delle celle verso il basso
//         *  7. Porzione da 1 -> il secondo elemento viene duplicato sotto TODO: vorrei rendere automatico il processo senza specificare
//         *  8. Le righe duplicate avranno lo stesso stile dell'originale
//         */
//        SimpleDuplicatorTemplateSetting<ReportStazioniDto> templateSetting2 = new SimpleDuplicatorTemplateSetting<>();
//        templateSetting2.setMagicParser(new SimpleDuplicatorParser());
//        templateSetting2.setObjectList(reportStazioniDtoList);
//        templateSetting2.setSheetIndex(2);
//        templateSetting2.setFirstRowToDuplicateIndex(10);
//        templateSetting2.setLastRowToDuplicateIndex(10);
//        templateSetting2.setDirection(BOTTOM);
//        templateSetting2.setObjectListPortion(1);
//        templateSetting2.setCopyStyle(true);
//
//        /** IMPOSTAZIONE 3 (PER OGNI SEZIONE)
//         *  1. Seleziono il ComplexDuplicatorParser per poter riempire celle di formato <symbol valore> in un certo range
//         *  2. Lista degli elementi da inserire nel report
//         *  3. L'indice del foglio che dobbiamo analizzare (dato che andremo a duplicare, e quindi modificare, il foglio è necessario specificare il suo indice) // TODO: farlo funzionare senza specificare? Magari con una string key
//         *  4. La prima riga della porzione da considerare, nel nostro esempio dobbiamo partire dalla riga 18
//         *  5. ...alla riga 23
//         *  6 ...
//         *  7. Le righe che riempio avranno lo stesso stile dell'originale
//         */
//        ComplexDuplicatorTemplateSetting<ReportStazioniSezioneDto> templateSetting3 = new ComplexDuplicatorTemplateSetting<>();
//        templateSetting3.setMagicParser(new ComplexDuplicatorParser());
//        templateSetting3.setObjectList(reportStazioniSezioneDtoList);
//        templateSetting3.setSheetIndex(1);
//        templateSetting3.setFirstRowIndex(18);
//        templateSetting3.setLastRowIndex(23);
////        templateSetting5.setGap(1); // gap è 0 perchè la porzione duplicata sarà subito sotto
//        templateSetting3.setCopyStyle(true);
//
//        List<TemplateSetting> templateSettingList = new ArrayList<>();
//        templateSettingList.add(templateSetting1);
//        templateSettingList.add(templateSetting2);
//        templateSettingList.add(templateSetting3);
//
//        InputStreamResource file = new InputStreamResource(magicParserService.exportExcel(templatePath, -1, templateSettingList, List.of(1, 2)));
//        return file;
//    }
//
//    public ResponseEntity<Resource> creaExcelStazioni(InputReportIacStazioniDto inputReportIacStazioniDto) throws Exception {
//        String reportPath = "Report_Stazioni_" + inputReportIacStazioniDto.getCodiceCantiere();
//        InputStreamResource file = getFileStazioni(inputReportIacStazioniDto);
//        return getResponse(file, reportPath);
//    }
//
//    public ResponseEntity<byte[]> creaPdfStazioni(InputReportIacStazioniDto inputReportIacStazioniDto) throws Exception {
//        String reportPath = "Report_Stazioni_" + inputReportIacStazioniDto.getCodiceCantiere();
//        InputStreamResource file = getFileStazioni(inputReportIacStazioniDto);
//        return getResponsePdf(file, reportPath); // pagina 1 e 2 landscape
//    }


/*** REPORT IAC: TEMPLATE ELETTR AEREO ***/
//    public InputStreamResource getFileElettrAereo() throws IOException, InvocationTargetException, IllegalAccessException, ExcelException {
//
//        String templatePath = "src/main/resources/templateExcel/Template_Report_ElettrAereo2.xlsx";
//
//        ReportElettrAereoDto reportElettrAereoDto1 = new ReportElettrAereoDto();
//        reportElettrAereoDto1.setNote("Nota 1");
//        reportElettrAereoDto1.setAtpTotale("ATP1 TOTALE 1");
//        reportElettrAereoDto1.setAtpPeriodo("ATP1 PERIODO 1");
//        DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
//        Calendar calendar = Calendar.getInstance();
//        Date today = calendar.getTime();
//        calendar.add(Calendar.DATE, -1);
//        Date yesterday = calendar.getTime();
//        reportElettrAereoDto1.setAvanzamentoDa(dateFormat.format(yesterday).toString());
//        reportElettrAereoDto1.setAvanzamentoA(dateFormat.format(today).toString());
//        reportElettrAereoDto1.setCodiceCantiere("CODICE CANTIERE 1");
//        reportElettrAereoDto1.setWbs("WBS 1");
//        reportElettrAereoDto1.setConsistenzaIntervento("CONSISTENZA 1");
//        reportElettrAereoDto1.setOdaLda("ODA 1");
//        reportElettrAereoDto1.setMainContractor("MAIN 1");
//        reportElettrAereoDto1.setNomeCantiere("NOME BELLO");
//        reportElettrAereoDto1.setDtRiferimento("DT 1");
//        reportElettrAereoDto1.setCommittente("COMMITENTE 1");
//        reportElettrAereoDto1.setRpe("RPE 1");
//        reportElettrAereoDto1.setDl("DIRETTORE 1");
//        reportElettrAereoDto1.setCsp("CSP 1");
//        reportElettrAereoDto1.setCse("CSE 1");
//        reportElettrAereoDto1.setProgettista("PROGETTISTA 1");
//        reportElettrAereoDto1.setCollaudatore("COLL 1");
//        reportElettrAereoDto1.setIac("IAC 1");
//        reportElettrAereoDto1.setIaca("IACA 1");
//        reportElettrAereoDto1.setRss("RSS1");
//        reportElettrAereoDto1.setSorvegliante("SORVEGLIANTE");
//        reportElettrAereoDto1.setRr("RR 1");
//        reportElettrAereoDto1.setDataInizioCantiere(new Date());
//        reportElettrAereoDto1.setDataFineCantiere(new Date());
//        reportElettrAereoDto1.setDataPresuntaInizioCantiere(new Date());
//        reportElettrAereoDto1.setDataPresuntaFineCantiere(new Date());
//        reportElettrAereoDto1.setRegione("REGIONE 1");
//        reportElettrAereoDto1.setComune("COMUNE 1");
//        reportElettrAereoDto1.setIndirizzoLoc("INDIRIZZO LOC. 1");
//        reportElettrAereoDto1.setAppaltoLavoriConsuntivato("APPALTO CONSUNTIVATO 1");
//        reportElettrAereoDto1.setAppaltoLavoriPresunto("APPALTO PRESUNTO 1");
//        reportElettrAereoDto1.setAppaltoLavoriP("APPALTO % 1");
//        reportElettrAereoDto1.setBobPeriodo("BOB PERIODO 1");
//        reportElettrAereoDto1.setBobTotale("BOB TOTALE 1");
//        reportElettrAereoDto1.setTrasportoPeriodo("TRASPORTO PERIODO 1");
//        reportElettrAereoDto1.setTrasportoTotale("TRASPORTO TOTALE 1");
//        reportElettrAereoDto1.setPortaleEsistente("PORTALE ESISTENTE 1");
//        reportElettrAereoDto1.setSottoFondazioniTotale("SOTTOFONDAZIONE TOTALE 1");
//        reportElettrAereoDto1.setSottoFondazioniPeriodo("SOTTOFONDAZIONE PERIODO 1");
//        reportElettrAereoDto1.setOggettoAppaltato("OGGETTO APPALTATO 1");
//        reportElettrAereoDto1.setNomeFornitore("NOME FORNITORE 1");
//        reportElettrAereoDto1.setNOda("NODA 1");
//        reportElettrAereoDto1.setImportoAppaltatore("IMPORTO APPALTATORE 1");
//        reportElettrAereoDto1.setRegioneSociale("REGIONE SOCIALE 1");
//        reportElettrAereoDto1.setImportoSubappaltatore("IMPORTO SUBAPPALTATORE 1");
//        ReportElettrAereoDto reportElettrAereoDto2 = new ReportElettrAereoDto();
//        reportElettrAereoDto2.setNote("Nota 2");
//        reportElettrAereoDto2.setAtpTotale("ATP1 TOTALE 2");
//        reportElettrAereoDto2.setAtpPeriodo("ATP1 PERIODO 2");
//        reportElettrAereoDto2.setCodiceCantiere("CODICE CANTIERE 2");
//        reportElettrAereoDto2.setOggettoAppaltato("OGGETTO APPALTATO 2");
//        reportElettrAereoDto2.setNomeFornitore("NOME FORNITORE 2");
//        reportElettrAereoDto2.setNOda("NODA 2");
//        reportElettrAereoDto2.setImportoAppaltatore("IMPORTO APPALTATORE 2");
//        reportElettrAereoDto2.setRegioneSociale("REGIONE SOCIALE 2");
//        reportElettrAereoDto2.setImportoSubappaltatore("IMPORTO SUBAPPALTATORE 2");
//
//        List<ReportElettrAereoDto> reportElettrAereoDtoList = new ArrayList<>();
//        reportElettrAereoDtoList.add(reportElettrAereoDto1);
//        reportElettrAereoDtoList.add(reportElettrAereoDto2);
//
//        ReportElettrAereoPicchettoDto reportElettrAereoPicchettoDto1 = new ReportElettrAereoPicchettoDto();
//        reportElettrAereoPicchettoDto1.setArcheologia("ARCH 1");
//        reportElettrAereoPicchettoDto1.setBob("BOB 1");
//        reportElettrAereoPicchettoDto1.setAsservimento("ASS 1");
//        reportElettrAereoPicchettoDto1.setAutorizzazione("AUT 1");
//        reportElettrAereoPicchettoDto1.setProcedibilitaAutorizzativa("PROC AUT 1");
//        reportElettrAereoPicchettoDto1.setScaviPerforazioni("Scavi 1");
//        reportElettrAereoPicchettoDto1.setFondazione("FOND 1");
//        reportElettrAereoPicchettoDto1.setSottoFondazione("SOTTO FOND 1");
//        reportElettrAereoPicchettoDto1.setMontaggio("MONT 1");
//        reportElettrAereoPicchettoDto1.setRinterro("RINT 1");
//        reportElettrAereoPicchettoDto1.setRipristino("RIPRISTINO 1");
//
//        ReportElettrAereoCampataDto reportElettrAereoCampataDto1 = new ReportElettrAereoCampataDto();
//        reportElettrAereoCampataDto1.setStendimento("STEND 1");
//        reportElettrAereoCampataDto1.setAsservimentoCampata("ASS CAMPATA 1");
//        reportElettrAereoCampataDto1.setAutorizzazioneCampata("AUT CAMPATA 1");
//        reportElettrAereoCampataDto1.setMorsetteriaAccessori("MORS 1");
//        reportElettrAereoCampataDto1.setRegolazione("REG 1");
//
//        List<ReportElettrAereoPicchettoDto> reportElettrAereoPicchettoDtoList = new ArrayList<>();
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        // 21esimo picchetto
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//        // 41esimo picchetto
//        reportElettrAereoPicchettoDtoList.add(reportElettrAereoPicchettoDto1);
//
//        List<ReportElettrAereoCampataDto> reportElettrAereoCampataDtoList = new ArrayList<>();
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        // 20esima campata
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//        // 40 esima campata
//        reportElettrAereoCampataDtoList.add(reportElettrAereoCampataDto1);
//
//        ReportElettrAereoRiepilogoDto reportElettrAereoRiepilogoDto1 = new ReportElettrAereoRiepilogoDto();
//        reportElettrAereoRiepilogoDto1.setAsservimentoAvanzamento("AA 1");
//        reportElettrAereoRiepilogoDto1.setAsservimentoCompletati("AC 1");
//        reportElettrAereoRiepilogoDto1.setAsservimentoDaCompletare("AD 1");
//        reportElettrAereoRiepilogoDto1.setAsservimentoCampataAvanzamento("AA C 1");
//        reportElettrAereoRiepilogoDto1.setAsservimentoCampataCompletati("AC C 1");
//        reportElettrAereoRiepilogoDto1.setAsservimentoCampataDaCompletare("AD C 1");
//        reportElettrAereoRiepilogoDto1.setAutorizzazioneAvanzamento("AUA 1");
//        reportElettrAereoRiepilogoDto1.setAutorizzazioneCompletati("AUC 1");
//        reportElettrAereoRiepilogoDto1.setAutorizzazioneDaCompletare("AUD 1");
//        ReportElettrAereoRiepilogoDto reportElettrAereoRiepilogoDto2 = new ReportElettrAereoRiepilogoDto();
//        reportElettrAereoRiepilogoDto2.setAsservimentoAvanzamento("AA 2");
//        reportElettrAereoRiepilogoDto2.setAsservimentoCompletati("AC 2");
//        reportElettrAereoRiepilogoDto2.setAsservimentoDaCompletare("AD 2");
//        reportElettrAereoRiepilogoDto2.setAsservimentoCampataAvanzamento("AA C 2");
//        reportElettrAereoRiepilogoDto2.setAsservimentoCampataCompletati("AC C 2");
//        reportElettrAereoRiepilogoDto2.setAsservimentoCampataDaCompletare("AD C 2");
//        reportElettrAereoRiepilogoDto2.setAutorizzazioneAvanzamento("AUA 2");
//        reportElettrAereoRiepilogoDto2.setAutorizzazioneCompletati("AUC 2");
//        reportElettrAereoRiepilogoDto2.setAutorizzazioneDaCompletare("AUD 2");
//
//        // ovviamente le dimensioni devono dipendere dalle campate/picchetti, se ci sono 21 picchetti allora la lista sarà di 2, se ce ne sono 41 di 3 ...
//        List<ReportElettrAereoRiepilogoDto> reportElettrAereoRiepilogoDtoList = new ArrayList<>();
//        reportElettrAereoRiepilogoDtoList.add(reportElettrAereoRiepilogoDto1);
//        reportElettrAereoRiepilogoDtoList.add(reportElettrAereoRiepilogoDto2);
//        reportElettrAereoRiepilogoDtoList.add(reportElettrAereoRiepilogoDto1);
//
//        /** IMPOSTAZIONE 1
//         *  1. Seleziono il ComplexParser per poter riempire celle di formato <symbol valore>
//         *  2. Elemento da inserire nel report (i valori dei suoi attributi sono il contenuto delle celle)
//         *  3. Le righe che riempio avranno lo stesso stile dell'originale
//         */
//        ComplexTemplateSetting<ReportElettrAereoDto> templateSetting1 = new ComplexTemplateSetting<>();
//        templateSetting1.setMagicParser(new ComplexParser());
//        templateSetting1.setObject(reportElettrAereoDto1);
////        templateSetting1.setSymbol("*"); // di default
//        templateSetting1.setCopyStyle(true);
//
//        /** IMPOSTAZIONE 2 (APPALTO/FORNITORE - SUBAPPALTATORE) SE AVESSIMO TUTTO IN UN FOGLIO ALLORA QUESTA IMPOSTAZIONE DOVREBBE ESSERE INSERITA PRIMA DI QUELLE PER I PICCHETTI E CAMPATE: VANNO A MODIFICARE GLI INDICI DELLE RIGHE
//         *  1. Seleziono il SimpleDuplicatorParser per poter riempire celle e duplicarle se mancano
//         *  2. Lista degli elementi da inserire nel report
//         *  3. L'indice del foglio che dobbiamo analizzare (dato che andremo a duplicare, e quindi modificare, il foglio è necessario specificare il suo indice) // TODO: farlo funzionare senza specificare? Magari con una string key
//         *  4. La prima riga della porzione da duplicare, nel nostro esempio dobbiamo partire dalla riga 38
//         *  5. Solo la riga stessa da duplicare, quindi 38
//         *  6. Scrivere il valore delle celle verso il basso
//         *  7. Porzione da 1 -> il secondo elemento viene duplicato sotto TODO: vorrei rendere automatico il processo senza specificare
//         *  8. Le righe duplicate avranno lo stesso stile dell'originale
//         */
//        SimpleDuplicatorTemplateSetting<ReportElettrAereoDto> templateSetting2 = new SimpleDuplicatorTemplateSetting<>();
//        templateSetting2.setMagicParser(new SimpleDuplicatorParser());
//        templateSetting2.setObjectList(reportElettrAereoDtoList);
//        templateSetting2.setSheetIndex(2);
//        templateSetting2.setFirstRowToDuplicateIndex(38);
//        templateSetting2.setLastRowToDuplicateIndex(38);
//        templateSetting2.setDirection(BOTTOM);
//        templateSetting2.setObjectListPortion(1);
//        templateSetting2.setCopyStyle(true);
//
//        /** IMPOSTAZIONE 3 (PICCHETTI)
//         *  1. Seleziono il SimpleDuplicatorParser per poter riempire celle e duplicarle se mancano
//         *  2. Lista degli elementi da inserire nel report
//         *  3. L'indice del foglio che dobbiamo analizzare (dato che andremo a duplicare, e quindi modificare, il foglio è necessario specificare il suo indice) // TODO: farlo funzionare senza specificare? Magari con una string key
//         *  4. La prima riga della porzione da duplicare, nel nostro esempio dobbiamo partire dalla riga 7
//         *  5. ..alla riga 18
//         *  6. Duplicazione con un certo gap, cioè un salto di 7 celle sotto (in questo esempio porzione occupata dalle campate) // TODO: dato il differente numero di elementi nelle liste ho dovuto gestire le porzioni come due cose diverse e fare dei salti... vorrei considerare la porzione (picchetti campate) come unica e gestirla a scacchi
//         *  7. Scrivere il valore delle celle verso destra
//         *  8. Scrivere il valore delle celle facendo salti di 2
//         *  9. Tutte le celle raggruppate logicamente come "PICCHETTO" nel DTO voglio gestirle diversamente
//         *  10. Le celle del gruppo logico "PICCHETTO" le sposto una cella indietro (questo permette un layout a scacchi con le campate)
//         *  11. Porzione da 20 -> il 21-esimo elemento viene duplicato sotto
//         *  12. Le righe che duplico avranno lo stesso stile dell'originale
//         */
//        SimpleDuplicatorTemplateSetting<ReportElettrAereoPicchettoDto> templateSetting3 = new SimpleDuplicatorTemplateSetting<>();
//        templateSetting3.setMagicParser(new SimpleDuplicatorParser());
//        templateSetting3.setObjectList(reportElettrAereoPicchettoDtoList);
//        templateSetting3.setSheetIndex(1);
//        templateSetting3.setFirstRowToDuplicateIndex(7);
//        templateSetting3.setLastRowToDuplicateIndex(18);
//        templateSetting3.setGap(7);
//        templateSetting3.setDirection(RIGHT);
//        templateSetting3.setSteps(2);
//        templateSetting3.setHeaderGroup("PICCHETTO");
//        templateSetting3.setHeaderGroupStartIndex(-1);
//        templateSetting3.setObjectListPortion(20);
//        templateSetting3.setCopyStyle(true);
//        templateSetting3.setMaxObjectForPage(40);
//
//        /** IMPOSTAZIONE 4 (CAMPATE)
//         *  1. Seleziono il SimpleDuplicatorParser per poter riempire celle e duplicarle se mancano
//         *  2. Lista degli elementi da inserire nel report
//         *  3. L'indice del foglio che dobbiamo analizzare (dato che andremo a duplicare, e quindi modificare, il foglio è necessario specificare il suo indice) // TODO: farlo funzionare senza specificare? Magari con una string key
//         *  4. La prima riga della porzione da duplicare, nel nostro esempio dobbiamo partire dalla riga 19
//         *  5. ..alla riga 25
//         *  6. Duplicazione con un certo gap, cioè un salto di 12 celle sotto (in questo esempio porzione occupata dai picchetti) // TODO: dato il differente numero di elementi nelle liste ho dovuto gestire le porzioni come due cose diverse e fare dei salti... vorrei considerare la porzione (picchetti campate) come unica e gestirla a scacchi
//         *  7. Scrivere il valore delle celle verso destra
//         *  8. Scrivere il valore delle celle facendo salti di 2
//         *  9. Porzione da 19 -> il 20-esimo elemento viene duplicato sotto
//         *  10. Le righe che duplico avranno lo stesso stile dell'originale
//         */
//        // ho dovuto dividere in due dto e non usare il layout a scacchi con un singolo DTO perchè campata ha autorizzazione e asservimento field anche di picchetto, inoltre le campate sono 19 e non 20, quindi per forza, TODO: vorrei non essere costretto a dividere in due dto nel caso di colonne uguali
//        SimpleDuplicatorTemplateSetting<ReportElettrAereoCampataDto> templateSetting4 = new SimpleDuplicatorTemplateSetting<>();
//        templateSetting4.setMagicParser(new SimpleDuplicatorParser());
//        templateSetting4.setObjectList(reportElettrAereoCampataDtoList);
//        templateSetting4.setSheetIndex(1);
//        templateSetting4.setFirstRowToDuplicateIndex(19);
//        templateSetting4.setLastRowToDuplicateIndex(25);
//        templateSetting4.setGap(12);
//        templateSetting4.setDirection(RIGHT);
//        templateSetting4.setSteps(2);
//        templateSetting4.setObjectListPortion(19);
//        templateSetting4.setCopyStyle(true);
//        templateSetting4.setMaxObjectForPage(38);
//
//        /** IMPOSTAZIONE 5 (RIEPILOGO)
//         *  1. Seleziono il ComplexDuplicatorParser per poter riempire celle di formato <symbol valore> in un certo range
//         *  2. Lista degli elementi da inserire nel report
//         *  3. L'indice del foglio che dobbiamo analizzare (dato che andremo a duplicare, e quindi modificare, il foglio è necessario specificare il suo indice) // TODO: farlo funzionare senza specificare? Magari con una string key
//         *  4. La prima riga della porzione da considerare, nel nostro esempio dobbiamo partire dalla riga 8
//         *  5. ...alla riga 25
//         *  6 ...
//         *  7. Le righe che riempio avranno lo stesso stile dell'originale
//         */
//        ComplexDuplicatorTemplateSetting<ReportElettrAereoRiepilogoDto> templateSetting5 = new ComplexDuplicatorTemplateSetting<>();
//        templateSetting5.setMagicParser(new ComplexDuplicatorParser());
//        templateSetting5.setObjectList(reportElettrAereoRiepilogoDtoList);
//        templateSetting5.setSheetIndex(1);
//        templateSetting5.setFirstRowIndex(8);
//        templateSetting5.setLastRowIndex(25);
//        templateSetting5.setDuplicate(false);
////        templateSetting5.setFirstColumn("AP");
////        templateSetting5.setGap(1); // per essere precisi con le porzioni da analizzare dovrei settare un gap, ma tanto...
//        templateSetting5.setCopyStyle(true);
//
//        List<TemplateSetting> templateSettingList = new ArrayList<>();
//        templateSettingList.add(templateSetting1);
//        templateSettingList.add(templateSetting2);
//        templateSettingList.add(templateSetting3);
//        templateSettingList.add(templateSetting4);
////        templateSettingList.add(templateSetting5);
//
//        InputStreamResource file = new InputStreamResource(magicParserService.exportExcel(templatePath, -1, templateSettingList, List.of(1)));
//        return file;
//    }
//
//    public ResponseEntity<Resource> creaExcelElettrAereo(Long idCantiere, String numOda) throws Exception {
//        String reportPath = "Report_ElettrAereo_" + idCantiere + "_" + numOda;
//        InputStreamResource file = getFileElettrAereo();
//        return getResponse(file, reportPath);
//    }
//
//    public ResponseEntity<byte[]> creaPdfElettrAereo(Long idCantiere, String numOda) throws Exception {
//        String reportPath = "Report_ElettrAereo_" + idCantiere + "_" + numOda;
//        InputStreamResource file = getFileElettrAereo();
//        return getResponsePdf(file, reportPath); // pagina 1 landscape
//    }
//
//
//
///*** REPORT IAC: TEMPLATE ELENCO AUTORIZZATO CANTIERE ***/
//    public InputStreamResource getFileElencoAutorizzatoCantiere() throws Exception {
//
//        String targetSheetPathAndName = "src/main/resources/templateExcel/ElencoAutorizzatoCantiere.xlsx";
//
//        ContrattiAppaltatoriDto contrattiAppaltatoriDto = new ContrattiAppaltatoriDto(2L, "descrizioneA1", "fornitoreA1", 1L, "numOdaA1", true);
//        // PERSONALE
//        List<VPersonaleDto> vPersonaleDtoList = new ArrayList<>();
//        VPersonaleDto p1 = new VPersonaleDto("CF1", "A", 1L, 1L, 1L, false, "A", "A", 1L);
//        VPersonaleDto p2 = new VPersonaleDto("CF2", "B", 2L, 1L, 1L, false, "B", "B", 1L);
//        VPersonaleDto p3 = new VPersonaleDto("CF3", "C", 1L, 1L, 1L, true, "C", "C", 1L);
//        VPersonaleDto p4 = new VPersonaleDto("CF4", "D", 2L, 1L, 1L, false, "D", "D", 1L);
//        vPersonaleDtoList.add(p1);
//        vPersonaleDtoList.add(p2);
//        vPersonaleDtoList.add(p3);
//
//        VSubappaltatoreCantiereDto vSubappaltatoreCantiereDto1 = new VSubappaltatoreCantiereDto("codeCantiere1", "codeId1", "denomImp1", 1L, 1L, "numOda1");
//        VSubappaltatoreCantiereDto vSubappaltatoreCantiereDto2 = new VSubappaltatoreCantiereDto("codeCantiere2", "codeId2", "denomImp2", 2L, 2L, "numOda2");
//        VSubappaltatoreCantiereDto vSubappaltatoreCantiereDto3 = new VSubappaltatoreCantiereDto("codeCantiere3", "codeId3", "denomImp3", 3L, 3L, "numOda3");
//        VSubappaltatoreCantiereDto vSubappaltatoreCantiereDto4 = new VSubappaltatoreCantiereDto("codeCantiere4", "codeId4", "denomImp4", 4L, 4L, "numOda4");
//        VSubappaltatoreCantiereDto vSubappaltatoreCantiereDto5 = new VSubappaltatoreCantiereDto("codeCantiere5", "codeId5", "denomImp5", 5L, 5L, "numOda5");
//        VSubappaltatoreCantiereDto vSubappaltatoreCantiereDto6 = new VSubappaltatoreCantiereDto("codeCantiere6", "codeId6", "denomImp6", 6L, 6L, "numOda6");
//        // PERSONALE
//        List<VPersonaleSubappDto> vPersonaleSubappDtoList1 = new ArrayList<>();
//        VPersonaleSubappDto s1 = new VPersonaleSubappDto("S_CF1", "S_COGNOME1", 1L, 1L, 1L, false, "mansione1", "nome1", 30L);
//        VPersonaleSubappDto s2 = new VPersonaleSubappDto("S_CF2", "S_COGNOME2", 2L, 1L, 2L, false, "mansione2", "nome2", 2L);
//        VPersonaleSubappDto s3 = new VPersonaleSubappDto("S_CF3", "S_COGNOME3", 1L, 1L, 3L, true, "mansione3", "nome3", 34L);
//        VPersonaleSubappDto s4 = new VPersonaleSubappDto("S_CF4", "S_COGNOME4", 2L, 1L, 4L, false, "mansione4", "nome4", 21L);
//        vPersonaleSubappDtoList1.add(s1);
//        vPersonaleSubappDtoList1.add(s2);
//        vPersonaleSubappDtoList1.add(s3);
//        vPersonaleSubappDtoList1.add(s4);
//        vPersonaleSubappDtoList1.add(s2);
//        List<VPersonaleSubappDto> vPersonaleSubappDtoList2 = new ArrayList<>();
//        vPersonaleSubappDtoList2.add(s1);
//        vPersonaleSubappDtoList2.add(s2);
//        List<VPersonaleSubappDto> vPersonaleSubappDtoList3 = new ArrayList<>();
//        vPersonaleSubappDtoList3.add(s4);
//        List<VPersonaleSubappDto> vPersonaleSubappDtoList4 = new ArrayList<>();
//        vPersonaleSubappDtoList4.add(s4);
//        vPersonaleSubappDtoList4.add(s1);
//        vPersonaleSubappDtoList4.add(s2);
//        vPersonaleSubappDtoList4.add(s4);
//
//        List<AppaltatorePersonaleDto> appaltatorePersonaleDtoList = new ArrayList<>();
//        AppaltatorePersonaleDto appaltatorePersonaleDto1 = new AppaltatorePersonaleDto(contrattiAppaltatoriDto, vPersonaleDtoList);
//        appaltatorePersonaleDtoList.add(appaltatorePersonaleDto1);
//
//        List<SubappaltatorePersonaleDto> subappaltatorePersonaleDtoList = new ArrayList<>();
//        SubappaltatorePersonaleDto subappaltatorePersonaleDto1 = new SubappaltatorePersonaleDto(vSubappaltatoreCantiereDto1, vPersonaleSubappDtoList1);
//        SubappaltatorePersonaleDto subappaltatorePersonaleDto2 = new SubappaltatorePersonaleDto(vSubappaltatoreCantiereDto2, vPersonaleSubappDtoList2);
//        SubappaltatorePersonaleDto subappaltatorePersonaleDto3 = new SubappaltatorePersonaleDto(vSubappaltatoreCantiereDto3, vPersonaleSubappDtoList3);
//        SubappaltatorePersonaleDto subappaltatorePersonaleDto4 = new SubappaltatorePersonaleDto(vSubappaltatoreCantiereDto4, vPersonaleSubappDtoList4);
//        SubappaltatorePersonaleDto subappaltatorePersonaleDto5 = new SubappaltatorePersonaleDto(vSubappaltatoreCantiereDto5, new ArrayList<>());
//        SubappaltatorePersonaleDto subappaltatorePersonaleDto6 = new SubappaltatorePersonaleDto(vSubappaltatoreCantiereDto6, new ArrayList<>());
//        subappaltatorePersonaleDtoList.add(subappaltatorePersonaleDto1);
//        subappaltatorePersonaleDtoList.add(subappaltatorePersonaleDto2);
//        subappaltatorePersonaleDtoList.add(subappaltatorePersonaleDto6);
//        subappaltatorePersonaleDtoList.add(subappaltatorePersonaleDto3);
//        subappaltatorePersonaleDtoList.add(subappaltatorePersonaleDto6);
//        subappaltatorePersonaleDtoList.add(subappaltatorePersonaleDto4);
//        subappaltatorePersonaleDtoList.add(subappaltatorePersonaleDto5);
//
//        /** IMPOSTAZIONE 1
//         *  1. Seleziono il CantiereParser per poter riempire quelle porzioni di formato (lista elementi dove ogni elemento ha una lista)
//         *  2. Classe dell'oggetto nella prima lista
//         *  3. Lista degli elementi da inserire nel report
//         *  4. Cerca la stringa nell'excel per avere l'indice della riga
//         *  5. Le righe che riempio avranno lo stesso stile dell'originale
//         *  6. Scrivere il valore delle celle verso il basso
//         */
//        CantiereTemplateSetting<AppaltatorePersonaleDto> templateSetting1 = new CantiereTemplateSetting<>();
//        templateSetting1.setMagicParser(new CantiereParser()); // TODO: rinominare per renderlo generico e utilizzabile su altri template
//        templateSetting1.setObjectClass(ContrattiAppaltatoriDto.class);
//        templateSetting1.setObjectList(appaltatorePersonaleDtoList); // 1 appaltatore con una lista personale di 3 elementi
//        templateSetting1.setKey("appaltatore:");
//        templateSetting1.setCopyStyle(true);
//        templateSetting1.setDirection(BOTTOM);
//
//        /** IMPOSTAZIONE 2
//         *  1. Seleziono il CantiereParser per poter riempire quelle porzioni di formato (lista elementi dove ogni elemento ha una lista)
//         *  2. Classe dell'oggetto nella prima lista
//         *  3. Lista degli elementi da inserire nel report
//         *  4. Cerca la stringa nell'excel per avere l'indice della riga
//         *  5. Le righe che riempio avranno lo stesso stile dell'originale
//         *  6. Scrivere il valore delle celle verso il basso
//         */
//        CantiereTemplateSetting<SubappaltatorePersonaleDto> templateSetting2 = new CantiereTemplateSetting<>();
//        templateSetting2.setMagicParser(new CantiereParser());
//        templateSetting2.setObjectClass(VSubappaltatoreCantiereDto.class);
//        templateSetting2.setObjectList(subappaltatorePersonaleDtoList); // 7 dubappaltatori dove ognuno ha un diverso numero di personale
//        templateSetting2.setKey("subappaltatore:");
//        templateSetting2.setCopyStyle(true);
//        templateSetting2.setDirection(BOTTOM);
//
//        List<TemplateSetting> templateSettingList = new ArrayList<>();
//        templateSettingList.add(templateSetting1);
//        templateSettingList.add(templateSetting2);
//
//        InputStreamResource file = new InputStreamResource(magicParserService.exportExcel(targetSheetPathAndName, -1, templateSettingList, List.of()));
//        return file;
//    }
//
//    public ResponseEntity<Resource> creaExcelElencoAutorizzatoCantiere(Long idCantiere, String numOda) throws Exception {
//        String reportPath = "Report_ElencoAutorizzatoCantiere_" + idCantiere + "_" + numOda;
//        InputStreamResource file = getFileElencoAutorizzatoCantiere();
//        return getResponse(file, reportPath);
//    }
//
//    public ResponseEntity<byte[]> creaPdfElencoAutorizzatoCantiere(Long idCantiere, String numOda) throws Exception {
//        String reportPath = "Report_ElencoAutorizzatoCantiere_" + idCantiere + "_" + numOda;
//        InputStreamResource file = getFileElencoAutorizzatoCantiere();
//        return getResponsePdf(file, reportPath);
//    }
}
