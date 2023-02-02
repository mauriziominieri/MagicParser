package org.parser.excel;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.util.List;

//@Service TODO per qualche motivo non funziona quando importato, devo capire il motivo
//public class ExcelService2 {
//
//    // con doppio tipo
////    public <T, P> ByteArrayInputStream exportExcel(Class<T> cls, Class<P> cls2, XLSXParser2<T, P> parser, List<T> objList, List<P> objList2, int numSheet, String path) throws Exception {
////        return parser.prova1(cls, cls2, objList, objList2, numSheet, path);
////    }
//
//    public ByteArrayInputStream exportExcel(String templatePath, int sheetIndex, List<XLSXTemplateSetting> xlsxTemplateSettingList) throws Exception {
//        XLSXParser2 parser = new XLSXParser2();
//        return parser.write(templatePath, sheetIndex, xlsxTemplateSettingList);
//    }
//
//    public <T> List<T> importExcel(MultipartFile file, XLSXParser2<T> parser, Class<T> cls, int numSheet, int headerRow) throws Exception {
//        return parser.fromExcelToObj(file, cls, numSheet, headerRow);
//    }
//}
