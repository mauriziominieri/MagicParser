package org.parser.excel;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Service
public class MagicParserService {

    // con doppio tipo
//    public <T, P> ByteArrayInputStream exportExcel(Class<T> cls, Class<P> cls2, XLSXParser2<T, P> parser, List<T> objList, List<P> objList2, int numSheet, String path) throws Exception {
//        return parser.prova1(cls, cls2, objList, objList2, numSheet, path);
//    }

    public ByteArrayInputStream exportExcel(String templatePath, int sheetIndex, List<TemplateSetting> templateSettingList) throws IOException, InvocationTargetException, IllegalAccessException, ExcelException {
        return new MagicParser().write(templatePath, sheetIndex, templateSettingList);
    }

//    public <T> List<T> importExcel(MultipartFile file, XLSXParser2<T> parser, Class<T> cls, int numSheet, int headerRow) throws Exception {
//        return parser.fromExcelToObj(file, cls, numSheet, headerRow);
//    }
}
