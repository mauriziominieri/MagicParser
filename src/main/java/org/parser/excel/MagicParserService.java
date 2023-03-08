package org.parser.excel;

import org.springframework.stereotype.Service;

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

    public ByteArrayInputStream exportExcel(String templatePath, int sheetIndex, List<TemplateSetting> templateSettingList, List<Integer> landscapePages) throws IOException, InvocationTargetException, IllegalAccessException, ExcelException {
        return new MagicParser().write(templatePath, sheetIndex, templateSettingList, landscapePages);
    }

//    public <T> List<T> importExcel(MultipartFile file, XLSXParser2<T> parser, Class<T> cls, int numSheet, int headerRow) throws Exception {
//        return parser.fromExcelToObj(file, cls, numSheet, headerRow);
//    }
}
