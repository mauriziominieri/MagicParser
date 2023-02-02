package org.parser.excel;

import lombok.*;

@NoArgsConstructor
@AllArgsConstructor
@Data
public class XLSXTemplateSetting {
    private XLSXParser2 xlsxParser; // parser da utilizzare per le operazioni excel
    private Class objectClass; // la classe dell'oggetto da usare per le operazioni excel
    private boolean copyStyle; // true per copiare lo stile della cella da sovrascrivere
}
