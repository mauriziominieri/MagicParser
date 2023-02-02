package org.parser.excel;

import lombok.*;

@NoArgsConstructor
@AllArgsConstructor
@Getter
@Setter
public class XLSXComplexTemplateSetting<T> extends XLSXTemplateSetting {
    private T object; // oggetto poi da scrivere nell'excel
    private String symbol = "*"; // simbolo da utilizzare per la ricerca delle celle nell'excel <symbol valore>
}
