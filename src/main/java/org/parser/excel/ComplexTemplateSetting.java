package org.parser.excel;

import lombok.*;

@NoArgsConstructor
@AllArgsConstructor
@Data
public class ComplexTemplateSetting<T> extends TemplateSetting {
    private T object; // oggetto poi da scrivere nell'excel
    private String symbol = "*"; // simbolo da utilizzare per la ricerca delle celle nell'excel <symbol valore>
}
