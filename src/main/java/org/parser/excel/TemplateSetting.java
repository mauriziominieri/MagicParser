package org.parser.excel;

import lombok.*;

@NoArgsConstructor
@AllArgsConstructor
@Data
public class TemplateSetting {
    private MagicParser magicParser; // parser da utilizzare per le operazioni excel
    private boolean copyStyle; // true per copiare lo stile della cella da sovrascrivere
}
