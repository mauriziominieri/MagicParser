package org.parser.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@NoArgsConstructor
@AllArgsConstructor
@Data
public class SimpleDuplicatorTemplateSetting<T> extends SimpleTemplateSetting<T> {
    private int firstRowToDuplicateIndex; // la prima riga della porzione poi da duplicare e riempire
    private int lastRowToDuplicateIndex; // il numero di righe da duplicare partendo dal firstRowToDuplicateIndex
    private int gap;
    private int sheetIndex; // dato che specifichiamo gli indici delle porzioni da duplicare dobbiamo settare sicuramente l'indice della pagina
    private int maxObjectForPage; // numero massimo di elementi in un foglio
}
