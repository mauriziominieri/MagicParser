package org.parser.excel;

import lombok.Data;

import java.util.List;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Data
public class ComplexDuplicatorTemplateSetting<T> extends ComplexTemplateSetting<T> {
    private List<T> objectList; // lista degli oggetti poi da scrivere nell'excel
    private String firstColumn; // la prima colonna della porzione da duplicare
    private String lastColumn;  // l'ultima colonna della porzione da duplicare
    private int firstRowIndex;  // l'ultima riga della porzione da duplicare
    private int lastRowIndex;   // l'ultima riga della porzione da duplicare
    private int gap;            // eventuale salto nella duplicazione
    private boolean duplicate = true; // TODO: da rimuovere
    private int sheetIndex;     // dato che specifichiamo gli indici delle porzioni da duplicare dobbiamo settare sicuramente l'indice della pagina
}
