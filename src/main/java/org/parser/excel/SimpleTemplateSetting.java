package org.parser.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@NoArgsConstructor
@AllArgsConstructor
@Data
public class SimpleTemplateSetting<T> extends TemplateSetting {
    private List<T> objectList; // lista degli oggetti poi da scrivere nell'excel
    private boolean fillDuplicateHeadersCells; // se riempire le celle duplicate, questa impostazione è INCOMPATIBILE con objectListPortion
    private int objectListPortion; // gli oggetti in lista possono essere divisi in porzioni da n, se fillDuplicateCells = true verrà ignorato
    private Direction direction; // fissata la cella header decide in che direzione scrivere il valore
    private int steps = 1; // ogni quante celle riempire in quella direzione (utile per un layout "a scacchi"), di default deve essere 1 in quanto 0 andrebbe a sostituire la cella dell'header
    private String headerGroup; // potrebbe essere utile differenziare alcune celle dell'header in un gruppo
    private int headerGroupStartIndex; // per dare un inizio diverso a quel gruppo
}
