package org.parser.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

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
public class AppaltatorePersonaleDto {
    ContrattiAppaltatoriDto contrattiAppaltatoriDto;
    List<VPersonaleDto> vPersonaleDtoList;
}
