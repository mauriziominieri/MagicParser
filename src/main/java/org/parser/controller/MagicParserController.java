package org.parser.controller;

import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.parser.excel.MagicParserWorker;
import org.parser.exception.ExcelException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Api("REST Controller per il MagicParser")
@RestController
@RequestMapping(value = "/magicParser")
public class MagicParserController {

    @Autowired
    MagicParserWorker magicParserWorker;

    @ApiOperation("Crea e scarica l'excel Report0")
    @GetMapping(value="report0")
    public ResponseEntity<Resource> report0() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        return magicParserWorker.report0();
    }

    @ApiOperation("Crea e scarica l'excel Report1")
    @GetMapping(value="report1")
    public ResponseEntity<Resource> report1() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        return magicParserWorker.report1();
    }

    @ApiOperation("Crea e scarica l'excel Report2")
    @GetMapping(value="report2")
    public ResponseEntity<Resource> report2() throws ExcelException, IOException, InvocationTargetException, IllegalAccessException {
        return magicParserWorker.report2();
    }

    @ApiOperation("Crea e scarica l'excel Report Stazioni")
    @GetMapping(value="stazioni")
    public ResponseEntity<Resource> reportStazioni() throws Exception {
        return magicParserWorker.creaExcelStazioni();
    }

    @ApiOperation("Crea e scarica l'excel Report Aereo")
    @GetMapping(value="aereo")
    public ResponseEntity<Resource> reportAereo() throws Exception {
        return magicParserWorker.creaExcelElettrAereo();
    }

    @ApiOperation("Crea e scarica l'excel Report Cantiere")
    @GetMapping(value="cantiere")
    public ResponseEntity<Resource> reportCantiere() throws Exception {
        return magicParserWorker.creaExcelElencoAutorizzatoCantiere();
    }
}
