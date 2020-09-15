/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Modelo;
import java.io.*;
import java.util.*;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
/**
 *
 * @author pogbom
 */
public class Modelo {
     Workbook wb;
    //private pantallaPrincipal view;
   // ventaProductos vP = new ventaProductos(view, true);

    public String Importar(File archivo, JTable tablaD) {
        String respuesta = "No se pudo cargar el archivo";
        DefaultTableModel modelo = new DefaultTableModel();
        tablaD.setModel(modelo);
        try {
            wb = WorkbookFactory.create(new FileInputStream(archivo));
            Sheet hoja = wb.getSheetAt(0);
            Iterator filaIterator = hoja.rowIterator();
            int indiceFila = -1;
            while (filaIterator.hasNext()) {
                indiceFila++;
                Row fila = (Row) filaIterator.next();
                Iterator columnaIterator = fila.cellIterator();
                Object[] listaColumna = new Object[5];
                int indiceColumna = -1;
                while (columnaIterator.hasNext()) {
                    indiceColumna++;
                    Cell celda = (Cell) columnaIterator.next();
                    if (indiceFila == 0) {
                        modelo.addColumn(celda.getStringCellValue());
                    } else {
                        if (celda != null) {
                            switch (celda.getCellType()) {
                                case Cell.CELL_TYPE_NUMERIC:
                                    listaColumna[indiceColumna] = (int) Math.round(celda.getNumericCellValue());
                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    listaColumna[indiceColumna] = celda.getStringCellValue();
                                    break;
                                case Cell.CELL_TYPE_BOOLEAN:
                                    listaColumna[indiceColumna] = celda.getBooleanCellValue();
                                    break;
                                default:
                                    listaColumna[indiceColumna] = celda.getDateCellValue();
                                    break;
                            }
                        }
                    }
                }
                if (indiceFila != 0) 
                    modelo.addRow(listaColumna);
            }
            respuesta = "Cargado correctamente";
        } catch (Exception e) {
        }
        return respuesta;
    }

    public String Exportar(File archivo, JTable tablaD) {
        String respuesta = "No se guardo el archivo";
        int numFila = tablaD.getRowCount(), numColumna = tablaD.getColumnCount();
        if (archivo.getName().endsWith("xls")) {
            wb = new HSSFWorkbook();
        } else {
            wb = new XSSFWorkbook();
        }
        Sheet hoja = wb.createSheet("prueba");
        try {
            for (int i = -1; i < numFila; i++) {
                Row fila = hoja.createRow(i + 1);
                for (int j = 0; j < numColumna; j++) {
                    Cell celda = fila.createCell(j);
                    if (i == -1) {
                        celda.setCellValue(String.valueOf(tablaD.getColumnName(j)));
                    } else {
                        celda.setCellValue(String.valueOf(tablaD.getValueAt(i, j)));
                    }
                    wb.write(new FileOutputStream(archivo));
                }
            }
            respuesta = "Guardado correctamente";
        } catch (Exception e) {
        }
        return respuesta;
    }
}
