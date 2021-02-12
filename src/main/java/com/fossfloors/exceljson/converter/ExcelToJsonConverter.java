package com.fossfloors.exceljson.converter;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import javax.annotation.PostConstruct;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Value;

import com.fossfloors.exceljson.pojo.ExcelWorkbook;
import com.fossfloors.exceljson.pojo.ExcelWorksheet;

public class ExcelToJsonConverter {

  @Value("${excelToJSON.numSheets:0}")
  private int        numSheets;
  @Value("${excelToJSON.rowLimit:0}")
  private int        rowLimit;
  @Value("${excelToJSON.rowOffset:0}")
  private int        rowOffset;

  @Value("${excelToJSON.omitEmpty:false}")
  private boolean    omitEmpty;
  @Value("${excelToJSON.fillColumns:false}")
  private boolean    fillColumns;

  @Value("${excelToJSON.dateFormat}")
  private String     dateFormatStr;

  private DateFormat dateFormat = null;

  @PostConstruct
  public void init() {
    if (dateFormatStr != null && !dateFormatStr.isEmpty()) {
      dateFormat = new SimpleDateFormat(dateFormatStr);
    }
  }

  public ExcelWorkbook convert(String sourceFile) throws InvalidFormatException, IOException {
    ExcelWorkbook book = new ExcelWorkbook();
    Workbook wb = WorkbookFactory.create(new FileInputStream(sourceFile));
    int loopLimit = wb.getNumberOfSheets();

    if (numSheets > 0 && loopLimit > numSheets) {
      loopLimit = numSheets;
    }

    int currentRowOffset = -1;
    int totalRowsAdded = 0;

    for (int i = 0; i < loopLimit; i++) {
      Sheet sheet = wb.getSheetAt(i);

      if (sheet == null) {
        continue;
      }

      ExcelWorksheet tmp = new ExcelWorksheet();
      tmp.setName(sheet.getSheetName());

      for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
        Row row = sheet.getRow(j);

        if (row == null) {
          continue;
        }

        boolean hasValues = false;
        ArrayList<Object> rowData = new ArrayList<Object>();

        for (int k = 0; k < row.getLastCellNum(); k++) {
          Cell cell = row.getCell(k);

          if (cell != null) {
            Object value = cellToObject(cell);
            hasValues = hasValues || value != null;
            rowData.add(value);
          } else {
            rowData.add(null);
          }
        }

        if (hasValues || !omitEmpty) {
          currentRowOffset++;

          if (rowLimit > 0 && totalRowsAdded == rowLimit) {
            break;
          }

          if (rowOffset > 0 && currentRowOffset < rowOffset) {
            continue;
          }

          tmp.addRow(rowData);
          totalRowsAdded++;
        }
      }

      if (fillColumns) {
        tmp.fillColumns();
      }

      book.addExcelWorksheet(tmp);
    }

    return book;
  }

  private Object cellToObject(Cell cell) {
    switch (cell.getCellType()) {
      case BOOLEAN:
        return cell.getBooleanCellValue();

      case FORMULA:
        switch (cell.getCachedFormulaResultType()) {
          case NUMERIC:
            return numeric(cell);

          case STRING:
            return cleanString(cell.getRichStringCellValue().toString());

          default:
            return null;
        }

      case NUMERIC: {
        if (cell.getCellStyle().getDataFormatString().contains("%")) {
          return cell.getNumericCellValue() * 100;
        }

        return numeric(cell);
      }

      case STRING:
        return cleanString(cell.getStringCellValue());

      default:
        return null;
    }
  }

  private String cleanString(String str) {
    return str.trim().replaceAll("[\\s\\n]+", " ").replaceAll("(/\\s)+", "/");
  }

  private Object numeric(Cell cell) {
    if (DateUtil.isCellDateFormatted(cell)) {
      if (dateFormat != null) {
        return dateFormat.format(cell.getDateCellValue());
      }

      return cell.getDateCellValue();
    }

    return cell.getNumericCellValue();
  }

}
