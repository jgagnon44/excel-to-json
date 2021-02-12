package com.fossfloors.exceljson.pojo;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;

import org.springframework.beans.factory.annotation.Value;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

public class ExcelWorkbook {

  @Value("${excelToJSON.prettyPrint:false}")
  private boolean                    prettyPrint;

  private Collection<ExcelWorksheet> sheets = new ArrayList<ExcelWorksheet>();

  public void addExcelWorksheet(ExcelWorksheet sheet) {
    sheets.add(sheet);
  }

  public String toJson() throws JsonGenerationException, JsonMappingException, IOException {
    ObjectMapper mapper = new ObjectMapper();

    if (prettyPrint) {
      mapper.enable(SerializationFeature.INDENT_OUTPUT);
    }

    return mapper.writeValueAsString(this);
  }

  public Collection<ExcelWorksheet> getSheets() {
    return sheets;
  }

  public void setSheets(Collection<ExcelWorksheet> sheets) {
    this.sheets = sheets;
  }

}
