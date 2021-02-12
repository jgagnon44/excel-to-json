package com.fossfloors.exceljson.pojo;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Value;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

public class ExcelWorksheet {

  @Value("${excelToJSON.prettyPrint:false}")
  private boolean            prettyPrint;

  private String             name;
  private List<List<Object>> data    = new ArrayList<>();
  private int                maxCols = 0;

  public String toJson() throws JsonGenerationException, JsonMappingException, IOException {
    ObjectMapper mapper = new ObjectMapper();

    if (prettyPrint) {
      mapper.enable(SerializationFeature.INDENT_OUTPUT);
    }

    return mapper.writeValueAsString(this);
  }

  public void addRow(ArrayList<Object> row) {
    data.add(row);

    if (maxCols < row.size()) {
      maxCols = row.size();
    }
  }

  public int getMaxRows() {
    return data.size();
  }

  public void fillColumns() {
    for (List<Object> tmp : data) {
      while (tmp.size() < maxCols) {
        tmp.add(null);
      }
    }
  }

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public List<List<Object>> getData() {
    return data;
  }

  public void setData(List<List<Object>> data) {
    this.data = data;
  }

  public int getMaxCols() {
    return maxCols;
  }

  public void setMaxCols(int maxCols) {
    this.maxCols = maxCols;
  }

}
