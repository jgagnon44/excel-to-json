package com.fossfloors.exceljson;

import com.fossfloors.exceljson.converter.ExcelToJsonConverter;

public class TestApplication {

  public static void main(String[] args) throws Exception {
    ExcelToJsonConverter converter = new ExcelToJsonConverter();
    String json = converter.convert(args[0]).toJson();

    if (json != null) {
      System.out.println(json);
    }
  }

}
