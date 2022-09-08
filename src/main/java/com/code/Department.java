package com.code;

class Department {
    String path;
    String parentCode;
    String code;
    String arabicName;
    String englishName;

    Department(String code, String arabicName, String englishName) {
        this.code = code == null ? "-1" : code;
        this.arabicName = arabicName;
        this.englishName = englishName;
    }

    Department(String code, String arabicName, String englishName, String path) {
        this(code, arabicName, englishName);
        this.path = path;
    }

    Department(String code, String arabicName, String englishName, String path, String parentCode) {
        this(code, arabicName, englishName, path);
        this.parentCode = parentCode;
    }
}