# java-excel-newbook-cell
Apache POI 5.0.0 : 新規ブック、新規シート、セル書込、保存

![image](https://user-images.githubusercontent.com/1501327/129868788-d248542a-6294-4e4e-8e56-e9d140f86b93.png)


```javascript
{
    "java.project.referencedLibraries": [
        "lib/**/*.jar",
        "C:\\app\\workspace\\lib\\poi-5.0.0.jar",
        "C:\\app\\workspace\\lib\\commons-collections4-4.4.jar",
        "C:\\app\\workspace\\lib\\xmlbeans-4.0.0.jar",
        "C:\\app\\workspace\\lib\\poi-ooxml-full-5.0.0.jar",
        "C:\\app\\workspace\\lib\\poi-ooxml-5.0.0.jar",
        "C:\\app\\workspace\\lib\\commons-compress-1.20.jar"
    ]
}
```

## [poi-bin-5.0.0-20210120.zip](https://www.apache.org/dyn/closer.lua/poi/release/bin/poi-bin-5.0.0-20210120.zip)

![image](https://user-images.githubusercontent.com/1501327/129868195-c8653358-3c1c-4e12-b237-88ce7660d0a5.png)

![image](https://user-images.githubusercontent.com/1501327/129868349-c1654a8d-e06d-44ca-ac22-795f4523bc81.png)

![image](https://user-images.githubusercontent.com/1501327/129868437-62274b8f-0f06-4f69-ab28-db54c70cb5e7.png)

```java
// セルに書き込む為のオブジェクト作成・取得処理
static void cell(XSSFSheet sheet, int row, int col, String value) {

    XSSFRow xslRow = sheet.getRow(row - 1);
    if ( xslRow == null ) {
        // 行を作成
        xslRow = sheet.createRow( row - 1 );
    }

    XSSFCell cell = xslRow.getCell( col - 1 );
    if ( cell == null ) {
        // セルを作成
        cell = xslRow.createCell( col - 1 );
    }
    cell.setCellValue(value);

}
```

[その他まとめ](https://logicalerror.seesaa.net/article/478723381.html)
