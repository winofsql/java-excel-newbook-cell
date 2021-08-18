import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    // *****************************************************
    // エントリポイント
    // *****************************************************
    public static void main(String[] args) {

        System.out.println( "新規ブック、新規シート、セル書込、保存" );

        try {
            // ブック作成
            XSSFWorkbook book = new XSSFWorkbook();

            XSSFSheet  sheet1 = book.createSheet("新しいシート");
            XSSFSheet  sheet2 = book.createSheet("追加のシート");

            cell( sheet1, 1, 2, "タイトル");
            cell( sheet1, 2, 2, "値1");
            cell( sheet1, 3, 2, "値2");
            cell( sheet1, 4, 2, "値3");

            FileOutputStream fos = new FileOutputStream("sample.xlsx");
            
            book.write(fos);
            
        } catch (Exception e) {

            e.printStackTrace();
            //TODO: handle exception
        }

    }

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

}
