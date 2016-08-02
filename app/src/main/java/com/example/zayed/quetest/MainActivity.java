package com.example.zayed.quetest;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.widget.TextView;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class MainActivity extends AppCompatActivity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        File file=new File("D:\\QueTEst\\app\\src\\main\\res\\surah bakarah.xlsx");

        try {
            FileInputStream fis=new FileInputStream(file);

            XSSFWorkbook xwb=new XSSFWorkbook(fis);
            XSSFSheet xSheet=xwb.getSheetAt(0);

            Iterator<Row> rowIterator=xSheet.rowIterator();

            while (rowIterator.hasNext())
            {
                Row row=rowIterator.next();

                Iterator<Cell> cellIterator=row.cellIterator();

                while (cellIterator.hasNext())
                {
                    Cell cell=cellIterator.next();

                    if(cell.getColumnIndex()==3)
                    {
                        TextView textView= (TextView) findViewById(R.id.text);
                        textView.setText(cell.getStringCellValue());
                    }
                }
            }
        } catch (FileNotFoundException | NullPointerException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
