package com.example.myapplication;

import android.os.Bundle;
import android.widget.Toast;

import androidx.activity.EdgeToEdge;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.graphics.Insets;
import androidx.core.view.ViewCompat;
import androidx.core.view.WindowInsetsCompat;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;

public class CreateExcelSheet extends AppCompatActivity {
    final String ExcelSheetFilePath = "/storage/emulated/0/Download/TaalKaraoke.xls";
    private static Cell cell;
    private static Sheet sheet;

    private static String EXCEL_SHEET_NAME = "Sheet1";
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_create_excel_sheet);
        if(!DoesFileExists(ExcelSheetFilePath)) {
            Toast.makeText(this, "File Doesn't Exists, Creating The File", Toast.LENGTH_SHORT).show();
            GenerateExecelSheet();
        }
        else
            Toast.makeText(this, "File Already Exists", Toast.LENGTH_SHORT).show();

    }
    protected Boolean DoesFileExists(String FilePath){

        File file = new File(FilePath);
        return file.exists();
    }
    public static void GenerateExecelSheet() {
        // New Workbook
        Workbook workbook = new HSSFWorkbook();

        cell = null;

        // Cell style for header row
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);

        // New Sheet
        sheet = null;
        sheet = workbook.createSheet(EXCEL_SHEET_NAME);

        // Generate column headings
        Row row = sheet.createRow(0);

        cell = row.createCell(0);
        cell.setCellValue("First Name");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(1);
        cell.setCellValue("Last Name");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(2);
        cell.setCellValue("Phone Number");
        cell.setCellStyle(cellStyle);

        cell = row.createCell(3);
        cell.setCellValue("Mail ID");
        cell.setCellStyle(cellStyle);
    }
}