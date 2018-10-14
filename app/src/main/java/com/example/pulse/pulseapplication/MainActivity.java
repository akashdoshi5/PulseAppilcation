package com.example.pulse.pulseapplication;

import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.os.Bundle;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.util.Log;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.EditText;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    Button button;
    EditText racfID;
    public static final String PREFS_NAME = "PulseTeam";
    Gson gson = new Gson();
    ExcelDetail selectedExcelDetail;
    TextView textView;
    Spinner spinner;

    @Override
    protected void onStart() {
        final SharedPreferences sharedPreferences = getSharedPreferences(PREFS_NAME, 0 );
        super.onStart();
        final String excelName = getString(R.string.ExcelName);
        setQuestionName(sharedPreferences, excelName);
    }

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        final SharedPreferences sharedPreferences = getSharedPreferences(PREFS_NAME, 0 );
        final String excelName = getString(R.string.ExcelName);
        sharedPreferences.getInt("sheetNumber"+excelName, 0);

        button = findViewById(R.id.button);
        racfID = findViewById(R.id.editText);
        textView =  findViewById(R.id.questionText);

        spinner = findViewById(R.id.loginSpinner);
        List<String> domains = new ArrayList<>();
        domains.add("Please select your Response!");
        domains.add("Agree");
        domains.add("Disagree");

        final ArrayAdapter<String> dataAdapter = new ArrayAdapter<String>(this,
                android.R.layout.simple_spinner_item, domains);
        dataAdapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
        spinner.setAdapter(dataAdapter);

        button.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                final String selectedObject = sharedPreferences.getString("selectedFrom"+excelName, null);
                if(selectedObject != null ) {
                    if(!racfID.getText().toString().trim().equals("")) {
                        if(!spinner.getSelectedItem().equals("Please select your Response!")) {
                            selectedExcelDetail = gson.fromJson(selectedObject, ExcelDetail.class);
                            readExcelFile(MainActivity.this, excelName);
                            Toast.makeText(MainActivity.this, "Data Inserted! Racf id: "+racfID.getText().toString().trim()+", Response: "+spinner.getSelectedItem().toString().trim() , Toast.LENGTH_SHORT).show();
                            racfID.setText("");
                            spinner.setAdapter(dataAdapter);
                        }else{
                            Toast.makeText(MainActivity.this, "Please select Response from Dropdown!", Toast.LENGTH_SHORT).show();
                        }
                    }else{
                        Toast.makeText(MainActivity.this, "Please Insert Racf Id!", Toast.LENGTH_SHORT).show();
                    }
                } else  {
                    Toast.makeText(MainActivity.this, "There is no question present to answer! Pulse Team Members needs to Login and add a Question!", Toast.LENGTH_SHORT).show();
                }
            }
        });
    }

    private void setQuestionName(SharedPreferences sharedPreferences, String excelName) {
        String selectedObject1 = sharedPreferences.getString("selectedFrom"+excelName, null);
        Gson gson=new Gson();;
        ExcelDetail excelDetail;
        Type type;
        type = new TypeToken<ExcelDetail>() {
        }.getType();
        excelDetail = gson.fromJson(selectedObject1, type);
        if(selectedObject1 != null ) {
            textView.setText("Question : "+ excelDetail.getQuestion());
        }else{
            textView.setText("No Question Available!");
        }
    }


    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        getMenuInflater().inflate(R.menu.menu, menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        Intent intent = new Intent(MainActivity.this, PulseLoginActivity.class);
        startActivity(intent);
        return super.onOptionsItemSelected(item);
    }
    private static boolean saveExcelFile(Context context, String fileName) {

        if (!isExternalStorageAvailable() || isExternalStorageReadOnly()) {
            Log.e("LOG", "Storage not available or read only");
            return false;
        }

        boolean success = false;

        Workbook wb = new HSSFWorkbook();

        Sheet sheet1 = wb.createSheet("Records");

        Row row1 = sheet1.createRow(0);
        Cell c1 = row1.createCell(0);
        c1.setCellValue("Question");

        sheet1.addMergedRegion(new CellRangeAddress(0,0,0,1));

        Row row2 = sheet1.createRow(1);
        Cell c = row2.createCell(0);
        c.setCellValue("Racf-IDs");

        c = row2.createCell(1);
        c.setCellValue("Response");

        sheet1.setColumnWidth(0, (15 * 500));
        sheet1.setColumnWidth(1, (15 * 500));

        File file = new File(context.getExternalFilesDir(null), fileName);
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            Log.w("FileUtils", "Writing file" + file);
            success = true;
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
            }
        }
        return success;
    }

    private void readExcelFile(Context context, String filename) {

        if (!isExternalStorageAvailable() || isExternalStorageReadOnly())
        {
            Log.e("LOG","Storage not available or read only");
            return;
        }

        try{
            File file = new File(context.getExternalFilesDir(null), filename);
            FileInputStream myInput = new FileInputStream(file);

            POIFSFileSystem myFileSystem = new POIFSFileSystem(myInput);

            HSSFWorkbook myWorkBook = new HSSFWorkbook(myFileSystem);

            HSSFSheet mySheet = myWorkBook.getSheet(selectedExcelDetail.getSheetName());

            int rowNum = mySheet.getLastRowNum();
            Row row = mySheet.createRow(++rowNum);

            Cell c = row.createCell(0);
            c.setCellValue(racfID.getText().toString().trim());

            c = row.createCell(1);
            c.setCellValue(spinner.getSelectedItem().toString().trim());


            FileOutputStream os = null;

            try {
                os = new FileOutputStream(file);
                myWorkBook.write(os);
                Log.w("FileUtils", "Writing file" + file);
            } catch (IOException e) {
                Log.w("FileUtils", "Error writing " + file, e);
            } catch (Exception e) {
                Log.w("FileUtils", "Failed to save file", e);
            } finally {
                try {
                    if (null != os)
                        os.close();
                } catch (Exception ex) {
                }
            }



        }catch (Exception e){e.printStackTrace();
            saveExcelFile(context, filename);
            readExcelFile(context, filename);
        }

        return;
    }

    public static boolean isExternalStorageReadOnly() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED_READ_ONLY.equals(extStorageState)) {
            return true;
        }
        return false;
    }

    public static boolean isExternalStorageAvailable() {
        String extStorageState = Environment.getExternalStorageState();
        if (Environment.MEDIA_MOUNTED.equals(extStorageState)) {
            return true;
        }
        return false;
    }
}

