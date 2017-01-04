import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

public class WriteExcelDemo {

    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Student data");

        Map<String, Object[]> data = new LinkedHashMap<String, Object[]>();
        data.put("FN", new Object[]{
                "First name", "Last name", "Email",
                "Age", "Group", "Grade1",
                "Grade2", "Grade3", "Grade4",
                "Phones"
        });

        BufferedReader br = new BufferedReader(new FileReader("StudentData.txt"));
        try {
            String line = br.readLine();

            while (line != null) {
                String[] inputData = line.split("\\s+");
                if (inputData[0].equals("FN")) {
                    line = br.readLine();
                    continue;
                }

                data.put(inputData[0], new Object[]{
                        inputData[1], inputData[2], inputData[3],
                        inputData[4], inputData[5], inputData[6],
                        inputData[7], inputData[8], inputData[9],
                        inputData[10]
                });
                line = br.readLine();
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            br.close();
        }

        Set<String> keySet = data.keySet();
        int rowNumber = 0;
        for (String key : keySet) {
            Row row = sheet.createRow(rowNumber++);
            Object[] objArr = data.get(key);
            int cellNumber = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellNumber++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Integer) {
                    cell.setCellValue((Integer) obj);
                }
            }
        }

        try {
            FileOutputStream out = new FileOutputStream(new File("StudentData.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("StudentData.xlsx written successfully on disk.");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}