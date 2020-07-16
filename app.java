import java.sql.*;
import java.util.ArrayList;
import java.util.Scanner;
import java.io.File;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

public class app {

    private static Scanner input = new Scanner(System.in);

    public static void main(String[] args) throws ClassNotFoundException, SQLException, FileNotFoundException,IOException,ClassNotFoundException {
        ArrayList<String> list = new ArrayList<String>();
        System.out.println("Enter Database name:");
        String database = input.nextLine();
        Class.forName("com.mysql.cj.jdbc.Driver");
        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306?useUnicode=true&useJDBCCompliantTimezoneShift=true&useLegacyDatetimeCode=false&serverTimezone=UTC", "root", "");
        Statement st = con.createStatement();
        DatabaseMetaData meta = con.getMetaData();
        ResultSet rs = meta.getCatalogs();
        while (rs.next()) {
            String listofDatabases = rs.getString("TABLE_CAT");
            list.add(listofDatabases);
        }
        if (list.contains(database)) {
            System.out.println("Database already exists");
            exceltodb();
        } else {
            st.executeUpdate("CREATE DATABASE " + database);
            System.out.println("Database is created");
            exceltodb();
        }
        rs.close();
    }

    public static void checkiftableexists() throws SQLException, FileNotFoundException,IOException,ClassNotFoundException {
        Class.forName("com.mysql.cj.jdbc.Driver");
        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/employee?useUnicode=true&useJDBCCompliantTimezoneShift=true&useLegacyDatetimeCode=false&serverTimezone=UTC", "root", "");
        DatabaseMetaData dbm = con.getMetaData();
        ResultSet rs = dbm.getTables(null, null, "employee", null);
        if (rs.next()) {
            System.out.println("Table exists");
        } else {
            String tableCommand = "CREATE TABLE employee (";
            System.out.println("Table does not exist");
            FileInputStream fis = new FileInputStream(new File("D:\\Projects\\Excel to Java\\exceldb.xls"));
            HSSFWorkbook wb = new HSSFWorkbook(fis);
            HSSFSheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(0);
            int count = 0;
            if (row == null) {
                System.out.println("Data not exist in excel!");
                System.exit(1);
            } else {
                for (int cn = row.getFirstCellNum(); cn < row.getLastCellNum(); cn++) {
                    Cell c = row.getCell(cn);
                    if (c == null) {
                        System.out.println("Cell has no data!");
                        System.exit(1);
                    } else {
                        if (count == 0) {
                            tableCommand = tableCommand + c + " varchar(255)";
                            count++;
                        } else {
                            tableCommand = tableCommand + "," + c + " varchar(255)";
                        }
                    }
                }
                tableCommand = tableCommand + ")";
                Statement stmt = null;
                stmt = con.createStatement();
                stmt.executeUpdate(tableCommand);
                System.out.println("Created table in given database...");
            }
        }
    }

    public static void exceltodb() throws FileNotFoundException, SQLException, IOException,ClassNotFoundException {
        Class.forName("com.mysql.cj.jdbc.Driver");
        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/employee?useUnicode=true&useJDBCCompliantTimezoneShift=true&useLegacyDatetimeCode=false&serverTimezone=UTC", "root", "");
        checkiftableexists();
        FileInputStream fis = new FileInputStream(new File("D:\\Projects\\Excel to Java\\exceldb.xls"));
        HSSFWorkbook wb = new HSSFWorkbook(fis);
        HSSFSheet sheet = wb.getSheetAt(0);
        for (int i = 1; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
               System.exit(1);
            }
            else {
                Statement stmt = con.createStatement();
                String sql = "INSERT INTO employee values (";
                int count = 0;
               for (int cn=row.getFirstCellNum(); cn<row.getLastCellNum(); cn++) {
                  Cell c = row.getCell(cn);
                  if (c == null) {
                         System.exit(1);
                } else {
                   if(count == 0)
                    {
                        sql = sql + "'" + c + "'";
                        count++;
                    }
                    else
                    {
                        sql = sql + ",'" + c + "'";
                    }
                  }
               }
               sql = sql + ")";
               System.out.println(sql);
               stmt.executeUpdate(sql);
            }
         }
    }
}