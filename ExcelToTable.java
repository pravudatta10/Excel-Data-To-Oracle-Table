import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToTable {
	public static void main(String[] args) throws IOException { 

		File file = new File("C:\\Users\\kpds0\\Downloads\\Financial Sample.xlsx");
		FileInputStream fis = new FileInputStream(file);

		Workbook workbook;
		if (file.getName().endsWith(".xls")) {
			workbook = new HSSFWorkbook(fis); // For .xls files
		} else if (file.getName().endsWith(".xlsx")) {
			workbook = new XSSFWorkbook(fis); // For .xlsx files
		} else {
			throw new IllegalArgumentException("The file is not a valid Excel file");
		}

		Sheet sheet = workbook.getSheetAt(0); // Assuming you want to read the first sheet

		String jdbcUrl = "jdbc:oracle:thin:@192.168.2.18:1521:orcl";
		String username = "training";
		String password = "training";
		String insertQuery = "INSERT INTO FINANCIAL"
				+ "(SEGMENT, COUNTRY, PRODUCT, DISCOUNT_BAND, UNITS_SOLD, MANUFACTURING_PRICE, SALE_PRICE, GROSS_SALES, "
				+ "DISCOUNTS, SALES, COGS, PROFIT, SDATE, MONTH_NUMBER, MONTH_NAME,YEAR) "
				+ "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ? ,?)";

		try (Connection connection = DriverManager.getConnection(jdbcUrl, username, password);
				PreparedStatement preparedStatement = connection.prepareStatement(insertQuery)) {

			// Iterate through rows and cells
			for (Row row : sheet) {
				int columnIndex = 1; // Start with the first column

				// Set values based on data types
				setStringValue(preparedStatement, columnIndex++, getStringValue(row.getCell(0)));
				setStringValue(preparedStatement, columnIndex++, getStringValue(row.getCell(1)));
				setStringValue(preparedStatement, columnIndex++, getStringValue(row.getCell(2)));
				setStringValue(preparedStatement, columnIndex++, getStringValue(row.getCell(3)));
				setNumericValue(preparedStatement, columnIndex++, row.getCell(4));
				setNumericValue(preparedStatement, columnIndex++, row.getCell(5));
				setNumericValue(preparedStatement, columnIndex++, row.getCell(6));
				setNumericValue(preparedStatement, columnIndex++, row.getCell(7));
				setNumericValue(preparedStatement, columnIndex++, row.getCell(8));
				setNumericValue(preparedStatement, columnIndex++, row.getCell(9));
				setNumericValue(preparedStatement, columnIndex++, row.getCell(10));
				setNumericValue(preparedStatement, columnIndex++, row.getCell(11));
				setDateValue(preparedStatement, columnIndex++, row.getCell(12));
				setNumericValue(preparedStatement, columnIndex++, row.getCell(13));
				setStringValue(preparedStatement, columnIndex++, getStringValue(row.getCell(14))); 
				setNumericValue(preparedStatement, columnIndex++, row.getCell(15));

				// Add the current row to the batch
				preparedStatement.executeUpdate();
			}

			// Execute the batch
			preparedStatement.executeBatch();

			// Commit the transaction (assuming auto-commit is turned off)
			connection.commit();
			System.out.println("ExcelToTable.main()");
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}

	private static String getStringValue(Cell cell) {
		return cell != null ? cell.getStringCellValue() : null;
	}

	private static void setStringValue(PreparedStatement preparedStatement, int index, String value)
			throws SQLException {
		preparedStatement.setString(index, value);
	}

	private static void setNumericValue(PreparedStatement preparedStatement, int index, Cell cell) throws SQLException {
		if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			preparedStatement.setDouble(index, cell.getNumericCellValue());
		} else {
			preparedStatement.setNull(index, java.sql.Types.DOUBLE);
		}
	}

	private static void setDateValue(PreparedStatement preparedStatement, int index, Cell cell) throws SQLException {
	
		if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC && DateUtil.isCellDateFormatted(cell)) {
			  java.util.Date dateValue = cell.getDateCellValue();
		      java.sql.Date sqlDate = new java.sql.Date(dateValue.getTime());
			preparedStatement.setDate(index, sqlDate);
		} else {
			preparedStatement.setNull(index, java.sql.Types.DATE);
		}
	}
}
