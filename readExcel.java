import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class utils {
	/**
	 * 读取后缀为xlsx文件
	 * @param filePath
	 * @return
	 */
	public static List<List<String>> readXlsx(String filePath){
		List<List<String>> result = new ArrayList<List<String>>();
		try {
			InputStream is = new FileInputStream(filePath);
			XSSFWorkbook xssfworkbook = new XSSFWorkbook(is);
			//循环每一页，并处理当前循环页
			for(XSSFSheet xssfsheet:xssfworkbook){
				if(xssfsheet==null) continue;
				for(int rowNum=1;rowNum<xssfsheet.getLastRowNum();rowNum++){
					XSSFRow xssfRow = xssfsheet.getRow(rowNum);
					int minColIx = xssfRow.getFirstCellNum();
					int maxColIx = xssfRow.getLastCellNum();
					List<String> rowList = new ArrayList<>();
					for(int colTx=minColIx;colTx<maxColIx;colTx++){
						XSSFCell xssfCell = xssfRow.getCell(colTx);
						if(xssfCell==null) continue;
						rowList.add(xssfCell.toString());
					}
					result.add(rowList);
				}
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
	/**
	 * 读取后缀为xls文件
	 * @param filePath
	 * @return
	 */
	public static List<List<String>> readXls(String filePath){
		List<List<String>> result = new ArrayList<List<String>>();
		try {
			InputStream is = new FileInputStream(filePath);
			HSSFWorkbook hssfworkbook = new HSSFWorkbook(is);
			//循环每一页，并处理当前循环页
			for(int numSheet=0;numSheet<hssfworkbook.getNumberOfSheets();numSheet++){
				HSSFSheet hssfSheet = hssfworkbook.getSheetAt(numSheet);
				if(hssfSheet==null) continue;
				for(int rowNum=1;rowNum<hssfSheet.getLastRowNum();rowNum++){
					HSSFRow hssfRow = hssfSheet.getRow(rowNum);
					int minColIx = hssfRow.getFirstCellNum();
					int maxColIx = hssfRow.getLastCellNum();
					List<String> rowList = new ArrayList<>();
					for(int colTx=minColIx;colTx<maxColIx;colTx++){
						HSSFCell hssfCell = hssfRow.getCell(colTx);
						if(hssfCell==null) continue;
						rowList.add(hssfCell.toString());
					}
					result.add(rowList);
				}
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}
}
