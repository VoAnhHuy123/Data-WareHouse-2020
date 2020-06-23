package Extract;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import java.util.StringTokenizer;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import com.mysql.jdbc.PreparedStatement;


public class ExtractData {
	BufferedReader read;
	Connection con;
	String tableName;
	String fileType;
	int numOfcol;
	int numOfdata;
	ArrayList<String> columnsList = new ArrayList<String>();
	int numberColumns;
	String dataPath;
	public void createConnection() throws SQLException {
		try {
			Class.forName("com.mysql.jdbc.Driver");
			con = DriverManager.getConnection("jdbc:mysql://localhost:3306/mydb", "root", "0126");
			System.out.println("thanh cong");
			
			
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	public void getConfig(String id) throws SQLException {
		PreparedStatement pre = (PreparedStatement) con.prepareStatement("SELECT * FROM mydb.configuration where id=?;");
		pre.setString(1, id);
		ResultSet tmp = pre.executeQuery();
		tmp.next();
		tableName = tmp.getString("filename");
		fileType = tmp.getString("filetype");
		numOfcol =	Integer.parseInt(tmp.getString("numofcol"));
		String listofcol = tmp.getString("listofcol");
		numOfdata = Integer.parseInt(tmp.getString("numofdata"));
		dataPath = tmp.getString("datapath");
		StringTokenizer tokens = new StringTokenizer(listofcol,"|");
		while(tokens.hasMoreTokens()) {
			columnsList.add(tokens.nextToken());
		}
		
		
	}
	public void execute() {
		String sqlCreateTable = "CREATE TABLE $tableName($column1 Interger not null,$column2 INTERGER )";
		try {
			PreparedStatement pre = (PreparedStatement) con.prepareStatement(sqlCreateTable);
			pre.setInt(1, 3);
			pre.setInt(2, 3);
			pre.setInt(3, 3);
			pre.execute();
			
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	public void ChangeWordToTxt() throws IOException {
//		File x = new File("D:\\Data WareHouse\\datafeed.txt");
//		x.createNewFile();
		try {
		      FileInputStream fis = new FileInputStream("D:\\Data WareHouse\\Data Feed Specification.docx");
		      XWPFDocument document = new XWPFDocument(OPCPackage.open(fis));
//		      List<XWPFParagraph> paragraphList = document.getParagraphs();
//		      for (XWPFParagraph paragraph : paragraphList) {
//		        System.out.println(paragraph.getText());
//		      }
		      XWPFWordExtractor wordExtractor = new XWPFWordExtractor(document);
		      PrintWriter out = new PrintWriter(new File("D:\\Data WareHouse\\datafeed.txt"));
		      out.println(wordExtractor.getText());
		      out.close();
//		      System.out.println(wordExtractor.getText());
		      wordExtractor.close();
		      document.close();
		    } catch (Exception ex) {
		      ex.printStackTrace();
		    }
	}
	public void getTableInfor() throws IOException {
		BufferedReader br = null;
		FileReader fr = null;
		try {
			fr = new FileReader("D:\\Data WareHouse\\datafeed.txt");
			br = new BufferedReader(fr);
			String sCurrentLine;
			br = new BufferedReader(new FileReader("D:\\Data WareHouse\\datafeed.txt"));
			while ((sCurrentLine = br.readLine()) != null) {
				StringTokenizer str = new StringTokenizer(sCurrentLine);
//				System.out.println(sCurrentLine);
				
				if (str.countTokens() !=0) {
					
//					System.out.println(str.hasMoreTokens());
//					String a="";
					while(str.hasMoreTokens()) {
						
						String vb = str.nextToken();
					
						if (vb.equals("NumberOfColumns:")) {
							
							numberColumns=Integer.parseInt(str.nextToken());
						}
						if (vb.equals("FileName:")) {
							System.out.println(str.nextToken());
						}
						if (vb.equals("ListOfColumns:")) {
							
							 for (int i = 0; i < numberColumns; i++) {
//								ColumnList.add(str.nextToken());
							}
						}
					
					}
				}

			}
 
		} catch (IOException e) {
 
			e.printStackTrace();
 
		} finally {
 
			try {
 
				if (br != null)
					br.close();
 
				if (fr != null)
					fr.close();
 
			} catch (IOException ex) {
 
				ex.printStackTrace();
 
			}
 
		}
 
	}
    
	public void editTableName(String name) {
		
	}
	public void readFile() throws FileNotFoundException {
		File g = new File(dataPath+"\\"+tableName+"."+fileType);
//		File g = new File("D:\\Data WareHouse\\datafeed.txt");
		  Scanner myReader = new Scanner(g);
		  while (myReader.hasNextLine()) {
		    String data = myReader.nextLine();
		    System.out.println(data);
		  }
		  myReader.close();
	}
	public void d() throws IOException {
		File g = new File(dataPath+"\\"+tableName+"."+fileType);
		FileInputStream fileIn=new FileInputStream(g);

		Workbook wb=WorkbookFactory.create(fileIn);      //this reads the file

		final Sheet sheet=(Sheet) wb.getSheet(tableName); 
		}  

	
	
	public static void main(String[] args) throws SQLException, IOException {
		ExtractData ex = new ExtractData();
		ex.createConnection();
		ex.getConfig("1");
		ex.d();
		
		
		
		
//		ex.execute();
//		ex.ChangeWordToTxt();
		
//		ex.getTableInfor();
//		System.out.println(ex.numberColumns);
//		for (int i = 0; i < ex.ColumnList.size(); i++) {
//			System.out.print(ex.ColumnList.get(i)+ " ");
//		}
	}
}