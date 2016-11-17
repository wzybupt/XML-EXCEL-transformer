import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
	
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.*;
import org.xml.sax.SAXException;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelUtils {

	/**
	 * 读取xml文件，并输出到xls文件 现阶段的xml文件里面有两种字符串， 一种是单条的字符串 string 一种是字符串数列
	 * string-array
	 * 
	 * @param sourcePath
	 *            所需要读取的xml文件的路径
	 */
	public static void readFromXML(String sourcePath) {
		DocumentBuilderFactory docBuilderFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder docBuilder = null;
		Document doc = null;
		ArrayList<String> stringNameList = new ArrayList<>();
		ArrayList<String> stringContentList = new ArrayList<>();
		ArrayList<String> stringArrayNameList = new ArrayList<>();
		ArrayList<NodeList> stringArrayContentList = new ArrayList<>();
		try {
			docBuilder = docBuilderFactory.newDocumentBuilder();
			doc = docBuilder.parse(new File(sourcePath));
		} catch (SAXException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		doc.getDocumentElement().normalize();
		NodeList listOfStrings = doc.getElementsByTagName("string");
		NodeList listOfStringArrays = doc.getElementsByTagName("string-array");
		int stringCount = listOfStrings.getLength();
		int stringArrayCount = listOfStringArrays.getLength();
		for (int i = 0; i < stringCount; i++) {
			stringNameList.add(listOfStrings.item(i).getAttributes().item(0).getNodeValue());
			stringContentList.add(listOfStrings.item(i).getTextContent());
		}
		for (int i = 0; i < stringArrayCount; i++) {
			stringArrayNameList.add(listOfStringArrays.item(i).getAttributes().item(0).getNodeValue());
			stringArrayContentList.add(((Element) listOfStringArrays.item(i)).getElementsByTagName("item"));
		}
		createSheet(stringNameList, stringContentList, stringArrayNameList, stringArrayContentList);
	}

	/**
	 * 在当前目录生成string.xls文件
	 * 
	 * @param stringNameList
	 *            单条字符串的名字
	 * @param stringContentList
	 *            单条字符串的内容
	 * @param stringArrayNameList
	 *            字符串数组的名字
	 * @param stringArrayContentList
	 *            字符串数组的内容
	 */
	private static void createSheet(ArrayList<String> stringNameList, ArrayList<String> stringContentList,
			ArrayList<String> stringArrayNameList, ArrayList<NodeList> stringArrayContentList) {
		FileOutputStream outputStream = null;
		try {
			outputStream = new FileOutputStream("string.xls");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet();
		Row row = null;
		workbook.setSheetName(0, "string");

		int rownum;
		for (rownum = 0; rownum < stringNameList.size(); rownum++) {
			row = sheet.createRow(rownum);
			row.createCell(0).setCellValue(stringNameList.get(rownum));
			row.createCell(1).setCellValue(stringContentList.get(rownum));
			row.createCell(2).setCellValue(stringContentList.get(rownum));
		}

		for (int i = 0; i < stringArrayNameList.size(); i++) {
			row = sheet.createRow(rownum);
			row.createCell(0).setCellValue(stringArrayNameList.get(i));
			rownum++;
			for (int j = 0; j < stringArrayContentList.get(i).getLength(); j++) {
				row = sheet.createRow(rownum);
				row.createCell(1).setCellValue(stringArrayContentList.get(i).item(j).getTextContent());
				row.createCell(2).setCellValue(stringArrayContentList.get(i).item(j).getTextContent());
				rownum++;
			}
		}

		try {
			workbook.write(outputStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			outputStream.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		try {
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/**
	 * 读取excel文件并输出为xml文件
	 * 
	 * @param sourcePath
	 *            源文件地址
	 */
	public static void readFromXLS(String sourcePath) {

		ArrayList<String> stringNameList = new ArrayList<>();
		ArrayList<String> stringContentList = new ArrayList<>();
		ArrayList<String> stringArrayNameList = new ArrayList<>();
		HashMap<Integer, ArrayList<String>> stringArrayContentMap = new HashMap<>();

		try {
			HSSFWorkbook hssfWorkbook = parseExcel(sourcePath);
			for (int i = 0; i < hssfWorkbook.getNumberOfSheets(); i++) {
				HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(i);
				int rowCount = hssfSheet.getPhysicalNumberOfRows();
				int stringArrayNum = 0;
				System.out.println(
						"Sheet " + i + " \"" + hssfWorkbook.getSheetName(i) + "\" has " + rowCount + " row(s).");
				for (int j = 0; j < rowCount; j++) {
					HSSFRow row = hssfSheet.getRow(j);
					if (row == null) {
						continue;
					}

					int cellNumInOneRow = row.getPhysicalNumberOfCells();
					if (cellNumInOneRow == 1) {// 对于string-array的名字，一行只有一个单元格
						stringArrayNum++;
						stringArrayNameList.add(row.getCell(0).getStringCellValue());
						continue;
					} else if (cellNumInOneRow == 2) {// 对于string-array的内容，一行有两个单元格
						ArrayList<String> list = stringArrayContentMap.get(stringArrayNum - 1);
						if (list == null) {
							list = new ArrayList<>();
						}
						list.add(row.getCell(1).getStringCellValue());
						stringArrayContentMap.put(stringArrayNum - 1, list);
					} else { // 普通string内容，一行有三个单元格
						stringNameList.add(row.getCell(0).getStringCellValue());
						stringContentList.add(row.getCell(2).getStringCellValue());
					}
					createXml(stringNameList, stringContentList, stringArrayNameList, stringArrayContentMap);
					hssfWorkbook.close();
				}
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * 将原excel文件转换成可处理文件
	 * 
	 * @param sourcePath
	 *            excel文件地址
	 * @return
	 * @throws IOException
	 */
	private static HSSFWorkbook parseExcel(String sourcePath) throws IOException {
		FileInputStream fileInputStream = new FileInputStream(sourcePath);
		try {
			return new HSSFWorkbook(fileInputStream);
		} finally {
			fileInputStream.close();
		}
	}

	public static void createXml(ArrayList<String> nameList, ArrayList<String> stringList,
			ArrayList<String> stringArrayNameList, HashMap<Integer, ArrayList<String>> stringArrayContentMap) {
		DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
		DocumentBuilder builder = null;
		try {
			builder = factory.newDocumentBuilder();
			Document document = builder.newDocument();
			Element mainRootElement = document.createElementNS("xmlns:xliff=\"urn:oasis:names:tc:xliff:document:1.2\"",
					"resources");
			document.appendChild(mainRootElement);
			for (int t = 0; t < nameList.size(); t++) {
				mainRootElement.appendChild(getNode(document, "string", nameList.get(t), stringList.get(t)));
			}

			for (int i = 0; i < stringArrayNameList.size(); i++) {
				Node stringArrayNode = getNode(document, "string-array", stringArrayNameList.get(i), null);
				for (int j = 0; j < stringArrayContentMap.get(i).size(); j++) {
					stringArrayNode.appendChild(getNode(document, "item", null, stringArrayContentMap.get(i).get(j)));
				}
				mainRootElement.appendChild(stringArrayNode);
			}

			Transformer transformer = TransformerFactory.newInstance().newTransformer();
			// transformer.setOutputProperty(name, value);
			DOMSource source = new DOMSource(document);
			StreamResult result = new StreamResult(new File("out.xml"));
			transformer.transform(source, result);
		} catch (ParserConfigurationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TransformerConfigurationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TransformerFactoryConfigurationError e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (TransformerException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private static Node getNode(Document document, String nodeName, String attrName, String value) {
		Element node = document.createElement(nodeName);
		if (attrName != null) {
			node.setAttribute("name", attrName);
		}
		node.setTextContent(value);
		return node;
	}
}
