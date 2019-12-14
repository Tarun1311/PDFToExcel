package com;

import org.apache.log4j.Logger;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.PDFTextStripperByArea;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class Extractor {

	static Logger log = Logger.getLogger(Extractor.class);
	static String[] documentName;

	public static void main(String[] args) throws IOException {

		File f = new File("C:/Users/VIP/Desktop/logging.log");
		FileOutputStream fout = new FileOutputStream(f);
		fout.flush();
		fout.close();

		List<String> files = new ArrayList<>();

		String FILE_NAME = "C:/Users/VIP/Desktop/amazon.xlsx";

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Amazon");

		String[][] datatypes = { { "Document No", "Order ID", "TransactionType", "SKU", "Product Details", "GrandTotal",
				"quantity", "Date", "Month", "Name", "Prime", "Contact", "Pincode", "State", "GST" } };

		int rowNum = 0;
		System.out.println("Creating excel");

		for (String[] datatype : datatypes) {
			Row row = sheet.createRow(rowNum++);
			int colNum = 0;
			for (String field : datatype) {
				Cell cell = row.createCell(colNum++);
				cell.setCellValue(field);
			}
		}
		int noOfFiles = 0;

		try (Stream<Path> walk = Files.walk(Paths.get("C:\\Users\\VIP\\Desktop\\PDFs"))) {

			List<String> result = walk.filter(Files::isRegularFile).map(x -> x.toString()).collect(Collectors.toList());
			noOfFiles = result.size();
			documentName = new String[noOfFiles];
			int i = 0;
			for (String file : result) {
				String location = file.replace('\\', '/');
				files.add(location);
				documentName[i] = location.split("/")[5];
				i++;
			}

		} catch (IOException e) {
			e.printStackTrace();
		}

		// int noOfFiles = files.size();

		String[] name = new String[noOfFiles];
		String[] state = new String[noOfFiles];
		String[] orderId = new String[noOfFiles];
		int[] quantity = new int[noOfFiles];
		List<String> sku[] = new ArrayList[noOfFiles];
		List<String> productDetails[] = new ArrayList[noOfFiles];
		float[] tax = new float[noOfFiles];
		float[] grandTotal = new float[noOfFiles];
		String[] transactionType = new String[noOfFiles];

		for (int i = 0; i < noOfFiles; i++) {
			sku[i] = new ArrayList<>();
			productDetails[i] = new ArrayList<>();
		}

		int fileCount = 0;
		boolean errorFlag=false;
		for (String file : files) {
			errorFlag=false;
			try (PDDocument document = PDDocument.load(new File(file))) {

				document.getClass();

				if (!document.isEncrypted()) {

					PDFTextStripperByArea stripper = new PDFTextStripperByArea();
					stripper.setSortByPosition(true);

					PDFTextStripper tStripper = new PDFTextStripper();

					String pdfFileInText = tStripper.getText(document);
					// System.out.println("Text:" + st);

					// split by whitespace
					String lines[] = pdfFileInText.split("\\r?\\n");
					int length = lines.length;

					for (int i = 0; i < length; i++) {
						// System.out.println(lines[i]);
						try {
							if (lines[i].equals("Delivery address:")) {
								name[fileCount] = lines[i + 1];
							} else if (lines[i].startsWith("Phone :")) {
								String[] address = lines[i - 1].split(" ");
								int len = address.length;
								String stateWithSpace = "";
								for (int j = 2; j < len - 2; j++)
									stateWithSpace += address[j] + ' ';
								state[fileCount] = stateWithSpace.substring(0, stateWithSpace.length() - 1);
								if (lines[i + 1].startsWith("COD Collectible Amount")) {
									transactionType[fileCount] = "COD";
									grandTotal[fileCount] = Float
											.valueOf(lines[i + 2].substring(2).replaceAll("[,]", ""));
								}
							} else if (lines[i].startsWith("Order ID:")) {
								orderId[fileCount] = lines[i].substring(10);
							} else if (lines[i].startsWith("SKU:")) {
								if (lines[i - 2].startsWith("Item total")) {
									String[] product = lines[i - 1].split(" ", 2);
									quantity[fileCount] += Integer.valueOf(product[0].replaceAll("[,]", ""));
									productDetails[fileCount].add(product[1]);
								} else if (lines[i - 2].startsWith("Quantity")) {
									String[] product = lines[i - 1].split(" ", 2);
									quantity[fileCount] += Integer.valueOf(product[0].replaceAll("[,]", ""));
									productDetails[fileCount].add(product[1]);
								} else {
									String[] product = lines[i - 2].split(" ", 2);
									quantity[fileCount] += Integer.valueOf(product[0].replaceAll("[,]", ""));
									productDetails[fileCount].add(product[1] + lines[i - 1]);
								}
								sku[fileCount].add(lines[i].split(" ")[1]);
							} else if (lines[i].startsWith("Tax")) {
								tax[fileCount] += Float.valueOf(lines[i].split(" ")[1].substring(2));
							} else if (lines[i].startsWith("Grand total:")) {
								grandTotal[fileCount] = Float
										.valueOf(lines[i].split(" ")[2].substring(2).replaceAll("[,]", ""));
								transactionType[fileCount] = "PREPAID";
								if (grandTotal[fileCount] == 0)
									transactionType[fileCount] = "";
							} else if (lines[i].startsWith("Thanks for buying on Amazon Marketplace.")) {
								break;
							}
						} catch (Exception ex) {
							log.error(documentName[fileCount] + " " + ex);
							errorFlag=true;
							fileCount++;
							break;
						}
					}
					if(errorFlag==true)
						continue;
				}

			}

			Object[][] data = { { documentName[fileCount], orderId[fileCount], transactionType[fileCount],
					sku[fileCount].toString(), productDetails[fileCount].toString(), grandTotal[fileCount],
					quantity[fileCount], "", "", name[fileCount], "", "", "", state[fileCount], tax[fileCount] } };

			System.out.println("Creating excel");

			for (Object[] datatype : data) {
				Row row = sheet.createRow(rowNum++);
				int colNum = 0;
				for (Object field : datatype) {
					Cell cell = row.createCell(colNum++);
					// cell.setCellValue(field);
					if (field instanceof String) {
						cell.setCellValue((String) field);
					} else if (field instanceof Integer) {
						cell.setCellValue((Integer) field);
					} else if (field instanceof Float) {
						cell.setCellValue((Float) field);
					}
				}
			}
			fileCount++;
		}
		try {
			FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
			workbook.write(outputStream);
			// workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("Done");
	}
}
