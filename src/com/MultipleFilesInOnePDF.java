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

public class MultipleFilesInOnePDF {

	static Logger log = Logger.getLogger(Extractor.class);
	static List<String> documentName = new ArrayList<>();

	// static String documentName;
	public static void main(String[] args) throws IOException {

		File f = new File("C:/Users/DELL/Desktop/logging.log");
		FileOutputStream fout = new FileOutputStream(f);
		fout.flush();
		fout.close();

		List<String> files = new ArrayList<>();

		String FILE_NAME = "C:/Users/DELL/Desktop/amazon.xlsx";

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

		try (Stream<Path> walk = Files.walk(Paths.get("C:\\Users\\DELL\\Desktop\\PDFs"))) {

			List<String> result = walk.filter(Files::isRegularFile).map(x -> x.toString()).collect(Collectors.toList());
			noOfFiles = result.size();
			// documentName = new String[noOfFiles];
			// int i = 0;
			for (String file : result) {
				String location = file.replace('\\', '/');
				files.add(location);
				documentName.add(location.split("/")[5]);
				// i++;
			}

		} catch (IOException e) {
			e.printStackTrace();
		}

		// int noOfFiles = files.size();

		// for (int i = 0; i < noOfFiles; i++) {
		// sku[i] = new ArrayList<>();
		// productDetails[i] = new ArrayList<>();
		// }

		int fileNo = 0;

		// boolean errorFlag = false;
		for (String file : files) {
			// errorFlag = false;
			int fileCount = 0;
			List<String> name = new ArrayList<>();
			List<String> state = new ArrayList<>();
			List<String> orderId = new ArrayList<>();
			List<Integer> quantity = new ArrayList<>();
			List<List<String>> sku = new ArrayList<List<String>>();
			List<List<String>> productDetails = new ArrayList<List<String>>();
			List<Float> tax = new ArrayList<>();
			List<Float> grandTotal = new ArrayList<>();
			List<String> transactionType = new ArrayList<>();
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
//						System.out.println(i + " " + lines[i]);
						try {
							if (lines[i].equals("Delivery address:")) {
								name.add(lines[i + 1]);
								sku.add(new ArrayList<>());
								productDetails.add(new ArrayList<>());
								while (true) {
//									System.out.println(i + " " + lines[i]);
									try {
										if (lines[i].startsWith("Phone :")) {
											String[] address = lines[i - 1].split(" ");
											int len = address.length;
											String stateWithSpace = "";
											for (int j = 2; j < len - 2; j++)
												stateWithSpace += address[j] + ' ';
											state.add(stateWithSpace.substring(0, stateWithSpace.length() - 1));
											if (lines[i + 1].startsWith("COD Collectible Amount")) {
												transactionType.add("COD");
												grandTotal.add(
														Float.valueOf(lines[i + 2].substring(2).replaceAll("[,]", "")));
											}
											// System.out.println(state.get(fileCount));
										} else if (lines[i].startsWith("Order ID:")) {
											orderId.add(lines[i].substring(10));
										} else if (lines[i].startsWith("SKU:")) {
											if (lines[i - 2].startsWith("Item total")) {
												String[] product = lines[i - 1].split(" ", 2);
												if (quantity.size() == fileCount)
													quantity.add(Integer.valueOf(product[0].replaceAll("[,]", "")));
												else
													quantity.set(fileCount, quantity.get(fileCount)
															+ Integer.valueOf(product[0].replaceAll("[,]", "")));
												productDetails.get(fileCount).add(product[1]);
											} else if (lines[i - 2].startsWith("Quantity")) {
												String[] product = lines[i - 1].split(" ", 2);
												if (quantity.size() == fileCount)
													quantity.add(Integer.valueOf(product[0].replaceAll("[,]", "")));
												else
													quantity.set(fileCount, quantity.get(fileCount)
															+ Integer.valueOf(product[0].replaceAll("[,]", "")));
												productDetails.get(fileCount).add(product[1]);
											} else {
												String[] product = lines[i - 2].split(" ", 2);
												if (quantity.size() == fileCount)
													quantity.add(Integer.valueOf(product[0].replaceAll("[,]", "")));
												else
													quantity.set(fileCount, quantity.get(fileCount)
															+ Integer.valueOf(product[0].replaceAll("[,]", "")));
												productDetails.get(fileCount).add(product[1] + lines[i - 1]);
											}
											sku.get(fileCount).add(lines[i].split(" ")[1]);
										} else if (lines[i].startsWith("Tax")) {
											if (tax.size() == fileCount)
												tax.add(Float.valueOf(lines[i].split(" ")[1].substring(2)));
											else
												tax.set(fileCount, tax.get(fileCount)
														+ Float.valueOf(lines[i].split(" ")[1].substring(2)));
										} else if (lines[i].startsWith("Grand total:")) {
											grandTotal.add(Float.valueOf(
													lines[i].split(" ")[2].substring(2).replaceAll("[,]", "")));
											transactionType.add("PREPAID");
											if (grandTotal.get(fileCount) == 0)
												transactionType.add("");
										} else if (lines[i].startsWith("Thanks for buying on Amazon Marketplace.")) {
											Object[][] data = { { documentName.get(fileNo), orderId.get(fileCount),
													transactionType.get(fileCount), sku.get(fileCount).toString(),
													productDetails.get(fileCount).toString(), grandTotal.get(fileCount),
													quantity.get(fileCount), "", "", name.get(fileCount), "", "", "",
													state.get(fileCount), tax.get(fileCount) } };
											System.out.println("Creating excel");
											fileCount++;
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
											i++;
											break;
										}
										i++;

									} catch (Exception ex) {
										log.error(documentName.get(fileNo) + " " + name.get(fileCount) + " " + ex + i);
										// errorFlag = true;
										name.remove(fileCount);
										state.remove(fileCount);
										orderId.remove(fileCount);
										quantity.remove(fileCount);
										sku.remove(fileCount);
										productDetails.remove(fileCount);
										tax.remove(fileCount);
										grandTotal.remove(fileCount);
										transactionType.remove(fileCount);
										break;
									}
								}
							}
						} catch (Exception ex) {
						}
						// if (errorFlag == true)
						// continue;

					}

				}

			}

			fileNo++;
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
