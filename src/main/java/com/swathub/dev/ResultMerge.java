package com.swathub.dev;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class ResultMerge {
	public static void main(String[] args) throws Exception {
		if (args.length <= 1) {
			System.out.println("Usage: java -jar ResultMerge.jar <base path> <path1> <path2> ...");
			return;
		}

		HashMap<String, List<File>> resultMap = new HashMap<String, List<File>>();
		File baseFolder = new File(args[0]);
		String[] extensions = {"xls"};
		for (int i = 1; i < args.length; i++) {
			File resultFolder = new File(baseFolder, args[i]);
			for (File resultFile : FileUtils.listFiles(resultFolder, extensions, false)){
				String[] nameSplit = resultFile.getName().split("-");
				String mergeName = nameSplit[0] + "-" + nameSplit[1];
				if (resultMap.containsKey(mergeName)) {
					resultMap.get(mergeName).add(resultFile);
				} else {
					List<File> fileList = new ArrayList<File>();
					fileList.add(resultFile);
					resultMap.put(mergeName, fileList);
				}
			}
		}

		File mergeFolder = new File(baseFolder, "merge");
		if (!mergeFolder.exists() || !mergeFolder.isDirectory()) {
			if (!mergeFolder.mkdir()) {
				throw new Exception("Can not create merge folder");
			}
		}

		for (String mergeName : resultMap.keySet()) {
			File mergerFile = new File(mergeFolder, mergeName + ".xls");

			HSSFWorkbook wbMerge = new HSSFWorkbook();
			HSSFCreationHelper creationHelper = wbMerge.getCreationHelper();
			HSSFSheet sheetMerge = wbMerge.createSheet("Result");

			HSSFFont boldFont = wbMerge.createFont();
			boldFont.setBold(true);
			HSSFCellStyle titleStyle = wbMerge.createCellStyle();
			titleStyle.setFont(boldFont);
			HSSFRow titleRow = sheetMerge.createRow(0);

			int colCnt = 0;
			for (File resultFile : resultMap.get(mergeName)) {
				titleRow.createCell(colCnt).setCellValue(resultFile.getName());
				titleRow.getCell(colCnt).setCellStyle(titleStyle);

				InputStream ins = new FileInputStream(resultFile);
				HSSFWorkbook wbResult = new HSSFWorkbook(ins);

				int rowCnt = 1;
				int maxColCnt = colCnt;
				for (HSSFPictureData picData : wbResult.getAllPictures()) {
					int pictureIdx = wbMerge.addPicture(picData.getData(), Workbook.PICTURE_TYPE_PNG);
					HSSFPatriarch drawing = sheetMerge.createDrawingPatriarch();

					HSSFClientAnchor anchor = creationHelper.createClientAnchor();
					anchor.setCol1(colCnt);
					anchor.setRow1(rowCnt);
					HSSFPicture picture = drawing.createPicture(anchor, pictureIdx);
					picture.resize();
					rowCnt = anchor.getRow2() + 1;
					if (anchor.getCol2() > maxColCnt) {
						maxColCnt = anchor.getCol2();
					}
				}

				colCnt = maxColCnt + 1;
				ins.close();
				wbResult.close();
			}

			ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
			wbMerge.write(outputStream);
			wbMerge.close();

			FileUtils.writeByteArrayToFile(mergerFile, outputStream.toByteArray());
			outputStream.close();
			System.out.println("Merge file is created. Name:" + mergeName + ".xls");
		}

	}
}
