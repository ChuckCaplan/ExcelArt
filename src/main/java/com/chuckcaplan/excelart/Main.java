package com.chuckcaplan.excelart;

import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.apache.commons.lang3.time.StopWatch;
import org.apache.commons.vfs2.FileObject;
import org.apache.commons.vfs2.VFS;
import org.apache.commons.vfs2.impl.DefaultFileSystemManager;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.imgscalr.Scalr;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.beust.jcommander.JCommander;
import com.beust.jcommander.Parameter;

/**
 * @author Chuck Caplan
 * 
 *         Takes an image as input and outputs a .xlsx Excel spreadsheet with
 *         the cell background colors matching the RGB values of each pixel.
 * 
 *         Idea from Matt Parker - https://www.youtube.com/watch?v=UBX2QQHlQ_I
 *
 */
public class Main {

	private final static Logger logger = LoggerFactory.getLogger(Main.class);

	// Command line parameters

	// read the image from either a file system or a url via Apache Commons VFS
	@Parameter(names = "-i", description = "The filesystem path or URL of the input image.", required = true, order = 0)
	private String imageFile;

	@Parameter(names = "-o", description = "The filename of the resulting .xlsx spreadsheet.", required = false, order = 1)
	private String outputFile = "output.xlsx";

	@Parameter(names = { "-h", "--help" }, help = true, description = "Show this help.", order = 2)
	private boolean help;

	// resize image to be no more than 200 pixels high or wide.
	// otherwise we may run out of memory or run for hours
	// allow an override via a hidden command line parameter
	@Parameter(names = "-m", hidden = true, description = "Hidden - The max number of rows or columns in the resulting spreadsheet.", order = 3)
	private int maxWidthHeight = 200;

	/**
	 * @param args
	 */
	public static void main(String... args) {
		StopWatch sw = new StopWatch();
		sw.start();
		Main main = new Main();
		// parse the command line arguments
		JCommander jc = JCommander.newBuilder().addObject(main).build();
		jc.setProgramName("java " + main.getClass().getName());
		try {
			jc.parse(args);
		} catch (Exception e) {
			logger.error(e.getMessage());
			jc.usage();
			System.exit(-1);
		}
		if (main.isHelp()) {
			jc.usage();
			System.exit(0);
		}
		// run the main logic
		main.run();
		sw.stop();
		logger.info("Total time in seconds: "+sw.getTime(TimeUnit.SECONDS));
	}

	/**
	 * @return whether or not to show help
	 */
	public boolean isHelp() {
		return help;
	}

	/**
	 * The main program logic
	 */
	private void run() {

		DefaultFileSystemManager fsManager = null;
		BufferedImage image = null;

		try {
			// load the image from the command line parameter. Using Apache Commons VFS so
			// we can accept a full path, relative path, just a file name, or also a URL.
			fsManager = (DefaultFileSystemManager) VFS.getManager();
			fsManager.setBaseFile(new File("."));
			FileObject file = fsManager.resolveFile(imageFile);
			image = ImageIO.read(file.getContent().getInputStream());
		} catch (IOException e) {
			logger.error("Error reading image", e);
			System.exit(-1);
		}

		// resize the image so the processing time is faster and so the resulting
		// spreadsheet will be a reasonable size.
		if (image.getHeight() > maxWidthHeight || image.getWidth() > maxWidthHeight) {
			// Scalr.resize is smart enough to figure out if the image is portrait or
			// landscape and resize appropriately, so just pass the max size in as both
			// width and height.
			image = Scalr.resize(image, Scalr.Method.AUTOMATIC, maxWidthHeight, maxWidthHeight);
			image.flush();
		}

		// create the excel file
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet();

		/*
		 * Loop through the image and write out rows and cells with the RGB info of each
		 * pixel. For each pixel we are actually writing out 3 rows - Red, Green, Blue
		 * (RGB). We do this because otherwise we would need to create thousands or
		 * millions of styles, which is too processor-intensive with POI.
		 */
		int rowNum = 0;
		logger.info("Rows to process: " + image.getHeight());
		for (int y = 0; y < image.getHeight(); y++) {
			logger.info("Processing row #" + (y + 1));
			Row redRow = sheet.createRow(rowNum);
			rowNum++;
			// write the RED info
			for (int x = 0; x < image.getWidth(); x++) {
				int rgb = image.getRGB(x, y); // always returns TYPE_INT_ARGB
				int red = (rgb >> 16) & 0xFF;
				// Getting pixel color by position x and y
				Cell cell = redRow.createCell(x);
				// set background color
				Color color = new Color(red, 0, 0);
				XSSFCellStyle style = getStyle(workbook, color);
				cell.setCellStyle(style);
			}
			// write the GREEN info
			Row greenRow = sheet.createRow(rowNum);
			rowNum++;
			for (int x = 0; x < image.getWidth(); x++) {
				int rgb = image.getRGB(x, y); // always returns TYPE_INT_ARGB
				int green = (rgb >> 8) & 0xFF;
				// Getting pixel color by position x and y
				Cell cell = greenRow.createCell(x);
				// set background color
				Color color = new Color(0, green, 0);
				XSSFCellStyle style = getStyle(workbook, color);
				cell.setCellStyle(style);
			}
			// write the BLUE info
			Row blueRow = sheet.createRow(rowNum);
			rowNum++;
			for (int x = 0; x < image.getWidth(); x++) {
				int rgb = image.getRGB(x, y); // always returns TYPE_INT_ARGB
				int blue = (rgb) & 0xFF;
				// Getting pixel color by position x and y
				Cell cell = blueRow.createCell(x);
				// set background color
				Color color = new Color(0, 0, blue);
				XSSFCellStyle style = getStyle(workbook, color);
				cell.setCellStyle(style);
			}
		}

		try {
			// Write the output to a file
			logger.info("Writing file " + outputFile);
			FileOutputStream fileOut = new FileOutputStream(outputFile);
			workbook.write(fileOut);
			fileOut.close();

			// Closing the workbook
			workbook.close();
		} catch (IOException e) {
			logger.error("Error writing file " + outputFile, e);
		}
		logger.info("Done! Exiting...");
	}

	// a cached map of styles to speed up processing
	private Map<Color, XSSFCellStyle> styles = new HashMap<Color, XSSFCellStyle>();

	/**
	 * Creating styles with POI is very slow, so I am caching them in a HashMap.
	 * Then as we need a new style we will see if one already exists for that color,
	 * and if so, return it. If not we will create a new one, add it to the map, and
	 * return it.
	 * 
	 * @param wb    The Workbook
	 * @param color The color to either create a new style or get a cached style
	 *              from the map
	 * @return The style for the color, regardless of whether it is newly created or
	 *         cached
	 */
	private XSSFCellStyle getStyle(Workbook wb, Color color) {
		// check if it exists in the map, and if so, return it.
		XSSFCellStyle style = styles.get(color);
		if (style != null) {
			return style;
		}
		// if not, create it, add it to the map, and return it
		style = (XSSFCellStyle) wb.createCellStyle();
		XSSFColor myColor = new XSSFColor(color, null);
		style.setFillForegroundColor(myColor);
		style.setFillBackgroundColor(myColor);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		// add to the map
		styles.put(color, style);
		return style;
	}

}
