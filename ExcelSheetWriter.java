package com.excel.report;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.List;
import java.util.StringTokenizer;
import net.htmlparser.jericho.Element;
import net.htmlparser.jericho.PHPTagTypes;
import net.htmlparser.jericho.Segment;
import net.htmlparser.jericho.Source;
import net.htmlparser.jericho.StartTagType;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.IndexedColors;



public class ExcelSheetWriter
{
  static String inputFolder = null;
  
  static int rowIndex = 1;
  static int totalFiles = 0;
  
  static final String dateFormat = "mm/dd/yyyy";
  static String errorReportFile = "error.txt";
  static String supplier = "";
  

  public ExcelSheetWriter() {}
  
  private void writeExcelFile(String reportInputDir, String outputFile)
  {
    FileOutputStream fileOutputStream = null;
    HSSFWorkbook sampleWorkbook = null;
    HSSFSheet sampleDataSheet = null;
    try
    {
      inputFolder = reportInputDir;

      sampleWorkbook = new HSSFWorkbook();
      sampleDataSheet = sampleWorkbook.createSheet("reportData");
      
      HSSFRow headerRow = sampleDataSheet.createRow(0);
      
      HSSFCellStyle cellStyle = setHeaderStyle(sampleWorkbook);


      HSSFCell skuIDHeaderCell = headerRow.createCell(0);
      skuIDHeaderCell.setCellStyle(cellStyle);
      skuIDHeaderCell.setCellValue(new HSSFRichTextString("SKU"));
      

      HSSFCell upcHeaderCell = headerRow.createCell(1);
      upcHeaderCell.setCellStyle(cellStyle);
      upcHeaderCell.setCellValue(new HSSFRichTextString("UPC"));
      
      HSSFCell vendorHeaderCell = headerRow.createCell(2);
      vendorHeaderCell.setCellStyle(cellStyle);
      vendorHeaderCell.setCellValue(new HSSFRichTextString("Vendor"));
      
      HSSFCell nameHeaderCell = headerRow.createCell(3);
      nameHeaderCell.setCellStyle(cellStyle);
      nameHeaderCell.setCellValue(new HSSFRichTextString("Name"));

      HSSFCell shortDescriptionHeaderCell = headerRow.createCell(4);
      shortDescriptionHeaderCell.setCellStyle(cellStyle);
      shortDescriptionHeaderCell.setCellValue(new HSSFRichTextString("Short Description"));
      
      HSSFCell descriptionHeaderCell = headerRow.createCell(5);
      descriptionHeaderCell.setCellStyle(cellStyle);
      descriptionHeaderCell.setCellValue(new HSSFRichTextString("Description"));
      
      HSSFCell supplierHeaderCell = headerRow.createCell(6);
      supplierHeaderCell.setCellStyle(cellStyle);
      supplierHeaderCell.setCellValue(new HSSFRichTextString("Supplier"));
      
      HSSFCell creatDTHeaderCell = headerRow.createCell(7);
      creatDTHeaderCell.setCellStyle(cellStyle);
      creatDTHeaderCell.setCellValue(new HSSFRichTextString("Creation_DT"));

      HSSFCell lastUpdtTSHeaderCell = headerRow.createCell(8);
      lastUpdtTSHeaderCell.setCellStyle(cellStyle);
      lastUpdtTSHeaderCell.setCellValue(new HSSFRichTextString("Last_Updated_DT"));

      HSSFCell lastUpdtIdHeaderCell = headerRow.createCell(9);
      lastUpdtIdHeaderCell.setCellStyle(cellStyle);
      lastUpdtIdHeaderCell.setCellValue(new HSSFRichTextString("Last_Updated_By_Id"));

      listFilesForFolder(reportInputDir, sampleDataSheet);
      
      System.out.println("totalFiles" + totalFiles);

      fileOutputStream = new FileOutputStream(outputFile);
      sampleWorkbook.write(fileOutputStream);
    }
    catch (Exception ex)
    {
      ex.printStackTrace();

      try
      {
        if (fileOutputStream != null)
        {
          fileOutputStream.close();
        }
      }
      catch (IOException ex)
      {
        ex.printStackTrace();
      }
    }
    finally
    {
      try
      {
        if (fileOutputStream != null)
        {
          fileOutputStream.close();
        }
      }
      catch (IOException ex)
      {
        ex.printStackTrace();
      }
    }
  }
  
  private HSSFCellStyle setHeaderStyle(HSSFWorkbook sampleWorkBook)
  {
    HSSFFont font = sampleWorkBook.createFont();
    font.setFontName("Arial");
    font.setColor(IndexedColors.PLUM.getIndex());
    font.setBoldweight((short)700);
    HSSFCellStyle cellStyle = sampleWorkBook.createCellStyle();
    cellStyle.setFont(font);
    return cellStyle;
  }
  
  public static void main(String[] args) {
    supplier = args[0];
    System.out.println("Input Supplier: " + supplier);
    System.out.println("Input folder: " + args[1]);
    System.out.println("Output file: " + args[2]);
    new ExcelSheetWriter().writeExcelFile(args[1], args[2]);
  }
  
  public static void listFilesForFolder(String baseFolder, HSSFSheet sampleDataSheet)
  {
    File folder = new File(baseFolder);
    
    for (File fileEntry : folder.listFiles()) {
      if (fileEntry.isDirectory()) {
        listFilesForFolder(fileEntry.getAbsolutePath(), sampleDataSheet);
      }
      else {
        System.out.println(fileEntry.getAbsoluteFile());
        totalFiles += 1;
        
        if ((fileEntry.getName().endsWith(".htm")) || (fileEntry.getName().endsWith(".html"))) {
          try
          {
            readHtml(fileEntry.toURL(), sampleDataSheet, fileEntry.getAbsolutePath());
          }
          catch (MalformedURLException e) {
            e.printStackTrace();
          }
        }
      }
    }
  }
  

  public static String readDoc(File f)
  {
    String text = "";
    int N = 1048576;
    char[] buffer = new char[N];
    FileReader fr = null;
    BufferedReader br = null;
    try {
      fr = new FileReader(f);
      br = new BufferedReader(fr);
      int read;
      do {
        read = br.read(buffer, 0, N);
        text = text + new String(buffer, 0, read);
      }
      while (read >= N);

    }
    catch (Exception ex)
    {
      ex.printStackTrace();
      try
      {
        br.close();
        fr.close();
      }
      catch (IOException e) {
        e.printStackTrace();
      }
    }
    finally
    {
      try
      {
        br.close();
        fr.close();
      }
      catch (IOException e) {
        e.printStackTrace();
      }
    }
    


    return text;
  }
  

  public static String getHeader(String headerName, String content)
  {
    return content.substring(0, 10);
  }
  
  public static void displaySegments(List<? extends Segment> segments) {
    for (Segment segment : segments) {
      System.out.println("-------------------------------------------------------------------------------");
      System.out.println(segment.getDebugInfo());
      System.out.println(segment.getFirstElement().getContent());
    }
    System.out.println("\n*******************************************************************************\n");
  }
  
  public static void readHtml(URL sourceUrlString, HSSFSheet sampleDataSheet, String filePath)
  {
    StringBuffer errorDetails = new StringBuffer(10);
    
    System.out.println("sourceUrlString" + sourceUrlString);
    
    try
    {
      Source source = null;
      try {
        source = new Source(sourceUrlString);
      }
      catch (MalformedURLException e) {
        errorDetails.append("error Reading file");
        e.printStackTrace();
      }
      catch (IOException e) {
        errorDetails.append(":\t error Reading file");
        e.printStackTrace();
      }
      System.out.println("\n*******************************************************************************\n");
      
      System.out.println("XML Declarations:");
      displaySegments(source.getAllTags(StartTagType.XML_DECLARATION));
      
      System.out.println("XML Processing instructions:");
      displaySegments(source.getAllTags(StartTagType.XML_PROCESSING_INSTRUCTION));
      
      PHPTagTypes.register();
      StartTagType.XML_DECLARATION.deregister();
      source = new Source(source);
      System.out.println("##################### PHP tag types now added to register #####################\n");
      
      System.out.println("H3 Elements:");
      
      List<? extends Segment> segments = source.getAllElements("h3");
      
      displaySegments(segments);
      
      HSSFRow dataRow1 = sampleDataSheet.createRow(rowIndex++);

      try
      {
        if (segments.get(0) != null) {
          if ( (((Segment)segments.get(0)).toString().contains("ID:")) || (((Segment)segments.get(0)).toString().contains("SKU:")) )
          {
            dataRow1.createCell(0).setCellValue(new HSSFRichTextString(((Segment)segments.get(0)).getFirstElement().getContent().toString()
              .split(":")[1]));
          }
          else {
            dataRow1.createCell(0).setCellValue(new HSSFRichTextString(((Segment)segments.get(0)).getFirstElement().getContent().toString()));
          }
        }
      } catch (Exception exec) {
        System.out.println("Error reading ID from <H3>");
        errorDetails.append(":\t  Error reading ID from <H3>");
        exec.printStackTrace();
      }
      
      try
      {
        if (segments.get(1) != null) {
          if (((Segment)segments.get(1)).toString().contains("UPC:")) {
            dataRow1.createCell(1).setCellValue(new HSSFRichTextString(((Segment)segments.get(1)).getFirstElement().getContent().toString()
              .split(":")[1]));
          }
          else {
            dataRow1.createCell(1).setCellValue(new HSSFRichTextString(((Segment)segments.get(1)).getFirstElement().getContent().toString()));
          }
        }
      } catch (Exception exec) {
        System.out.println("Error reading UPC from <H3>");
        errorDetails.append(":\t Error reading UPC from <H3>");
        
        exec.printStackTrace();
      }
      

      try
      {
        if (segments.get(2) != null) {
          if (((Segment)segments.get(2)).toString().contains("Vendor:")) {
            dataRow1.createCell(2).setCellValue(new HSSFRichTextString(((Segment)segments.get(2)).getFirstElement().getContent().toString()
              .split(":")[1]));
          } else {
            dataRow1.createCell(2).setCellValue(new HSSFRichTextString(((Segment)segments.get(2)).getFirstElement().getContent().toString()));
          }
        }
      }
      catch (Exception exec)
      {
        System.out.println("Error reading Vendor from <H3>");
        errorDetails.append(":\t Error reading Vendor from <H3>");
        
        exec.printStackTrace();
      }
      


      System.out.println("H2 Elements:");
      List<? extends Segment> h2Segments = source.getAllElements("h2");
      displaySegments(h2Segments);
      try
      {
        if ((h2Segments != null) && (h2Segments.get(0) != null) && (!((Segment)h2Segments.get(0)).isWhiteSpace())) {
          dataRow1.createCell(3).setCellValue(new HSSFRichTextString(((Segment)h2Segments.get(0)).getFirstElement().getContent().toString()));
        }
      } catch (Exception exec) {
        System.out.println("Error reading  <H2> elements");
        errorDetails.append(":\t Error reading <H2> elements ");
        
        exec.printStackTrace();
      }
      
      System.out.println("Paragraph Elements:");
      List<? extends Segment> paraSegments = source.getAllElements("p");
      displaySegments(paraSegments);
      
      try
      {
        if ((paraSegments != null) && (paraSegments.size() > 0) && (paraSegments.get(0) != null) && (!((Segment)paraSegments.get(0)).isWhiteSpace()))
        {
          dataRow1.createCell(4).setCellValue(new HSSFRichTextString(((Segment)paraSegments.get(0)).toString()));
        }
      } catch (Exception exec) {
        System.out.println("Error reading  <p> elements");
        errorDetails.append(":\t Error reading <p> elements ");
        
        exec.printStackTrace();
      }
 
 
 
 
 
      
      System.out.println("UL Elements:");
      List<? extends Segment> ulSegments = source.getAllElements("ul");
      
      displaySegments(ulSegments);
      try {
        if ((ulSegments != null) && (ulSegments.size() > 0) && (ulSegments.get(0) != null) && (!((Segment)ulSegments.get(0)).isWhiteSpace()))
        {
          dataRow1.createCell(5).setCellValue(new HSSFRichTextString(((Segment)ulSegments.get(0)).getFirstElement().getContent().toString()));
        }
      } catch (Exception exec) {
        System.out.println("Error reading  <ul> elements");
        errorDetails.append(":\t Error reading <ul> elements ");
        
        exec.printStackTrace();
      }
      


      dataRow1.createCell(6).setCellValue(supplier);
      


      Process proc = Runtime.getRuntime().exec("cmd /c dir " + filePath + " /tc");
      
      System.out.println("filePath" + filePath);
      
      BufferedReader br = new BufferedReader(new InputStreamReader(proc.getInputStream()));
      String tempdata = "";
      for (int i = 0; i < 6; i++)
      {
        tempdata = br.readLine();
      }
      
      StringTokenizer st = new StringTokenizer(tempdata);
      String date = st.nextToken();
      String time = st.nextToken();
      
      dataRow1.createCell(7).setCellValue(date);

    } catch (Exception exec) {
      System.out.println("Error reading  html file" + sourceUrlString);
      errorDetails.append(":\t Error reading html file" + sourceUrlString);
      
      exec.printStackTrace();
    }
    

    if (errorDetails.toString().trim().length() > 0) {
      writeErrorReport(errorDetails.toString(), filePath);
    }
  }
  



  public static void writeErrorReport(String errorDetails, String filePath)
  {
    if (errorDetails.toString().trim().length() > 0) {
      File file = null;
      try {
        file = new File(inputFolder + "/" + errorReportFile);
        

        if (!file.exists()) {
          file.createNewFile();
        }
        

        FileWriter fileWritter = new FileWriter(file.getAbsoluteFile(), true);
        BufferedWriter bufferWritter = new BufferedWriter(fileWritter);
        bufferWritter.write("\n" + filePath + ":\t " + errorDetails.toString());
        bufferWritter.close();
        fileWritter.close();
      }
      catch (IOException e)
      {
        e.printStackTrace();
      }
    }
  }
}
