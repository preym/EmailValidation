package com.prem.emailcheck;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.naming.NamingEnumeration;
import javax.naming.NamingException;
import javax.naming.directory.Attribute;
import javax.naming.directory.Attributes;
import javax.naming.directory.DirContext;
import javax.naming.directory.InitialDirContext;
import java.io.*;
import java.net.Socket;
import java.util.*;

public class SMTPMXLookup {


  private static ArrayList<String> testData = new ArrayList<String>();
  private static HSSFWorkbook workbook = new HSSFWorkbook();
  private static HSSFSheet sheet = workbook.createSheet("Sheet1");
  private static int rownum = 0;

  private static int hear(BufferedReader in) throws IOException {
    String line = null;
    int res = 0;
    while ((line = in.readLine()) != null) {
      String pfx = line.substring(0, 3);
      try {
        res = Integer.parseInt(pfx);
      } catch (Exception ex) {
        res = -1;
      }
      if (line.charAt(3) != '-') break;
    }
    return res;
  }

  private static void say(BufferedWriter wr, String text)
      throws IOException {
    wr.write(text + "\r\n");
    wr.flush();
    return;
  }

  private static ArrayList getMX(String hostName)
      throws NamingException {
    // Perform a DNS lookup for MX records in the domain
    Hashtable env = new Hashtable();
    env.put("java.naming.factory.initial",
        "com.sun.jndi.dns.DnsContextFactory");
    DirContext ictx = new InitialDirContext(env);
    Attributes attrs = ictx.getAttributes
        (hostName, new String[]{"MX"});
    Attribute attr = attrs.get("MX");
    // if we don't have an MX record, try the machine itself
    if ((attr == null) || (attr.size() == 0)) {
      attrs = ictx.getAttributes(hostName, new String[]{"A"});
      attr = attrs.get("A");
      if (attr == null)
        throw new NamingException
            ("No match for name '" + hostName + "'");
    }
    ArrayList resultList = new ArrayList();
    NamingEnumeration enumeration = attr.getAll();
    while (enumeration.hasMore()) {
      String mailhost;
      String x = (String) enumeration.next();
      String f[] = x.split(" ");
      //  THE fix *************
      if (f.length == 1)
        mailhost = f[0];
      else if (f[1].endsWith("."))
        mailhost = f[1].substring(0, (f[1].length() - 1));
      else
        mailhost = f[1];
      resultList.add(mailhost);
    }
    return resultList;
  }

  public static boolean isAddressValid(String address) {
    int pos = address.indexOf('@');
    if (pos == -1) return false;
    String domain = address.substring(++pos);
    ArrayList mxList = null;
    try {
      mxList = getMX(domain);
    } catch (NamingException ex) {
      return false;
    }
    if (mxList.size() == 0) return false;
    for (int mx = 0; mx < mxList.size(); mx++) {
      boolean valid = false;
      try {
        int res;
        Socket skt = new Socket((String) mxList.get(mx), 25);
        BufferedReader rdr = new BufferedReader
            (new InputStreamReader(skt.getInputStream()));
        BufferedWriter wtr = new BufferedWriter
            (new OutputStreamWriter(skt.getOutputStream()));
        res = hear(rdr);
        if (res != 220) throw new Exception("Invalid header");
        say(wtr, "EHLO rgagnon.com");
        res = hear(rdr);
        if (res != 250) throw new Exception("Not ESMTP");
        // validate the sender address
        say(wtr, "MAIL FROM: <tim@orbaker.com>");
        res = hear(rdr);
        if (res != 250) throw new Exception("Sender rejected");
        say(wtr, "RCPT TO: <" + address + ">");
        res = hear(rdr);
        say(wtr, "RSET");
        hear(rdr);
        say(wtr, "QUIT");
        hear(rdr);
        if (res != 250)
          throw new Exception(address + "  Address is not valid!");
        valid = true;
        rdr.close();
        wtr.close();
        skt.close();
      } catch (Exception ex) {
        ex.printStackTrace();
      } finally {
        if (valid) {
          writeFileData(address);
          return true;
        }
      }
    }
    return false;
  }

  public static void main(String args[]) {
    readFileData();
    for (int ctr = 0; ctr < testData.size(); ctr++) {
      System.out.println(testData.get(ctr) + " is valid? " +
          isAddressValid(testData.get(ctr)));
    }
    return;
  }

  private static void readFileData() {
    try {
      FileInputStream file = new FileInputStream(new File("./input.xlsx"));
      XSSFWorkbook workbook = new XSSFWorkbook(file);
      XSSFSheet sheet = workbook.getSheetAt(0);
      Iterator<Row> rowIterator = sheet.iterator();
      while (rowIterator.hasNext()) {
        Row row = rowIterator.next();
        Iterator<Cell> cellIterator = row.cellIterator();
        if (cellIterator.hasNext()) {
          Cell cell = cellIterator.next();
          System.out.println(cell.getStringCellValue());
          testData.add(cell.getStringCellValue());
        }
      }
      file.close();
      System.out.println("Reading completed. Your Input is:\n" + testData);
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }


  private static void writeFileData(String email) {
    Row row = sheet.createRow(rownum++);
    Cell cell = row.createCell(0);
    cell.setCellValue(email);
    try {
      FileOutputStream out =
          new FileOutputStream(new File("./output.xls"));
      workbook.write(out);
      out.close();
      System.out.println("Excel written successfully..");
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
} 
